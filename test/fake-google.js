'use strict';

/**
 * Имитация Sheets API v4 и части Drive API v3 в памяти — для тестов без сети.
 * Передаётся в GSheets как authClient (метод request в стиле gaxios).
 */

const { letterToCol } = require('../src/a1');

function parseRange(range) {
    const m = /^'((?:[^']|'')+)'(?:!(.*))?$/.exec(range) || /^([^!]+)(?:!(.*))?$/.exec(range);
    const title = m[1].replace(/''/g, "'");
    const ref = m[2];
    if (!ref) return { title, r1: 1, c1: 1, r2: Infinity, c2: Infinity };
    const [a, b] = ref.split(':');
    const cell = (s, end) => {
        const mm = /^([A-Z]*)(\d*)$/.exec(s);
        return {
            c: mm[1] ? letterToCol(mm[1]) : (end ? Infinity : 1),
            r: mm[2] ? Number(mm[2]) : (end ? Infinity : 1),
        };
    };
    const s = cell(a, false);
    const e = b === undefined ? { r: s.r, c: s.c } : cell(b, true);
    return { title, r1: s.r, c1: s.c, r2: e.r, c2: e.c };
}

class FakeGoogle {
    constructor() {
        this.docs = new Map();
        this.files = new Map();
        this.calls = [];
        this.fail = [];      // [{ status, times, match }]
        this.seq = 100;
    }

    addDoc(id, title, sheets = [{ title: 'Лист1' }]) {
        const doc = { id, title, locale: 'ru_RU', timeZone: 'Europe/Kyiv', sheets: [] };
        sheets.forEach((s, i) => doc.sheets.push(this._newSheet(s.title, i, s.rows ?? 1000, s.cols ?? 26, s.data)));
        this.docs.set(id, doc);
        this.files.set(id, { id, name: title, mimeType: 'application/vnd.google-apps.spreadsheet', parents: ['root'], permissions: [] });
        return doc;
    }

    _newSheet(title, index, rows, cols, data = []) {
        return { sheetId: this.seq++, title, index, rows, cols, frozen: 0, data: data.map((r) => r.slice()), formats: [] };
    }

    _sheet(doc, title) {
        const s = doc.sheets.find((x) => x.title === title);
        if (!s) throw this._err(400, 'INVALID_ARGUMENT', `Unable to parse range: ${title}`);
        return s;
    }

    _err(status, code, message, reason) {
        const e = new Error(message);
        e.fake = { status, data: { error: { code: status, status: code, message, errors: reason ? [{ reason }] : undefined } } };
        return e;
    }

    _trim(rows) {
        const out = rows.map((r) => {
            const c = (r || []).slice();
            while (c.length && (c[c.length - 1] === '' || c[c.length - 1] === undefined)) c.pop();
            return c.map((v) => (v === undefined ? '' : v));
        });
        while (out.length && out[out.length - 1].length === 0) out.pop();
        return out;
    }

    _get(doc, range) {
        const { title, r1, c1, r2, c2 } = parseRange(range);
        const s = this._sheet(doc, title);
        const rows = [];
        const rEnd = Math.min(r2, s.data.length);
        for (let r = r1; r <= rEnd; r++) {
            const row = s.data[r - 1] || [];
            const cEnd = Math.min(c2, row.length);
            const out = [];
            for (let c = c1; c <= cEnd; c++) out.push(row[c - 1] ?? '');
            rows.push(out);
        }
        return this._trim(rows);
    }

    _set(doc, range, values, input) {
        const { title, r1, c1 } = parseRange(range);
        const s = this._sheet(doc, title);
        const lastRow = r1 + values.length - 1;
        const lastCol = c1 + Math.max(...values.map((v) => v.length)) - 1;
        if (lastRow > s.rows || lastCol > s.cols) {
            throw this._err(400, 'INVALID_ARGUMENT', `Range (${range}) exceeds grid limits. Max rows: ${s.rows}, max columns: ${s.cols}`);
        }
        values.forEach((row, i) => {
            const r = r1 - 1 + i;
            while (s.data.length <= r) s.data.push([]);
            row.forEach((v, j) => {
                let val = v;
                if (input === 'USER_ENTERED' && typeof v === 'string' && /^-?\d+(\.\d+)?$/.test(v)) val = Number(v);
                s.data[r][c1 - 1 + j] = val;
            });
        });
        return values.reduce((n, r) => n + r.length, 0);
    }

    async request(o) {
        const url = new URL(o.url);
        const method = o.method || 'GET';
        this.calls.push({ method, url: o.url, data: o.data, headers: o.headers });

        const f = this.fail.find((x) => x.times > 0 && (!x.match || x.match.test(o.url)));
        if (f) {
            f.times--;
            return { status: f.status, headers: f.retryAfter ? { 'retry-after': String(f.retryAfter) } : {}, data: { error: { code: f.status, status: 'UNAVAILABLE', message: 'fake failure' } } };
        }

        try {
            const data = this._route(url, method, o);
            if (data && data.__full) return data.__full;
            return { status: 200, headers: {}, data };
        } catch (e) {
            if (e.fake) return { status: e.fake.status, headers: {}, data: e.fake.data };
            throw e;
        }
    }

    _route(url, method, o) {
        const p = decodeURIComponent(url.pathname);
        const q = url.searchParams;
        let m;

        // ── Sheets ──
        if (url.host === 'sheets.googleapis.com') {
            if (p === '/v4/spreadsheets' && method === 'POST') {
                const id = `doc${this.seq++}xxxxxxxx`;
                this.addDoc(id, o.data.properties.title);
                return { spreadsheetId: id };
            }
            m = /^\/v4\/spreadsheets\/([^/:]+)(.*)$/.exec(p);
            if (!m) throw this._err(404, 'NOT_FOUND', 'not found');
            const doc = this.docs.get(m[1]);
            if (!doc) throw this._err(404, 'NOT_FOUND', 'Requested entity was not found.');
            const rest = m[2];

            if (rest === '' && method === 'GET') {
                return {
                    spreadsheetId: doc.id,
                    properties: { title: doc.title, locale: doc.locale, timeZone: doc.timeZone },
                    sheets: doc.sheets.map((s) => ({
                        properties: { sheetId: s.sheetId, title: s.title, index: s.index, sheetType: 'GRID', gridProperties: { rowCount: s.rows, columnCount: s.cols, frozenRowCount: s.frozen } },
                    })),
                };
            }
            if (rest === ':batchUpdate') return { replies: o.data.requests.map((r) => this._batch(doc, r)) };
            if (rest === '/values:batchGet') {
                return { valueRanges: q.getAll('ranges').map((r) => ({ range: r, values: this._get(doc, r) })) };
            }
            if (rest === '/values:batchUpdate') {
                let total = 0;
                for (const d of o.data.data) total += this._set(doc, d.range, d.values, o.data.valueInputOption);
                return { totalUpdatedCells: total };
            }
            if (rest === '/values:batchClear') {
                for (const r of o.data.ranges) {
                    const { title, r1, c1, r2, c2 } = parseRange(r);
                    const s = this._sheet(doc, title);
                    s.data.forEach((row, i) => {
                        if (i + 1 < r1 || i + 1 > r2) return;
                        for (let c = c1; c <= Math.min(c2, row.length); c++) row[c - 1] = '';
                    });
                }
                return {};
            }
            if ((m = /^\/values\/(.+):append$/.exec(rest))) {
                const { title } = parseRange(m[1]);
                const s = this._sheet(doc, title);
                const start = this._trim(s.data).length + 1;
                const need = start + o.data.values.length - 1;
                if (need > s.rows) s.rows = need;
                const maxC = Math.max(...o.data.values.map((v) => v.length));
                if (maxC > s.cols) s.cols = maxC;
                this._set(doc, `'${title.replace(/'/g, "''")}'!A${start}`, o.data.values, q.get('valueInputOption'));
                return { updates: { updatedRows: o.data.values.length } };
            }
            if ((m = /^\/values\/(.+)$/.exec(rest)) && method === 'GET') {
                return { range: m[1], values: this._get(doc, m[1]) };
            }
            if ((m = /^\/sheets\/(\d+):copyTo$/.exec(rest))) {
                const src = doc.sheets.find((s) => s.sheetId === Number(m[1]));
                const dst = this.docs.get(o.data.destinationSpreadsheetId);
                const copy = this._newSheet(`Копия ${src.title}`, dst.sheets.length, src.rows, src.cols, src.data);
                dst.sheets.push(copy);
                return { sheetId: copy.sheetId, title: copy.title };
            }
        }

        // ── Docs export ──
        if (url.host === 'docs.google.com') {
            m = /^\/spreadsheets\/d\/([^/]+)\/export$/.exec(p);
            const doc = this.docs.get(m[1]);
            const s = doc.sheets.find((x) => x.sheetId === Number(q.get('gid')));
            const csv = this._trim(s.data).map((r) => r.join(q.get('format') === 'tsv' ? '\t' : ',')).join('\n');
            return Buffer.from(csv);
        }

        // ── Drive ──
        if (url.host === 'www.googleapis.com') {
            if (p === '/drive/v3/files' && method === 'POST') {
                const id = `file${this.seq++}xxxxxxxx`;
                const meta = o.data;
                if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') this.addDoc(id, meta.name);
                this.files.set(id, { id, name: meta.name, mimeType: meta.mimeType, parents: meta.parents || ['root'], permissions: [] });
                return { id, name: meta.name };
            }
            if (p === '/drive/v3/files' && method === 'GET') {
                const qq = q.get('q') || '';
                const name = /name='((?:\\'|[^'])*)'/.exec(qq)?.[1]?.replace(/\\'/g, "'");
                const mime = /mimeType='([^']+)'/.exec(qq)?.[1];
                const files = [...this.files.values()].filter((f) => (!name || f.name === name) && (!mime || f.mimeType === mime) && !f.trashed);
                return { files: files.map(({ permissions, content, ...f }) => f) };
            }
            if ((m = /^\/drive\/v3\/files\/([^/]+)\/copy$/.exec(p))) {
                const src = this.docs.get(m[1]);
                const id = `copy${this.seq++}xxxxxxxx`;
                const d = this.addDoc(id, o.data.name, []);
                d.sheets = src.sheets.map((s) => ({ ...s, data: s.data.map((r) => r.slice()) }));
                return { id };
            }
            if ((m = /^\/drive\/v3\/files\/([^/]+)\/permissions$/.exec(p))) {
                const f = this.files.get(m[1]);
                if (method === 'GET') return { permissions: f.permissions };
                const perm = { id: `perm${this.seq++}`, ...o.data };
                f.permissions.push(perm);
                return perm;
            }
            if ((m = /^\/drive\/v3\/files\/([^/]+)\/export$/.exec(p))) {
                return Buffer.from(`EXPORT:${q.get('mimeType')}`);
            }
            if ((m = /^\/drive\/v3\/files\/([^/]+)$/.exec(p))) {
                const f = this.files.get(m[1]);
                if (!f) throw this._err(404, 'NOT_FOUND', 'File not found', 'notFound');
                if (method === 'GET' && q.get('alt') === 'media') return f.content;
                if (method === 'GET') return f;
                if (method === 'PATCH') {
                    if (o.data?.trashed) f.trashed = true;
                    if (q.get('addParents')) f.parents = [q.get('addParents')];
                    return { id: f.id, parents: f.parents };
                }
                if (method === 'DELETE') { this.files.delete(f.id); return {}; }
            }
            if (p.startsWith('/upload/drive/v3/files')) {
                const type = q.get('uploadType');
                if (type === 'multipart') {
                    const body = o.data.toString('utf8');
                    const meta = JSON.parse(/\r\n\r\n(\{.*?\})\r\n/s.exec(body)[1]);
                    const content = Buffer.from(body.split(/\r\n\r\n/).slice(2).join('\r\n\r\n').replace(/\r\n--[^\r\n]+--$/, ''));
                    const fileId = p.split('/')[5];
                    const id = fileId || `up${this.seq++}xxxxxxxx`;
                    if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
                        const rows = content.toString('utf8').trim().split('\n').map((l) => l.split(','));
                        this.addDoc(id, meta.name, [{ title: meta.name, data: rows }]);
                    }
                    const prev = this.files.get(id) || { permissions: [] };
                    this.files.set(id, { ...prev, id, name: meta.name || prev.name, mimeType: meta.mimeType || /--[^\r\n]+\r\nContent-Type: ([^\r\n]+)\r\n\r\n(?!\{)/.exec(body)[1], parents: meta.parents || prev.parents || ['root'], content });
                    return { id, name: meta.name || prev.name, mimeType: meta.mimeType };
                }
                if (type === 'resumable') {
                    this.pendingUpload = o.data;
                    return { __full: { status: 200, headers: { location: 'https://www.googleapis.com/upload/session/xyz' }, data: {} } };
                }
            }
            if (p === '/upload/session/xyz' && method === 'PUT') {
                const id = `big${this.seq++}xxxxxxxx`;
                this.files.set(id, { id, name: this.pendingUpload.name, mimeType: o.headers['Content-Type'], content: o.data, permissions: [] });
                return { id, name: this.pendingUpload.name, size: String(o.data.length) };
            }
        }

        throw this._err(404, 'NOT_FOUND', `fake: нет маршрута ${method} ${url}`);
    }

    _batch(doc, r) {
        const byId = (id) => doc.sheets.find((s) => s.sheetId === id);
        if (r.addSheet) {
            const pr = r.addSheet.properties;
            if (doc.sheets.some((s) => s.title === pr.title)) throw this._err(400, 'INVALID_ARGUMENT', `Sheet ${pr.title} already exists`);
            const s = this._newSheet(pr.title, pr.index ?? doc.sheets.length, pr.gridProperties?.rowCount ?? 1000, pr.gridProperties?.columnCount ?? 26);
            doc.sheets.push(s);
            return { addSheet: { properties: { sheetId: s.sheetId, title: s.title, index: s.index, gridProperties: { rowCount: s.rows, columnCount: s.cols } } } };
        }
        if (r.deleteSheet) { doc.sheets = doc.sheets.filter((s) => s.sheetId !== r.deleteSheet.sheetId); return {}; }
        if (r.updateSheetProperties) {
            const pr = r.updateSheetProperties.properties;
            const s = byId(pr.sheetId);
            if (pr.title) s.title = pr.title;
            if (pr.index !== undefined) s.index = pr.index;
            const gp = pr.gridProperties || {};
            if (gp.rowCount) { s.rows = gp.rowCount; s.data.length = Math.min(s.data.length, gp.rowCount); }
            if (gp.columnCount) { s.cols = gp.columnCount; s.data.forEach((row) => { row.length = Math.min(row.length, gp.columnCount); }); }
            if (gp.frozenRowCount !== undefined) s.frozen = gp.frozenRowCount;
            return {};
        }
        if (r.updateSpreadsheetProperties) { Object.assign(doc, r.updateSpreadsheetProperties.properties); return {}; }
        if (r.deleteDimension) {
            const { sheetId, startIndex, endIndex } = r.deleteDimension.range;
            const s = byId(sheetId);
            s.data.splice(startIndex, endIndex - startIndex);
            s.rows -= endIndex - startIndex;
            return {};
        }
        if (r.duplicateSheet) {
            const src = byId(r.duplicateSheet.sourceSheetId);
            const s = this._newSheet(r.duplicateSheet.newSheetName, doc.sheets.length, src.rows, src.cols, src.data);
            doc.sheets.push(s);
            return { duplicateSheet: { properties: { sheetId: s.sheetId, title: s.title, index: s.index } } };
        }
        if (r.repeatCell || r.autoResizeDimensions || r.setBasicFilter) {
            const sid = (r.repeatCell?.range || r.autoResizeDimensions?.dimensions || r.setBasicFilter?.filter.range).sheetId;
            byId(sid).formats.push(r);
            return {};
        }
        throw this._err(400, 'INVALID_ARGUMENT', `fake: неизвестный запрос ${Object.keys(r)[0]}`);
    }
}

module.exports = { FakeGoogle, parseRange };
