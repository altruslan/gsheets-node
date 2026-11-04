'use strict';

const fs = require('fs');
const { Sheet } = require('./sheet');
const { GSheetsError } = require('./http');

const MAX_PAYLOAD = 1_500_000; // Google рекомендует до 2 МБ на запрос

const EXPORT_MIME = {
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ods: 'application/x-vnd.oasis.opendocument.spreadsheet',
    pdf: 'application/pdf',
    csv: 'text/csv',
    tsv: 'text/tab-separated-values',
    zip: 'application/zip', // HTML
};

function qs(params) {
    const parts = [];
    for (const [k, v] of Object.entries(params)) {
        if (v === undefined || v === null) continue;
        for (const item of Array.isArray(v) ? v : [v]) parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(item)}`);
    }
    return parts.length ? `?${parts.join('&')}` : '';
}

/** Документ Google Sheets. Создаётся через GSheets: client.open(), client.create(). */
class Spreadsheet {
    constructor(client, id) {
        this.client = client;
        this.id = id;
        this.title = '';
        this.locale = '';
        this.timeZone = '';
        this.sheets = [];
    }

    get url()  { return `https://docs.google.com/spreadsheets/d/${this.id}/edit`; }
    get http() { return this.client.http; }
    get base() { return `${this.client.SHEETS}/${this.id}`; }

    /** Перечитывает метаданные (название, листы, размеры). */
    async refresh() {
        const meta = await this.http.request({
            url: `${this.base}${qs({ fields: 'spreadsheetId,properties(title,locale,timeZone),sheets(properties)' })}`,
        });
        this._apply(meta);
        return this;
    }

    _apply(meta) {
        this.title = meta.properties?.title ?? this.title;
        this.locale = meta.properties?.locale ?? this.locale;
        this.timeZone = meta.properties?.timeZone ?? this.timeZone;
        const old = new Map(this.sheets.map((s) => [s.sheetId, s]));
        this.sheets = (meta.sheets || []).map((s) => {
            const ex = old.get(s.properties.sheetId);
            if (ex) { ex._setProps(s.properties); return ex; }
            return new Sheet(this, s.properties);
        });
    }

    // ── Листы ────────────────────────────────────────────────────────────────

    /** Лист по названию, sheetId (gid) или индексу { index: 0 }. Из кэша, без запроса. */
    sheet(ref) {
        if (ref === undefined) return this.sheets[0] || null;
        if (typeof ref === 'number') return this.sheets.find((s) => s.sheetId === ref) || null;
        if (ref && typeof ref === 'object' && 'index' in ref) return this.sheets.find((s) => s.index === ref.index) || null;
        return this.sheets.find((s) => s.title === ref) || null;
    }

    /** То же, что sheet(), но обновляет метаданные, если лист не найден, и бросает NOT_FOUND. */
    async getSheet(ref) {
        let s = this.sheet(ref);
        if (!s) { await this.refresh(); s = this.sheet(ref); }
        if (!s) throw new GSheetsError(`Лист не найден: ${JSON.stringify(ref)}`, { status: 404, code: 'NOT_FOUND' });
        return s;
    }

    /**
     * Новый лист.
     * @param {string} title
     * @param {object} [opts] { rows=1000, cols=26, index, headers, formatHeader=true }
     */
    async addSheet(title, opts = {}) {
        const properties = { title, gridProperties: { rowCount: opts.rows ?? 1000, columnCount: opts.cols ?? Math.max(26, opts.headers?.length || 0) } };
        if (opts.index !== undefined) properties.index = opts.index;
        const [reply] = await this.batchUpdate([{ addSheet: { properties } }]);
        const sheet = new Sheet(this, reply.addSheet.properties);
        this.sheets.push(sheet);
        if (opts.headers?.length) {
            await sheet.setHeaders(opts.headers);
            if (opts.formatHeader !== false) await sheet.formatHeader();
        }
        return sheet;
    }

    /** Возвращает лист, создавая его при отсутствии. Если лист пуст, а headers заданы — записывает заголовки. */
    async ensureSheet(title, opts = {}) {
        let s = this.sheet(title);
        if (!s) { await this.refresh(); s = this.sheet(title); }
        if (!s) return this.addSheet(title, opts);
        if (opts.headers?.length && (await s.getHeaders()).length === 0) {
            await s.setHeaders(opts.headers);
            if (opts.formatHeader !== false) await s.formatHeader();
        }
        return s;
    }

    async deleteSheet(ref) {
        const s = ref instanceof Sheet ? ref : await this.getSheet(ref);
        await this.batchUpdate([{ deleteSheet: { sheetId: s.sheetId } }]);
        this.sheets = this.sheets.filter((x) => x.sheetId !== s.sheetId);
    }

    /** Порядок листов: reorderSheets(['Итог', 'Товары']) — перечисленные встают первыми. */
    async reorderSheets(titles) {
        const requests = titles.map((t, i) => {
            const s = this.sheet(t);
            if (!s) throw new GSheetsError(`Лист не найден: ${t}`, { status: 404, code: 'NOT_FOUND' });
            return { updateSheetProperties: { properties: { sheetId: s.sheetId, index: i }, fields: 'index' } };
        });
        await this.batchUpdate(requests);
        await this.refresh();
    }

    // ── Документ ─────────────────────────────────────────────────────────────

    async rename(title) {
        await this.batchUpdate([{ updateSpreadsheetProperties: { properties: { title }, fields: 'title' } }]);
        this.title = title;
        return this;
    }

    /** Произвольные запросы spreadsheets.batchUpdate. Возвращает массив replies. */
    async batchUpdate(requests) {
        if (!requests?.length) return [];
        const res = await this.http.request({ url: `${this.base}:batchUpdate`, method: 'POST', data: { requests } });
        return res.replies || [];
    }

    // ── Значения ─────────────────────────────────────────────────────────────

    _valueOpts(opts = {}) {
        return {
            valueRenderOption: opts.render || 'UNFORMATTED_VALUE',
            dateTimeRenderOption: opts.dateTimeRender || 'FORMATTED_STRING',
            majorDimension: opts.majorDimension || 'ROWS',
        };
    }

    /** Значения одного диапазона в A1-нотации ('Лист'!A1:C10). */
    async getValues(range, opts) {
        const res = await this.http.request({ url: `${this.base}/values/${encodeURIComponent(range)}${qs(this._valueOpts(opts))}` });
        return res.values || [];
    }

    /** Значения нескольких диапазонов одним запросом → { [range]: values } */
    async getValuesMany(ranges, opts) {
        const res = await this.http.request({ url: `${this.base}/values:batchGet${qs({ ranges, ...this._valueOpts(opts) })}` });
        const out = {};
        (res.valueRanges || []).forEach((vr, i) => { out[ranges[i]] = vr.values || []; });
        return out;
    }

    /**
     * Запись нескольких диапазонов (values.batchUpdate) с разбиением на запросы до ~1.5 МБ.
     * @param {Array<{range:string, values:any[][]}>} data
     * @param {object} [opts] { input: 'RAW' | 'USER_ENTERED' }
     */
    async writeValues(data, opts = {}) {
        const valueInputOption = opts.input || this.client.valueInput;
        const batches = [];
        let cur = [];
        let size = 0;

        for (const entry of data) {
            for (const part of splitEntry(entry)) {
                const s = JSON.stringify(part.values).length + part.range.length;
                if (cur.length && size + s > MAX_PAYLOAD) { batches.push(cur); cur = []; size = 0; }
                cur.push(part);
                size += s;
            }
        }
        if (cur.length) batches.push(cur);

        let updatedCells = 0;
        for (const batch of batches) {
            const res = await this.http.request({
                url: `${this.base}/values:batchUpdate`,
                method: 'POST',
                data: { valueInputOption, data: batch },
            });
            updatedCells += res.totalUpdatedCells || 0;
        }
        return { updatedCells, requests: batches.length };
    }

    /** values.append — добавление строк после таблицы, с разбиением на части. */
    async appendValues(range, values, opts = {}) {
        const valueInputOption = opts.input || this.client.valueInput;
        for (const part of splitEntry({ range, values })) {
            await this.http.request({
                url: `${this.base}/values/${encodeURIComponent(range)}:append${qs({ valueInputOption, insertDataOption: 'INSERT_ROWS' })}`,
                method: 'POST',
                data: { values: part.values },
            });
        }
    }

    async clearValues(ranges) {
        await this.http.request({ url: `${this.base}/values:batchClear`, method: 'POST', data: { ranges } });
    }

    // ── Drive: доступ, копии, экспорт ────────────────────────────────────────

    /**
     * Доступ к документу.
     *   share('user@mail.com', 'writer')   роли: reader | commenter | writer | owner
     *   share({ anyone: true }, 'reader')  по ссылке
     *   share({ domain: 'eurooptica.com' }, 'reader')
     */
    async share(who, role = 'reader', { notify = false, message } = {}) {
        return this.client.share(this.id, who, role, { notify, message });
    }

    async permissions() {
        return this.client.listPermissions(this.id);
    }

    /** Копия документа (Drive files.copy). Возвращает новый Spreadsheet. */
    async copy(name, { folderId } = {}) {
        const res = await this.http.request({
            url: `${this.client.DRIVE}/files/${this.id}/copy${qs({ supportsAllDrives: true, fields: 'id' })}`,
            method: 'POST',
            kind: 'none',
            data: { name: name || `${this.title} (копия)`, ...(folderId ? { parents: [folderId] } : {}) },
        });
        return this.client.open(res.id);
    }

    /** Перемещает документ в папку Drive. */
    async moveTo(folderId) {
        return this.client.moveFile(this.id, folderId);
    }

    /**
     * Экспорт: 'xlsx' | 'ods' | 'pdf' | 'csv' | 'tsv' | 'zip'. Для csv/tsv/pdf можно указать gid листа.
     * @returns {Promise<Buffer>}
     */
    async export(format = 'xlsx', { gid } = {}) {
        if (!EXPORT_MIME[format]) throw new GSheetsError(`Неизвестный формат экспорта: ${format}`, { code: 'INVALID_ARGUMENT' });
        if (gid === undefined && format !== 'csv' && format !== 'tsv') {
            const data = await this.http.request({
                url: `${this.client.DRIVE}/files/${this.id}/export${qs({ mimeType: EXPORT_MIME[format] })}`,
                responseType: 'arraybuffer',
                kind: 'none',
            });
            return Buffer.from(data);
        }
        // csv/tsv через Drive отдают только первый лист — используем экспорт Sheets по gid
        const sheetGid = gid ?? this.sheets[0]?.sheetId ?? 0;
        const data = await this.http.request({
            url: `https://docs.google.com/spreadsheets/d/${this.id}/export${qs({ format, gid: sheetGid })}`,
            responseType: 'arraybuffer',
            kind: 'read',
        });
        return Buffer.from(data);
    }

    /** Экспорт в файл. */
    async download(filePath, format, opts) {
        const fmt = format || String(filePath).split('.').pop().toLowerCase();
        const buf = await this.export(fmt, opts);
        await fs.promises.writeFile(filePath, buf);
        return filePath;
    }

    /** В корзину (trash=true) или удалить навсегда. */
    async delete({ trash = true } = {}) {
        return this.client.deleteFile(this.id, { trash });
    }
}

/** Делит запись на части, чтобы каждая влезала в MAX_PAYLOAD. */
function splitEntry(entry) {
    const total = JSON.stringify(entry.values).length;
    if (total <= MAX_PAYLOAD || entry.values.length <= 1) return [entry];

    const m = /^(.*!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+)?)?$/.exec(entry.range);
    const perRow = Math.max(1, Math.floor(entry.values.length * MAX_PAYLOAD / total));
    const parts = [];
    for (let i = 0; i < entry.values.length; i += perRow) {
        const values = entry.values.slice(i, i + perRow);
        let range = entry.range;
        if (m) {
            const startRow = Number(m[3]) + i;
            range = `${m[1] || ''}${m[2]}${startRow}${m[4] ? `:${m[4]}${startRow + values.length - 1}` : ''}`;
        }
        parts.push({ range, values });
    }
    return parts;
}

module.exports = { Spreadsheet, EXPORT_MIME, qs, splitEntry };
