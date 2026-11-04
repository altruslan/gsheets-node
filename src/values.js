'use strict';

/** Преобразование значений JS ↔ ячейки, CSV. */

const CELL_MAX_CHARS = 50000; // лимит Google Sheets на ячейку

/**
 * JS-значение → значение ячейки.
 *   null/undefined → ''      Date → 'YYYY-MM-DD HH:mm:ss' (UTC) или ISO
 *   массив примитивов → 'a, b'   объект → JSON
 */
function toCell(v, opts = {}) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return Number.isFinite(v) ? v : '';
    if (typeof v === 'boolean') return v;
    if (typeof v === 'bigint') return v.toString();
    if (v instanceof Date) {
        if (Number.isNaN(v.getTime())) return '';
        return opts.dateFormat === 'iso' ? v.toISOString() : v.toISOString().replace('T', ' ').slice(0, 19);
    }
    let s;
    if (Array.isArray(v)) {
        s = v.every((x) => x === null || ['string', 'number', 'boolean'].includes(typeof x))
            ? v.filter((x) => x !== null && x !== '').join(opts.arraySeparator ?? ', ')
            : JSON.stringify(v);
    } else if (typeof v === 'object') {
        s = JSON.stringify(v);
    } else {
        s = String(v);
    }
    if (s.length > CELL_MAX_CHARS) s = s.slice(0, CELL_MAX_CHARS);
    return s;
}

/** Строковое представление для сравнения «изменилось или нет». */
function cellKey(v) {
    if (v === null || v === undefined) return '';
    return String(v);
}

function padRow(row, len) {
    const out = Array.isArray(row) ? row.slice(0, Math.max(len, row.length)) : [];
    while (out.length < len) out.push('');
    return out;
}

// ─── CSV ─────────────────────────────────────────────────────────────────────

/** RFC 4180: кавычки, переводы строк внутри полей, "" как экранирование. */
function parseCsv(text, { delimiter = ',' } = {}) {
    const s = String(text).replace(/^\uFEFF/, '');
    const rows = [];
    let row = [];
    let field = '';
    let i = 0;
    let quoted = false;

    while (i < s.length) {
        const ch = s[i];
        if (quoted) {
            if (ch === '"') {
                if (s[i + 1] === '"') { field += '"'; i += 2; continue; }
                quoted = false; i++; continue;
            }
            field += ch; i++; continue;
        }
        if (ch === '"' && field === '') { quoted = true; i++; continue; }
        if (ch === delimiter) { row.push(field); field = ''; i++; continue; }
        if (ch === '\r') { i++; continue; }
        if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
        field += ch; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
}

function toCsv(rows, { delimiter = ',', eol = '\n', bom = false } = {}) {
    const esc = (v) => {
        const s = v === null || v === undefined ? '' : String(v);
        return /["\r\n]/.test(s) || s.includes(delimiter) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    return (bom ? '\uFEFF' : '') + rows.map((r) => r.map(esc).join(delimiter)).join(eol) + (rows.length ? eol : '');
}

module.exports = { toCell, cellKey, padRow, parseCsv, toCsv, CELL_MAX_CHARS };
