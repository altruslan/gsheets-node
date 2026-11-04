'use strict';

/** Утилиты A1-нотации и разбора ссылок на таблицы. */

/** 1 → A, 27 → AA */
function colToLetter(n) {
    if (!Number.isInteger(n) || n < 1) throw new RangeError(`Некорректный номер столбца: ${n}`);
    let s = '';
    while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

/** A → 1, AA → 27 */
function letterToCol(s) {
    const str = String(s).toUpperCase();
    if (!/^[A-Z]+$/.test(str)) throw new RangeError(`Некорректная буква столбца: ${s}`);
    let n = 0;
    for (const ch of str) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
}

/** Имя листа в кавычках для A1: My Sheet → 'My Sheet', O'Neil → 'O''Neil' */
function quoteSheet(title) {
    return `'${String(title).replace(/'/g, "''")}'`;
}

/**
 * Собирает диапазон: range('Products', 1, 1, 10, 5) → 'Products'!A1:E10
 * row/col начинаются с 1; endRow/endCol можно не указывать (открытый диапазон).
 */
function range(sheet, startRow, startCol = 1, endRow, endCol) {
    const q = quoteSheet(sheet);
    if (!startRow) return q;
    const a = `${colToLetter(startCol)}${startRow}`;
    if (endRow === undefined && endCol === undefined) return `${q}!${a}`;
    const b = `${endCol ? colToLetter(endCol) : ''}${endRow ?? ''}`;
    return `${q}!${a}:${b}`;
}

/** 'B3' → { row: 3, col: 2 } */
function parseCell(cell) {
    const m = /^([A-Za-z]+)(\d+)$/.exec(String(cell).trim());
    if (!m) throw new RangeError(`Некорректная ячейка: ${cell}`);
    return { row: Number(m[2]), col: letterToCol(m[1]) };
}

/**
 * Извлекает ID таблицы и gid листа из ссылки или возвращает ID как есть.
 *   https://docs.google.com/spreadsheets/d/<ID>/edit#gid=123 → { id: '<ID>', gid: 123 }
 */
function parseSpreadsheetUrl(input) {
    const s = String(input || '').trim();
    const m = /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/.exec(s);
    const id = m ? m[1] : s;
    if (!/^[a-zA-Z0-9_-]{10,}$/.test(id)) throw new RangeError(`Не удалось определить ID таблицы: ${input}`);
    const g = /[#&?]gid=(\d+)/.exec(s);
    return { id, gid: g ? Number(g[1]) : undefined };
}

module.exports = { colToLetter, letterToCol, quoteSheet, range, parseCell, parseSpreadsheetUrl };
