'use strict';

const { colToLetter, quoteSheet, parseCell } = require('./a1');
const { toCell, cellKey, padRow, parseCsv, toCsv } = require('./values');
const { GSheetsError } = require('./http');

/**
 * Лист (вкладка) таблицы. Создаётся через Spreadsheet: doc.sheet(), doc.addSheet(), doc.ensureSheet().
 */
class Sheet {
    constructor(doc, properties) {
        this.doc = doc;
        this._setProps(properties);
        this._headers = null;
    }

    _setProps(p) {
        this.props = p;
        this.sheetId = p.sheetId;
        this.title = p.title;
        this.index = p.index;
    }

    get rowCount()    { return this.props.gridProperties?.rowCount ?? 0; }
    get columnCount() { return this.props.gridProperties?.columnCount ?? 0; }
    get frozenRows()  { return this.props.gridProperties?.frozenRowCount ?? 0; }
    get url()         { return `${this.doc.url}#gid=${this.sheetId}`; }
    get logger()      { return this.doc.client.logger; }

    /** A1-диапазон на этом листе: a1('A1:C10') → 'Sheet'!A1:C10 */
    a1(range) {
        return range ? `${quoteSheet(this.title)}!${range}` : quoteSheet(this.title);
    }

    async refresh() {
        await this.doc.refresh();
        const fresh = this.doc.sheets.find((s) => s.sheetId === this.sheetId);
        if (!fresh) throw new GSheetsError(`Лист ${this.title} удалён`, { status: 404, code: 'NOT_FOUND' });
        this._setProps(fresh.props);
        return this;
    }

    // ── Чтение ───────────────────────────────────────────────────────────────

    /**
     * Значения диапазона (массив строк).
     * @param {string} [range]  'A1:D20', 'A:A', '2:2'; без аргумента — весь лист
     * @param {object} [opts]
     * @param {'UNFORMATTED_VALUE'|'FORMATTED_VALUE'|'FORMULA'} [opts.render='UNFORMATTED_VALUE']
     */
    async read(range, opts = {}) {
        return this.doc.getValues(this.a1(range), opts);
    }

    /** Значение одной ячейки: getCell('B2') */
    async getCell(cell, opts) {
        const v = await this.read(`${cell}:${cell}`, opts);
        return v[0]?.[0] ?? '';
    }

    /** Заголовки (строка headerRow). Кэшируются до записи заголовков через этот объект. */
    async getHeaders({ headerRow = 1, refresh = false } = {}) {
        if (this._headers && !refresh && this._headers.row === headerRow) return this._headers.list.slice();
        const v = await this.read(`${headerRow}:${headerRow}`);
        const list = (v[0] || []).map((h) => String(h ?? '').trim());
        while (list.length && list[list.length - 1] === '') list.pop();
        this._headers = { row: headerRow, list };
        return list.slice();
    }

    /** Записывает строку заголовков. */
    async setHeaders(headers, { headerRow = 1 } = {}) {
        await this.write(`A${headerRow}`, [headers]);
        this._headers = { row: headerRow, list: headers.slice() };
        return headers;
    }

    /**
     * Строки как объекты по заголовкам. Пустые строки пропускаются.
     * У каждого объекта есть неперечисляемое свойство _row — номер строки на листе.
     */
    async readObjects({ headerRow = 1, render } = {}) {
        const values = await this.read(undefined, { render });
        const headers = (values[headerRow - 1] || []).map((h) => String(h ?? '').trim());
        this._headers = { row: headerRow, list: headers.slice() };
        const out = [];
        for (let r = headerRow; r < values.length; r++) {
            const row = values[r] || [];
            if (!row.some((c) => c !== '' && c !== null && c !== undefined)) continue;
            const obj = {};
            headers.forEach((h, i) => { if (h) obj[h] = row[i] ?? ''; });
            Object.defineProperty(obj, '_row', { value: r + 1, enumerable: false });
            out.push(obj);
        }
        return out;
    }

    /** Поиск объектов по условию: findObjects(o => o.brand === 'Maui Jim') или { brand: 'Maui Jim' } */
    async findObjects(where, opts) {
        const pred = typeof where === 'function'
            ? where
            : (o) => Object.entries(where).every(([k, v]) => cellKey(o[k]) === cellKey(v));
        return (await this.readObjects(opts)).filter(pred);
    }

    // ── Запись ───────────────────────────────────────────────────────────────

    /**
     * Записывает двумерный массив, начиная с ячейки (по умолчанию A1). Лист расширяется при необходимости.
     * Большие объёмы автоматически делятся на запросы до ~1.5 МБ.
     */
    async write(startCell = 'A1', rows, opts = {}) {
        if (!Array.isArray(rows) || rows.length === 0) return { updatedRows: 0 };
        const { row, col } = parseCell(startCell);
        const width = Math.max(...rows.map((r) => r.length), 1);
        await this.ensureSize(row + rows.length - 1, col + width - 1);
        const values = rows.map((r) => r.map((v) => toCell(v, this.doc.client.cellOpts)));
        await this.doc.writeValues([{ range: this._block(row, col, values), values }], opts);
        return { updatedRows: rows.length };
    }

    /** Записывает значение в одну ячейку. */
    async setCell(cell, value, opts) {
        return this.write(cell, [[value]], opts);
    }

    /**
     * Несколько разрозненных диапазонов одним запросом:
     *   writeMany({ 'A1': [['x']], 'D5:E6': [[1,2],[3,4]] })
     */
    async writeMany(map, opts) {
        const data = Object.entries(map).map(([range, values]) => ({
            range: this.a1(range),
            values: values.map((r) => r.map((v) => toCell(v, this.doc.client.cellOpts))),
        }));
        return this.doc.writeValues(data, opts);
    }

    /** Добавляет строки после последней заполненной (values.append, лист растёт сам). */
    async append(rows, opts = {}) {
        if (!rows?.length) return { appended: 0 };
        const values = rows.map((r) => r.map((v) => toCell(v, this.doc.client.cellOpts)));
        await this.doc.appendValues(this.a1('A1'), values, opts);
        return { appended: rows.length };
    }

    /** Очищает значения (форматирование сохраняется). Без аргумента — весь лист. */
    async clear(range) {
        await this.doc.clearValues([this.a1(range)]);
        if (!range) this._headers = null;
    }

    // ── Объекты ──────────────────────────────────────────────────────────────

    _collectColumns(items, base = [], opts = {}) {
        const cols = base.slice();
        const set = new Set(cols);
        const extra = opts.columns || [];
        for (const c of extra) if (!set.has(c)) { cols.push(c); set.add(c); }
        if (opts.addMissingColumns === false && base.length) return cols;
        for (const it of items) {
            for (const k of Object.keys(it)) if (!set.has(k)) { cols.push(k); set.add(k); }
        }
        return cols;
    }

    _rowFromObject(obj, headers) {
        return headers.map((h) => (Object.prototype.hasOwnProperty.call(obj, h) ? toCell(obj[h], this.doc.client.cellOpts) : ''));
    }

    /**
     * Добавляет объекты в конец. Заголовки создаются из ключей, новые ключи добавляются новыми столбцами
     * (addMissingColumns=false — лишние ключи игнорируются).
     * @param {object[]} items
     * @param {object} [opts] { columns, addMissingColumns=true, headerRow=1 }
     */
    async appendObjects(items, opts = {}) {
        if (!items?.length) return { appended: 0 };
        const headerRow = opts.headerRow ?? 1;
        const current = await this.getHeaders({ headerRow });
        const headers = this._collectColumns(items, current, opts);
        if (headers.length !== current.length) await this.setHeaders(headers, { headerRow });
        return this.append(items.map((o) => this._rowFromObject(o, headers)), opts);
    }

    /**
     * Полная замена содержимого листа: заголовки + строки. Удобно для «снимка» каталога.
     * @param {object} [opts] { columns, trim=true — удалить лишние пустые строки/столбцы }
     */
    async replaceObjects(items, opts = {}) {
        const headers = this._collectColumns(items || [], [], opts);
        const rows = [headers, ...(items || []).map((o) => this._rowFromObject(o, headers))];
        await this.clear();
        await this.write('A1', rows, opts);
        this._headers = { row: 1, list: headers.slice() };
        if (opts.trim !== false) {
            await this.resize(Math.max(rows.length, 2), Math.max(headers.length, 1));
        }
        return { written: items?.length || 0, columns: headers.length };
    }

    /**
     * Обновляет строки по ключу и добавляет новые.
     *
     * @param {object[]} items
     * @param {object} opts
     * @param {string|string[]} opts.key          Столбец(ы) ключа, например 'sku' или ['brand', 'sku']
     * @param {boolean} [opts.merge=true]         true — меняются только переданные поля, остальные ячейки строки сохраняются;
     *                                            false — строка перезаписывается целиком
     * @param {boolean} [opts.addMissingColumns=true]
     * @param {boolean} [opts.insert=true]        false — только обновлять существующие
     * @param {string}  [opts.timestampColumn]    Столбец, куда пишется время изменения (только при реальном изменении)
     * @param {string}  [opts.createdColumn]      Столбец со временем добавления строки
     * @param {string[]}[opts.columns]            Порядок столбцов для нового листа
     * @param {number}  [opts.headerRow=1]
     * @returns {Promise<{inserted:number, updated:number, unchanged:number, skipped:number}>}
     */
    async upsertObjects(items, opts = {}) {
        const keys = Array.isArray(opts.key) ? opts.key : [opts.key];
        if (!keys.length || keys.some((k) => !k)) throw new GSheetsError('upsertObjects: не указан key', { code: 'INVALID_ARGUMENT' });
        const headerRow = opts.headerRow ?? 1;
        const merge = opts.merge !== false;
        const insert = opts.insert !== false;
        const stats = { inserted: 0, updated: 0, unchanged: 0, skipped: 0 };
        if (!items?.length) return stats;

        const values = await this.read();
        const current = (values[headerRow - 1] || []).map((h) => String(h ?? '').trim());
        while (current.length && current[current.length - 1] === '') current.pop();

        const baseCols = current.length ? current : [...keys, ...(opts.columns || []).filter((c) => !keys.includes(c))];
        let headers = this._collectColumns(items, baseCols, { ...opts, columns: [...keys, ...(opts.columns || [])] });
        for (const c of [opts.createdColumn, opts.timestampColumn]) if (c && !headers.includes(c)) headers.push(c);
        const headersChanged = headers.length !== current.length;

        const keyIdx = keys.map((k) => headers.indexOf(k));
        const keyOfRow = (row) => keyIdx.map((i) => cellKey(row[i]).trim()).join('\u0001');
        const keyOfObj = (o) => keys.map((k) => cellKey(toCell(o[k], this.doc.client.cellOpts)).trim()).join('\u0001');
        const emptyKey = keys.map(() => '').join('\u0001');

        // Существующие строки
        const existing = new Map();
        let dupes = 0;
        for (let r = headerRow; r < values.length; r++) {
            const k = keyOfRow(values[r] || []);
            if (k === emptyKey) continue;
            if (existing.has(k)) { dupes++; continue; }
            existing.set(k, r);
        }
        if (dupes) this.logger.warn?.(`[gsheets] ${this.title}: ${dupes} дубликатов ключа на листе — обновляется первая строка`);

        // Последний объект с одинаковым ключом побеждает
        const incoming = new Map();
        for (const it of items) {
            const k = keyOfObj(it);
            if (k === emptyKey) { stats.skipped++; continue; }
            incoming.set(k, it);
        }

        const now = toCell(new Date(), this.doc.client.cellOpts);
        const tsIdx = opts.timestampColumn ? headers.indexOf(opts.timestampColumn) : -1;
        const crIdx = opts.createdColumn ? headers.indexOf(opts.createdColumn) : -1;
        const changed = new Map(); // rowIndex(0-based) → row
        const appended = [];

        for (const [k, obj] of incoming) {
            const fresh = this._rowFromObject(obj, headers);
            if (existing.has(k)) {
                const r = existing.get(k);
                const old = padRow(values[r], headers.length);
                const next = merge
                    ? old.map((v, i) => (Object.prototype.hasOwnProperty.call(obj, headers[i]) ? fresh[i] : v))
                    : fresh.map((v, i) => (i === crIdx ? old[i] : v));
                const diff = next.some((v, i) => i !== tsIdx && cellKey(v) !== cellKey(old[i]));
                if (!diff) { stats.unchanged++; continue; }
                if (tsIdx >= 0) next[tsIdx] = now;
                changed.set(r, next);
                stats.updated++;
            } else if (insert) {
                if (tsIdx >= 0) fresh[tsIdx] = now;
                if (crIdx >= 0) fresh[crIdx] = now;
                appended.push(fresh);
                stats.inserted++;
            } else {
                stats.skipped++;
            }
        }

        const lastRow = Math.max(values.length, headerRow); // 1-based номер последней заполненной строки
        const needRows = lastRow + appended.length;
        await this.ensureSize(needRows, headers.length);

        const data = [];
        if (headersChanged) data.push({ range: this._block(headerRow, 1, [headers]), values: [headers] });

        // Подряд идущие изменённые строки — одним диапазоном
        const idx = [...changed.keys()].sort((a, b) => a - b);
        for (let i = 0; i < idx.length;) {
            let j = i;
            while (j + 1 < idx.length && idx[j + 1] === idx[j] + 1) j++;
            const block = idx.slice(i, j + 1).map((r) => changed.get(r));
            data.push({ range: this._block(idx[i] + 1, 1, block), values: block });
            i = j + 1;
        }
        if (appended.length) data.push({ range: this._block(lastRow + 1, 1, appended), values: appended });

        if (data.length) await this.doc.writeValues(data, opts);
        this._headers = { row: headerRow, list: headers.slice() };
        return stats;
    }

    /** Удаляет строки по номерам (1-based). Подряд идущие удаляются одним запросом. */
    async deleteRows(rowNumbers) {
        const rows = [...new Set(rowNumbers)].filter((n) => Number.isInteger(n) && n > 0).sort((a, b) => b - a);
        if (!rows.length) return 0;
        const requests = [];
        for (let i = 0; i < rows.length;) {
            let j = i;
            while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++;
            requests.push({
                deleteDimension: {
                    range: { sheetId: this.sheetId, dimension: 'ROWS', startIndex: rows[j] - 1, endIndex: rows[i] },
                },
            });
            i = j + 1;
        }
        await this.doc.batchUpdate(requests);
        this.props.gridProperties.rowCount -= rows.length;
        return rows.length;
    }

    /** Удаляет строки-объекты, подходящие под условие. Возвращает число удалённых. */
    async deleteObjects(where, opts) {
        const found = await this.findObjects(where, opts);
        return this.deleteRows(found.map((o) => o._row));
    }

    // ── Структура листа ──────────────────────────────────────────────────────

    /** Увеличивает лист, если он меньше rows × cols (никогда не уменьшает). */
    async ensureSize(rows, cols) {
        if (rows <= this.rowCount && cols <= this.columnCount) return;
        await this.refresh(); // данные о размере могли устареть (append расширяет лист сам)
        if (rows <= this.rowCount && cols <= this.columnCount) return;
        await this.resize(Math.max(rows, this.rowCount), Math.max(cols, this.columnCount));
    }

    /** Устанавливает точный размер листа. Уменьшение удаляет данные за пределами! */
    async resize(rows, cols) {
        const gp = {};
        const fields = [];
        if (rows) { gp.rowCount = rows; fields.push('gridProperties.rowCount'); }
        if (cols) { gp.columnCount = cols; fields.push('gridProperties.columnCount'); }
        if (!fields.length) return;
        await this.doc.batchUpdate([{
            updateSheetProperties: { properties: { sheetId: this.sheetId, gridProperties: gp }, fields: fields.join(',') },
        }]);
        this.props.gridProperties = { ...this.props.gridProperties, ...gp };
    }

    async rename(title) {
        await this.doc.batchUpdate([{ updateSheetProperties: { properties: { sheetId: this.sheetId, title }, fields: 'title' } }]);
        this.title = title;
        this.props.title = title;
        return this;
    }

    /** Закрепляет строки/столбцы. */
    async freeze(rows = 1, cols = 0) {
        await this.doc.batchUpdate([{
            updateSheetProperties: {
                properties: { sheetId: this.sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
                fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
            },
        }]);
        Object.assign(this.props.gridProperties, { frozenRowCount: rows, frozenColumnCount: cols });
    }

    /** Оформляет строку заголовков: жирный шрифт, фон, закрепление, фильтр. */
    async formatHeader({ headerRow = 1, bold = true, background = { red: 0.93, green: 0.93, blue: 0.93 }, freeze = true, filter = false } = {}) {
        const cols = Math.max((await this.getHeaders({ headerRow })).length, 1);
        const requests = [{
            repeatCell: {
                range: { sheetId: this.sheetId, startRowIndex: headerRow - 1, endRowIndex: headerRow, startColumnIndex: 0, endColumnIndex: cols },
                cell: { userEnteredFormat: { textFormat: { bold }, backgroundColor: background } },
                fields: 'userEnteredFormat(textFormat.bold,backgroundColor)',
            },
        }];
        if (freeze) {
            requests.push({
                updateSheetProperties: {
                    properties: { sheetId: this.sheetId, gridProperties: { frozenRowCount: headerRow } },
                    fields: 'gridProperties.frozenRowCount',
                },
            });
        }
        if (filter) {
            requests.push({ setBasicFilter: { filter: { range: { sheetId: this.sheetId, startRowIndex: headerRow - 1, startColumnIndex: 0, endColumnIndex: cols } } } });
        }
        await this.doc.batchUpdate(requests);
    }

    /** Автоширина столбцов (по умолчанию — всех). */
    async autoResize(startCol = 1, endCol) {
        await this.doc.batchUpdate([{
            autoResizeDimensions: {
                dimensions: { sheetId: this.sheetId, dimension: 'COLUMNS', startIndex: startCol - 1, endIndex: endCol ?? this.columnCount },
            },
        }]);
    }

    /**
     * Формат чисел для столбца: setNumberFormat('price', '#,##0.00') или setNumberFormat(3, '0%').
     * type: NUMBER | CURRENCY | PERCENT | DATE | DATE_TIME | TEXT
     */
    async setNumberFormat(column, pattern, { type = 'NUMBER', headerRow = 1 } = {}) {
        const col = typeof column === 'number' ? column : (await this.getHeaders({ headerRow })).indexOf(column) + 1;
        if (col < 1) throw new GSheetsError(`Столбец ${column} не найден`, { code: 'INVALID_ARGUMENT' });
        await this.doc.batchUpdate([{
            repeatCell: {
                range: { sheetId: this.sheetId, startRowIndex: headerRow, startColumnIndex: col - 1, endColumnIndex: col },
                cell: { userEnteredFormat: { numberFormat: { type, pattern } } },
                fields: 'userEnteredFormat.numberFormat',
            },
        }]);
    }

    /** Копия листа в этой же таблице. */
    async duplicate(newTitle) {
        const [reply] = await this.doc.batchUpdate([{ duplicateSheet: { sourceSheetId: this.sheetId, newSheetName: newTitle } }]);
        await this.doc.refresh();
        return this.doc.sheet(reply.duplicateSheet.properties.sheetId);
    }

    /** Копия листа в другую таблицу (по ID). */
    async copyTo(spreadsheetId) {
        return this.doc.client.http.request({
            url: `${this.doc.client.SHEETS}/${this.doc.id}/sheets/${this.sheetId}:copyTo`,
            method: 'POST',
            data: { destinationSpreadsheetId: spreadsheetId },
        });
    }

    async delete() {
        return this.doc.deleteSheet(this);
    }

    // ── Импорт / экспорт ─────────────────────────────────────────────────────

    /** Загружает CSV-текст или Buffer в лист. replace=true (по умолчанию) очищает лист перед загрузкой. */
    async importCsv(csv, { delimiter = ',', startCell = 'A1', replace = true, ...opts } = {}) {
        const rows = parseCsv(Buffer.isBuffer(csv) ? csv.toString('utf8') : csv, { delimiter });
        if (replace) await this.clear();
        await this.write(startCell, rows, opts);
        this._headers = null;
        return { rows: rows.length };
    }

    /** CSV-строка из данных листа (собирается локально, без ограничений экспорта Drive). */
    async toCsv({ delimiter = ',', bom = false, render = 'FORMATTED_VALUE' } = {}) {
        return toCsv(await this.read(undefined, { render }), { delimiter, bom });
    }

    /** Экспорт листа средствами Google: 'csv' | 'tsv' | 'pdf' | 'xlsx'. Возвращает Buffer. */
    async export(format = 'csv') {
        return this.doc.export(format, { gid: this.sheetId });
    }

    // ── Внутреннее ───────────────────────────────────────────────────────────

    _block(row, col, values) {
        const width = Math.max(...values.map((r) => r.length), 1);
        return this.a1(`${colToLetter(col)}${row}:${colToLetter(col + width - 1)}${row + values.length - 1}`);
    }
}

module.exports = { Sheet };
