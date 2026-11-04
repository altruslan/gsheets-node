'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { GSheets, GSheetsError, RateLimiter, a1, parseCsv, toCsv, toCell } = require('..');
const { FakeGoogle } = require('./fake-google');

const silent = { warn() {}, error() {}, debug() {} };
const DOC = 'testDoc1234567890';

function setup(opts = {}) {
    const g = new FakeGoogle();
    g.addDoc(DOC, 'Каталог', [{ title: 'Лист1', rows: 5, cols: 3 }]);
    const gs = new GSheets({ authClient: g, logger: silent, retryDelay: 1, readsPerMinute: 0, writesPerMinute: 0, ...opts });
    return { g, gs };
}

const sheetData = (g, title, id = DOC) => g._trim(g.docs.get(id).sheets.find((s) => s.title === title).data);

test('a1: столбцы, диапазоны, ссылки', () => {
    assert.equal(a1.colToLetter(1), 'A');
    assert.equal(a1.colToLetter(26), 'Z');
    assert.equal(a1.colToLetter(27), 'AA');
    assert.equal(a1.colToLetter(18278), 'ZZZ');
    assert.equal(a1.letterToCol('AA'), 27);
    assert.equal(a1.quoteSheet("O'Neil list"), "'O''Neil list'");
    assert.equal(a1.range('Товары', 2, 1, 10, 5), "'Товары'!A2:E10");
    assert.equal(a1.range('T', 1, 1), "'T'!A1");
    assert.deepEqual(a1.parseCell('C12'), { row: 12, col: 3 });
    assert.deepEqual(
        a1.parseSpreadsheetUrl('https://docs.google.com/spreadsheets/d/1AbC_def-GHIjklmnop/edit#gid=42'),
        { id: '1AbC_def-GHIjklmnop', gid: 42 },
    );
    assert.equal(a1.parseSpreadsheetUrl('1AbC_def-GHIjklmnop').id, '1AbC_def-GHIjklmnop');
    assert.throws(() => a1.parseSpreadsheetUrl('bad'));
});

test('values: toCell и CSV', () => {
    assert.equal(toCell(null), '');
    assert.equal(toCell(NaN), '');
    assert.equal(toCell(['a', 'b', null]), 'a, b');
    assert.equal(toCell({ x: 1 }), '{"x":1}');
    assert.equal(toCell(new Date('2026-11-04T10:20:30Z')), '2026-11-04 10:20:30');
    assert.equal(toCell('x'.repeat(60000)).length, 50000);

    const csv = 'sku,name\r\nA1,"Очки, ""Aviator"""\nB2,"две\nстроки"\n';
    const rows = parseCsv(csv);
    assert.deepEqual(rows, [['sku', 'name'], ['A1', 'Очки, "Aviator"'], ['B2', 'две\nстроки']]);
    assert.deepEqual(parseCsv(toCsv(rows)), rows);
    assert.deepEqual(parseCsv('a;b\n1;2', { delimiter: ';' }), [['a', 'b'], ['1', '2']]);
});

test('open: по ссылке, листы из кэша', async () => {
    const { gs } = setup();
    const doc = await gs.open(`https://docs.google.com/spreadsheets/d/${DOC}/edit#gid=0`);
    assert.equal(doc.title, 'Каталог');
    assert.equal(doc.sheets.length, 1);
    assert.equal(doc.sheet('Лист1').rowCount, 5);
    assert.equal(doc.sheet(), doc.sheets[0]);
    await assert.rejects(gs.open('missingDoc12345'), (e) => e instanceof GSheetsError && e.isNotFound);
});

test('листы: add / ensure / rename / duplicate / delete / reorder', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);

    const p = await doc.addSheet('Товары', { headers: ['sku', 'price'] });
    assert.deepEqual(sheetData(g, 'Товары'), [['sku', 'price']]);
    assert.equal(g.docs.get(DOC).sheets.find((s) => s.title === 'Товары').frozen, 1);

    const same = await doc.ensureSheet('Товары', { headers: ['other'] });
    assert.equal(same.sheetId, p.sheetId);
    assert.deepEqual(await same.getHeaders(), ['sku', 'price']); // не перезаписаны

    await p.rename('Products');
    assert.ok(doc.sheet('Products'));
    const copy = await p.duplicate('Products-copy');
    assert.equal(copy.title, 'Products-copy');
    await doc.reorderSheets(['Products-copy']);
    assert.equal(doc.sheet('Products-copy').index, 0);
    await doc.deleteSheet('Products-copy');
    assert.equal(doc.sheet('Products-copy'), null);
    await assert.rejects(doc.getSheet('Нет такого'), (e) => e.isNotFound);
});

test('write расширяет сетку, read/getCell/setCell', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = doc.sheet('Лист1'); // 5 × 3

    const rows = Array.from({ length: 12 }, (_, i) => [i, `r${i}`, i * 2, 'x', 'y']);
    await s.write('A1', rows);
    assert.equal(sheetData(g, 'Лист1').length, 12);
    assert.ok(s.rowCount >= 12 && s.columnCount >= 5);

    await s.setCell('B2', 'новое');
    assert.equal(await s.getCell('B2'), 'новое');
    const part = await s.read('A1:B2');
    assert.deepEqual(part, [[0, 'r0'], [1, 'новое']]);

    await s.writeMany({ 'D1': [['d1']], 'E2:E3': [['e2'], ['e3']] });
    assert.equal(await s.getCell('E3'), 'e3');
});

test('appendObjects: заголовки из ключей, новые столбцы добавляются', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = await doc.addSheet('Items', { rows: 2, cols: 2 });

    await s.appendObjects([{ sku: 'A', price: 10 }, { sku: 'B', price: 20 }]);
    await s.appendObjects([{ sku: 'C', price: 30, color: 'red' }]);
    assert.deepEqual(sheetData(g, 'Items'), [
        ['sku', 'price', 'color'],
        ['A', 10], ['B', 20],
        ['C', 30, 'red'],
    ]);

    const objs = await s.readObjects();
    assert.equal(objs.length, 3);
    assert.deepEqual(objs[2], { sku: 'C', price: 30, color: 'red' });
    assert.equal(objs[2]._row, 4);
    assert.equal(JSON.stringify(objs[0]).includes('_row'), false);

    const found = await s.findObjects({ sku: 'B' });
    assert.equal(found[0].price, 20);
});

test('upsertObjects: обновление, вставка, без изменений, merge, timestamp', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = await doc.addSheet('P', { rows: 3, cols: 3 });

    let r = await s.upsertObjects([
        { sku: 'A', price: 10, qty: 1 },
        { sku: 'B', price: 20, qty: 2 },
    ], { key: 'sku', timestampColumn: 'updated_at', createdColumn: 'created_at' });
    assert.deepEqual(r, { inserted: 2, updated: 0, unchanged: 0, skipped: 0 });
    let data = sheetData(g, 'P');
    assert.deepEqual(data[0], ['sku', 'price', 'qty', 'created_at', 'updated_at']);

    // ручная заметка в отдельном столбце не должна пропасть
    await s.write('F1', [['note'], ['важно']]);

    r = await s.upsertObjects([
        { sku: 'A', price: 11 },          // изменилась цена
        { sku: 'B', price: 20, qty: 2 },  // без изменений
        { sku: 'C', price: 30, qty: 3 },  // новая
        { sku: '', price: 1 },            // пустой ключ
    ], { key: 'sku', timestampColumn: 'updated_at', createdColumn: 'created_at' });
    assert.deepEqual(r, { inserted: 1, updated: 1, unchanged: 1, skipped: 1 });

    data = sheetData(g, 'P');
    assert.equal(data[1][0], 'A');
    assert.equal(data[1][1], 11);
    assert.equal(data[1][2], 1);          // qty сохранён (merge)
    assert.equal(data[1][5], 'важно');    // ручной столбец сохранён
    assert.equal(data[3][0], 'C');
    assert.match(String(data[3][3]), /^\d{4}-\d{2}-\d{2} /);

    // merge:false — строка перезаписывается, created_at сохраняется
    const created = data[1][3];
    await s.upsertObjects([{ sku: 'A', price: 12 }], { key: 'sku', merge: false, createdColumn: 'created_at' });
    data = sheetData(g, 'P');
    assert.equal(data[1][2], '');
    assert.equal(data[1][3], created);

    // insert:false — только обновление
    r = await s.upsertObjects([{ sku: 'Z', price: 1 }], { key: 'sku', insert: false });
    assert.equal(r.skipped, 1);
});

test('upsertObjects: составной ключ и подряд идущие строки одним диапазоном', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = await doc.addSheet('C');
    const items = Array.from({ length: 6 }, (_, i) => ({ brand: i < 3 ? 'X' : 'Y', sku: `S${i % 3}`, v: i }));
    await s.upsertObjects(items, { key: ['brand', 'sku'] });
    assert.equal(sheetData(g, 'C').length, 7);

    g.calls.length = 0;
    await s.upsertObjects(items.map((it) => ({ ...it, v: it.v + 100 })), { key: ['brand', 'sku'] });
    const upd = g.calls.filter((c) => c.url.includes('values:batchUpdate'));
    assert.equal(upd.length, 1);
    assert.equal(upd[0].data.data.length, 1); // строки 2..7 — один блок
    assert.equal(upd[0].data.data[0].range, "'C'!A2:C7");
});

test('replaceObjects очищает и обрезает лист', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = doc.sheet('Лист1');
    await s.write('A1', Array.from({ length: 10 }, () => ['old', 'old', 'old', 'old']));
    await s.replaceObjects([{ a: 1 }, { a: 2, b: 3 }]);
    assert.deepEqual(sheetData(g, 'Лист1'), [['a', 'b'], [1], [2, 3]]);
    assert.equal(s.rowCount, 3);
    assert.equal(s.columnCount, 2);
});

test('deleteRows / deleteObjects', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = await doc.addSheet('D');
    await s.appendObjects(['a', 'b', 'c', 'd', 'e'].map((sku) => ({ sku })));
    assert.equal(await s.deleteObjects((o) => ['b', 'c', 'e'].includes(o.sku)), 3);
    assert.deepEqual(sheetData(g, 'D'), [['sku'], ['a'], ['d']]);
    const del = g.calls.filter((c) => c.data?.requests?.[0]?.deleteDimension);
    assert.equal(del.at(-1).data.requests.length, 2); // b-c одним запросом, e отдельно
});

test('форматирование: formatHeader, freeze, autoResize, setNumberFormat', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = await doc.addSheet('F', { headers: ['sku', 'price'], formatHeader: false });
    await s.formatHeader({ filter: true });
    await s.freeze(1, 1);
    await s.autoResize();
    await s.setNumberFormat('price', '#,##0.00');
    const f = g.docs.get(DOC).sheets.find((x) => x.title === 'F').formats;
    assert.ok(f.some((r) => r.setBasicFilter));
    assert.ok(f.some((r) => r.repeatCell?.cell.userEnteredFormat.numberFormat?.pattern === '#,##0.00'));
    await assert.rejects(s.setNumberFormat('missing', '0'), GSheetsError);
});

test('CSV: импорт в лист и экспорт', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = doc.sheet('Лист1');
    await s.importCsv('sku;name\nA;"x;y"\nB;z', { delimiter: ';' });
    assert.deepEqual(sheetData(g, 'Лист1'), [['sku', 'name'], ['A', 'x;y'], ['B', 'z']]);
    assert.equal(await s.toCsv(), 'sku,name\nA,x;y\nB,z\n');
    const exported = await s.export('csv');
    assert.ok(Buffer.isBuffer(exported));
    assert.match(exported.toString(), /^sku,name/);
    const xlsx = await doc.export('xlsx');
    assert.match(xlsx.toString(), /spreadsheetml/);
});

test('create: с листами, в папке, share; ошибка квоты сервисного аккаунта', async () => {
    const { g, gs } = setup();
    const doc = await gs.create('Отчёт', { sheets: [{ title: 'Итог', headers: ['k', 'v'] }, 'Данные'], shareWith: ['boss@eurooptica.com'] });
    assert.deepEqual(doc.sheets.map((s) => s.title), ['Итог', 'Данные']);
    assert.equal(g.files.get(doc.id).permissions[0].emailAddress, 'boss@eurooptica.com');
    assert.equal(g.files.get(doc.id).permissions[0].role, 'writer');

    const inFolder = await gs.create('В папке', { folderId: 'folderX', timeZone: 'Europe/Kyiv' });
    assert.deepEqual(g.files.get(inFolder.id).parents, ['folderX']);

    g.request = ((orig) => async (o) => {
        if (o.url.endsWith('/v4/spreadsheets') && o.method === 'POST') {
            return { status: 403, data: { error: { code: 403, message: "The user's Drive storage quota has been exceeded.", errors: [{ reason: 'storageQuotaExceeded' }] } } };
        }
        return orig.call(g, o);
    })(g.request);
    await assert.rejects(gs.create('X'), (e) => e.isQuotaExceeded && /общем диске/.test(e.message));
});

test('find / openOrCreate / list / copy / delete', async () => {
    const { g, gs } = setup();
    const a = await gs.openOrCreate('Maui Jim', { sheets: ['Товары'] });
    const b = await gs.openOrCreate('Maui Jim', { sheets: ['Товары', 'Остатки'] });
    assert.equal(a.id, b.id);
    assert.ok(b.sheet('Остатки'));
    assert.equal((await gs.find("O'Reilly")).length, 0);
    assert.ok((await gs.list()).length >= 2);

    const c = await a.copy('Maui Jim копия');
    assert.notEqual(c.id, a.id);
    await c.delete();
    assert.equal(g.files.get(c.id).trashed, true);
});

test('Drive: загрузка, импорт CSV как таблицы, скачивание, папки, перемещение', async () => {
    const { g, gs } = setup();
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gs-'));
    const csvPath = path.join(tmp, 'stock.csv');
    fs.writeFileSync(csvPath, 'sku,qty\nA,1\nB,2');

    const folder = await gs.ensureFolder('Выгрузки');
    const again = await gs.ensureFolder('Выгрузки');
    assert.equal(folder.id, again.id);

    const up = await gs.uploadFile(csvPath, { folderId: folder.id });
    assert.equal(up.name, 'stock.csv');
    assert.equal(g.files.get(up.id).mimeType, 'text/csv');
    assert.equal((await gs.downloadFile(up.id)).toString(), 'sku,qty\nA,1\nB,2');

    const doc = await gs.importFile(csvPath, { name: 'Остатки' });
    assert.equal(doc.title, 'Остатки');
    assert.deepEqual((await doc.sheet().readObjects())[1], { sku: 'B', qty: '2' });

    await assert.rejects(gs.uploadFile(Buffer.from('x'), { name: 'a.png', convert: true }), GSheetsError);

    const big = Buffer.alloc(6 * 1024 * 1024, 1);
    const bigUp = await gs.uploadFile(big, { name: 'big.zip' });
    assert.equal(Number(bigUp.size), big.length);

    await gs.moveFile(up.id, 'otherFolder');
    assert.deepEqual(g.files.get(up.id).parents, ['otherFolder']);

    await gs.share(up.id, { anyone: true }, 'reader');
    assert.equal((await gs.listPermissions(up.id))[0].type, 'anyone');
    await assert.rejects(gs.share(up.id, {}, 'reader'), GSheetsError);
});

test('повторы: 503 и 429 с Retry-After, без повтора на 400', async () => {
    const { g, gs } = setup();
    g.fail.push({ status: 503, times: 2 });
    const doc = await gs.open(DOC);
    assert.equal(doc.title, 'Каталог');

    g.fail.push({ status: 429, times: 1, retryAfter: 0.01 });
    await doc.refresh();

    const before = g.calls.length;
    await assert.rejects(doc.getValues("'Нет'!A1"), (e) => e.status === 400 && !e.isRetryable);
    assert.equal(g.calls.length - before, 1);
});

test('RateLimiter ограничивает частоту', async () => {
    const rl = new RateLimiter(3, 200);
    const t0 = Date.now();
    for (let i = 0; i < 4; i++) await rl.take();
    assert.ok(Date.now() - t0 >= 190);
});

test('большие объёмы делятся на несколько запросов', async () => {
    const { g, gs } = setup();
    const doc = await gs.open(DOC);
    const s = await doc.addSheet('Big', { rows: 10, cols: 5 });
    const long = 'x'.repeat(300);
    const rows = Array.from({ length: 8000 }, (_, i) => [i, long, long]);
    g.calls.length = 0;
    await s.write('A1', rows);
    const writes = g.calls.filter((c) => c.url.includes('values:batchUpdate'));
    assert.ok(writes.length > 1, `ожидалось несколько запросов, было ${writes.length}`);
    for (const w of writes) assert.ok(JSON.stringify(w.data).length < 2_000_000);
    assert.equal(sheetData(g, 'Big').length, 8000);
    assert.equal(sheetData(g, 'Big')[7999][0], 7999);
});

test('USER_ENTERED передаётся и через опции метода', async () => {
    const { g, gs } = setup({ valueInput: 'USER_ENTERED' });
    const doc = await gs.open(DOC);
    await doc.sheet().write('A1', [['42']]);
    assert.equal(sheetData(g, 'Лист1')[0][0], 42);
    await doc.sheet().write('A2', [['42']], { input: 'RAW' });
    assert.equal(sheetData(g, 'Лист1')[1][0], '42');
});
