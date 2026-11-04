'use strict';

/**
 * Пример: сохранение результатов парсера поставщика в Google Таблицу.
 *
 * Таблица одна на поставщика, лист — на бренд. Строки обновляются по SKU,
 * товары, которых нет в текущем прогоне, помечаются в столбце status.
 *
 *   node examples/parser-results.js
 */

const { GSheets } = require('..');

const cfg = {
    keyFile: process.env.GOOGLE_KEY_FILE || 'config/google-service-account.json',
    spreadsheet: process.env.GSHEET_ID, // таблица, открытая для e-mail сервисного аккаунта (Редактор)
};

async function saveResults(vendor, brand, items) {
    const gs = new GSheets({ keyFile: cfg.keyFile });
    const doc = await gs.open(cfg.spreadsheet);

    const sheet = await doc.ensureSheet(brand, {
        headers: ['sku', 'model', 'color', 'size', 'price', 'quantity', 'status', 'created_at', 'updated_at'],
    });

    const runStarted = new Date();
    const res = await sheet.upsertObjects(
        items.map((it) => ({ ...it, status: 'active' })),
        { key: 'sku', timestampColumn: 'updated_at', createdColumn: 'created_at' },
    );

    // Товары, пропавшие у поставщика
    const seen = new Set(items.map((i) => String(i.sku)));
    const gone = (await sheet.readObjects())
        .filter((o) => o.status === 'active' && !seen.has(String(o.sku)))
        .map((o) => ({ sku: o.sku, status: 'missing', quantity: 0 }));
    if (gone.length) await sheet.upsertObjects(gone, { key: 'sku', insert: false, timestampColumn: 'updated_at' });

    // Сводка прогонов
    const log = await doc.ensureSheet('Журнал', { headers: ['time', 'vendor', 'brand', 'inserted', 'updated', 'unchanged', 'missing'] });
    await log.appendObjects([{ time: runStarted, vendor, brand, ...res, missing: gone.length }]);

    console.log(`${brand}: +${res.inserted} ~${res.updated} =${res.unchanged} −${gone.length} → ${sheet.url}`);
}

if (require.main === module) {
    saveResults('mauijim', 'Maui Jim', [
        { sku: 'B100-02', model: 'Peahi', color: 'Gloss Black', size: '65', price: 289, quantity: 4 },
        { sku: 'H202-10', model: 'Hookipa', color: 'Tortoise', size: '64', price: 229, quantity: 0 },
    ]).catch((err) => { console.error(err.message); process.exit(1); });
}

module.exports = { saveResults };
