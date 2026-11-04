'use strict';

const { GSheets, SCOPES } = require('./src/client');
const { Spreadsheet } = require('./src/spreadsheet');
const { Sheet } = require('./src/sheet');
const { GSheetsError, RateLimiter } = require('./src/http');
const a1 = require('./src/a1');
const values = require('./src/values');

module.exports = {
    GSheets,
    Spreadsheet,
    Sheet,
    GSheetsError,
    RateLimiter,
    SCOPES,
    a1,
    toCell: values.toCell,
    parseCsv: values.parseCsv,
    toCsv: values.toCsv,
    createClient: (opts) => new GSheets(opts),
};
