'use strict';

const fs = require('fs');
const path = require('path');
const { HttpClient, GSheetsError } = require('./http');
const { Spreadsheet, qs } = require('./spreadsheet');
const { parseSpreadsheetUrl } = require('./a1');

const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
];

const SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
const FOLDER_MIME = 'application/vnd.google-apps.folder';
const RESUMABLE_THRESHOLD = 5 * 1024 * 1024;

const MIME_BY_EXT = {
    csv: 'text/csv', tsv: 'text/tab-separated-values', txt: 'text/plain', json: 'application/json',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xls: 'application/vnd.ms-excel',
    ods: 'application/vnd.oasis.opendocument.spreadsheet', pdf: 'application/pdf', zip: 'application/zip',
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', webp: 'image/webp', gif: 'image/gif',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};
const CONVERTIBLE = new Set(['text/csv', 'text/tab-separated-values', MIME_BY_EXT.xlsx, MIME_BY_EXT.xls, MIME_BY_EXT.ods]);

const escQ = (s) => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");

/**
 * Клиент Google Sheets + Drive.
 *
 * Авторизация (одно из):
 *   keyFile:      путь к JSON-ключу сервисного аккаунта
 *   credentials:  объект ключа сервисного аккаунта или authorized_user
 *   oauth:        { clientId, clientSecret, refreshToken } — от имени пользователя Google
 *   authClient:   готовый клиент google-auth-library (или любой объект с методом request)
 *   (ничего)      Application Default Credentials / GOOGLE_APPLICATION_CREDENTIALS
 */
class GSheets {
    /**
     * @param {object} [opts]
     * @param {string} [opts.keyFile]
     * @param {object} [opts.credentials]
     * @param {{clientId:string, clientSecret:string, refreshToken:string}} [opts.oauth]
     * @param {object} [opts.authClient]
     * @param {string} [opts.subject]         Domain-wide delegation: действовать от имени пользователя Workspace
     * @param {string[]} [opts.scopes]
     * @param {'RAW'|'USER_ENTERED'} [opts.valueInput='RAW']  RAW — как есть; USER_ENTERED — как ввод в интерфейсе (формулы, даты)
     * @param {object} [opts.cell]            { dateFormat: 'iso'|'sheet', arraySeparator: ', ' }
     * @param {number} [opts.retries=6]
     * @param {number} [opts.readsPerMinute=60]
     * @param {number} [opts.writesPerMinute=60]
     * @param {object} [opts.logger=console]
     */
    constructor(opts = {}) {
        this.logger = opts.logger || console;
        this.valueInput = opts.valueInput || 'RAW';
        this.cellOpts = opts.cell || {};
        this.SHEETS = opts.sheetsUrl || 'https://sheets.googleapis.com/v4/spreadsheets';
        this.DRIVE = opts.driveUrl || 'https://www.googleapis.com/drive/v3';
        this.UPLOAD = opts.uploadUrl || 'https://www.googleapis.com/upload/drive/v3';

        this.authClient = opts.authClient || createAuth(opts);
        this.http = new HttpClient({ ...opts, authClient: this.authClient, logger: this.logger });
        this._serviceEmail = opts.credentials?.client_email
            || (opts.keyFile ? safeReadJson(opts.keyFile)?.client_email : undefined);
    }

    /** E-mail сервисного аккаунта — с ним нужно поделиться таблицей (роль «Редактор»). */
    get serviceAccountEmail() { return this._serviceEmail; }

    // ── Таблицы ──────────────────────────────────────────────────────────────

    /** Открывает таблицу по ID или ссылке. */
    async open(idOrUrl) {
        const { id } = parseSpreadsheetUrl(idOrUrl);
        const doc = new Spreadsheet(this, id);
        try {
            await doc.refresh();
        } catch (err) {
            if (err instanceof GSheetsError && (err.isNotFound || err.isForbidden) && this._serviceEmail) {
                err.message += ` — проверьте, что таблица открыта для ${this._serviceEmail}`;
            }
            throw err;
        }
        return doc;
    }

    /**
     * Создаёт таблицу.
     * @param {string} title
     * @param {object} [opts]
     * @param {Array<string|{title:string, headers?:string[]}>} [opts.sheets]  Листы (первый заменяет «Лист1»)
     * @param {string} [opts.folderId]  Папка Drive (для сервисного аккаунта — папка на общем диске)
     * @param {string} [opts.locale]    'ru_RU', 'uk_UA'
     * @param {string} [opts.timeZone]  'Europe/Kyiv'
     * @param {Array<{email:string, role?:string}>|string[]} [opts.shareWith]
     */
    async create(title, opts = {}) {
        let id;
        try {
            if (opts.folderId) {
                const f = await this.http.request({
                    url: `${this.DRIVE}/files${qs({ supportsAllDrives: true, fields: 'id' })}`,
                    method: 'POST',
                    kind: 'none',
                    data: { name: title, mimeType: SHEET_MIME, parents: [opts.folderId] },
                });
                id = f.id;
                if (opts.locale || opts.timeZone) {
                    const doc = new Spreadsheet(this, id);
                    const properties = {};
                    if (opts.locale) properties.locale = opts.locale;
                    if (opts.timeZone) properties.timeZone = opts.timeZone;
                    await doc.batchUpdate([{ updateSpreadsheetProperties: { properties, fields: Object.keys(properties).join(',') } }]);
                }
            } else {
                const properties = { title };
                if (opts.locale) properties.locale = opts.locale;
                if (opts.timeZone) properties.timeZone = opts.timeZone;
                const res = await this.http.request({ url: this.SHEETS, method: 'POST', data: { properties } });
                id = res.spreadsheetId;
            }
        } catch (err) {
            if (err instanceof GSheetsError && (err.isQuotaExceeded || /storage quota/i.test(err.message))) {
                err.message += ' — у сервисного аккаунта нет своего места в Drive: укажите folderId папки на общем диске, '
                    + 'используйте OAuth от имени пользователя или создайте таблицу вручную и поделитесь ею с сервисным аккаунтом';
            }
            throw err;
        }

        const doc = await this.open(id);
        const sheets = (opts.sheets || []).map((s) => (typeof s === 'string' ? { title: s } : s));
        if (sheets.length) {
            const first = doc.sheets[0];
            await first.rename(sheets[0].title);
            if (sheets[0].headers?.length) { await first.setHeaders(sheets[0].headers); await first.formatHeader(); }
            for (const s of sheets.slice(1)) await doc.addSheet(s.title, s);
        }
        for (const s of opts.shareWith || []) {
            const email = typeof s === 'string' ? s : s.email;
            await doc.share(email, (typeof s === 'object' && s.role) || 'writer', { notify: s.notify });
        }
        return doc;
    }

    /** Ищет таблицы по точному названию (в папке, если указана). */
    async find(title, { folderId } = {}) {
        const q = [`mimeType='${SHEET_MIME}'`, `name='${escQ(title)}'`, 'trashed=false'];
        if (folderId) q.push(`'${escQ(folderId)}' in parents`);
        return this.listFiles({ q: q.join(' and ') });
    }

    /** Открывает таблицу по названию или создаёт новую. */
    async openOrCreate(title, opts = {}) {
        const [found] = await this.find(title, opts);
        if (found) {
            const doc = await this.open(found.id);
            for (const s of opts.sheets || []) {
                const sd = typeof s === 'string' ? { title: s } : s;
                await doc.ensureSheet(sd.title, sd);
            }
            return doc;
        }
        return this.create(title, opts);
    }

    /** Список таблиц: { folderId, nameContains, limit } */
    async list({ folderId, nameContains, limit = 100 } = {}) {
        const q = [`mimeType='${SHEET_MIME}'`, 'trashed=false'];
        if (folderId) q.push(`'${escQ(folderId)}' in parents`);
        if (nameContains) q.push(`name contains '${escQ(nameContains)}'`);
        return this.listFiles({ q: q.join(' and '), limit });
    }

    // ── Drive: файлы и папки ─────────────────────────────────────────────────

    /** Список файлов Drive по запросу q (синтаксис Drive). */
    async listFiles({ q, limit = 1000, fields = 'id,name,mimeType,modifiedTime,webViewLink,parents,size' } = {}) {
        const out = [];
        let pageToken;
        do {
            const res = await this.http.request({
                url: `${this.DRIVE}/files${qs({
                    q, pageToken, pageSize: Math.min(limit - out.length, 1000),
                    fields: `nextPageToken,files(${fields})`,
                    supportsAllDrives: true, includeItemsFromAllDrives: true, orderBy: 'modifiedTime desc',
                })}`,
                kind: 'none',
            });
            out.push(...(res.files || []));
            pageToken = res.nextPageToken;
        } while (pageToken && out.length < limit);
        return out;
    }

    async createFolder(name, { parentId } = {}) {
        return this.http.request({
            url: `${this.DRIVE}/files${qs({ supportsAllDrives: true, fields: 'id,name,webViewLink' })}`,
            method: 'POST',
            kind: 'none',
            data: { name, mimeType: FOLDER_MIME, ...(parentId ? { parents: [parentId] } : {}) },
        });
    }

    /** Папка по имени (в родительской), создаётся при отсутствии. */
    async ensureFolder(name, { parentId } = {}) {
        const q = [`mimeType='${FOLDER_MIME}'`, `name='${escQ(name)}'`, 'trashed=false'];
        if (parentId) q.push(`'${escQ(parentId)}' in parents`);
        const [found] = await this.listFiles({ q: q.join(' and '), limit: 1 });
        return found || this.createFolder(name, { parentId });
    }

    /**
     * Загружает файл в Drive.
     * @param {string|Buffer} source       Путь к файлу или Buffer
     * @param {object} [opts]
     * @param {string} [opts.name]         Имя в Drive (по умолчанию — имя файла)
     * @param {string} [opts.folderId]
     * @param {string} [opts.mimeType]     Определяется по расширению
     * @param {boolean}[opts.convert=false] CSV/XLSX/ODS → Google Таблица
     * @param {string} [opts.fileId]       Заменить содержимое существующего файла
     * @returns {Promise<{id,name,mimeType,webViewLink}>}
     */
    async uploadFile(source, opts = {}) {
        const isBuf = Buffer.isBuffer(source);
        const data = isBuf ? source : await fs.promises.readFile(source);
        const name = opts.name || (isBuf ? 'upload' : path.basename(source));
        const ext = (path.extname(name).slice(1) || (isBuf ? '' : path.extname(source).slice(1))).toLowerCase();
        const mimeType = opts.mimeType || MIME_BY_EXT[ext] || 'application/octet-stream';

        const meta = {};
        if (!opts.fileId) meta.name = name;
        if (opts.convert) {
            if (!CONVERTIBLE.has(mimeType)) throw new GSheetsError(`Файл ${mimeType} нельзя преобразовать в таблицу`, { code: 'INVALID_ARGUMENT' });
            meta.mimeType = SHEET_MIME;
            if (opts.fileId) meta.name = name;
        }
        if (opts.folderId && !opts.fileId) meta.parents = [opts.folderId];

        const method = opts.fileId ? 'PATCH' : 'POST';
        const target = opts.fileId ? `${this.UPLOAD}/files/${opts.fileId}` : `${this.UPLOAD}/files`;
        const common = { supportsAllDrives: true, fields: 'id,name,mimeType,webViewLink,size' };

        if (data.length <= RESUMABLE_THRESHOLD) {
            const boundary = `gsheets${Date.now().toString(16)}${Math.random().toString(16).slice(2)}`;
            const body = Buffer.concat([
                Buffer.from(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n`),
                Buffer.from(`--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`),
                data,
                Buffer.from(`\r\n--${boundary}--`),
            ]);
            return this.http.request({
                url: `${target}${qs({ uploadType: 'multipart', ...common })}`,
                method,
                kind: 'none',
                headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
                data: body,
            });
        }

        // Возобновляемая загрузка для больших файлов
        const init = await this.http.request({
            url: `${target}${qs({ uploadType: 'resumable', ...common })}`,
            method,
            kind: 'none',
            headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Upload-Content-Type': mimeType, 'X-Upload-Content-Length': String(data.length) },
            data: meta,
            full: true,
        });
        const location = init.headers?.location || init.headers?.get?.('location');
        if (!location) throw new GSheetsError('Drive не вернул адрес для загрузки', { code: 'UPLOAD' });
        return this.http.request({ url: location, method: 'PUT', kind: 'none', headers: { 'Content-Type': mimeType }, data });
    }

    /** Импорт CSV/XLSX/ODS как новой Google Таблицы. Возвращает Spreadsheet. */
    async importFile(source, opts = {}) {
        const f = await this.uploadFile(source, { ...opts, convert: true });
        return this.open(f.id);
    }

    /** Скачивает файл Drive (не Google-документ) в Buffer или в файл. */
    async downloadFile(fileId, filePath) {
        const data = Buffer.from(await this.http.request({
            url: `${this.DRIVE}/files/${fileId}${qs({ alt: 'media', supportsAllDrives: true })}`,
            responseType: 'arraybuffer',
            kind: 'none',
        }));
        if (filePath) { await fs.promises.writeFile(filePath, data); return filePath; }
        return data;
    }

    async getFile(fileId, fields = 'id,name,mimeType,modifiedTime,webViewLink,parents,size') {
        return this.http.request({ url: `${this.DRIVE}/files/${fileId}${qs({ fields, supportsAllDrives: true })}`, kind: 'none' });
    }

    async moveFile(fileId, folderId) {
        const f = await this.getFile(fileId, 'parents');
        return this.http.request({
            url: `${this.DRIVE}/files/${fileId}${qs({ addParents: folderId, removeParents: (f.parents || []).join(','), supportsAllDrives: true, fields: 'id,parents' })}`,
            method: 'PATCH',
            kind: 'none',
            data: {},
        });
    }

    async deleteFile(fileId, { trash = true } = {}) {
        if (trash) {
            await this.http.request({ url: `${this.DRIVE}/files/${fileId}${qs({ supportsAllDrives: true })}`, method: 'PATCH', kind: 'none', data: { trashed: true } });
        } else {
            await this.http.request({ url: `${this.DRIVE}/files/${fileId}${qs({ supportsAllDrives: true })}`, method: 'DELETE', kind: 'none' });
        }
        return true;
    }

    /**
     * Доступ к файлу Drive.
     *   share(id, 'user@mail.com', 'writer')
     *   share(id, { anyone: true }, 'reader')
     *   share(id, { domain: 'eurooptica.com' }, 'reader')
     */
    async share(fileId, who, role = 'reader', { notify = false, message } = {}) {
        let perm;
        if (typeof who === 'string') perm = { type: 'user', emailAddress: who, role };
        else if (who?.anyone) perm = { type: 'anyone', role };
        else if (who?.domain) perm = { type: 'domain', domain: who.domain, role };
        else if (who?.group) perm = { type: 'group', emailAddress: who.group, role };
        else throw new GSheetsError('share: укажите e-mail, { anyone: true }, { domain } или { group }', { code: 'INVALID_ARGUMENT' });

        const params = { supportsAllDrives: true, fields: 'id,type,role,emailAddress' };
        if (perm.type === 'user' || perm.type === 'group') {
            params.sendNotificationEmail = notify;
            if (notify && message) params.emailMessage = message;
        }
        if (role === 'owner') params.transferOwnership = true;
        return this.http.request({ url: `${this.DRIVE}/files/${fileId}/permissions${qs(params)}`, method: 'POST', kind: 'none', data: perm });
    }

    async listPermissions(fileId) {
        const res = await this.http.request({
            url: `${this.DRIVE}/files/${fileId}/permissions${qs({ supportsAllDrives: true, fields: 'permissions(id,type,role,emailAddress,domain)' })}`,
            kind: 'none',
        });
        return res.permissions || [];
    }

    async unshare(fileId, permissionId) {
        await this.http.request({ url: `${this.DRIVE}/files/${fileId}/permissions/${permissionId}${qs({ supportsAllDrives: true })}`, method: 'DELETE', kind: 'none' });
        return true;
    }
}

function safeReadJson(file) {
    try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch (_) { return undefined; }
}

function createAuth(opts) {
    const { GoogleAuth, OAuth2Client } = require('google-auth-library');
    const scopes = opts.scopes || SCOPES;

    if (opts.oauth) {
        const { clientId, clientSecret, refreshToken, redirectUri } = opts.oauth;
        if (!clientId || !clientSecret || !refreshToken) {
            throw new GSheetsError('oauth: нужны clientId, clientSecret и refreshToken', { code: 'INVALID_ARGUMENT' });
        }
        const client = new OAuth2Client({ clientId, clientSecret, redirectUri });
        client.setCredentials({ refresh_token: refreshToken });
        return client;
    }

    return new GoogleAuth({
        keyFile: opts.keyFile,
        credentials: opts.credentials,
        scopes,
        clientOptions: opts.subject ? { subject: opts.subject } : undefined,
    });
}

module.exports = { GSheets, SCOPES, SHEET_MIME, FOLDER_MIME, MIME_BY_EXT };
