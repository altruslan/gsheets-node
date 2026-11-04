'use strict';

/**
 * HTTP-слой: авторизованные запросы, повторы с экспоненциальной задержкой,
 * ограничение частоты под квоты Sheets API (60 чтений и 60 записей в минуту на пользователя).
 */

class GSheetsError extends Error {
    constructor(message, { status = 0, code = 'unknown', reason, details, cause } = {}) {
        super(message);
        this.name = 'GSheetsError';
        this.status = status;   // HTTP-статус, 0 — сетевая ошибка
        this.code = code;       // google status: INVALID_ARGUMENT, PERMISSION_DENIED, NOT_FOUND, RESOURCE_EXHAUSTED...
        this.reason = reason;   // Drive reason: storageQuotaExceeded, notFound...
        this.details = details;
        if (cause) this.cause = cause;
    }

    get isNotFound()     { return this.status === 404; }
    get isForbidden()    { return this.status === 403; }
    get isUnauthorized() { return this.status === 401; }
    get isRateLimited()  { return this.status === 429 || this.reason === 'rateLimitExceeded' || this.reason === 'userRateLimitExceeded'; }
    get isQuotaExceeded(){ return this.reason === 'storageQuotaExceeded'; }
    get isRetryable() {
        return this.status === 0 || this.isRateLimited || [408, 500, 502, 503, 504].includes(this.status);
    }
}

/** Скользящее окно: не более `limit` вызовов за `windowMs`. */
class RateLimiter {
    constructor(limit, windowMs = 60000) {
        this.limit = limit;
        this.windowMs = windowMs;
        this.stamps = [];
        this.queue = Promise.resolve();
    }

    take() {
        if (!this.limit) return Promise.resolve();
        const run = async () => {
            for (;;) {
                const now = Date.now();
                while (this.stamps.length && now - this.stamps[0] >= this.windowMs) this.stamps.shift();
                if (this.stamps.length < this.limit) {
                    this.stamps.push(now);
                    return;
                }
                await sleep(this.windowMs - (now - this.stamps[0]) + 5);
            }
        };
        const p = this.queue.then(run);
        this.queue = p.catch(() => {});
        return p;
    }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function toError(res, err) {
    if (!res) {
        return new GSheetsError(`Сетевая ошибка: ${err?.message || err}`, { code: 'NETWORK', cause: err });
    }
    let data = res.data;
    if (Buffer.isBuffer(data) || data instanceof ArrayBuffer) {
        try { data = JSON.parse(Buffer.from(data).toString('utf8')); } catch (_) { data = {}; }
    } else if (typeof data === 'string') {
        try { data = JSON.parse(data); } catch (_) { data = { error: { message: data } }; }
    }
    const e = (data && data.error) || {};
    const reason = Array.isArray(e.errors) && e.errors[0] ? e.errors[0].reason : undefined;
    return new GSheetsError(`HTTP ${res.status}${e.status ? ` ${e.status}` : ''}: ${e.message || 'ошибка запроса'}`, {
        status: res.status,
        code: e.status || `HTTP_${res.status}`,
        reason,
        details: e.details || e.errors,
    });
}

class HttpClient {
    /**
     * @param {object} opts
     * @param {{request:Function}} opts.authClient  google-auth-library клиент (GoogleAuth/JWT/OAuth2Client)
     * @param {number} [opts.retries=6]
     * @param {number} [opts.retryDelay=1000]
     * @param {number} [opts.maxRetryDelay=64000]
     * @param {number} [opts.readsPerMinute=60]     0 — без ограничения
     * @param {number} [opts.writesPerMinute=60]
     * @param {number} [opts.timeout=180000]        Сервер обрывает запросы дольше 180 с
     * @param {object} [opts.logger]
     */
    constructor(opts) {
        this.auth = opts.authClient;
        this.retries = opts.retries ?? 6;
        this.retryDelay = opts.retryDelay ?? 1000;
        this.maxRetryDelay = opts.maxRetryDelay ?? 64000;
        this.timeout = opts.timeout ?? 180000;
        this.logger = opts.logger || console;
        this.readLimiter = new RateLimiter(opts.readsPerMinute ?? 60);
        this.writeLimiter = new RateLimiter(opts.writesPerMinute ?? 60);
        this.requestCount = { read: 0, write: 0 };
    }

    /**
     * @param {object} o
     * @param {string} o.url
     * @param {string} [o.method='GET']
     * @param {object} [o.params]
     * @param {*}      [o.data]
     * @param {object} [o.headers]
     * @param {'json'|'arraybuffer'|'text'} [o.responseType='json']
     * @param {'read'|'write'|'none'} [o.kind]  тип для лимитера; по умолчанию GET — read, остальное — write
     * @param {boolean} [o.full=false]           вернуть весь ответ (headers + data)
     */
    async request(o) {
        const method = (o.method || 'GET').toUpperCase();
        const kind = o.kind || (method === 'GET' ? 'read' : 'write');
        const limiter = kind === 'read' ? this.readLimiter : kind === 'write' ? this.writeLimiter : null;

        let lastErr;
        for (let attempt = 0; attempt <= this.retries; attempt++) {
            if (limiter) await limiter.take();
            if (kind !== 'none') this.requestCount[kind]++;

            let res;
            try {
                res = await this.auth.request({
                    url: o.url,
                    method,
                    params: o.params,
                    data: o.data,
                    headers: o.headers,
                    responseType: o.responseType || 'json',
                    timeout: this.timeout,
                    validateStatus: () => true,
                    retry: false,
                });
            } catch (err) {
                res = err && err.response; // gaxios может бросить даже с validateStatus
                if (!res) {
                    lastErr = toError(null, err);
                }
            }

            if (res && res.status >= 200 && res.status < 300) return o.full ? res : res.data;
            if (res) lastErr = toError(res);

            if (!lastErr.isRetryable || attempt === this.retries) break;

            const retryAfter = Number(res?.headers?.['retry-after'] ?? res?.headers?.get?.('retry-after'));
            const delay = Number.isFinite(retryAfter) && retryAfter > 0
                ? retryAfter * 1000
                : Math.min(this.retryDelay * 2 ** attempt, this.maxRetryDelay) + Math.floor(Math.random() * 1000);
            this.logger.warn?.(`[gsheets] ${lastErr.message} — повтор ${attempt + 1}/${this.retries} через ${Math.round(delay / 1000)} с`);
            await sleep(delay);
        }
        throw lastErr;
    }
}

module.exports = { HttpClient, GSheetsError, RateLimiter, sleep };
