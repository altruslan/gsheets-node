#!/usr/bin/env node
'use strict';

/**
 * Получение refresh token для работы от имени пользователя Google (OAuth 2.0).
 * Нужен, если таблицы должны создаваться в обычном «Моём диске»: у сервисного аккаунта своего места в Drive нет.
 *
 * 1. Google Cloud Console → APIs & Services → Credentials → Create credentials → OAuth client ID → Desktop app.
 * 2. Включите Google Sheets API и Google Drive API.
 * 3. node tools/oauth-token.js <client_id> <client_secret>
 *    (или GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET в окружении)
 * 4. Откройте ссылку, войдите, разрешите доступ — refresh token появится в консоли.
 *    На сервере без браузера: откройте ссылку на своём компьютере и пробросьте порт
 *    ssh -L 53682:localhost:53682 user@server
 */

const http = require('http');
const { OAuth2Client } = require('google-auth-library');
const { SCOPES } = require('../src/client');

const clientId = process.argv[2] || process.env.GOOGLE_CLIENT_ID;
const clientSecret = process.argv[3] || process.env.GOOGLE_CLIENT_SECRET;
const port = Number(process.env.OAUTH_PORT || 53682);

if (!clientId || !clientSecret) {
    console.error('Использование: node tools/oauth-token.js <client_id> <client_secret>');
    process.exit(1);
}

const redirectUri = `http://localhost:${port}/oauth2callback`;
const client = new OAuth2Client({ clientId, clientSecret, redirectUri });
const url = client.generateAuthUrl({ access_type: 'offline', prompt: 'consent', scope: SCOPES });

const server = http.createServer(async (req, res) => {
    const u = new URL(req.url, redirectUri);
    if (u.pathname !== '/oauth2callback') { res.writeHead(404); return res.end(); }
    const code = u.searchParams.get('code');
    if (!code) { res.end(`Ошибка: ${u.searchParams.get('error') || 'нет кода'}`); return; }
    try {
        const { tokens } = await client.getToken(code);
        res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Готово. Можно закрыть вкладку и вернуться в консоль.');
        console.log('\nДобавьте в конфигурацию (не публикуйте этот токен):\n');
        console.log(JSON.stringify({ oauth: { clientId, clientSecret, refreshToken: tokens.refresh_token } }, null, 2));
    } catch (err) {
        res.end(`Ошибка: ${err.message}`);
        console.error(err.message);
    } finally {
        server.close();
    }
});

server.listen(port, () => {
    console.log(`Откройте в браузере:\n\n${url}\n\nОжидаю ответ на ${redirectUri} ...`);
});
