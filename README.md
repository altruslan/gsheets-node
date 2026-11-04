# gsheets-node

Модуль для работы с Google Таблицами и Google Drive из парсеров на Node.js. Нужен, чтобы сохранять результаты парсинга: создавать и открывать таблицы, управлять листами, записывать строки как объекты, обновлять их по ключу (SKU), импортировать и экспортировать CSV/XLSX, загружать файлы в Drive и открывать к ним доступ.

Зависимость одна — `google-auth-library`. Требуется Node.js 18+.

## Возможности

- **Документы:** открыть по ID или ссылке, создать (в т. ч. в папке), найти по названию, `openOrCreate`, переименовать, скопировать, переместить, удалить, экспортировать (XLSX, ODS, PDF, CSV/TSV конкретного листа).
- **Листы:** добавить, `ensureSheet`, переименовать, дублировать, скопировать в другую таблицу, удалить, упорядочить, изменить размер.
- **Данные:** чтение/запись диапазонов и отдельных ячеек, `append`, очистка, пакетное чтение/запись.
- **Объекты:** `readObjects`, `findObjects`, `appendObjects`, `replaceObjects`, `upsertObjects` (по ключу, в т. ч. составному), `deleteObjects`.
- **Оформление:** заголовок (жирный, фон, закрепление, фильтр), автоширина столбцов, числовые форматы.
- **Drive:** загрузка файлов (multipart до 5 МБ, возобновляемая — для больших), импорт CSV/XLSX как таблицы, скачивание, папки, доступ (пользователь, домен, группа, по ссылке).
- **Надёжность:** повторы с экспоненциальной задержкой при 429/5xx и сетевых сбоях (учитывается `Retry-After`), ограничитель частоты под квоты Sheets API, автоматическое расширение листа перед записью, разбиение больших записей на запросы до ~1,5 МБ.

## Установка

```bash
npm install github:altruslan/gsheets-node
```

## Авторизация

### Сервисный аккаунт (рекомендуется для серверных парсеров)

1. [Google Cloud Console](https://console.cloud.google.com) → создать проект → включить **Google Sheets API** и **Google Drive API**.
2. IAM & Admin → Service Accounts → создать → Keys → Add key → JSON. Положить ключ, например, в `config/google-service-account.json` (каталог `config/` в `.gitignore`).
3. Открыть доступ к таблице (или папке) для `client_email` из ключа с ролью **Редактор**.

```js
const { GSheets } = require('gsheets-node');
const gs = new GSheets({ keyFile: 'config/google-service-account.json' });
console.log(gs.serviceAccountEmail); // с этим адресом нужно поделиться таблицей
```

> **Важно:** у сервисного аккаунта нет своего места в Google Drive, поэтому `gs.create()` без папки завершится ошибкой `storageQuotaExceeded`. Варианты: создать таблицу вручную и поделиться ею с сервисным аккаунтом; создавать в папке на **общем диске** (`folderId`, Google Workspace); или использовать OAuth от имени пользователя.

### OAuth от имени пользователя

Подходит, если таблицы должны создаваться в обычном «Моём диске».

1. Cloud Console → Credentials → Create credentials → OAuth client ID → **Desktop app**.
2. Получить refresh token: `npm run oauth-token -- <client_id> <client_secret>` и пройти по ссылке.

```js
const gs = new GSheets({ oauth: { clientId, clientSecret, refreshToken } });
```

Другие варианты: `credentials` (объект ключа), `subject` (domain-wide delegation в Workspace), `authClient` (готовый клиент google-auth-library), без параметров — `GOOGLE_APPLICATION_CREDENTIALS`.

## Быстрый старт

```js
const { GSheets } = require('gsheets-node');
const gs = new GSheets({ keyFile: 'config/google-service-account.json' });

const doc = await gs.open('https://docs.google.com/spreadsheets/d/1AbC.../edit');
const sheet = await doc.ensureSheet('Maui Jim', {
  headers: ['sku', 'model', 'price', 'quantity', 'created_at', 'updated_at'],
});

const stats = await sheet.upsertObjects(items, {
  key: 'sku',
  timestampColumn: 'updated_at',
  createdColumn: 'created_at',
});
// { inserted: 12, updated: 3, unchanged: 240, skipped: 0 }

const rows = await sheet.readObjects();         // [{ sku, model, price, ... }], у каждой строки есть _row
const cheap = await sheet.findObjects((r) => r.price < 100);
```

Полный пример для парсера — [`examples/parser-results.js`](examples/parser-results.js): лист на бренд, обновление по SKU, пометка пропавших товаров, журнал прогонов.

## Работа с документами

```js
const doc = await gs.create('Остатки', {
  sheets: [{ title: 'Товары', headers: ['sku', 'qty'] }, 'Журнал'],
  folderId: '1FolderId...',          // обязательно для сервисного аккаунта (общий диск)
  locale: 'uk_UA', timeZone: 'Europe/Kyiv',
  shareWith: ['manager@eurooptica.com'],
});

const same = await gs.openOrCreate('Остатки', { folderId: '1FolderId...' });
const files = await gs.list({ nameContains: 'Остатки' });

await doc.rename('Остатки 2026');
const copy = await doc.copy('Остатки — архив');
await doc.share('user@mail.com', 'reader');
await doc.share({ anyone: true }, 'reader');
await doc.download('ostatki.xlsx');            // формат по расширению
const pdf = await doc.export('pdf');           // Buffer
await copy.delete();                           // в корзину; { trash: false } — навсегда
```

## Работа с листами

```js
const s = await doc.addSheet('Новый', { rows: 100, cols: 10, headers: ['a', 'b'] });
await s.rename('Товары');
await s.duplicate('Товары — копия');
await doc.reorderSheets(['Журнал', 'Товары']);
await doc.deleteSheet('Товары — копия');

doc.sheet('Товары');      // по названию (из кэша)
doc.sheet(123456);        // по gid
doc.sheet({ index: 0 });  // по позиции
await doc.getSheet('Товары'); // с обновлением метаданных, NOT_FOUND если нет
```

## Данные

```js
await s.write('A1', [['sku', 'qty'], ['A1', 5]]);   // лист расширяется сам
await s.setCell('C2', 'проверено');
await s.writeMany({ 'E1': [['итого']], 'E2': [['=SUM(B:B)']] }, { input: 'USER_ENTERED' });
await s.append([['B2', 7]]);
const values = await s.read('A1:C10');
const one = await s.getCell('B2');
await s.clear('A2:C');      // без аргумента — весь лист
```

### Объекты

| Метод | Что делает |
|---|---|
| `readObjects()` | Строки как объекты по заголовкам; пустые строки пропускаются; `_row` — номер строки |
| `findObjects(where)` | Фильтр: объект `{ brand: 'X' }` или функция |
| `appendObjects(items)` | Добавить в конец; недостающие столбцы создаются (`addMissingColumns: false` — игнорировать) |
| `replaceObjects(items)` | Очистить лист и записать заново; лишние строки/столбцы обрезаются |
| `upsertObjects(items, { key })` | Обновить по ключу, добавить новые |
| `deleteObjects(where)` | Удалить подходящие строки |

Параметры `upsertObjects`:

- `key` — столбец или массив столбцов (`['brand', 'sku']`).
- `merge` (по умолчанию `true`) — меняются только переданные поля; остальные ячейки строки, включая столбцы, добавленные вручную, сохраняются. `false` — строка перезаписывается целиком.
- `insert: false` — только обновлять существующие строки.
- `timestampColumn` / `createdColumn` — время изменения / добавления. Время изменения ставится, только если данные действительно изменились.
- Если ключ встречается во входных данных несколько раз, побеждает последний. Если на листе есть дубликаты ключа, обновляется первая строка (с предупреждением в логе).

За один вызов выполняется одно чтение листа и одна запись (или несколько, если данных больше ~1,5 МБ). Подряд идущие изменённые строки записываются одним диапазоном.

### Преобразование значений

| JS | Ячейка |
|---|---|
| `null`, `undefined`, `NaN` | пусто |
| `Date` | `2026-11-04 10:20:30` (UTC); `cell: { dateFormat: 'iso' }` — ISO 8601 |
| массив примитивов | `a, b, c` (`cell: { arraySeparator }`) |
| объект / вложенный массив | JSON |
| строка > 50 000 символов | обрезается (лимит Google) |

`valueInput: 'RAW'` (по умолчанию) записывает значения как есть: `00123` остаётся строкой, `=SUM()` — текстом. `USER_ENTERED` разбирает их как ввод в интерфейсе (формулы, числа, даты). Можно указать глобально в `new GSheets({ valueInput })` или для одного вызова `{ input: 'USER_ENTERED' }`.

При чтении по умолчанию используется `UNFORMATTED_VALUE` (числа — числами). `{ render: 'FORMATTED_VALUE' }` — как отображается в таблице.

## Оформление

```js
await s.formatHeader({ filter: true });          // жирный, серый фон, закрепление, фильтр
await s.freeze(1, 1);
await s.autoResize();
await s.setNumberFormat('price', '#,##0.00');
await s.setNumberFormat('discount', '0%', { type: 'PERCENT' });
await doc.batchUpdate([ /* любые запросы spreadsheets.batchUpdate */ ]);
```

## Импорт и экспорт

```js
await s.importCsv(fs.readFileSync('stock.csv'), { delimiter: ';' }); // заменяет содержимое листа
const csv = await s.toCsv();                    // собирается локально из данных листа
const buf = await s.export('csv');              // экспорт Google по gid листа

const imported = await gs.importFile('price.xlsx', { name: 'Прайс', folderId });  // XLSX → Google Таблица
```

## Drive

```js
const folder = await gs.ensureFolder('Выгрузки', { parentId: '1Root...' });
const file = await gs.uploadFile('out/lindberg.zip', { folderId: folder.id });
await gs.uploadFile(buffer, { name: 'report.pdf', fileId: file.id });   // заменить содержимое
await gs.downloadFile(file.id, 'local.zip');
await gs.moveFile(file.id, otherFolderId);
await gs.share(file.id, { domain: 'eurooptica.com' }, 'reader');
const perms = await gs.listPermissions(file.id);
```

## Ошибки и квоты

Все методы бросают `GSheetsError`:

```js
try { await gs.open(id); } catch (e) {
  e.status;  // HTTP (0 — сеть)
  e.code;    // NOT_FOUND, PERMISSION_DENIED, INVALID_ARGUMENT, RESOURCE_EXHAUSTED...
  e.reason;  // storageQuotaExceeded, notFound... (Drive)
  e.isNotFound; e.isForbidden; e.isRateLimited; e.isQuotaExceeded; e.isRetryable;
}
```

Квоты [Sheets API](https://developers.google.com/workspace/sheets/api/limits): 300 запросов на чтение и 300 на запись в минуту на проект, по 60 в минуту на пользователя; пакетный запрос считается одним. Поэтому по умолчанию включены ограничители `readsPerMinute: 60` и `writesPerMinute: 60` (0 — отключить). Если несколько парсеров работают под одним сервисным аккаунтом одновременно, уменьшите эти значения или запускайте их по очереди. Лимит таблицы — 10 млн ячеек: для больших каталогов лучше отдельная таблица на поставщика.

Счётчик запросов за время работы клиента: `gs.http.requestCount` → `{ read, write }`.

## Тесты

```bash
npm test
```

Тесты используют имитацию Sheets/Drive API в памяти (`test/fake-google.js`) и не требуют ключей.
