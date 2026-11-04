/// <reference types="node" />
// Типы для gsheets-node

export type ValueInput = 'RAW' | 'USER_ENTERED';
export type ValueRender = 'UNFORMATTED_VALUE' | 'FORMATTED_VALUE' | 'FORMULA';
export type Cell = string | number | boolean;
export type Role = 'reader' | 'commenter' | 'writer' | 'owner';
export type Who = string | { anyone: true } | { domain: string } | { group: string };

export interface Logger { warn?(...a: any[]): void; error?(...a: any[]): void; debug?(...a: any[]): void; }

export interface GSheetsOptions {
  keyFile?: string;
  credentials?: object;
  oauth?: { clientId: string; clientSecret: string; refreshToken: string; redirectUri?: string };
  authClient?: { request(opts: any): Promise<any> };
  subject?: string;
  scopes?: string[];
  valueInput?: ValueInput;
  cell?: { dateFormat?: 'iso' | 'sheet'; arraySeparator?: string };
  retries?: number;
  retryDelay?: number;
  maxRetryDelay?: number;
  readsPerMinute?: number;
  writesPerMinute?: number;
  timeout?: number;
  logger?: Logger;
}

export interface DriveFile { id: string; name: string; mimeType?: string; modifiedTime?: string; webViewLink?: string; parents?: string[]; size?: string; }
export interface SheetSpec { title: string; headers?: string[]; rows?: number; cols?: number; index?: number; formatHeader?: boolean; }
export interface CreateOptions { sheets?: Array<string | SheetSpec>; folderId?: string; locale?: string; timeZone?: string; shareWith?: Array<string | { email: string; role?: Role; notify?: boolean }>; }
export interface ReadOptions { render?: ValueRender; dateTimeRender?: 'SERIAL_NUMBER' | 'FORMATTED_STRING'; majorDimension?: 'ROWS' | 'COLUMNS'; }
export interface WriteOptions { input?: ValueInput; }
export interface ObjectsOptions extends WriteOptions { columns?: string[]; addMissingColumns?: boolean; headerRow?: number; }
export interface UpsertOptions extends ObjectsOptions {
  key: string | string[];
  merge?: boolean;
  insert?: boolean;
  timestampColumn?: string;
  createdColumn?: string;
}
export interface UpsertResult { inserted: number; updated: number; unchanged: number; skipped: number; }
export type Row<T = Record<string, Cell>> = T & { readonly _row: number };

export class GSheetsError extends Error {
  status: number; code: string; reason?: string; details?: any; cause?: Error;
  readonly isNotFound: boolean; readonly isForbidden: boolean; readonly isUnauthorized: boolean;
  readonly isRateLimited: boolean; readonly isQuotaExceeded: boolean; readonly isRetryable: boolean;
}

export class RateLimiter { constructor(limit: number, windowMs?: number); take(): Promise<void>; }

export class Sheet {
  readonly doc: Spreadsheet;
  sheetId: number; title: string; index: number;
  readonly rowCount: number; readonly columnCount: number; readonly frozenRows: number; readonly url: string;

  a1(range?: string): string;
  refresh(): Promise<this>;

  read(range?: string, opts?: ReadOptions): Promise<Cell[][]>;
  getCell(cell: string, opts?: ReadOptions): Promise<Cell>;
  getHeaders(opts?: { headerRow?: number; refresh?: boolean }): Promise<string[]>;
  setHeaders(headers: string[], opts?: { headerRow?: number }): Promise<string[]>;
  readObjects<T = Record<string, Cell>>(opts?: { headerRow?: number; render?: ValueRender }): Promise<Row<T>[]>;
  findObjects<T = Record<string, Cell>>(where: Record<string, any> | ((o: Row<T>) => boolean), opts?: { headerRow?: number }): Promise<Row<T>[]>;

  write(startCell: string, rows: any[][], opts?: WriteOptions): Promise<{ updatedRows: number }>;
  setCell(cell: string, value: any, opts?: WriteOptions): Promise<{ updatedRows: number }>;
  writeMany(map: Record<string, any[][]>, opts?: WriteOptions): Promise<{ updatedCells: number; requests: number }>;
  append(rows: any[][], opts?: WriteOptions): Promise<{ appended: number }>;
  clear(range?: string): Promise<void>;

  appendObjects(items: object[], opts?: ObjectsOptions): Promise<{ appended: number }>;
  replaceObjects(items: object[], opts?: ObjectsOptions & { trim?: boolean }): Promise<{ written: number; columns: number }>;
  upsertObjects(items: object[], opts: UpsertOptions): Promise<UpsertResult>;
  deleteRows(rowNumbers: number[]): Promise<number>;
  deleteObjects(where: Record<string, any> | ((o: Row) => boolean), opts?: { headerRow?: number }): Promise<number>;

  ensureSize(rows: number, cols: number): Promise<void>;
  resize(rows?: number, cols?: number): Promise<void>;
  rename(title: string): Promise<this>;
  freeze(rows?: number, cols?: number): Promise<void>;
  formatHeader(opts?: { headerRow?: number; bold?: boolean; background?: { red: number; green: number; blue: number }; freeze?: boolean; filter?: boolean }): Promise<void>;
  autoResize(startCol?: number, endCol?: number): Promise<void>;
  setNumberFormat(column: string | number, pattern: string, opts?: { type?: 'NUMBER' | 'CURRENCY' | 'PERCENT' | 'DATE' | 'DATE_TIME' | 'TEXT'; headerRow?: number }): Promise<void>;
  duplicate(newTitle: string): Promise<Sheet>;
  copyTo(spreadsheetId: string): Promise<{ sheetId: number; title: string }>;
  delete(): Promise<void>;

  importCsv(csv: string | Buffer, opts?: WriteOptions & { delimiter?: string; startCell?: string; replace?: boolean }): Promise<{ rows: number }>;
  toCsv(opts?: { delimiter?: string; bom?: boolean; render?: ValueRender }): Promise<string>;
  export(format?: 'csv' | 'tsv' | 'pdf' | 'xlsx'): Promise<Buffer>;
}

export class Spreadsheet {
  readonly client: GSheets;
  id: string; title: string; locale: string; timeZone: string; sheets: Sheet[];
  readonly url: string;

  refresh(): Promise<this>;
  sheet(ref?: string | number | { index: number }): Sheet | null;
  getSheet(ref: string | number | { index: number }): Promise<Sheet>;
  addSheet(title: string, opts?: Omit<SheetSpec, 'title'>): Promise<Sheet>;
  ensureSheet(title: string, opts?: Omit<SheetSpec, 'title'>): Promise<Sheet>;
  deleteSheet(ref: Sheet | string | number): Promise<void>;
  reorderSheets(titles: string[]): Promise<void>;
  rename(title: string): Promise<this>;
  batchUpdate(requests: object[]): Promise<any[]>;

  getValues(range: string, opts?: ReadOptions): Promise<Cell[][]>;
  getValuesMany(ranges: string[], opts?: ReadOptions): Promise<Record<string, Cell[][]>>;
  writeValues(data: Array<{ range: string; values: any[][] }>, opts?: WriteOptions): Promise<{ updatedCells: number; requests: number }>;
  appendValues(range: string, values: any[][], opts?: WriteOptions): Promise<void>;
  clearValues(ranges: string[]): Promise<void>;

  share(who: Who, role?: Role, opts?: { notify?: boolean; message?: string }): Promise<any>;
  permissions(): Promise<any[]>;
  copy(name?: string, opts?: { folderId?: string }): Promise<Spreadsheet>;
  moveTo(folderId: string): Promise<any>;
  export(format?: 'xlsx' | 'ods' | 'pdf' | 'csv' | 'tsv' | 'zip', opts?: { gid?: number }): Promise<Buffer>;
  download(filePath: string, format?: string, opts?: { gid?: number }): Promise<string>;
  delete(opts?: { trash?: boolean }): Promise<boolean>;
}

export class GSheets {
  constructor(opts?: GSheetsOptions);
  readonly serviceAccountEmail: string | undefined;
  readonly http: { requestCount: { read: number; write: number } };

  open(idOrUrl: string): Promise<Spreadsheet>;
  create(title: string, opts?: CreateOptions): Promise<Spreadsheet>;
  find(title: string, opts?: { folderId?: string }): Promise<DriveFile[]>;
  openOrCreate(title: string, opts?: CreateOptions): Promise<Spreadsheet>;
  list(opts?: { folderId?: string; nameContains?: string; limit?: number }): Promise<DriveFile[]>;

  listFiles(opts?: { q?: string; limit?: number; fields?: string }): Promise<DriveFile[]>;
  createFolder(name: string, opts?: { parentId?: string }): Promise<DriveFile>;
  ensureFolder(name: string, opts?: { parentId?: string }): Promise<DriveFile>;
  uploadFile(source: string | Buffer, opts?: { name?: string; folderId?: string; mimeType?: string; convert?: boolean; fileId?: string }): Promise<DriveFile>;
  importFile(source: string | Buffer, opts?: { name?: string; folderId?: string; mimeType?: string }): Promise<Spreadsheet>;
  downloadFile(fileId: string): Promise<Buffer>;
  downloadFile(fileId: string, filePath: string): Promise<string>;
  getFile(fileId: string, fields?: string): Promise<DriveFile>;
  moveFile(fileId: string, folderId: string): Promise<any>;
  deleteFile(fileId: string, opts?: { trash?: boolean }): Promise<boolean>;
  share(fileId: string, who: Who, role?: Role, opts?: { notify?: boolean; message?: string }): Promise<any>;
  listPermissions(fileId: string): Promise<any[]>;
  unshare(fileId: string, permissionId: string): Promise<boolean>;
}

export const SCOPES: string[];
export function createClient(opts?: GSheetsOptions): GSheets;
export function toCell(v: any, opts?: { dateFormat?: 'iso' | 'sheet'; arraySeparator?: string }): Cell;
export function parseCsv(text: string, opts?: { delimiter?: string }): string[][];
export function toCsv(rows: any[][], opts?: { delimiter?: string; eol?: string; bom?: boolean }): string;
export const a1: {
  colToLetter(n: number): string;
  letterToCol(s: string): number;
  quoteSheet(title: string): string;
  range(sheet: string, startRow?: number, startCol?: number, endRow?: number, endCol?: number): string;
  parseCell(cell: string): { row: number; col: number };
  parseSpreadsheetUrl(input: string): { id: string; gid?: number };
};
