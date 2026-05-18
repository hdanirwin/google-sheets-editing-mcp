import type { sheets_v4 } from "googleapis";
import { initSheetsClient, getServiceAccountEmail } from "./auth.js";
import { resolveSheetId } from "./ids.js";
import { readCsvFile } from "./csv.js";

type ValueInputOption = "USER_ENTERED" | "RAW";
type ValueRenderOption = "FORMATTED_VALUE" | "UNFORMATTED_VALUE" | "FORMULA";

export function errorStatus(err: unknown): number | null {
  if (!err || typeof err !== "object") return null;
  const e = err as {
    status?: unknown;
    code?: unknown;
    response?: { status?: unknown };
  };
  for (const candidate of [e.status, e.response?.status, e.code]) {
    if (typeof candidate === "number") return candidate;
    if (typeof candidate === "string" && /^\d+$/.test(candidate)) {
      return Number(candidate);
    }
  }
  return null;
}

function wrapError(action: string, err: unknown): Error {
  const msg = err instanceof Error ? err.message : String(err);
  const status = errorStatus(err);
  const email = getServiceAccountEmail();
  const hint =
    email && (status === 403 || status === 404)
      ? ` (Make sure the sheet is shared with ${email} as Editor.)`
      : "";
  return new Error(`${action} failed: ${msg}${hint}`);
}

async function withSheet<T>(
  sheet: string,
  fn: (sheets: sheets_v4.Sheets, spreadsheetId: string) => Promise<T>
): Promise<T> {
  const sheets = await initSheetsClient();
  const spreadsheetId = resolveSheetId(sheet);
  return fn(sheets, spreadsheetId);
}

export interface TabInfo {
  title: string;
  sheetId: number;
  index: number;
  rowCount: number;
  columnCount: number;
}

export interface SheetInfo {
  spreadsheetId: string;
  title: string;
  url: string;
  tabs: TabInfo[];
}

export async function getSheetInfo(sheet: string): Promise<SheetInfo> {
  return withSheet(sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.get({
        spreadsheetId,
        fields:
          "spreadsheetId,properties.title,spreadsheetUrl,sheets.properties(sheetId,title,index,gridProperties.rowCount,gridProperties.columnCount)",
      });
      const data = res.data;
      const tabs: TabInfo[] = (data.sheets ?? []).map((s) => ({
        title: s.properties?.title ?? "",
        sheetId: s.properties?.sheetId ?? 0,
        index: s.properties?.index ?? 0,
        rowCount: s.properties?.gridProperties?.rowCount ?? 0,
        columnCount: s.properties?.gridProperties?.columnCount ?? 0,
      }));
      return {
        spreadsheetId: data.spreadsheetId ?? spreadsheetId,
        title: data.properties?.title ?? "",
        url: data.spreadsheetUrl ?? "",
        tabs,
      };
    } catch (err) {
      throw wrapError("get_sheet_info", err);
    }
  });
}

function rangeFor(tab: string | undefined, range: string | undefined): string {
  if (range && tab) return `${quoteTab(tab)}!${range}`;
  if (range) return range;
  if (tab) return quoteTab(tab);
  throw new Error("Either tab or range (or both) must be provided.");
}

export function quoteTab(tab: string): string {
  return /[^A-Za-z0-9_]/.test(tab) ? `'${tab.replace(/'/g, "''")}'` : tab;
}

export function trailingRanges(
  tab: string,
  rows: number,
  cols: number
): { rows: string; cols: string } {
  const q = quoteTab(tab);
  return {
    rows: `${q}!A${rows + 1}:ZZZ`,
    cols: `${q}!${colLetter(cols + 1)}1:ZZZ${rows}`,
  };
}

export function colLetter(n: number): string {
  let s = "";
  let i = n;
  while (i > 0) {
    const r = (i - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s || "A";
}

export interface ReadRangeResult {
  range: string;
  values: string[][];
}

export async function readRange(opts: {
  sheet: string;
  tab?: string;
  range?: string;
  valueRenderOption?: ValueRenderOption;
}): Promise<ReadRangeResult> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: rangeFor(opts.tab, opts.range),
        valueRenderOption: opts.valueRenderOption ?? "FORMATTED_VALUE",
      });
      return {
        range: res.data.range ?? "",
        values: (res.data.values as string[][]) ?? [],
      };
    } catch (err) {
      throw wrapError("read_range", err);
    }
  });
}

export interface UpdateRangeResult {
  updatedRange: string;
  updatedRows: number;
  updatedColumns: number;
  updatedCells: number;
}

export async function updateRange(opts: {
  sheet: string;
  tab?: string;
  range: string;
  values: (string | number | boolean | null)[][];
  valueInputOption?: ValueInputOption;
}): Promise<UpdateRangeResult> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: rangeFor(opts.tab, opts.range),
        valueInputOption: opts.valueInputOption ?? "USER_ENTERED",
        requestBody: { values: opts.values as unknown[][] },
      });
      return {
        updatedRange: res.data.updatedRange ?? "",
        updatedRows: res.data.updatedRows ?? 0,
        updatedColumns: res.data.updatedColumns ?? 0,
        updatedCells: res.data.updatedCells ?? 0,
      };
    } catch (err) {
      throw wrapError("update_range", err);
    }
  });
}

export interface AppendRowsResult {
  updatedRange: string;
  updatedRows: number;
}

export async function appendRows(opts: {
  sheet: string;
  tab: string;
  values: (string | number | boolean | null)[][];
  valueInputOption?: ValueInputOption;
}): Promise<AppendRowsResult> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: quoteTab(opts.tab),
        valueInputOption: opts.valueInputOption ?? "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: opts.values as unknown[][] },
      });
      return {
        updatedRange: res.data.updates?.updatedRange ?? "",
        updatedRows: res.data.updates?.updatedRows ?? 0,
      };
    } catch (err) {
      throw wrapError("append_rows", err);
    }
  });
}

export interface ClearRangeResult {
  clearedRange: string;
}

export async function clearRange(opts: {
  sheet: string;
  tab?: string;
  range?: string;
}): Promise<ClearRangeResult> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: rangeFor(opts.tab, opts.range),
      });
      return { clearedRange: res.data.clearedRange ?? "" };
    } catch (err) {
      throw wrapError("clear_range", err);
    }
  });
}

async function getTabIdByTitle(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  title: string
): Promise<number> {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties(sheetId,title)",
  });
  const match = (res.data.sheets ?? []).find(
    (s) => s.properties?.title === title
  );
  if (!match || match.properties?.sheetId == null) {
    throw new Error(`Tab "${title}" not found.`);
  }
  return match.properties.sheetId;
}

export async function createTab(opts: {
  sheet: string;
  title: string;
  rowCount?: number;
  columnCount?: number;
}): Promise<TabInfo> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: opts.title,
                  gridProperties: {
                    rowCount: opts.rowCount ?? 1000,
                    columnCount: opts.columnCount ?? 26,
                  },
                },
              },
            },
          ],
        },
      });
      const props = res.data.replies?.[0]?.addSheet?.properties;
      return {
        title: props?.title ?? opts.title,
        sheetId: props?.sheetId ?? 0,
        index: props?.index ?? 0,
        rowCount: props?.gridProperties?.rowCount ?? opts.rowCount ?? 1000,
        columnCount:
          props?.gridProperties?.columnCount ?? opts.columnCount ?? 26,
      };
    } catch (err) {
      throw wrapError("create_tab", err);
    }
  });
}

export async function renameTab(opts: {
  sheet: string;
  oldTitle: string;
  newTitle: string;
}): Promise<void> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const tabId = await getTabIdByTitle(sheets, spreadsheetId, opts.oldTitle);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId: tabId, title: opts.newTitle },
                fields: "title",
              },
            },
          ],
        },
      });
    } catch (err) {
      throw wrapError("rename_tab", err);
    }
  });
}

export async function deleteTab(opts: {
  sheet: string;
  title: string;
}): Promise<void> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const tabId = await getTabIdByTitle(sheets, spreadsheetId, opts.title);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ deleteSheet: { sheetId: tabId } }],
        },
      });
    } catch (err) {
      throw wrapError("delete_tab", err);
    }
  });
}

export interface CsvReplaceResult {
  tab: string;
  rows: number;
  columns: number;
}

async function resolveDefaultTab(sheet: string): Promise<string> {
  const info = await getSheetInfo(sheet);
  const first = info.tabs[0]?.title;
  if (!first) throw new Error("Spreadsheet has no tabs.");
  return first;
}

export async function replaceTabWithCsv(opts: {
  sheet: string;
  csvPath: string;
  tab?: string;
}): Promise<CsvReplaceResult> {
  const rows = await readCsvFile(opts.csvPath);
  if (!rows.length) throw new Error("CSV is empty.");
  const csvRows = rows.length;
  const csvCols = Math.max(...rows.map((r) => r.length), 0);
  const targetTab = opts.tab ?? (await resolveDefaultTab(opts.sheet));

  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      // Write new data first — if a transient failure interrupts the operation,
      // existing data is preserved rather than wiped.
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${quoteTab(targetTab)}!A1`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: rows },
      });
      // Then clear any rows/columns that extended past the new data.
      const trailing = trailingRanges(targetTab, csvRows, csvCols);
      await sheets.spreadsheets.values.batchClear({
        spreadsheetId,
        requestBody: { ranges: [trailing.rows, trailing.cols] },
      });
      return { tab: targetTab, rows: csvRows, columns: csvCols };
    } catch (err) {
      throw wrapError("replace_tab_with_csv", err);
    }
  });
}

export interface CsvAppendResult {
  tab: string;
  appendedRows: number;
}

export async function appendCsv(opts: {
  sheet: string;
  csvPath: string;
  tab?: string;
  includeHeader?: boolean;
}): Promise<CsvAppendResult> {
  const rows = await readCsvFile(opts.csvPath);
  const body = opts.includeHeader ? rows : rows.slice(1);
  if (!body.length) {
    throw new Error(
      opts.includeHeader
        ? "CSV is empty."
        : "Nothing to append (CSV had only a header, and includeHeader=false)."
    );
  }

  const targetTab = opts.tab ?? (await resolveDefaultTab(opts.sheet));
  const result = await appendRows({
    sheet: opts.sheet,
    tab: targetTab,
    values: body,
  });
  return { tab: targetTab, appendedRows: result.updatedRows };
}

export async function batchUpdate(opts: {
  sheet: string;
  requests: sheets_v4.Schema$Request[];
}): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: opts.requests },
      });
      return res.data;
    } catch (err) {
      throw wrapError("batch_update", err);
    }
  });
}
