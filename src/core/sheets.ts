import type { sheets_v4 } from "googleapis";
import { initSheetsClient, getServiceAccountEmail } from "./auth.js";
import { resolveSheetId } from "./ids.js";
import { readCsvFile } from "./csv.js";
import {
  classifyColumns,
  classifySequence,
  colLetter,
  colNumber,
  shiftColumnsInA1,
  shiftRowsInA1,
  type ColumnCheck,
  type ColumnCheckStatus,
  type SequenceCheck,
} from "./columns.js";

// Re-exported so existing importers (and tests) keep finding colLetter here.
export { colLetter } from "./columns.js";

type ValueInputOption = "USER_ENTERED" | "RAW";
type ValueRenderOption = "FORMATTED_VALUE" | "UNFORMATTED_VALUE" | "FORMULA";

/** Optional column-drift guard shared by the data tools. */
export interface ColumnGuardOptions {
  /**
   * Header values you expect in the header row, in order, beginning at
   * `headerStartColumn`. When provided, the tool reads the actual header row
   * before acting and compares. Omit to skip the check entirely.
   */
  expectedHeaders?: string[];
  /** 1-based row holding the headers. Default: 1. */
  headerRow?: number;
  /** Column letter where `expectedHeaders` begins. Default: "A". */
  headerStartColumn?: string;
}

/** What the column guard found, attached to a tool's result so the caller sees it. */
export interface ColumnCheckSummary {
  status: ColumnCheckStatus;
  offset: number;
  expectedHeaders: string[];
  actualHeaders: string[];
  adjustedRange?: string;
  note: string;
}

/**
 * Optional row-drift guard. When set, the tool reads the key column before
 * acting and confirms the target rows still hold the expected keys — catching
 * rows deleted or reordered since the caller last looked. Only the range-based
 * tools support it, since the target rows are anchored by the range's start row.
 */
export interface RowGuardOptions {
  /**
   * Key values you expect at the target rows, in order, starting at the range's
   * first row. The tool reads `keyColumn` and locates this block. Omit to skip.
   */
  expectedKeys?: string[];
  /** Column letter holding the row identity (id/email/name). Default: "A". */
  keyColumn?: string;
}

/** What the row guard found, attached to a tool's result. */
export interface RowCheckSummary {
  status: ColumnCheckStatus;
  offset: number;
  expectedKeys: string[];
  adjustedRange?: string;
  note: string;
}

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

/** Pull the tab name out of an A1 range that carries a "Tab!..." prefix. */
function tabFromRange(range: string | undefined): string | undefined {
  if (!range) return undefined;
  const bang = range.lastIndexOf("!");
  if (bang < 0) return undefined;
  let t = range.slice(0, bang);
  if (t.startsWith("'") && t.endsWith("'")) {
    t = t.slice(1, -1).replace(/''/g, "'");
  }
  return t || undefined;
}

function columnMismatchMessage(
  check: ColumnCheck,
  tab: string,
  headerRow: number
): string {
  return [
    `Column check stopped the operation: the header row on "${tab}" (row ${headerRow}) no longer matches the expected columns, and the change is more than a simple shift.`,
    `Expected (starting at column ${colLetter(check.expectedStartColumn)}): ${JSON.stringify(check.expectedHeaders)}`,
    `Found in that row: ${JSON.stringify(check.actualHeaders)}`,
    check.reason ? `Reason: ${check.reason}` : "",
    `Re-read the sheet (get_sheet_info / read_range) and re-plan before writing.`,
  ]
    .filter(Boolean)
    .join("\n");
}

/**
 * Run the column guard for a data tool. Returns null when no expectedHeaders
 * were supplied (check skipped). Throws on a mismatch. Otherwise returns the
 * classification (exact or offset) so the caller can adjust its target.
 */
async function runColumnGuard(
  sheet: string,
  tab: string | undefined,
  range: string | undefined,
  guard: ColumnGuardOptions
): Promise<ColumnCheck | null> {
  if (!guard.expectedHeaders || guard.expectedHeaders.length === 0) return null;
  const headerRow = guard.headerRow ?? 1;
  const startCol = guard.headerStartColumn ? colNumber(guard.headerStartColumn) : 1;
  const targetTab =
    tab ?? tabFromRange(range) ?? (await resolveDefaultTab(sheet));
  const res = await readRange({
    sheet,
    range: `${quoteTab(targetTab)}!${headerRow}:${headerRow}`,
    valueRenderOption: "FORMATTED_VALUE",
  });
  const actualRow = res.values[0] ?? [];
  const check = classifyColumns(guard.expectedHeaders, actualRow, startCol);
  if (check.status === "mismatch") {
    throw new Error(columnMismatchMessage(check, targetTab, headerRow));
  }
  return check;
}

function summarizeCheck(
  check: ColumnCheck | null,
  adjustedRange?: string
): ColumnCheckSummary | undefined {
  if (!check) return undefined;
  if (check.status === "exact") {
    return {
      status: "exact",
      offset: 0,
      expectedHeaders: check.expectedHeaders,
      actualHeaders: check.actualHeaders,
      note: "Header row matches the expected columns.",
    };
  }
  const dir = check.offset > 0 ? "right" : "left";
  return {
    status: "offset",
    offset: check.offset,
    expectedHeaders: check.expectedHeaders,
    actualHeaders: check.actualHeaders,
    adjustedRange,
    note:
      `Header row is shifted ${Math.abs(check.offset)} column(s) ${dir} from expected; ` +
      (adjustedRange
        ? `target auto-adjusted to ${adjustedRange}.`
        : `target auto-adjusted to compensate.`),
  };
}

/** First row number referenced by an A1 range, or null when it has none (e.g. "A:C"). */
function firstRowOf(range: string | undefined): number | null {
  if (!range) return null;
  const bang = range.lastIndexOf("!");
  const body = bang >= 0 ? range.slice(bang + 1) : range;
  const m = body.split(":")[0].match(/\d+/);
  return m ? parseInt(m[0], 10) : null;
}

function rowMismatchMessage(
  check: SequenceCheck,
  tab: string,
  keyColumn: string
): string {
  const present = new Set(check.actual.map((v) => v.trim()));
  const missing = check.expected.filter((k) => !present.has(k.trim()));
  return [
    `Row check stopped the operation: the rows on "${tab}" (keyed by column ${keyColumn}) no longer line up with the expected keys, and the change is more than a simple shift.`,
    `Expected keys (starting at row ${check.expectedStart}): ${JSON.stringify(check.expected)}`,
    missing.length
      ? `Keys not found in column ${keyColumn}: ${JSON.stringify(missing)} — rows may have been deleted.`
      : `The keys are all present but not as one contiguous block — rows may have been reordered.`,
    `Re-read column ${keyColumn} and re-plan before writing.`,
  ]
    .filter(Boolean)
    .join("\n");
}

/**
 * Run the row guard. Returns null when no expectedKeys were supplied. Throws on
 * a mismatch (deleted/reordered rows). Otherwise returns the classification so
 * the caller can shift its target rows by the offset. `colOffset` is the column
 * shift already detected, so the key column is read from its current position.
 */
async function runRowGuard(
  sheet: string,
  tab: string | undefined,
  range: string | undefined,
  colOffset: number,
  guard: RowGuardOptions
): Promise<SequenceCheck | null> {
  if (!guard.expectedKeys || guard.expectedKeys.length === 0) return null;
  const startRow = firstRowOf(range);
  if (startRow == null) {
    throw new Error(
      "expectedKeys requires a range with a specific start row (e.g. 'B5:D20') so the rows can be anchored."
    );
  }
  let keyCol = guard.keyColumn ?? "A";
  if (colOffset !== 0) {
    const shifted = colNumber(keyCol) + colOffset;
    if (shifted < 1) {
      throw new Error(
        `A column shift moves the key column ${keyCol} before column A; re-read the sheet and re-plan.`
      );
    }
    keyCol = colLetter(shifted);
  }
  const targetTab =
    tab ?? tabFromRange(range) ?? (await resolveDefaultTab(sheet));
  const res = await readRange({
    sheet,
    range: `${quoteTab(targetTab)}!${keyCol}:${keyCol}`,
    valueRenderOption: "FORMATTED_VALUE",
  });
  const actualColumn = res.values.map((r) => r[0] ?? "");
  const check = classifySequence(guard.expectedKeys, actualColumn, startRow);
  if (check.status === "mismatch") {
    throw new Error(rowMismatchMessage(check, targetTab, keyCol));
  }
  return check;
}

function summarizeRowCheck(
  check: SequenceCheck | null,
  adjustedRange?: string
): RowCheckSummary | undefined {
  if (!check) return undefined;
  if (check.status === "exact") {
    return {
      status: "exact",
      offset: 0,
      expectedKeys: check.expected,
      note: "Target rows match the expected keys.",
    };
  }
  const dir = check.offset > 0 ? "down" : "up";
  return {
    status: "offset",
    offset: check.offset,
    expectedKeys: check.expected,
    adjustedRange,
    note:
      `Target rows are shifted ${Math.abs(check.offset)} row(s) ${dir} from expected ` +
      `(rows likely ${check.offset > 0 ? "added" : "removed"} above); ` +
      (adjustedRange
        ? `target auto-adjusted to ${adjustedRange}.`
        : `target auto-adjusted to compensate.`),
  };
}

/**
 * Run both guards for a range-based tool and return the effective range with any
 * column and row offsets applied. Throws on a mismatch in either dimension.
 */
async function resolveTarget(
  opts: {
    sheet: string;
    tab?: string;
    range?: string;
  } & ColumnGuardOptions &
    RowGuardOptions
): Promise<{
  colCheck: ColumnCheck | null;
  rowCheck: SequenceCheck | null;
  range: string | undefined;
}> {
  const colCheck = await runColumnGuard(opts.sheet, opts.tab, opts.range, opts);
  const rowCheck = await runRowGuard(
    opts.sheet,
    opts.tab,
    opts.range,
    colCheck?.offset ?? 0,
    opts
  );
  let range = opts.range;
  if (range && colCheck && colCheck.offset !== 0) {
    range = shiftColumnsInA1(range, colCheck.offset);
  }
  if (range && rowCheck && rowCheck.offset !== 0) {
    range = shiftRowsInA1(range, rowCheck.offset);
  }
  return { colCheck, rowCheck, range };
}

export interface ReadRangeResult {
  range: string;
  values: string[][];
  columnCheck?: ColumnCheckSummary;
  rowCheck?: RowCheckSummary;
}

export async function readRange(
  opts: {
    sheet: string;
    tab?: string;
    range?: string;
    valueRenderOption?: ValueRenderOption;
  } & ColumnGuardOptions &
    RowGuardOptions
): Promise<ReadRangeResult> {
  // resolveTarget short-circuits when neither expectedHeaders nor expectedKeys
  // is set, so the guards' own internal reads don't recurse back through here.
  const { colCheck, rowCheck, range } = await resolveTarget(opts);
  const adjusted = range !== opts.range ? range : undefined;
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: rangeFor(opts.tab, range),
        valueRenderOption: opts.valueRenderOption ?? "FORMATTED_VALUE",
      });
      return {
        range: res.data.range ?? "",
        values: (res.data.values as string[][]) ?? [],
        columnCheck: summarizeCheck(colCheck, colCheck?.offset ? adjusted : undefined),
        rowCheck: summarizeRowCheck(rowCheck, rowCheck?.offset ? adjusted : undefined),
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
  columnCheck?: ColumnCheckSummary;
  rowCheck?: RowCheckSummary;
}

export async function updateRange(
  opts: {
    sheet: string;
    tab?: string;
    range: string;
    values: (string | number | boolean | null)[][];
    valueInputOption?: ValueInputOption;
  } & ColumnGuardOptions &
    RowGuardOptions
): Promise<UpdateRangeResult> {
  const { colCheck, rowCheck, range } = await resolveTarget(opts);
  const adjusted = range !== opts.range ? range : undefined;
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: rangeFor(opts.tab, range),
        valueInputOption: opts.valueInputOption ?? "USER_ENTERED",
        requestBody: { values: opts.values as unknown[][] },
      });
      return {
        updatedRange: res.data.updatedRange ?? "",
        updatedRows: res.data.updatedRows ?? 0,
        updatedColumns: res.data.updatedColumns ?? 0,
        updatedCells: res.data.updatedCells ?? 0,
        columnCheck: summarizeCheck(colCheck, colCheck?.offset ? adjusted : undefined),
        rowCheck: summarizeRowCheck(rowCheck, rowCheck?.offset ? adjusted : undefined),
      };
    } catch (err) {
      throw wrapError("update_range", err);
    }
  });
}

export interface AppendRowsResult {
  updatedRange: string;
  updatedRows: number;
  columnCheck?: ColumnCheckSummary;
}

export async function appendRows(
  opts: {
    sheet: string;
    tab: string;
    values: (string | number | boolean | null)[][];
    valueInputOption?: ValueInputOption;
  } & ColumnGuardOptions
): Promise<AppendRowsResult> {
  const check = await runColumnGuard(opts.sheet, opts.tab, undefined, opts);
  // Append always starts at column A, so a rightward shift is compensated by
  // padding each row with leading blanks. A leftward shift can't be expressed
  // without dropping data, so stop and let the caller re-plan.
  let values = opts.values;
  if (check && check.offset > 0) {
    const pad = new Array(check.offset).fill(null);
    values = opts.values.map((row) => [...pad, ...row]);
  } else if (check && check.offset < 0) {
    throw new Error(
      `Column check stopped append: the header row is shifted ${Math.abs(check.offset)} column(s) left of expected, which append cannot compensate for without dropping data. Re-read the sheet and re-plan.`
    );
  }
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: quoteTab(opts.tab),
        valueInputOption: opts.valueInputOption ?? "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: values as unknown[][] },
      });
      return {
        updatedRange: res.data.updates?.updatedRange ?? "",
        updatedRows: res.data.updates?.updatedRows ?? 0,
        columnCheck: summarizeCheck(check),
      };
    } catch (err) {
      throw wrapError("append_rows", err);
    }
  });
}

export interface ClearRangeResult {
  clearedRange: string;
  columnCheck?: ColumnCheckSummary;
  rowCheck?: RowCheckSummary;
}

export async function clearRange(
  opts: {
    sheet: string;
    tab?: string;
    range?: string;
  } & ColumnGuardOptions &
    RowGuardOptions
): Promise<ClearRangeResult> {
  const { colCheck, rowCheck, range } = await resolveTarget(opts);
  const adjusted = range !== opts.range ? range : undefined;
  return withSheet(opts.sheet, async (sheets, spreadsheetId) => {
    try {
      const res = await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: rangeFor(opts.tab, range),
      });
      return {
        clearedRange: res.data.clearedRange ?? "",
        columnCheck: summarizeCheck(colCheck, colCheck?.offset ? adjusted : undefined),
        rowCheck: summarizeRowCheck(rowCheck, rowCheck?.offset ? adjusted : undefined),
      };
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
  columnCheck?: ColumnCheckSummary;
}

async function resolveDefaultTab(sheet: string): Promise<string> {
  const info = await getSheetInfo(sheet);
  const first = info.tabs[0]?.title;
  if (!first) throw new Error("Spreadsheet has no tabs.");
  return first;
}

export async function replaceTabWithCsv(
  opts: {
    sheet: string;
    csvPath: string;
    tab?: string;
  } & ColumnGuardOptions
): Promise<CsvReplaceResult> {
  const rows = await readCsvFile(opts.csvPath);
  if (!rows.length) throw new Error("CSV is empty.");
  const csvRows = rows.length;
  const csvCols = Math.max(...rows.map((r) => r.length), 0);
  const targetTab = opts.tab ?? (await resolveDefaultTab(opts.sheet));

  // Verify we're overwriting the tab/layout we think we are. A full-tab rewrite
  // starts at A1, so an offset can't be auto-adjusted — it's reported, not applied.
  const check = await runColumnGuard(opts.sheet, targetTab, undefined, opts);

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
      return {
        tab: targetTab,
        rows: csvRows,
        columns: csvCols,
        columnCheck: summarizeCheck(check),
      };
    } catch (err) {
      throw wrapError("replace_tab_with_csv", err);
    }
  });
}

export interface CsvAppendResult {
  tab: string;
  appendedRows: number;
  columnCheck?: ColumnCheckSummary;
}

export async function appendCsv(
  opts: {
    sheet: string;
    csvPath: string;
    tab?: string;
    includeHeader?: boolean;
  } & ColumnGuardOptions
): Promise<CsvAppendResult> {
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
  // The guard runs inside appendRows (which also applies any offset padding).
  const result = await appendRows({
    sheet: opts.sheet,
    tab: targetTab,
    values: body,
    expectedHeaders: opts.expectedHeaders,
    headerRow: opts.headerRow,
    headerStartColumn: opts.headerStartColumn,
  });
  return {
    tab: targetTab,
    appendedRows: result.updatedRows,
    columnCheck: result.columnCheck,
  };
}

export async function batchUpdate(
  opts: {
    sheet: string;
    requests: sheets_v4.Schema$Request[];
    /** Tab whose header row to confirm when expectedHeaders is set. */
    tab?: string;
  } & ColumnGuardOptions
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  // Raw requests reference columns by index, so an offset can't be safely
  // auto-applied here — verify only, and stop on any drift (incl. a shift).
  const check = await runColumnGuard(opts.sheet, opts.tab, undefined, opts);
  if (check && check.offset !== 0) {
    throw new Error(
      `Column check stopped batch_update: the header row is shifted ${Math.abs(check.offset)} column(s) from expected. Raw batch requests target columns by index and can't be auto-adjusted — re-read the sheet and rebuild the requests for the current layout.`
    );
  }
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
