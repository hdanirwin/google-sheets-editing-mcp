// Unit tests for pure helpers. No network access, no env vars required.
// Run via `npm test` (which builds first) or directly:
//   node --test tests/unit.test.mjs   # if dist/ is already built
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  colLetter,
  errorStatus,
  quoteTab,
  trailingRanges,
} from "../dist/core/sheets.js";
import { resolveSheetId } from "../dist/core/ids.js";
import { parseCsv } from "../dist/core/csv.js";
import {
  classifyColumns,
  classifySequence,
  colNumber,
  shiftColumnsInA1,
  shiftRowsInA1,
} from "../dist/core/columns.js";

test("colLetter: single-letter columns", () => {
  assert.equal(colLetter(1), "A");
  assert.equal(colLetter(2), "B");
  assert.equal(colLetter(26), "Z");
});

test("colLetter: double-letter columns", () => {
  assert.equal(colLetter(27), "AA");
  assert.equal(colLetter(28), "AB");
  assert.equal(colLetter(52), "AZ");
  assert.equal(colLetter(702), "ZZ");
});

test("colLetter: triple-letter columns", () => {
  assert.equal(colLetter(703), "AAA");
  // ZZZ = 26^3 + 26^2 + 26 = 18278, the Sheets per-tab column limit.
  assert.equal(colLetter(18278), "ZZZ");
});

test("colLetter: zero/negative falls back to A", () => {
  assert.equal(colLetter(0), "A");
  assert.equal(colLetter(-5), "A");
});

test("resolveSheetId: extracts ID from full URL", () => {
  const id = "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_-AB";
  assert.equal(
    resolveSheetId(`https://docs.google.com/spreadsheets/d/${id}/edit#gid=0`),
    id
  );
  assert.equal(
    resolveSheetId(`https://docs.google.com/spreadsheets/d/${id}/`),
    id
  );
});

test("resolveSheetId: passes through bare IDs", () => {
  const id = "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_-AB";
  assert.equal(resolveSheetId(id), id);
});

test("resolveSheetId: returns input on no match", () => {
  // Defensive — invalid input is returned as-is for the API to reject.
  assert.equal(resolveSheetId("not-a-url-or-id"), "not-a-url-or-id");
});

test("errorStatus: reads top-level status", () => {
  assert.equal(errorStatus({ status: 404 }), 404);
  assert.equal(errorStatus({ status: 503 }), 503);
});

test("errorStatus: reads nested response.status", () => {
  assert.equal(errorStatus({ response: { status: 403 } }), 403);
});

test("errorStatus: parses numeric string codes", () => {
  assert.equal(errorStatus({ code: "429" }), 429);
});

test("errorStatus: prefers status over code", () => {
  assert.equal(errorStatus({ status: 500, code: "200" }), 500);
});

test("errorStatus: ignores non-numeric values", () => {
  assert.equal(errorStatus({ code: "ECONNRESET" }), null);
  assert.equal(errorStatus({ status: "broken" }), null);
});

test("errorStatus: handles null/non-object", () => {
  assert.equal(errorStatus(null), null);
  assert.equal(errorStatus(undefined), null);
  assert.equal(errorStatus("just a string"), null);
  assert.equal(errorStatus(42), null);
});

test("parseCsv: parses basic rows", () => {
  assert.deepStrictEqual(
    parseCsv("a,b,c\n1,2,3\n4,5,6"),
    [
      ["a", "b", "c"],
      ["1", "2", "3"],
      ["4", "5", "6"],
    ]
  );
});

test("parseCsv: drops trailing blank line", () => {
  // The most common CSV-on-disk shape: ends with a newline. We must NOT emit
  // an empty trailing row, because that gets written into the sheet.
  assert.deepStrictEqual(
    parseCsv("a,b\n1,2\n"),
    [
      ["a", "b"],
      ["1", "2"],
    ]
  );
});

test("parseCsv: drops whitespace-only rows", () => {
  assert.deepStrictEqual(
    parseCsv("a,b\n\n   \n1,2\n"),
    [
      ["a", "b"],
      ["1", "2"],
    ]
  );
});

test("parseCsv: handles quoted fields with commas", () => {
  assert.deepStrictEqual(
    parseCsv('a,b\n"hello, world",2'),
    [
      ["a", "b"],
      ["hello, world", "2"],
    ]
  );
});

test("parseCsv: leaves numbers as strings (no dynamic typing)", () => {
  const rows = parseCsv("n,m\n42,3.14");
  assert.equal(typeof rows[1][0], "string");
  assert.equal(rows[1][0], "42");
  assert.equal(typeof rows[1][1], "string");
  assert.equal(rows[1][1], "3.14");
});

test("quoteTab: leaves simple identifiers bare", () => {
  assert.equal(quoteTab("Sheet1"), "Sheet1");
  assert.equal(quoteTab("my_tab_2"), "my_tab_2");
});

test("quoteTab: wraps and escapes tabs with special characters", () => {
  assert.equal(quoteTab("My Tab"), "'My Tab'");
  assert.equal(quoteTab("price ($)"), "'price ($)'");
  // Embedded single quotes must be doubled per Sheets A1 quoting rules.
  assert.equal(quoteTab("a'b"), "'a''b'");
});

test("trailingRanges: produces row-tail and column-tail for replace_tab_with_csv", () => {
  // 3 rows × 2 cols of CSV → clear A4:ZZZ (rows below) and C1:ZZZ3 (cols right).
  assert.deepStrictEqual(trailingRanges("Sheet1", 3, 2), {
    rows: "Sheet1!A4:ZZZ",
    cols: "Sheet1!C1:ZZZ3",
  });
});

test("trailingRanges: quotes tabs with spaces", () => {
  assert.deepStrictEqual(trailingRanges("My Tab", 1, 1), {
    rows: "'My Tab'!A2:ZZZ",
    cols: "'My Tab'!B1:ZZZ1",
  });
});

test("colNumber: inverts colLetter", () => {
  assert.equal(colNumber("A"), 1);
  assert.equal(colNumber("Z"), 26);
  assert.equal(colNumber("AA"), 27);
  assert.equal(colNumber("ZZ"), 702);
  assert.equal(colNumber("aa"), 27); // case-insensitive
});

test("colNumber: non-letters return 0", () => {
  assert.equal(colNumber(""), 0);
  assert.equal(colNumber("A1"), 0);
  assert.equal(colNumber("1"), 0);
});

test("classifyColumns: exact match at expected start", () => {
  const r = classifyColumns(["Name", "Email", "Status"], ["Name", "Email", "Status"]);
  assert.equal(r.status, "exact");
  assert.equal(r.offset, 0);
});

test("classifyColumns: trims whitespace before comparing", () => {
  const r = classifyColumns(["Name", "Status"], ["Name ", " Status"]);
  assert.equal(r.status, "exact");
});

test("classifyColumns: empty expected is a no-op match", () => {
  const r = classifyColumns([], ["whatever"]);
  assert.equal(r.status, "exact");
  assert.equal(r.offset, 0);
});

test("classifyColumns: detects a uniform rightward shift", () => {
  // A column was inserted at the front, so the block moved one column right.
  const r = classifyColumns(["Name", "Email"], ["New", "Name", "Email"]);
  assert.equal(r.status, "offset");
  assert.equal(r.offset, 1);
  assert.equal(r.matchedStartColumn, 2);
});

test("classifyColumns: honors a non-default expected start column", () => {
  // Headers expected to start at column C (3); found starting at column D (4).
  const r = classifyColumns(["X", "Y"], ["", "", "", "X", "Y"], 3);
  assert.equal(r.status, "offset");
  assert.equal(r.offset, 1);
});

test("classifyColumns: a rename is a mismatch, not a shift", () => {
  const r = classifyColumns(["Name", "Email"], ["Name", "E-mail"]);
  assert.equal(r.status, "mismatch");
  assert.ok(r.reason);
});

test("classifyColumns: missing block is a mismatch", () => {
  const r = classifyColumns(["Name", "Email"], ["Region", "Owner"]);
  assert.equal(r.status, "mismatch");
});

test("classifyColumns: a block present at the expected start wins over duplicates", () => {
  // "X" sits exactly where expected (col A) AND repeats later — still exact.
  const r = classifyColumns(["X"], ["X", "Y", "X"]);
  assert.equal(r.status, "exact");
});

test("classifyColumns: ambiguous shift (not at expected start) is a mismatch", () => {
  // Expected at col A, but "X" appears only at cols B and D → can't pick a shift.
  const r = classifyColumns(["X"], ["Y", "X", "Z", "X"]);
  assert.equal(r.status, "mismatch");
  assert.ok(r.reason);
});

test("shiftColumnsInA1: shifts a cell range, leaves rows alone", () => {
  assert.equal(shiftColumnsInA1("A1:C10", 1), "B1:D10");
  assert.equal(shiftColumnsInA1("B2:D5", 2), "D2:F5");
});

test("shiftColumnsInA1: shifts full-column ranges", () => {
  assert.equal(shiftColumnsInA1("A:C", 1), "B:D");
});

test("shiftColumnsInA1: preserves a tab prefix", () => {
  assert.equal(shiftColumnsInA1("'My Tab'!A1:B2", 1), "'My Tab'!B1:C2");
  assert.equal(shiftColumnsInA1("Sheet1!C3", 3), "Sheet1!F3");
});

test("shiftColumnsInA1: offset 0 is a no-op", () => {
  assert.equal(shiftColumnsInA1("A1:C10", 0), "A1:C10");
});

test("shiftColumnsInA1: shifting before column A throws", () => {
  assert.throws(() => shiftColumnsInA1("A1:C10", -1), /before column A/);
});

// --- Row drift: classifySequence reused for key columns ---

test("classifySequence: key column exact at expected start row", () => {
  // Keys expected at rows 2-4; key column is [header, a, b, c] → indices 0..3.
  const r = classifySequence(["a", "b", "c"], ["Key", "a", "b", "c"], 2);
  assert.equal(r.status, "exact");
  assert.equal(r.offset, 0);
});

test("classifySequence: rows deleted above → upward offset", () => {
  // Expected keys at rows 5-6, but two rows above were deleted → now rows 3-4.
  const r = classifySequence(["acme", "beta"], ["Key", "x", "acme", "beta"], 5);
  assert.equal(r.status, "offset");
  assert.equal(r.offset, -2); // moved up two rows
  assert.equal(r.matchedStart, 3);
});

test("classifySequence: rows inserted above → downward offset", () => {
  const r = classifySequence(["acme"], ["Key", "acme"], 1);
  assert.equal(r.status, "offset");
  assert.equal(r.offset, 1);
});

test("classifySequence: a deleted target row is a mismatch", () => {
  // Expected acme,beta,gamma but beta is gone → not contiguous.
  const r = classifySequence(["acme", "beta", "gamma"], ["acme", "gamma"], 1);
  assert.equal(r.status, "mismatch");
});

test("classifySequence: reordered rows are a mismatch", () => {
  const r = classifySequence(["acme", "beta"], ["beta", "acme"], 1);
  assert.equal(r.status, "mismatch");
});

test("shiftRowsInA1: shifts rows, leaves columns alone", () => {
  assert.equal(shiftRowsInA1("B5:D20", 2), "B7:D22");
  assert.equal(shiftRowsInA1("B5:D20", -2), "B3:D18");
});

test("shiftRowsInA1: shifts full-row ranges", () => {
  assert.equal(shiftRowsInA1("5:8", 1), "6:9");
});

test("shiftRowsInA1: preserves a tab prefix with digits", () => {
  // The "1" in "Sheet1" must not be shifted — only the body after "!".
  assert.equal(shiftRowsInA1("Sheet1!A5:B6", 1), "Sheet1!A6:B7");
  assert.equal(shiftRowsInA1("'2024 Data'!C3", 2), "'2024 Data'!C5");
});

test("shiftRowsInA1: offset 0 is a no-op", () => {
  assert.equal(shiftRowsInA1("B5:D20", 0), "B5:D20");
});

test("shiftRowsInA1: shifting above row 1 throws", () => {
  assert.throws(() => shiftRowsInA1("A2:B3", -5), /above row 1/);
});
