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
