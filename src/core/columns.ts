// Drift detection. Pure, network-free helpers used by the guards that every
// data tool runs before writing. Two axes share one algorithm:
//   - Column drift: compare the current header ROW to the headers you expected.
//   - Row drift:    compare the current KEY COLUMN (an ID/email/name column) to
//                   the row keys you expected — catches rows deleted or reordered.
//
// Both reduce to: "is this expected sequence still present, contiguous, and in
// order — and if so, where?" Three outcomes:
//   - exact:    the expected sequence sits at the expected start position.
//   - offset:   the sequence is intact and contiguous, but shifted by a uniform
//               N (a column inserted to the left; rows deleted above). The caller
//               can compensate by shifting its target by N.
//   - mismatch: anything else — an entry was renamed/deleted/reordered, or the
//               block is ambiguous. The caller should stop and re-plan.

export type DriftStatus = "exact" | "offset" | "mismatch";
/** Back-compat alias — column drift used this name first. */
export type ColumnCheckStatus = DriftStatus;

/** Result of classifying one sequence (a header row or a key column). */
export interface SequenceCheck {
  status: DriftStatus;
  /** How far the actual block is shifted from the expected start (actual - expected). 0 unless status === "offset". */
  offset: number;
  expected: string[];
  actual: string[];
  /** 1-based position where the caller expected the block to begin. */
  expectedStart: number;
  /** 1-based position where the contiguous block was actually found, or null when not found. */
  matchedStart: number | null;
  /** Human-readable explanation, set for mismatches. */
  reason?: string;
}

/** Column-flavored view of a SequenceCheck (header-named fields). */
export interface ColumnCheck {
  status: ColumnCheckStatus;
  offset: number;
  expectedHeaders: string[];
  actualHeaders: string[];
  expectedStartColumn: number;
  matchedStartColumn: number | null;
  reason?: string;
}

/** 1 → "A", 27 → "AA", 18278 → "ZZZ". Non-positive falls back to "A". */
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

/** Inverse of colLetter: "A" → 1, "AA" → 27. Returns 0 for non-letter input. */
export function colNumber(letters: string): number {
  if (!letters) return 0;
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const d = ch.charCodeAt(0) - 64; // 'A' === 65
    if (d < 1 || d > 26) return 0;
    n = n * 26 + d;
  }
  return n;
}

// Entries are compared trimmed so trailing whitespace in a cell doesn't read as
// drift. Comparison is otherwise exact (case- and content-sensitive).
const norm = (s: string): string => (s ?? "").trim();

/**
 * Classify how an actual sequence relates to an expected one: find the expected
 * block contiguously within the actual values and report exact / offset /
 * mismatch. Used for both header rows (columns) and key columns (rows).
 *
 * @param expected      Values the caller expects, in order.
 * @param actual        The values as read from the sheet, from position 1.
 * @param expectedStart 1-based position where `expected` should begin (default 1).
 */
export function classifySequence(
  expected: string[],
  actual: string[],
  expectedStart = 1
): SequenceCheck {
  const base = { expected, actual, expectedStart };
  if (expected.length === 0) {
    // Nothing to confirm — treat as a no-op match.
    return { ...base, status: "exact", offset: 0, matchedStart: expectedStart };
  }

  const exp = expected.map(norm);
  const act = actual.map(norm);

  // Every starting index where the expected block appears contiguously (0-based).
  const positions: number[] = [];
  for (let j = 0; j + exp.length <= act.length; j++) {
    let hit = true;
    for (let k = 0; k < exp.length; k++) {
      if (act[j + k] !== exp[k]) {
        hit = false;
        break;
      }
    }
    if (hit) positions.push(j);
  }

  const expectedIdx = expectedStart - 1; // 0-based

  if (positions.includes(expectedIdx)) {
    return { ...base, status: "exact", offset: 0, matchedStart: expectedStart };
  }
  if (positions.length === 1) {
    const j = positions[0];
    return { ...base, status: "offset", offset: j - expectedIdx, matchedStart: j + 1 };
  }

  return {
    ...base,
    status: "mismatch",
    offset: 0,
    matchedStart: null,
    reason:
      positions.length > 1
        ? "The expected values appear in more than one position, so the shift is ambiguous."
        : "The expected values were not found as a contiguous block.",
  };
}

/**
 * Column-flavored wrapper around classifySequence: compares an actual header row
 * to expected headers and returns the result with header-named fields.
 */
export function classifyColumns(
  expected: string[],
  actualRow: string[],
  expectedStartColumn = 1
): ColumnCheck {
  const r = classifySequence(expected, actualRow, expectedStartColumn);
  return {
    status: r.status,
    offset: r.offset,
    expectedHeaders: r.expected,
    actualHeaders: r.actual,
    expectedStartColumn: r.expectedStart,
    matchedStartColumn: r.matchedStart,
    reason: r.reason,
  };
}

/**
 * Shift every column reference in an A1 range by `offset` columns, leaving rows
 * and any "Tab!" prefix untouched. "A1:C10" + 1 → "B1:D10"; "A:C" + 1 → "B:D".
 * Throws if the shift would move a column before column A.
 */
export function shiftColumnsInA1(a1: string, offset: number): string {
  if (offset === 0) return a1;
  const bang = a1.lastIndexOf("!");
  const prefix = bang >= 0 ? a1.slice(0, bang + 1) : "";
  const body = bang >= 0 ? a1.slice(bang + 1) : a1;
  const shifted = body.replace(/[A-Za-z]+/g, (letters) => {
    const col = colNumber(letters);
    if (col === 0) return letters; // not a column token; leave as-is
    const next = col + offset;
    if (next < 1) {
      throw new Error(
        `Cannot apply a ${offset}-column shift: it moves column ${letters} before column A.`
      );
    }
    return colLetter(next);
  });
  return prefix + shifted;
}

/**
 * Shift every row reference in an A1 range by `offset` rows, leaving columns and
 * any "Tab!" prefix untouched. "B5:D20" + 2 → "B7:D22"; "5:8" + 1 → "6:9".
 * Throws if the shift would move a row above row 1.
 */
export function shiftRowsInA1(a1: string, offset: number): string {
  if (offset === 0) return a1;
  const bang = a1.lastIndexOf("!");
  const prefix = bang >= 0 ? a1.slice(0, bang + 1) : "";
  const body = bang >= 0 ? a1.slice(bang + 1) : a1;
  const shifted = body.replace(/\d+/g, (digits) => {
    const next = parseInt(digits, 10) + offset;
    if (next < 1) {
      throw new Error(
        `Cannot apply a ${offset}-row shift: it moves row ${digits} above row 1.`
      );
    }
    return String(next);
  });
  return prefix + shifted;
}
