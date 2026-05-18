import { readFile } from "node:fs/promises";
import Papa from "papaparse";

export async function readCsvFile(path: string): Promise<string[][]> {
  const text = await readFile(path, "utf8");
  return parseCsv(text);
}

export function parseCsv(text: string): string[][] {
  const result = Papa.parse<string[]>(text, {
    skipEmptyLines: "greedy",
    dynamicTyping: false,
  });
  if (result.errors.length) {
    const first = result.errors[0];
    throw new Error(`CSV parse error at row ${first.row}: ${first.message}`);
  }
  return result.data as string[][];
}
