const URL_RE = /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/;

export function resolveSheetId(input: string): string {
  const m = input.match(URL_RE);
  return m ? m[1] : input;
}
