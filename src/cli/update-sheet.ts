import { parseArgs } from "node:util";
import { appendCsv, replaceTabWithCsv } from "../core/sheets.js";

const USAGE = `Usage:
  update-sheet <sheet_id_or_url> <csv_path> [--tab NAME] [--mode replace|append]

Modes:
  replace (default)  Clear the tab and write the CSV from A1.
  append             Append CSV rows (skipping the first row as header) to the tab.

Auth:
  Reads a 1Password service-account token from $OP_SERVICE_ACCOUNT_TOKEN
  or ~/.config/1password/service-account-token, then fetches the Google
  service-account JSON at $SHEETS_SA_OP_REF (an op:// secret reference).
`;

async function main() {
  let parsed;
  try {
    parsed = parseArgs({
      args: process.argv.slice(2),
      options: {
        tab: { type: "string" },
        mode: { type: "string", default: "replace" },
        help: { type: "boolean", short: "h", default: false },
      },
      allowPositionals: true,
      strict: true,
    });
  } catch (e) {
    process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n\n${USAGE}`);
    process.exit(2);
  }

  if (parsed.values.help || parsed.positionals.length < 2) {
    process.stdout.write(USAGE);
    process.exit(parsed.values.help ? 0 : 2);
  }

  const [sheet, csvPath] = parsed.positionals;
  const tab = parsed.values.tab;
  const mode = parsed.values.mode;

  if (mode !== "replace" && mode !== "append") {
    process.stderr.write(`Invalid --mode: ${mode}. Must be 'replace' or 'append'.\n`);
    process.exit(2);
  }

  try {
    if (mode === "replace") {
      const r = await replaceTabWithCsv({ sheet, csvPath, tab });
      process.stdout.write(
        `Replaced '${r.tab}': ${r.rows} rows × ${r.columns} cols.\n`
      );
    } else {
      const r = await appendCsv({ sheet, csvPath, tab });
      process.stdout.write(`Appended ${r.appendedRows} rows to '${r.tab}'.\n`);
    }
  } catch (e) {
    process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
    process.exit(1);
  }
}

main().catch((e) => {
  process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
