# google-sheets-editing-mcp

[![ci](https://github.com/hdanirwin/google-sheets-editing-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/hdanirwin/google-sheets-editing-mcp/actions/workflows/ci.yml)
[![license: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Claude's native Google Drive MCP can **read** Google Sheets. This one lets it
**edit** them — write CSVs into tabs, append rows, freeze headers, create new
worksheets, apply conditional formatting — all from a chat prompt.

Built for RevOps and GTM engineers who already live in Sheets: lead lists,
pipeline reports, territory plans, quota trackers. Now Claude Code can not
just manipulate data, but also update spreadsheets as desired.

## Things you can ask Claude Code to do with this

- "Replace the **Pipeline** tab with this HubSpot export and freeze the
  header."
- "Add a **Status** column to the **Accounts** tab and fill it based on
  these rules."
- "Create one tab per AE in this list, populated with their assigned
  accounts."
- "Append last week's inbound leads to the **Inbound** tab — skip the
  header row."
- "Highlight any row in **Q3 Forecast** where amount > $50k and stage =
  'Closed Lost'."

## Tools

| Tool | What it does |
|---|---|
| `get_sheet_info` | List every tab with its dimensions and gid |
| `read_range` | Read values from a range |
| `update_range` | Write values to a range |
| `append_rows` | Append rows after existing data |
| `clear_range` | Clear cells (preserves formatting) |
| `replace_tab_with_csv` | Overwrite a tab from a CSV file |
| `append_csv` | Append CSV rows (skips header by default) |
| `create_tab` / `rename_tab` / `delete_tab` | Worksheet lifecycle |
| `batch_update` | Raw Sheets API for formatting, freezes, merges, conditional formatting |

Data tools also accept optional `expectedHeaders` (column drift) and, on the
range-based tools, `expectedKeys` (row drift) to confirm the layout hasn't
changed before they act — see
[Confirming columns haven't changed](#confirming-columns-havent-changed).

## Confirming columns haven't changed

Every data tool (`read_range`, `update_range`, `append_rows`, `clear_range`,
`replace_tab_with_csv`, `append_csv`, `batch_update`) takes an optional
`expectedHeaders` argument. When you pass it, the tool reads the actual header
row **before** acting and compares it to what you expect — so a write that
assumed last turn's layout can't silently land in the wrong columns.

```jsonc
{
  "sheet": "…",
  "tab": "Accounts",
  "range": "C2:C500",
  "values": [["Active"], ["Churned"]],
  "expectedHeaders": ["Account", "Owner", "Status"]  // what column A→C should be
}
```

Three outcomes:

- **Exact match** — the headers are where you expect; the operation proceeds
  unchanged.
- **Uniform shift** — the headers are intact but moved by N columns (e.g.
  someone inserted a column to the left). The operation **proceeds with the
  target auto-adjusted**: ranges shift by N, appended rows are padded to stay
  aligned. The response's `columnCheck` reports the offset and the adjusted
  range.
- **Anything else** — a header was renamed, deleted, reordered, or the match is
  ambiguous. The operation **stops** and returns the expected-vs-actual rows so
  you can re-read the sheet and decide the next step.

Optional companions: `headerRow` (1-based, default 1) if the headers aren't on
row 1, and `headerStartColumn` (default `A`) if `expectedHeaders` begins
somewhere other than column A. `batch_update` is verify-only — because raw
requests target columns by index, even a uniform shift stops the call so you
can rebuild the requests for the current layout.

### Rows: catching deleted or reordered rows

Columns drift sideways; rows drift up and down. If a row above your target was
deleted, or rows got reordered between turns, "row 5" now points at a different
record. The range-based tools (`read_range`, `update_range`, `clear_range`) take
an optional `expectedKeys` — the row-identity values you expect at the target
rows — plus `keyColumn`, the column that holds those identities (default `A`):

```jsonc
{
  "sheet": "…",
  "tab": "Accounts",
  "range": "C5:C7",                       // you think these are rows 5–7
  "values": [["Active"], ["Churned"], ["Active"]],
  "keyColumn": "A",                       // the Account-ID column
  "expectedKeys": ["acme", "beta", "gamma"]  // what column A should read at 5–7
}
```

The tool reads `keyColumn`, finds where those keys actually are now, and:

- **Exact** — the keys are at the rows you targeted; proceeds.
- **Uniform shift** — the keys moved together by N rows (e.g. a row was deleted
  above, so everything below slid up). The target rows are **auto-adjusted** by N
  and it proceeds; `rowCheck` reports the offset.
- **Anything else** — a key is missing (row deleted) or out of order (rows
  reordered). The operation **stops** and tells you which keys are missing or
  that the block isn't contiguous, so you can re-read and re-plan.

`expectedKeys` requires a `range` with a concrete start row (e.g. `C5:C7`), since
that's what anchors the rows. Column and row guards compose: if both drifted, the
target is shifted on both axes (and the key column is read from its shifted
position). Each result carries `columnCheck` and/or `rowCheck` describing what
happened.

## Setup

### 1. Google service account

In a Google Cloud project, enable the Sheets API, create a service account,
and download a JSON key. **Share every sheet you want Claude to edit with
the SA's email address as Editor** — same flow as sharing with a teammate.

### 2. Tell the server where the SA JSON lives

The server has to load that JSON file at startup. You have two options — pick
whichever fits how you already manage secrets.

**Option 1 — 1Pass** 

Storing key in 1Pass allows for easy portability and security. 
[`@1password/sdk`](https://www.npmjs.com/package/@1password/sdk) package
(no `op` CLI required).

1. Put the SA JSON in a 1Pass item (file attachment or text field).
2. Create a 1Pass **service account** with Read access to that vault.
3. Configure the env:

   ```bash
   export SHEETS_SA_OP_REF="op://<vault>/<item>/<field-or-filename>"

   # Token: either env var…
   export OP_SERVICE_ACCOUNT_TOKEN='ops_...'
   # …or a 600-mode file (no quotes, no trailing newline):
   printf '%s' 'ops_...' > ~/.config/1password/service-account-token
   chmod 600 ~/.config/1password/service-account-token
   ```

If both `SHEETS_SA_JSON_PATH` and `SHEETS_SA_OP_REF` are set, the path
wins.


**Option 2 — file path on disk** (simplest)

Put the JSON anywhere readable by the server process and point at it:

```bash
export SHEETS_SA_JSON_PATH="$HOME/.config/google/sheets-sa.json"
chmod 600 "$SHEETS_SA_JSON_PATH"
```

Works the same way if the file is mounted from Google Secret Manager, AWS
Secrets Manager, HashiCorp Vault, an encrypted volume, or anything else
that lands a file on disk.


### 3. Build and register

```bash
npm install
npm run build
```

Register with Claude Code in `~/.claude.json`:

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "/absolute/path/to/bin/google-sheets-mcp",
      "env": {
        "SHEETS_SA_JSON_PATH": "/Users/you/.config/google/sheets-sa.json"
      }
    }
  }
}
```

(Swap in `SHEETS_SA_OP_REF` + `OP_SERVICE_ACCOUNT_TOKEN` if you went with
Option B.)

## CLI

For ad-hoc use without Claude — the most common operation is bulk-writing a
CSV to a tab:

```bash
bin/update-sheet <sheet_url_or_id> <csv_path> [--tab NAME] [--mode replace|append]
```

`replace` (default) writes the CSV from A1 and clears trailing rows/cols.
`append` skips the header and adds the rest after the last existing row.

## Reliability

- 30-second timeout per Sheets API call.
- Idempotent ops (GET/PUT/DELETE) retry up to 3× on 429/5xx.
- Non-idempotent ops (append, batchUpdate) never retry, so you never see
  duplicate writes — your weekly lead import won't accidentally land twice.
- `replace_tab_with_csv` writes new data **before** clearing trailing rows
  and columns, so a transient failure preserves the original tab rather
  than wiping it.

## Concurrency

Multiple Claude conversations — or any clients sharing the service account — can
hit the same spreadsheet at once. There is **no locking anywhere in this stack**:
each conversation runs its own server process, and the server only forwards calls
to Google. The Sheets API is the sole arbiter. In practice:

- **Reads never conflict.** Any number of readers can read the same tab at once;
  each gets a consistent snapshot of the moment its request was served. Reads
  never block reads or writes.
- **Edits to different tabs are safe.** Different cells, no conflict — both
  writes land.
- **Edits to the same range are last-write-wins, silently.** The Sheets values
  API has no optimistic concurrency (no "write only if unchanged"), so two
  writers racing the same cells get no conflict error — the later write wins.
- **Only a single `batch_update` is atomic** (all-or-nothing). No transaction
  spans multiple calls.
- **Quota is shared.** All clients on one service account draw the same per-user
  Sheets quota, so heavy concurrent use can hit `429`. Idempotent ops retry;
  `append_rows` / `batch_update` don't, surfacing an error rather than risking a
  duplicate.

### The guard is drift detection, not a lock

`expectedHeaders` / `expectedKeys` read the current layout and *then* write — two
separate API calls. Another writer can change the sheet in the gap between them,
so the guard **shrinks** the race window and catches drift that happened *before*
the call; it does **not** make read-then-write atomic. Use it for "did the layout
change since I last looked," not for "are two writers racing right now." Any true
mutual exclusion has to be cooperative and live outside the Sheet — and can only
bind clients that opt into it (the web UI and other users never will), which is
exactly why the guard exists alongside it.

## Tests

```bash
npm test
```

## License

[MIT](LICENSE)
