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

## Setup

### 1. Google service account

In a Google Cloud project, enable the Sheets API, create a service account,
and download a JSON key. **Share every sheet you want Claude to edit with
the SA's email address as Editor** — same flow as sharing with a teammate.

### 2. Tell the server where the SA JSON lives

The server has to load that JSON file at startup. You have two options — pick
whichever fits how you already manage secrets.

**Option A — file path on disk** (simplest)

Put the JSON anywhere readable by the server process and point at it:

```bash
export SHEETS_SA_JSON_PATH="$HOME/.config/google/sheets-sa.json"
chmod 600 "$SHEETS_SA_JSON_PATH"
```

Works the same way if the file is mounted from Google Secret Manager, AWS
Secrets Manager, HashiCorp Vault, an encrypted volume, or anything else
that lands a file on disk.

**Option B — 1Password** (what I use)

I prefer not to keep the JSON sitting on disk, so I store it in 1Password
and let the server fetch it at startup via the
[`@1password/sdk`](https://www.npmjs.com/package/@1password/sdk) package
(no `op` CLI required).

1. Put the SA JSON in a 1Password item (file attachment or text field).
2. Create a 1Password **service account** with Read access to that vault.
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

## Tests

```bash
npm test
```

## License

[MIT](LICENSE)
