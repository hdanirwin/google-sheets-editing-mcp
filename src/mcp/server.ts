import { McpServer, ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z, ZodRawShape } from "zod";
import {
  appendCsv,
  appendRows,
  batchUpdate,
  clearRange,
  createTab,
  deleteTab,
  getSheetInfo,
  readRange,
  renameTab,
  replaceTabWithCsv,
  updateRange,
} from "../core/sheets.js";
import { initSheetsClient } from "../core/auth.js";

const sheet = z.string().describe("Google Sheets URL or bare spreadsheet ID.");
const cellValues = z
  .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
  .describe("2D array of cell values, rows of cells.");
const valueInputOption = z
  .enum(["USER_ENTERED", "RAW"])
  .optional()
  .describe(
    "USER_ENTERED parses formulas/dates like the UI does. RAW writes literal strings. Default: USER_ENTERED."
  );
const TAB_OR_RANGE_REQUIRED = "Either tab or range (or both) must be provided.";
const tabOrRangePresent = (args: { tab?: string; range?: string }) =>
  !!(args.tab || args.range);

function ok(data: unknown) {
  return {
    content: [
      {
        type: "text" as const,
        text: typeof data === "string" ? data : JSON.stringify(data, null, 2),
      },
    ],
  };
}

function err(e: unknown) {
  return {
    isError: true,
    content: [
      {
        type: "text" as const,
        text: e instanceof Error ? e.message : String(e),
      },
    ],
  };
}

type ArgsOf<S extends ZodRawShape> = { [K in keyof S]: z.infer<S[K]> };

function registerTool<S extends ZodRawShape>(
  server: McpServer,
  name: string,
  description: string,
  schema: S,
  fn: (args: ArgsOf<S>) => Promise<unknown>
) {
  const cb = (async (args: ArgsOf<S>) => {
    try {
      return ok(await fn(args));
    } catch (e) {
      return err(e);
    }
  }) as unknown as ToolCallback<S>;
  server.tool(name, description, schema, cb);
}

// Tools whose schemas need the cross-field tab-or-range rule. McpServer's
// tool() takes a ZodRawShape, not a refined ZodObject, so we apply the guard
// inside the handler instead of at schema-build time.
function registerToolWithGuard<S extends ZodRawShape>(
  server: McpServer,
  name: string,
  description: string,
  schema: S,
  guard: (args: ArgsOf<S>) => boolean,
  guardMessage: string,
  fn: (args: ArgsOf<S>) => Promise<unknown>
) {
  const cb = (async (args: ArgsOf<S>) => {
    if (!guard(args)) return err(new Error(guardMessage));
    try {
      return ok(await fn(args));
    } catch (e) {
      return err(e);
    }
  }) as unknown as ToolCallback<S>;
  server.tool(name, description, schema, cb);
}

async function main() {
  // Eagerly load credentials so we fail fast at startup, not on first tool call.
  await initSheetsClient();
  process.stderr.write("[google-sheets-editor] ready\n");

  const server = new McpServer({
    name: "google-sheets-editor",
    version: "0.1.0",
  });

  registerTool(
    server,
    "get_sheet_info",
    "Get spreadsheet title and a list of all tabs with their dimensions. Use this first when you don't know the sheet structure.",
    { sheet },
    (args) => getSheetInfo(args.sheet)
  );

  registerToolWithGuard(
    server,
    "read_range",
    "Read cell values from a range. Pass either tab (whole tab), range (A1 notation with or without tab prefix), or both.",
    {
      sheet,
      tab: z.string().optional().describe("Tab name. Optional if range includes it."),
      range: z
        .string()
        .optional()
        .describe("A1 notation, e.g. 'A1:C10' or 'MyTab!A1:C10'."),
      valueRenderOption: z
        .enum(["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"])
        .optional()
        .describe("How to render cell values. Default: FORMATTED_VALUE."),
    },
    tabOrRangePresent,
    TAB_OR_RANGE_REQUIRED,
    (args) => readRange(args)
  );

  registerTool(
    server,
    "update_range",
    "Write values to a specific range. Existing cells outside the written range are not affected.",
    {
      sheet,
      tab: z.string().optional(),
      range: z
        .string()
        .describe(
          "A1 notation for the top-left or full range to write, e.g. 'A1' or 'A1:C3'."
        ),
      values: cellValues,
      valueInputOption,
    },
    (args) => updateRange(args)
  );

  registerTool(
    server,
    "append_rows",
    "Append rows after the last row of existing data on a tab.",
    {
      sheet,
      tab: z.string(),
      values: cellValues,
      valueInputOption,
    },
    (args) => appendRows(args)
  );

  registerToolWithGuard(
    server,
    "clear_range",
    "Clear cell values in a range (formatting is preserved). Provide tab to clear an entire tab.",
    {
      sheet,
      tab: z.string().optional(),
      range: z.string().optional(),
    },
    tabOrRangePresent,
    TAB_OR_RANGE_REQUIRED,
    (args) => clearRange(args)
  );

  registerTool(
    server,
    "replace_tab_with_csv",
    "Write a CSV file's contents to a tab, replacing the existing data. New data is written first; any rows/columns past the new data's edge are cleared after. Defaults to the first tab if none given.",
    {
      sheet,
      csvPath: z.string().describe("Absolute path to the CSV file."),
      tab: z.string().optional(),
    },
    (args) => replaceTabWithCsv(args)
  );

  registerTool(
    server,
    "append_csv",
    "Append a CSV file's rows to a tab. By default skips the first row (assumed header).",
    {
      sheet,
      csvPath: z.string(),
      tab: z.string().optional(),
      includeHeader: z
        .boolean()
        .optional()
        .describe("If true, append the first row too. Default: false."),
    },
    (args) => appendCsv(args)
  );

  registerTool(
    server,
    "create_tab",
    "Create a new worksheet (tab) in the spreadsheet.",
    {
      sheet,
      title: z.string(),
      rowCount: z.number().int().positive().optional(),
      columnCount: z.number().int().positive().optional(),
    },
    (args) => createTab(args)
  );

  registerTool(
    server,
    "rename_tab",
    "Rename an existing worksheet.",
    {
      sheet,
      oldTitle: z.string(),
      newTitle: z.string(),
    },
    async (args) => {
      await renameTab(args);
      return { renamed: args.oldTitle, to: args.newTitle };
    }
  );

  registerTool(
    server,
    "delete_tab",
    "Delete a worksheet by title.",
    {
      sheet,
      title: z.string(),
    },
    async (args) => {
      await deleteTab(args);
      return { deleted: args.title };
    }
  );

  // Escape-hatch tool: the schema is intentionally permissive so callers can
  // pass any Sheets API Request shape. The `as` cast at the call site reflects
  // that we trust the caller, not that we've narrowed the type.
  registerTool(
    server,
    "batch_update",
    "Escape hatch: send raw Sheets API batchUpdate requests for formatting, freezing, merging, conditional formatting, etc. See https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request",
    {
      sheet,
      requests: z
        .array(z.record(z.unknown()))
        .describe("Array of Sheets API Request objects."),
    },
    (args) =>
      batchUpdate({
        sheet: args.sheet,
        requests: args.requests as Parameters<typeof batchUpdate>[0]["requests"],
      })
  );

  await server.connect(new StdioServerTransport());
}

main().catch((e) => {
  process.stderr.write(
    `[google-sheets-editor] fatal: ${e instanceof Error ? e.message : String(e)}\n`
  );
  process.exit(1);
});
