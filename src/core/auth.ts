import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import { createClient } from "@1password/sdk";
import { google, sheets_v4 } from "googleapis";

const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
const REQUEST_TIMEOUT_MS = 30_000;
const TOKEN_FILE = join(homedir(), ".config/1password/service-account-token");

interface ServiceAccountKey {
  client_email: string;
  private_key: string;
  [key: string]: unknown;
}

let cachedSheets: sheets_v4.Sheets | null = null;
let cachedEmail: string | null = null;
let initPromise: Promise<sheets_v4.Sheets> | null = null;

async function loadOpToken(): Promise<string> {
  const fromEnv = process.env.OP_SERVICE_ACCOUNT_TOKEN?.trim();
  if (fromEnv) return fromEnv;

  try {
    const fromDisk = (await readFile(TOKEN_FILE, "utf8")).trim();
    if (!fromDisk) {
      throw new Error(`1Password service-account token file is empty: ${TOKEN_FILE}`);
    }
    return fromDisk;
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === "ENOENT") {
      throw new Error(
        `Missing 1Password service-account token. Set $OP_SERVICE_ACCOUNT_TOKEN ` +
          `or create ${TOKEN_FILE} (mode 600).`
      );
    }
    throw err;
  }
}

async function loadSaJson(): Promise<string> {
  // Direct file path takes precedence — simplest option for users who keep
  // the SA JSON on disk (e.g., behind GCP Secret Manager, AWS, Vault, etc.).
  const path = process.env.SHEETS_SA_JSON_PATH?.trim();
  if (path) {
    try {
      return (await readFile(path, "utf8")).trim();
    } catch (err) {
      throw new Error(
        `Failed to read SHEETS_SA_JSON_PATH=${path}: ` +
          `${err instanceof Error ? err.message : String(err)}`
      );
    }
  }

  // Otherwise fall back to 1Password.
  const opRef = process.env.SHEETS_SA_OP_REF?.trim();
  if (opRef) return fetchSaJsonFrom1Password(opRef);

  throw new Error(
    "No credentials configured. Set either SHEETS_SA_JSON_PATH (path to your " +
      "Google service-account JSON on disk) or SHEETS_SA_OP_REF (1Password " +
      "secret reference, e.g. op://<vault>/<item>/<filename>.json). " +
      "See .env.example."
  );
}

async function fetchSaJsonFrom1Password(opRef: string): Promise<string> {
  const token = await loadOpToken();
  let client;
  try {
    client = await createClient({
      auth: token,
      integrationName: "google-sheets-editor",
      integrationVersion: "0.1.0",
    });
  } catch (err) {
    throw new Error(
      `Failed to initialize 1Password SDK client: ${err instanceof Error ? err.message : String(err)}`
    );
  }
  try {
    return (await client.secrets.resolve(opRef)).trim();
  } catch (err) {
    throw new Error(
      `Failed to fetch SA JSON from 1Password (ref: ${opRef}): ` +
        `${err instanceof Error ? err.message : String(err)}. ` +
        `Check: (1) the token in $OP_SERVICE_ACCOUNT_TOKEN (or ${TOKEN_FILE}) is current; ` +
        `(2) the service account has Read access to the vault; ` +
        `(3) the SHEETS_SA_OP_REF env var points at a valid op:// path.`
    );
  }
}

export function initSheetsClient(): Promise<sheets_v4.Sheets> {
  if (cachedSheets) return Promise.resolve(cachedSheets);
  if (initPromise) return initPromise;

  initPromise = (async () => {
    const saJson = await loadSaJson();

    let parsed: ServiceAccountKey;
    try {
      parsed = JSON.parse(saJson);
    } catch (err) {
      throw new Error(
        `Loaded SA credential is not valid JSON ` +
          `(${saJson.length} bytes received): ` +
          `${err instanceof Error ? err.message : String(err)}`
      );
    }

    const auth = new google.auth.GoogleAuth({
      credentials: parsed,
      scopes: SCOPES,
    });

    cachedSheets = google.sheets({
      version: "v4",
      auth,
      timeout: REQUEST_TIMEOUT_MS,
      retry: true,
      retryConfig: {
        retry: 3,
        retryDelay: 200,
        // Only retry idempotent methods so we never duplicate writes/appends.
        httpMethodsToRetry: ["GET", "PUT", "DELETE"],
        statusCodesToRetry: [
          [429, 429],
          [500, 599],
        ],
      },
    });
    cachedEmail = parsed.client_email ?? null;
    return cachedSheets;
  })();

  initPromise.catch(() => {
    initPromise = null;
  });
  return initPromise;
}

export function getServiceAccountEmail(): string | null {
  return cachedEmail;
}
