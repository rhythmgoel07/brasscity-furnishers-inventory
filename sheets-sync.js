// ============================================================================
// Google Sheets I/O for two-way sync.
//
// Uses the `drive.file` scope only. That scope is classified NON-SENSITIVE by
// Google, so this app never has to go through OAuth verification or a security
// assessment. The tradeoff is that it can only touch files it created itself —
// which is exactly what we want: this app can see its own inventory sheet and
// nothing else in your Drive. Broader scopes like drive.readonly would drag the
// project into a paid annual security review.
//
// All merge decisions live in sync-engine.js. This file only reads, writes, and
// remembers — it never decides who wins.
// ============================================================================

import { SYNC_FIELDS, FIELD_HEADERS, ID_HEADER, normalizeRecord } from "./sync-engine.js";

const GIS_SRC = "https://accounts.google.com/gsi/client";
const SCOPE = "https://www.googleapis.com/auth/drive.file";
const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
const SHEET_NAME = "Inventory";
const HEADERS = [ID_HEADER, ...SYNC_FIELDS.map((f) => FIELD_HEADERS[f])];

let tokenClient = null;
let accessToken = null;
let tokenExpiresAt = 0;

function loadGis() {
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = GIS_SRC; s.async = true;
    s.onload = resolve;
    s.onerror = () => reject(new Error("Couldn't load Google sign-in"));
    document.head.appendChild(s);
  });
}

// Access tokens last about an hour. Once you've consented, re-requesting is
// silent (prompt: ""), so this only shows a popup the very first time.
export async function getAccessToken(clientId, { interactive = false } = {}) {
  if (accessToken && Date.now() < tokenExpiresAt - 60_000) return accessToken;
  await loadGis();

  if (!tokenClient) {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: clientId, scope: SCOPE, callback: () => {},
    });
  }

  return new Promise((resolve, reject) => {
    tokenClient.callback = (resp) => {
      if (resp.error) {
        // A silent attempt failing just means consent is needed — not a real error.
        return reject(Object.assign(new Error(resp.error), { needsConsent: true }));
      }
      accessToken = resp.access_token;
      tokenExpiresAt = Date.now() + (Number(resp.expires_in) || 3600) * 1000;
      resolve(accessToken);
    };
    tokenClient.requestAccessToken({ prompt: interactive ? "consent" : "" });
  });
}

export function forgetToken() { accessToken = null; tokenExpiresAt = 0; }

async function api(url, token, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) },
  });
  if (res.status === 401 || res.status === 403) {
    forgetToken();
    throw new Error("Google access expired — reconnect and try again");
  }
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json())?.error?.message || ""; } catch { /* body wasn't JSON */ }
    throw new Error(`Sheets API ${res.status}: ${detail || res.statusText}`);
  }
  return res.json();
}

function colLetter(index) {
  let s = "", n = index + 1;
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ----------------------------------------------------------------------------
// Spreadsheet creation
// ----------------------------------------------------------------------------
export async function createSpreadsheet(token, title = "Furniture Inventory (synced)") {
  const created = await api(SHEETS_API, token, {
    method: "POST",
    body: JSON.stringify({
      properties: { title },
      sheets: [{ properties: { title: SHEET_NAME, gridProperties: { frozenRowCount: 1 } } }],
    }),
  });
  const spreadsheetId = created.spreadsheetId;
  const sheetId = created.sheets[0].properties.sheetId;

  await api(`${SHEETS_API}/${spreadsheetId}/values/${SHEET_NAME}!A1?valueInputOption=RAW`, token, {
    method: "PUT",
    body: JSON.stringify({ values: [HEADERS] }),
  });

  // Hide the _id column and bold the header. The ID column is machine-managed:
  // it's what lets you rename a SKU in the sheet without the app reading it as
  // "one item deleted, a different one added" and losing the attached photo.
  await api(`${SHEETS_API}/${spreadsheetId}:batchUpdate`, token, {
    method: "POST",
    body: JSON.stringify({
      requests: [
        { updateDimensionProperties: {
            range: { sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1 },
            properties: { hiddenByUser: true }, fields: "hiddenByUser" } },
        { repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: { userEnteredFormat: { textFormat: { bold: true } } },
            fields: "userEnteredFormat.textFormat.bold" } },
      ],
    }),
  });

  return { spreadsheetId, sheetId, url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit` };
}

export async function getSheetId(token, spreadsheetId) {
  const meta = await api(`${SHEETS_API}/${spreadsheetId}?fields=sheets.properties`, token);
  const sheet = meta.sheets.find((s) => s.properties.title === SHEET_NAME) || meta.sheets[0];
  return sheet.properties.sheetId;
}

// ----------------------------------------------------------------------------
// Reading
// ----------------------------------------------------------------------------
export async function readSheet(token, spreadsheetId) {
  const data = await api(
    `${SHEETS_API}/${spreadsheetId}/values/${SHEET_NAME}!A1:ZZ100000?valueRenderOption=UNFORMATTED_VALUE`,
    token
  );
  const rows = data.values || [];
  if (rows.length === 0) return { records: new Map(), rowOf: new Map(), colOf: {}, issues: ["The sheet is empty."] };

  const header = rows[0].map((h) => String(h || "").trim());

  // Columns are matched BY NAME, never by position. Someone inserting a column
  // in the middle of the sheet would otherwise shift every value one field to
  // the right and silently rewrite the whole catalogue.
  const colOf = {};
  const issues = [];
  const idCol = header.indexOf(ID_HEADER);
  if (idCol === -1) issues.push(`Missing the hidden "${ID_HEADER}" column — don't delete or rename it.`);
  for (const f of SYNC_FIELDS) {
    const i = header.indexOf(FIELD_HEADERS[f]);
    if (i === -1) issues.push(`Missing column "${FIELD_HEADERS[f]}".`);
    else colOf[f] = i;
  }
  if (issues.length) return { records: new Map(), rowOf: new Map(), colOf, issues };

  const records = new Map();
  const rowOf = new Map();
  const seen = new Set();
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every((c) => c === "" || c === null || c === undefined)) continue; // blank spacer row
    const id = String(row[idCol] ?? "").trim();
    if (!id) { issues.push(`Row ${r + 1} has no ${ID_HEADER} — add rows through the app, or leave that cell for the app to fill.`); continue; }
    if (seen.has(id)) { issues.push(`Row ${r + 1} duplicates id ${id}. Remove the duplicate before syncing.`); continue; }
    seen.add(id);
    const rec = {};
    for (const f of SYNC_FIELDS) rec[f] = row[colOf[f]];
    records.set(id, normalizeRecord(rec));
    rowOf.set(id, r + 1); // 1-based sheet row number
  }
  return { records, rowOf, colOf, issues };
}

// ----------------------------------------------------------------------------
// Writing
// ----------------------------------------------------------------------------

// Writes only the specific cells the merge decided should change. Rewriting whole
// rows would clobber columns the sheet changed and the app didn't.
export async function writeCells(token, spreadsheetId, edits, rowOf, colOf) {
  const data = [];
  for (const { id, fields } of edits) {
    const row = rowOf.get(id);
    if (!row) continue;
    for (const [f, value] of Object.entries(fields)) {
      if (colOf[f] === undefined) continue;
      data.push({ range: `${SHEET_NAME}!${colLetter(colOf[f])}${row}`, values: [[value]] });
    }
  }
  if (!data.length) return 0;
  // Chunked: a single batch with thousands of ranges gets rejected as too large.
  for (let i = 0; i < data.length; i += 500) {
    await api(`${SHEETS_API}/${spreadsheetId}/values:batchUpdate`, token, {
      method: "POST",
      body: JSON.stringify({ valueInputOption: "RAW", data: data.slice(i, i + 500) }),
    });
  }
  return data.length;
}

export async function appendRows(token, spreadsheetId, creations) {
  if (!creations.length) return 0;
  const values = creations.map(({ id, record }) => {
    const r = normalizeRecord(record);
    return [id, ...SYNC_FIELDS.map((f) => r[f])];
  });
  for (let i = 0; i < values.length; i += 500) {
    await api(
      `${SHEETS_API}/${spreadsheetId}/values/${SHEET_NAME}!A1:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
      token,
      { method: "POST", body: JSON.stringify({ values: values.slice(i, i + 500) }) }
    );
  }
  return values.length;
}

export async function deleteRows(token, spreadsheetId, sheetId, ids, rowOf) {
  const rows = ids.map(({ id }) => rowOf.get(id)).filter(Boolean);
  if (!rows.length) return 0;
  // Bottom-up: deleting row 5 first would shift row 9 up to 8, so a top-down
  // pass would delete the wrong rows.
  rows.sort((a, b) => b - a);
  const requests = rows.map((row) => ({
    deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: row - 1, endIndex: row } },
  }));
  for (let i = 0; i < requests.length; i += 200) {
    await api(`${SHEETS_API}/${spreadsheetId}:batchUpdate`, token, {
      method: "POST", body: JSON.stringify({ requests: requests.slice(i, i + 200) }),
    });
  }
  return rows.length;
}
