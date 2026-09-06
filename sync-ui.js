// ============================================================================
// Sync orchestration.
//
// Ties together three things that are deliberately kept apart:
//   sync-engine.js  decides WHAT should change (pure, tested, no I/O)
//   sheets-sync.js  talks to Google
//   this file       persists the baseline, drives the UI, applies the plan
//
// Nothing is ever written without the plan being shown first. That preview is
// the main safety property of the whole feature: two-way sync goes wrong
// silently or not at all, so it must not be silent.
// ============================================================================

import {
  doc, getDoc, setDoc, deleteDoc, writeBatch, collection, getDocs,
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";
import {
  SYNC_FIELDS, FIELD_HEADERS, mergeAll, resolveConflicts, nextBaseline, planSummary, normalizeRecord,
} from "./sync-engine.js";
import {
  getAccessToken, forgetToken, createSpreadsheet, getSheetId,
  readSheet, writeCells, appendRows, deleteRows,
} from "./sheets-sync.js";

const CONFIG_DOC = ["sync", "config"];
const BASELINE_PREFIX = "sync_baseline_";
const CHUNK_ITEMS = 800; // keeps each baseline document well under Firestore's 1 MB cap

let currentPlan = null;
let currentContext = null;
const conflictChoices = {};

// ----------------------------------------------------------------------------
// Config + baseline persistence
// ----------------------------------------------------------------------------
export async function loadConfig(db) {
  const snap = await getDoc(doc(db, ...CONFIG_DOC));
  return snap.exists() ? snap.data() : null;
}

export async function saveConfig(db, cfg) {
  await setDoc(doc(db, ...CONFIG_DOC), cfg, { merge: true });
}

// The baseline is the memory of what both sides looked like after the last
// successful sync. Without it there is no way to tell an edit from a stale
// value, so it is stored server-side rather than in browser storage — losing it
// to a cleared cache would silently downgrade every future sync to guesswork.
export async function loadBaseline(db) {
  const snap = await getDocs(collection(db, "syncBaseline"));
  const map = new Map();
  const chunks = [];
  snap.forEach((d) => { if (d.id.startsWith(BASELINE_PREFIX)) chunks.push(d.data()); });
  for (const c of chunks) {
    let parsed;
    try { parsed = JSON.parse(c.json || "{}"); }
    catch { continue; } // a corrupt chunk degrades to "no baseline", never to bad data
    for (const [id, rec] of Object.entries(parsed)) map.set(id, rec);
  }
  return map;
}

export async function saveBaseline(db, baseline) {
  const entries = [...baseline.entries()];
  const batch = writeBatch(db);

  const existing = await getDocs(collection(db, "syncBaseline"));
  const keep = new Set();
  for (let i = 0; i < entries.length; i += CHUNK_ITEMS) {
    const slice = entries.slice(i, i + CHUNK_ITEMS);
    const obj = {};
    for (const [id, rec] of slice) obj[id] = rec;
    const key = `${BASELINE_PREFIX}${String(i / CHUNK_ITEMS).padStart(4, "0")}`;
    keep.add(key);
    batch.set(doc(db, "syncBaseline", key), { json: JSON.stringify(obj), count: slice.length });
  }
  // Remove chunks left over from a previously larger catalogue.
  existing.forEach((d) => { if (!keep.has(d.id)) batch.delete(doc(db, "syncBaseline", d.id)); });
  await batch.commit();
}

// ----------------------------------------------------------------------------
// Running a sync
// ----------------------------------------------------------------------------
function itemsToMap(items) {
  const m = new Map();
  for (const it of items) m.set(it.id, normalizeRecord(it));
  return m;
}

export async function connectGoogle(db, clientId, { interactive = true } = {}) {
  const token = await getAccessToken(clientId, { interactive });
  let cfg = await loadConfig(db);
  if (!cfg || !cfg.spreadsheetId) {
    const created = await createSpreadsheet(token);
    cfg = { spreadsheetId: created.spreadsheetId, sheetId: created.sheetId, url: created.url };
    await saveConfig(db, cfg);
  } else if (cfg.sheetId === undefined) {
    cfg.sheetId = await getSheetId(token, cfg.spreadsheetId);
    await saveConfig(db, cfg);
  }
  return cfg;
}

// Builds the plan. Reads only — nothing is written until applyPlan runs.
export async function planSync(db, clientId, items) {
  const cfg = await connectGoogle(db, clientId, { interactive: false });
  const token = await getAccessToken(clientId);
  const { records, rowOf, colOf, issues } = await readSheet(token, cfg.spreadsheetId);

  // A structurally broken sheet aborts before any comparison. Syncing against a
  // sheet whose columns were renamed would read every field as blank and look
  // exactly like "the user cleared the whole catalogue".
  if (issues.length) return { blocked: issues, cfg };

  const baseline = await loadBaseline(db);
  const local = itemsToMap(items);
  const plan = mergeAll(baseline, local, records);

  currentPlan = plan;
  currentContext = { cfg, rowOf, colOf, baseline, token };
  for (const k of Object.keys(conflictChoices)) delete conflictChoices[k];

  return { plan, summary: planSummary(plan), cfg };
}

// Applies an already-reviewed plan. `confirmedDeletions` is the set of item ids
// the user explicitly agreed to delete because their sheet row is missing.
export async function applyPlan(db, clientId, { confirmedDeletions = new Set(), onProgress } = {}) {
  if (!currentPlan || !currentContext) throw new Error("Nothing to apply — run a sync first");
  const { cfg, rowOf, colOf, baseline } = currentContext;
  const token = await getAccessToken(clientId);
  const resolved = resolveConflicts(currentPlan, conflictChoices);
  const step = (msg) => { if (onProgress) onProgress(msg); };

  // --- Firestore side ---
  step("Applying sheet changes…");
  const writes = [];
  for (const { id, fields } of resolved.toFirestore) writes.push({ id, fields });
  for (const { id, record } of resolved.createInFirestore) {
    writes.push({ id, fields: { ...record, id, image: null, hasPhoto: false }, create: true });
  }
  for (let i = 0; i < writes.length; i += 400) {
    const batch = writeBatch(db);
    for (const w of writes.slice(i, i + 400)) {
      batch.set(doc(db, "items", w.id), w.fields, { merge: true });
    }
    await batch.commit();
  }

  // Deletions only ever happen for ids the user ticked.
  const toDelete = resolved.pendingSheetDeletions.filter((d) => confirmedDeletions.has(d.id));
  if (toDelete.length) {
    step(`Deleting ${toDelete.length} item(s)…`);
    for (let i = 0; i < toDelete.length; i += 400) {
      const batch = writeBatch(db);
      for (const d of toDelete.slice(i, i + 400)) batch.delete(doc(db, "items", d.id));
      await batch.commit();
    }
    for (const d of toDelete) { try { await deleteDoc(doc(db, "itemPhotos", d.id)); } catch { /* already gone */ } }
  }

  // --- Sheet side ---
  step("Updating the sheet…");
  await writeCells(token, cfg.spreadsheetId, resolved.toSheet, rowOf, colOf);
  await appendRows(token, cfg.spreadsheetId, resolved.createInSheet);
  if (resolved.deleteInSheet.length) {
    const sheetId = cfg.sheetId ?? await getSheetId(token, cfg.spreadsheetId);
    await deleteRows(token, cfg.spreadsheetId, sheetId, resolved.deleteInSheet, rowOf);
  }

  // --- Baseline last ---
  // Only after both sides succeeded. If anything above threw, the baseline stays
  // where it was and the next sync simply re-proposes the same work, rather than
  // believing changes landed when they didn't.
  step("Saving sync state…");
  const nb = nextBaseline(resolved, baseline, confirmedDeletions);
  await saveBaseline(db, nb);

  const summary = planSummary(resolved);
  currentPlan = null;
  currentContext = null;
  return { ...summary, deleted: toDelete.length };
}

export function setConflictChoice(id, field, pick) { conflictChoices[`${id}::${field}`] = pick; }
export function getConflictChoices() { return { ...conflictChoices }; }
export function unresolvedCount() {
  if (!currentPlan) return 0;
  return currentPlan.conflicts.filter((c) => !conflictChoices[`${c.id}::${c.field}`]).length;
}
export function getCurrentPlan() { return currentPlan; }
export function disconnect() { forgetToken(); currentPlan = null; currentContext = null; }

// ----------------------------------------------------------------------------
// Rendering helpers
// ----------------------------------------------------------------------------
export function renderPlanHtml(plan, esc) {
  const s = planSummary(plan);
  const line = (n, label, cls) =>
    `<div class="sync-line ${cls || ""}"><span class="sync-num">${n}</span><span>${label}</span></div>`;

  let html = `<div class="sync-summary sans">`;
  html += line(s.fromSheet, "change(s) coming from the sheet");
  html += line(s.fromApp, "change(s) going to the sheet");
  if (s.conflicts) html += line(s.conflicts, "conflict(s) need your decision", "warn");
  if (s.needsDeleteConfirm) html += line(s.needsDeleteConfirm, "item(s) missing from the sheet", "warn");
  html += line(s.unchanged, "unchanged");
  html += `</div>`;

  if (plan.conflicts.length) {
    html += `<div class="sync-section-title sans">Conflicts — both sides changed the same field</div>`;
    for (const c of plan.conflicts) {
      const key = `${c.id}::${c.field}`;
      html += `
        <div class="conflict-row sans" data-conflict="${esc(key)}">
          <div class="conflict-head"><strong>${esc(c.sku)}</strong> · ${esc(FIELD_HEADERS[c.field] || c.field)}</div>
          <div class="conflict-was">was: ${esc(String(c.base ?? "—"))}</div>
          <div class="conflict-options">
            <button class="conflict-opt" data-pick="app" data-key="${esc(key)}">App<br><b>${esc(String(c.local))}</b></button>
            <button class="conflict-opt" data-pick="sheet" data-key="${esc(key)}">Sheet<br><b>${esc(String(c.remote))}</b></button>
          </div>
        </div>`;
    }
  }

  if (plan.pendingSheetDeletions.length) {
    html += `<div class="sync-section-title sans">Missing from the sheet — tick to delete, otherwise they're kept</div>`;
    for (const d of plan.pendingSheetDeletions) {
      html += `
        <label class="delete-row sans">
          <input type="checkbox" class="del-check" data-id="${esc(d.id)}" />
          <span>${esc(d.record.sku)} — ${esc(d.record.name)}</span>
        </label>`;
    }
  }
  return html;
}
