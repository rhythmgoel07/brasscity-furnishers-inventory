// ============================================================================
// Three-way merge engine for two-way Google Sheets sync.
//
// This file is deliberately PURE: no Firestore, no Google APIs, no DOM. It takes
// three plain snapshots and returns a plan. That makes it testable, and this is
// the one piece that absolutely must be correct — everything else is plumbing.
//
// WHY THREE-WAY:
// A naive two-way sync compares the app against the sheet and picks a winner,
// usually "most recently changed". That cannot tell the difference between
// "the sheet has 40 because someone typed 40" and "the sheet has 40 because
// that's the old value nobody touched". So it happily overwrites a real edit
// with a stale one, silently.
//
// Keeping a BASELINE — the exact state at the end of the last successful sync —
// removes the guesswork. If a value equals the baseline, that side did not
// change. Only when BOTH sides moved away from the baseline, to different
// values, is there a genuine conflict. Those get surfaced, never guessed.
//
// Merging is per FIELD, not per row: changing price in the sheet and stock in
// the app on the same item is not a conflict, and shouldn't be treated as one.
// ============================================================================

// Columns that participate in sync. Order here is the column order in the sheet.
export const SYNC_FIELDS = [
  "sku", "name", "tags", "material", "supplier",
  "warehouse", "cost", "price", "stock", "reorder",
];

export const NUMERIC_FIELDS = new Set(["cost", "price", "stock", "reorder"]);

// Human-readable sheet headers, matched by NAME on read so that reordering or
// inserting columns in the sheet can't silently shift data into the wrong field.
export const FIELD_HEADERS = {
  sku: "SKU",
  name: "Item Name",
  tags: "Tags",
  material: "Material",
  supplier: "Supplier",
  warehouse: "Warehouse/Location",
  cost: "Cost Price",
  price: "Selling Price",
  stock: "Stock Qty",
  reorder: "Reorder Level",
};

// Every row also carries a hidden stable ID column. Without it, editing a SKU in
// the sheet looks exactly like "deleted one item, added a different one", which
// would drop the photo and history attached to the original.
export const ID_HEADER = "_id";

// Normalises a value so that "5" from a sheet cell and 5 from Firestore compare
// equal. Sheets returns everything as strings unless the cell is typed, and an
// untouched empty cell can come back as "", null, or be missing entirely.
export function normalize(field, value) {
  // Tags are one field holding a set. Sorting and lowercasing here means the
  // same set always compares equal regardless of the order it was typed —
  // otherwise every sync would report phantom changes on untouched rows.
  if (field === "tags") {
    if (value === null || value === undefined) return "";
    return [...new Set(String(value).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean))].sort().join(", ");
  }
  if (NUMERIC_FIELDS.has(field)) {
    if (value === null || value === undefined || value === "") return 0;
    // Strips currency symbols, thousands separators and stray spaces, so a cell
    // someone formatted as "₹12,500" still compares equal to 12500.
    const cleaned = String(value).replace(/[^0-9.\-]/g, "");
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

export function normalizeRecord(rec) {
  const out = {};
  for (const f of SYNC_FIELDS) out[f] = normalize(f, rec ? rec[f] : undefined);
  return out;
}

// ----------------------------------------------------------------------------
// mergeAll — the whole decision procedure.
//
//   baseline : Map<id, record>  state at the end of the last successful sync
//   local    : Map<id, record>  what Firestore holds right now
//   remote   : Map<id, record>  what the sheet holds right now
//
// Returns a PLAN. Nothing is applied here; the caller decides whether to run it,
// which is what makes a preview/confirmation step possible.
// ----------------------------------------------------------------------------
export function mergeAll(baseline, local, remote) {
  const plan = {
    toFirestore: [],      // {id, fields} — sheet edits to write into the app
    toSheet: [],          // {id, record} — app changes to write into the sheet
    createInFirestore: [],// {id, record} — rows added in the sheet
    createInSheet: [],    // {id, record} — items added in the app
    deleteInSheet: [],    // {id} — deleted in the app, safe to remove from sheet
    pendingSheetDeletions: [], // {id, record} — MISSING from sheet; needs confirmation
    conflicts: [],        // {id, sku, field, base, local, remote}
    agreed: [],           // {id, record} — already identical, just refresh baseline
  };

  const ids = new Set([...baseline.keys(), ...local.keys(), ...remote.keys()]);

  for (const id of ids) {
    const b = baseline.get(id);
    const l = local.get(id);
    const r = remote.get(id);

    // --- Creations -----------------------------------------------------------
    if (!b && l && !r) { plan.createInSheet.push({ id, record: normalizeRecord(l) }); continue; }
    if (!b && !l && r) { plan.createInFirestore.push({ id, record: normalizeRecord(r) }); continue; }

    // --- Deletions -----------------------------------------------------------
    if (b && !l && r) {
      // Deleted in the app. Removing the row from the sheet is the safe
      // direction: the app is where deletion is a deliberate, confirmed action.
      plan.deleteInSheet.push({ id });
      continue;
    }
    if (b && l && !r) {
      // Missing from the sheet. This is NOT auto-applied. A row can vanish from
      // a sheet by accident far too easily — a stray Ctrl+Z, a sort that left
      // rows behind, a filtered view saved wrong. Deleting real inventory on
      // that basis is unacceptable, so it's queued for explicit confirmation.
      plan.pendingSheetDeletions.push({ id, record: normalizeRecord(l) });
      continue;
    }
    if (!l && !r) continue; // gone from both sides; drop it from the baseline

    // --- Both sides present: merge field by field ----------------------------
    const ln = normalizeRecord(l);
    const rn = normalizeRecord(r);
    // No baseline but present on both sides: the same SKU was added independently
    // in each place. There's no "unchanged" reference, so any difference is a
    // genuine conflict rather than something we can silently resolve.
    const bn = b ? normalizeRecord(b) : null;

    const fieldsForFirestore = {};
    const fieldsForSheet = {};
    // The agreed end-state for this row, field by field. This is what makes the
    // sync converge: writing the raw local row back to the sheet would overwrite
    // the sheet's own newer values in columns the app didn't touch.
    const merged = {};
    let conflicted = false;

    for (const f of SYNC_FIELDS) {
      const lv = ln[f], rv = rn[f];
      if (lv === rv) { merged[f] = lv; continue; } // both sides agree already

      const bv = bn ? bn[f] : undefined;
      if (bn && lv === bv) {
        merged[f] = rv;
        fieldsForFirestore[f] = rv;      // only the sheet moved
      } else if (bn && rv === bv) {
        merged[f] = lv;
        fieldsForSheet[f] = lv;          // only the app moved
      } else {
        merged[f] = lv;                  // provisional; nothing is written yet
        conflicted = true;
        plan.conflicts.push({ id, sku: ln.sku || rn.sku || id, field: f, base: bv, local: lv, remote: rv });
      }
    }

    // Writes carry an explicit field list. The sheet writer updates only those
    // cells, so a row that has one conflicted column can still safely receive an
    // update to a different column.
    if (Object.keys(fieldsForFirestore).length) {
      plan.toFirestore.push({ id, fields: fieldsForFirestore, record: merged });
    }
    if (Object.keys(fieldsForSheet).length) {
      plan.toSheet.push({ id, fields: fieldsForSheet, record: merged });
    }
    if (!conflicted && !Object.keys(fieldsForFirestore).length && !Object.keys(fieldsForSheet).length) {
      plan.agreed.push({ id, record: merged });
    }
  }

  return plan;
}

// Applies the user's conflict choices ("app" or "sheet" per conflict) on top of
// a plan, folding them into the normal write lists.
export function resolveConflicts(plan, choices) {
  const byId = new Map();
  for (const c of plan.conflicts) {
    const pick = choices[`${c.id}::${c.field}`];
    if (pick !== "app" && pick !== "sheet") continue; // unresolved: leave it alone
    if (!byId.has(c.id)) byId.set(c.id, { toFirestore: {}, toSheet: {} });
    const entry = byId.get(c.id);
    if (pick === "sheet") entry.toFirestore[c.field] = c.remote;
    else entry.toSheet[c.field] = c.local;
  }

  const out = {
    ...plan,
    toFirestore: plan.toFirestore.map((x) => ({ ...x, fields: { ...x.fields }, record: { ...x.record } })),
    toSheet: plan.toSheet.map((x) => ({ ...x, fields: { ...x.fields }, record: { ...x.record } })),
    conflicts: plan.conflicts.filter((c) => {
      const pick = choices[`${c.id}::${c.field}`];
      return pick !== "app" && pick !== "sheet";
    }),
  };

  for (const [id, entry] of byId) {
    // A resolution also settles the merged end-state for that field, so `record`
    // has to move with it — otherwise the baseline written afterwards would
    // still hold the losing value and the next sync would resurrect it.
    const settled = { ...entry.toFirestore, ...entry.toSheet };

    if (Object.keys(entry.toFirestore).length) {
      const existing = out.toFirestore.find((x) => x.id === id);
      if (existing) { Object.assign(existing.fields, entry.toFirestore); Object.assign(existing.record, settled); }
      else out.toFirestore.push({ id, fields: entry.toFirestore, record: settled });
    }
    if (Object.keys(entry.toSheet).length) {
      const existing = out.toSheet.find((x) => x.id === id);
      if (existing) { Object.assign(existing.fields, entry.toSheet); Object.assign(existing.record, settled); }
      else out.toSheet.push({ id, fields: entry.toSheet, record: settled });
    }
    // Keep the merged record on the opposite list coherent too.
    for (const list of [out.toFirestore, out.toSheet]) {
      const e = list.find((x) => x.id === id);
      if (e) Object.assign(e.record, settled);
    }
  }
  return out;
}

// Builds the baseline to persist after a sync actually succeeds. Rows still in
// conflict keep their OLD baseline: pretending they synced would make the next
// run think the losing side was never edited, which is exactly the silent
// overwrite this design exists to prevent.
export function nextBaseline(plan, prevBaseline, appliedIds) {
  const next = new Map();
  const stillConflicted = new Set(plan.conflicts.map((c) => c.id));

  const put = (id, record) => {
    if (stillConflicted.has(id)) {
      const old = prevBaseline.get(id);
      if (old) next.set(id, normalizeRecord(old));
      return;
    }
    next.set(id, normalizeRecord(record));
  };

  for (const { id, record } of plan.agreed) put(id, record);
  for (const { id, record } of plan.createInSheet) put(id, record);
  for (const { id, record } of plan.createInFirestore) put(id, record);
  for (const { id, record } of plan.toSheet) put(id, record);
  for (const { id, fields, record } of plan.toFirestore) {
    const base = prevBaseline.get(id) || {};
    put(id, { ...normalizeRecord(base), ...(record || {}), ...fields });
  }
  // An item whose ONLY change is an unresolved conflict appears in no write list
  // at all, so the loops above never touch it and it would silently fall out of
  // the baseline. Losing the baseline is the worst possible outcome here: the
  // next sync has no reference point, so it can no longer tell which side moved
  // and every remaining difference escalates to a conflict. Carry it forward.
  for (const c of plan.conflicts) {
    if (next.has(c.id)) continue;
    const old = prevBaseline.get(c.id);
    if (old) next.set(c.id, normalizeRecord(old));
  }
  // Items whose sheet row is missing but which were NOT confirmed for deletion
  // must stay in the baseline, otherwise the next sync re-reads them as brand
  // new rows and re-adds them to the sheet in a loop.
  for (const { id, record } of plan.pendingSheetDeletions) {
    if (!appliedIds || !appliedIds.has(id)) put(id, record);
  }
  return next;
}

export function planSummary(plan) {
  return {
    fromSheet: plan.toFirestore.length + plan.createInFirestore.length,
    fromApp: plan.toSheet.length + plan.createInSheet.length + plan.deleteInSheet.length,
    conflicts: plan.conflicts.length,
    needsDeleteConfirm: plan.pendingSheetDeletions.length,
    unchanged: plan.agreed.length,
  };
}
