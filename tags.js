// ============================================================================
// Item tags — replaces the free-text Category field.
//
// Category was typed by hand on every item, which is how "Dining", "Dining Set"
// and "BistroSet" ended up as three separate, half-empty filters. Tags fix that
// at the source: a tag exists only if a manager defined it, and everywhere else
// in the app a tag can only be SELECTED, never invented.
//
// Tags are single words, letters and digits only, stored lowercase. That one
// constraint buys a lot:
//   • a comma can never appear inside a name, so the Excel column is safely
//     comma-separated
//   • no case ambiguity — SOFA, Sofa and sofa are one tag
//   • parsing can split on spaces as well as commas, so a hand-typed cell like
//     "dining tables" still yields two correct tags
// The UI capitalises them for display; the stored value stays lowercase.
// ============================================================================

import {
  doc, setDoc, deleteDoc, collection, getDocs, writeBatch,
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";

export const TAG_PATTERN = /^[a-z0-9]+$/;
export const MAX_TAG_LENGTH = 24;

// Returns an error string, or null when the name is acceptable.
export function validateTagName(raw) {
  const name = String(raw || "").trim().toLowerCase();
  if (!name) return "Enter a tag name.";
  if (name.length > MAX_TAG_LENGTH) return `Keep tags under ${MAX_TAG_LENGTH} characters.`;
  if (/\s/.test(name)) return "Tags must be a single word — no spaces.";
  if (!TAG_PATTERN.test(name)) return "Letters and digits only — no punctuation or symbols.";
  return null;
}

export function normalizeTagName(raw) { return String(raw || "").trim().toLowerCase(); }

export function dedupeSort(tags) {
  return [...new Set((tags || []).map(normalizeTagName).filter(Boolean))].sort();
}

// Splits a free-typed cell into tag names. Deliberately permissive about the
// separator: since a tag can never contain a space, "dining, tables",
// "dining tables" and "Dining;Tables" all mean the same thing.
export function parseTagCell(cell) {
  if (cell === null || cell === undefined) return [];
  return dedupeSort(String(cell).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean));
}

// The canonical written form: deduplicated and sorted, so the same set always
// serialises identically. Without the sort, reordering tags would look like a
// real edit to the Sheets sync and raise phantom conflicts on every run.
export function formatTagCell(tags) { return dedupeSort(tags).join(", "); }

export function splitKnownUnknown(tags, registryNames) {
  const known = [], unknown = [];
  for (const t of dedupeSort(tags)) (registryNames.has(t) ? known : unknown).push(t);
  return { known, unknown };
}

// AND matching: an item must carry EVERY selected tag, so each chip narrows.
export function itemMatchesTags(itemTags, selected) {
  if (!selected || selected.length === 0) return true;
  const have = new Set((itemTags || []).map(normalizeTagName));
  return selected.every((t) => have.has(normalizeTagName(t)));
}

// Suggestions for the picker. Already-chosen tags are excluded so the dropdown
// never offers something that's already a chip. Prefix matches rank above
// substring matches, so typing "so" puts "sofa" above "cushions".
export function suggestTags(registry, queryStr, chosen) {
  const q = normalizeTagName(queryStr);
  const taken = new Set((chosen || []).map(normalizeTagName));
  const pool = registry.filter((t) => !taken.has(t.name));
  if (!q) return pool.slice(0, 40);
  const starts = [], contains = [];
  for (const t of pool) {
    if (t.name.startsWith(q)) starts.push(t);
    else if (t.name.includes(q)) contains.push(t);
  }
  return [...starts, ...contains].slice(0, 40);
}

export function usageCounts(items, registry) {
  const counts = new Map(registry.map((t) => [t.name, 0]));
  for (const it of items || []) {
    for (const t of it.tags || []) counts.set(t, (counts.get(t) || 0) + 1);
  }
  return counts;
}

// ----------------------------------------------------------------------------
// Registry persistence. One document per tag, id === name, so a tag cannot be
// duplicated by casing and the id stays readable.
// ----------------------------------------------------------------------------
export function tagDocRef(db, name) { return doc(db, "tags", normalizeTagName(name)); }

export async function loadTagRegistry(db) {
  const snap = await getDocs(collection(db, "tags"));
  const list = [];
  snap.forEach((d) => list.push({ name: d.id, ...d.data() }));
  return list.sort((a, b) => a.name.localeCompare(b.name));
}

export async function createTag(db, rawName) {
  const name = normalizeTagName(rawName);
  const err = validateTagName(name);
  if (err) throw new Error(err);
  await setDoc(tagDocRef(db, name), { name });
  return name;
}

// Deleting a tag also strips it from every item. Leaving it behind would create
// ghost tags that filter to nothing and can't be cleared from the UI.
export async function deleteTag(db, name, items) {
  const tag = normalizeTagName(name);
  const affected = (items || []).filter((it) => (it.tags || []).includes(tag));
  for (let i = 0; i < affected.length; i += 400) {
    const batch = writeBatch(db);
    for (const it of affected.slice(i, i + 400)) {
      batch.update(doc(db, "items", it.id), { tags: (it.tags || []).filter((t) => t !== tag) });
    }
    await batch.commit();
  }
  await deleteDoc(tagDocRef(db, tag));
  return affected.length;
}

// Merges one tag into another — the repair for a catalogue that already drifted,
// e.g. folding "diningset" into "dining" across every item in one action.
export async function mergeTags(db, fromName, toName, items) {
  const from = normalizeTagName(fromName), to = normalizeTagName(toName);
  if (from === to) return 0;
  const affected = (items || []).filter((it) => (it.tags || []).includes(from));
  for (let i = 0; i < affected.length; i += 400) {
    const batch = writeBatch(db);
    for (const it of affected.slice(i, i + 400)) {
      batch.update(doc(db, "items", it.id), {
        tags: dedupeSort([...(it.tags || []).filter((t) => t !== from), to]),
      });
    }
    await batch.commit();
  }
  await deleteDoc(tagDocRef(db, from));
  return affected.length;
}

// ----------------------------------------------------------------------------
// One-time migration off Category.
//
// Every distinct category becomes a tag and each item's category becomes its
// tags, so nothing is retyped. Multi-word categories like "Dining Set" split
// into two valid single-word tags rather than being dropped.
// ----------------------------------------------------------------------------
export function categoriesToTags(items) {
  const tagSet = new Set();
  const perItem = new Map();
  for (const it of items || []) {
    const parsed = parseTagCell(it.category);
    if (!parsed.length) continue;
    perItem.set(it.id, parsed);
    parsed.forEach((t) => tagSet.add(t));
  }
  return { tags: [...tagSet].sort(), perItem };
}

export async function migrateCategories(db, items, onProgress) {
  const { tags, perItem } = categoriesToTags(items);

  const tagBatch = writeBatch(db);
  for (const t of tags) tagBatch.set(tagDocRef(db, t), { name: t });
  await tagBatch.commit();

  const entries = [...perItem.entries()];
  let done = 0;
  for (let i = 0; i < entries.length; i += 400) {
    const batch = writeBatch(db);
    for (const [id, tagList] of entries.slice(i, i + 400)) {
      // tags are written and category cleared in the SAME update, so an
      // interruption can never leave an item with neither.
      batch.update(doc(db, "items", id), { tags: tagList, category: null });
    }
    await batch.commit();
    done += Math.min(400, entries.length - i);
    if (onProgress) onProgress(done, entries.length);
  }
  return { tagsCreated: tags.length, itemsUpdated: entries.length };
}
