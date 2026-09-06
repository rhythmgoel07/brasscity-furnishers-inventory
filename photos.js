// ============================================================================
// Photo storage, split out of the item documents.
//
// THE PROBLEM THIS SOLVES:
// Photos used to live inline in each item document as base64. Firestore's
// onSnapshot sends whole documents, so every device downloaded every photo on
// every cold start — roughly 10 MB for 500 items — and any change to any item
// re-sent that item's photo along with it. Staff on the shop floor were paying
// for photos they never scrolled to.
//
// THE FIX:
// Item documents keep only text plus `avgRgb` (three small numbers, needed by
// photo search so it can rank every item WITHOUT loading any photos). The image
// bytes move to a parallel `itemPhotos` collection, fetched one document at a
// time, only when a card actually scrolls into view.
//
// Cost for a 500-item catalogue on a cold start: about 300 KB instead of 10 MB.
// Photos then arrive lazily and are cached in memory and by Firestore's own
// offline cache, so scrolling back up costs nothing.
//
// Staying on Firestore rather than Cloud Storage is deliberate: since February
// 2026 Cloud Storage requires the Blaze plan and a linked card, and at this
// catalogue size the only thing it would add is CDN caching.
// ============================================================================

import {
  doc, getDoc, setDoc, deleteDoc, collection, getDocs, writeBatch,
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";

// Two sizes. The grid only ever needs the small one; the large one loads solely
// when a photo is tapped, which is what makes big photos affordable at all.
export const THUMB_DIM = 200;
export const FULL_DIM = 900;

const memCache = new Map();     // id -> {thumb, full}
const inflight = new Map();     // id -> Promise, so N cards for one id fetch once

export function photoDocRef(db, id) { return doc(db, "itemPhotos", id); }

export async function loadPhoto(db, id) {
  if (memCache.has(id)) return memCache.get(id);
  if (inflight.has(id)) return inflight.get(id);

  const p = (async () => {
    try {
      const snap = await getDoc(photoDocRef(db, id));
      const data = snap.exists() ? snap.data() : null;
      memCache.set(id, data);
      return data;
    } catch (e) {
      console.error("photo load failed", id, e);
      inflight.delete(id);  // let a later scroll retry rather than caching failure
      return null;
    } finally {
      inflight.delete(id);
    }
  })();

  inflight.set(id, p);
  return p;
}

export function cachedPhoto(id) { return memCache.get(id) || null; }

export async function savePhoto(db, id, { thumb, full, avgRgb }) {
  memCache.set(id, { thumb, full, avgRgb });
  await setDoc(photoDocRef(db, id), { thumb, full, avgRgb });
}

export async function deletePhoto(db, id) {
  memCache.delete(id);
  try { await deleteDoc(photoDocRef(db, id)); }
  catch (e) { console.error("photo delete failed", id, e); }
}

export function forgetPhoto(id) { memCache.delete(id); }

// ----------------------------------------------------------------------------
// Lazy loading. One shared IntersectionObserver for the whole grid rather than
// one per card — thousands of observers is itself a performance problem.
// ----------------------------------------------------------------------------
let observer = null;

export function initLazyPhotos(db) {
  if (observer) return observer;
  observer = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      const el = entry.target;
      observer.unobserve(el);
      hydrate(db, el);
    }
  }, {
    // Start fetching slightly before a card is visible so photos are usually
    // there by the time they scroll into frame.
    rootMargin: "300px 0px",
    threshold: 0.01,
  });
  return observer;
}

async function hydrate(db, el) {
  const id = el.dataset.photoId;
  if (!id) return;
  const photo = await loadPhoto(db, id);
  if (!photo || !photo.thumb) return;      // no photo: the colour placeholder stays
  if (!el.isConnected) return;             // card was re-rendered while we waited
  const img = document.createElement("img");
  img.className = "thumb";
  img.alt = "";
  img.decoding = "async";
  img.src = photo.thumb;
  img.style.cssText = el.style.cssText;
  el.replaceWith(img);
}

// Call after each grid render to pick up newly created placeholders.
export function observeAll(root = document) {
  if (!observer) return;
  root.querySelectorAll("[data-photo-id]:not([data-observed])").forEach((el) => {
    el.dataset.observed = "1";
    observer.observe(el);
  });
}

// ----------------------------------------------------------------------------
// Migration off the old inline format.
//
// Runs once, in small batches, and only clears the old `image` field AFTER the
// photo document has been written — so an interruption halfway through leaves
// duplicated data, never missing data. Re-running it is harmless.
// ----------------------------------------------------------------------------
export async function countLegacyPhotos(db) {
  const snap = await getDocs(collection(db, "items"));
  let n = 0;
  snap.forEach((d) => { const v = d.data(); if (typeof v.image === "string" && v.image.startsWith("data:")) n++; });
  return n;
}

export async function migrateLegacyPhotos(db, { onProgress, makeSizes } = {}) {
  const snap = await getDocs(collection(db, "items"));
  const legacy = [];
  snap.forEach((d) => {
    const v = d.data();
    if (typeof v.image === "string" && v.image.startsWith("data:")) legacy.push({ id: d.id, image: v.image, avgRgb: v.avgRgb });
  });

  let done = 0;
  for (const { id, image, avgRgb } of legacy) {
    try {
      // The old inline images were capped at 160px, so there's no larger version
      // to recover — the same image serves as both until the photo is retaken.
      const sizes = makeSizes ? await makeSizes(image) : { thumb: image, full: image };
      await savePhoto(db, id, { thumb: sizes.thumb, full: sizes.full, avgRgb: avgRgb || null });
      const batch = writeBatch(db);
      batch.update(doc(db, "items", id), { image: null });
      await batch.commit();
    } catch (e) {
      console.error("migration failed for", id, e);
    }
    done++;
    if (onProgress) onProgress(done, legacy.length);
  }
  return { migrated: done, total: legacy.length };
}
