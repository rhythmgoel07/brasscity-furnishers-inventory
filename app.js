// ============================================================================
// Firebase setup. This is an ES module (see index.html: <script type="module">),
// which is what lets us use `import` directly in a plain browser with no build step.
// ============================================================================
import { firebaseConfig } from "./firebase-config.js";
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js";
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  collection, doc, getDoc, onSnapshot, setDoc, updateDoc, deleteDoc, increment, writeBatch,
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";
import {
  getAuth, signInWithEmailAndPassword, onAuthStateChanged, signOut,
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-auth.js";
import {
  THUMB_DIM, FULL_DIM, loadPhoto, cachedPhoto, savePhoto, deletePhoto, forgetPhoto,
  initLazyPhotos, observeAll, countLegacyPhotos, migrateLegacyPhotos,
} from "./photos.js";
import {
  planSync, applyPlan, renderPlanHtml, setConflictChoice, unresolvedCount,
  getCurrentPlan, disconnect as disconnectSheets,
} from "./sync-ui.js";

if (String(firebaseConfig.apiKey || "").includes("PASTE_YOUR")) {
  document.body.innerHTML = `<div style="font-family:system-ui,sans-serif;padding:40px 20px;max-width:480px;margin:0 auto;text-align:center;color:#2B2420;">
    <h2>Setup needed</h2>
    <p>Open <code>firebase-config.js</code> and paste in your own Firebase project's
    values — see README.md for exactly where to find them. The app will work normally
    once that's done.</p></div>`;
  throw new Error("firebase-config.js not yet configured");
}

const app = initializeApp(firebaseConfig);

// Offline cache. The old enableIndexedDbPersistence() is deprecated AND silently
// gave up whenever a second tab was open — so a manager with the app open on two
// tabs lost offline support without any warning. persistentMultipleTabManager
// shares one cache across every tab instead.
const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
});
const auth = getAuth(app);

// ============================================================================
// Lazy script loading. The four helper libraries (QR generator, QR reader, Excel
// reader, Excel writer) total roughly 2 MB. Loading them up front meant every
// staff member on a phone downloaded 2 MB just to look at stock levels. Now each
// one is fetched only the first time a feature that needs it is actually used.
// ============================================================================
const LIB = {
  qrcode: "https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js",
  jsqr: "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js",
  xlsx: "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
  exceljs: "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
};
const scriptPromises = new Map();
function loadScript(src) {
  if (scriptPromises.has(src)) return scriptPromises.get(src);
  const p = new Promise((resolve, reject) => {
    const el = document.createElement("script");
    el.src = src;
    el.async = true;
    el.onload = () => resolve();
    el.onerror = () => { scriptPromises.delete(src); reject(new Error("Failed to load " + src)); };
    document.head.appendChild(el);
  });
  scriptPromises.set(src, p);
  return p;
}

// ============================================================================
// Small shared helpers
// ============================================================================

// Every image operation used to share two <canvas> elements in the HTML. During
// an Excel import dozens of photos were resized at the same time, all drawing
// onto that one canvas — so photos ended up attached to the wrong items. Each
// call now gets its own throwaway canvas.
function scratchCanvas(w, h) {
  const c = document.createElement("canvas");
  c.width = w; c.height = h;
  return c;
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Could not decode image"));
    img.src = src;
  });
}

function fitDims(img, maxDim) {
  let w = img.width, h = img.height;
  if (w > h) { if (w > maxDim) { h = Math.round((h * maxDim) / w); w = maxDim; } }
  else if (h > maxDim) { w = Math.round((w * maxDim) / h); h = maxDim; }
  return { w, h };
}

function avgRgbOfImage(img) {
  const c = scratchCanvas(24, 24);
  const ctx = c.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(img, 0, 0, 24, 24);
  const data = ctx.getImageData(0, 0, 24, 24).data;
  let r = 0, g = 0, b = 0, n = 0;
  for (let i = 0; i < data.length; i += 4) { r += data[i]; g += data[i + 1]; b += data[i + 2]; n++; }
  return [Math.round(r / n), Math.round(g / n), Math.round(b / n)];
}

// Runs an async job over a list a few at a time instead of all at once. Importing
// 300 photos in parallel could exhaust memory on a phone and crash the tab.
async function mapLimit(list, limit, fn) {
  const out = new Array(list.length);
  let cursor = 0;
  const workers = Array.from({ length: Math.min(limit, list.length) }, async () => {
    while (cursor < list.length) {
      const i = cursor++;
      out[i] = await fn(list[i], i);
    }
  });
  await Promise.all(workers);
  return out;
}

// ============================================================================
// State & constants
// ============================================================================
let ITEMS = [];
let isManager = false;
let mode = "showroom";
let searchType = "text";
let query = "";
let catFilter = "All";
let editingId = null;
let formPhoto = null;
// Tracks WHERE the current photo came from. The old code used a single
// `unchanged` flag, which conflated two very different situations: "this photo is
// already safely in itemPhotos" and "this photo is still inline on the item and
// has never been migrated". Treating the second as unchanged meant saving an
// edited legacy item cleared its inline image without ever writing the photo
// document — destroying the photo. Hence an explicit origin.
//   "none"     no photo
//   "new"      user just picked one; must be written
//   "legacy"   still inline on the item doc; must be migrated on save
//   "existing" already in itemPhotos; leave it alone
//   "loading"  in itemPhotos but not fetched yet; leave it alone
let formPhotoOrigin = "none";
let formPhotoRemoved = false;
let confirmDeleteId = null;

const CATS = {
  Sofas: "#4A6C6F", Dining: "#B5502A", Bedroom: "#6B4A6B", Chairs: "#5B7048",
  Tables: "#B8863B", Storage: "#6B4226", Office: "#4A5A6B", Outdoor: "#C0713F",
};

function hashColor(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = str.charCodeAt(i) + ((h << 5) - h);
  return `hsl(${h % 360}, 35%, 40%)`;
}
function colorFor(cat) { return CATS[cat] || hashColor(cat || "Other"); }

function hslToRgb(h, s, l) {
  h = ((h % 360) + 360) % 360; s /= 100; l /= 100;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  const t = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x]
    : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return t.map((v) => Math.round((v + m) * 255));
}

// Previously every hsl() colour collapsed to one hard-coded grey, so all custom
// categories shared an identical fallback tone and photo search ranked them the same.
function hexToRgb(hex) {
  const hsl = /^hsl\(\s*(-?[\d.]+)\s*,\s*([\d.]+)%\s*,\s*([\d.]+)%\s*\)$/.exec(hex);
  if (hsl) return hslToRgb(parseFloat(hsl[1]), parseFloat(hsl[2]), parseFloat(hsl[3]));
  const h = hex.replace("#", "");
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}
function mixColor(hex1, hex2, t) {
  const a = hexToRgb(hex1), b = hexToRgb(hex2);
  return a.map((v, i) => Math.round(v * (1 - t) + b[i] * t));
}
function stockInfo(it) {
  const stock = Number(it.stock) || 0;
  const reorder = Number(it.reorder) || 0;
  if (stock <= 0) return { text: "OUT OF STOCK", color: "#A83232" };
  if (stock <= reorder) return { text: "LOW STOCK", color: "#B8863B" };
  return { text: "IN STOCK", color: "#5B7048" };
}
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function fmtINR(n) { return "₹" + Number(n || 0).toLocaleString("en-IN"); }
function ensureAvgRgb(it) {
  const ok = Array.isArray(it.avgRgb) && it.avgRgb.length === 3 && it.avgRgb.every((n) => Number.isFinite(n));
  if (!ok) it.avgRgb = mixColor(colorFor(it.category), "#FFFFFF", 0.55);
  return it.avgRgb;
}
function toast(msg) {
  const el = document.getElementById("toast");
  el.textContent = "✓ " + msg;
  el.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove("show"), 2400);
}

const SEED_ITEMS = [
  { id: "FRN-001", sku: "FRN-001", name: "Oakwood 3-Seater Sofa", category: "Sofas", material: "Oak & Linen", supplier: "WoodCraft Suppliers", warehouse: "Warehouse A", cost: 18000, price: 27500, stock: 12, reorder: 5, image: null },
  { id: "FRN-002", sku: "FRN-002", name: "Maple Dining Table (6-seat)", category: "Dining", material: "Maple Wood", supplier: "Timberline Furnishings", warehouse: "Warehouse A", cost: 15500, price: 23900, stock: 6, reorder: 4, image: null },
  { id: "FRN-003", sku: "FRN-003", name: "Elmwood Queen Bed Frame", category: "Bedroom", material: "Elm Wood", supplier: "WoodCraft Suppliers", warehouse: "Warehouse B", cost: 12000, price: 19500, stock: 3, reorder: 5, image: null },
  { id: "FRN-004", sku: "FRN-004", name: "Cushion Armchair - Grey", category: "Chairs", material: "Fabric & Pine", supplier: "ComfortSeating Co.", warehouse: "Warehouse A", cost: 4200, price: 7500, stock: 20, reorder: 8, image: null },
  { id: "FRN-005", sku: "FRN-005", name: "Glass Top Coffee Table", category: "Tables", material: "Glass & Steel", supplier: "Timberline Furnishings", warehouse: "Warehouse A", cost: 3200, price: 5800, stock: 15, reorder: 6, image: null },
  { id: "FRN-006", sku: "FRN-006", name: "Teakwood Bookshelf 5-Tier", category: "Storage", material: "Teak Wood", supplier: "WoodCraft Suppliers", warehouse: "Warehouse B", cost: 6800, price: 10900, stock: 2, reorder: 4, image: null },
  { id: "FRN-007", sku: "FRN-007", name: "Recliner Sofa - Leather", category: "Sofas", material: "Leather", supplier: "ComfortSeating Co.", warehouse: "Warehouse A", cost: 21000, price: 32000, stock: 4, reorder: 3, image: null },
  { id: "FRN-008", sku: "FRN-008", name: "Wardrobe 3-Door", category: "Bedroom", material: "MDF & Laminate", supplier: "Timberline Furnishings", warehouse: "Warehouse B", cost: 13500, price: 21000, stock: 7, reorder: 5, image: null },
  { id: "FRN-009", sku: "FRN-009", name: "Study Desk with Drawer", category: "Office", material: "Engineered Wood", supplier: "ComfortSeating Co.", warehouse: "Warehouse A", cost: 5200, price: 8900, stock: 0, reorder: 5, image: null },
  { id: "FRN-010", sku: "FRN-010", name: "Outdoor Patio Chair Set (2)", category: "Outdoor", material: "Rattan & Aluminium", supplier: "GreenScape Furniture", warehouse: "Warehouse C", cost: 7600, price: 12500, stock: 10, reorder: 6, image: null },
].map((it) => ({ ...it, avgRgb: mixColor(colorFor(it.category), "#FFFFFF", 0.55) }));

// ============================================================================
// Image helpers
// ============================================================================
function renderToDataUrl(img, maxDim, quality) {
  const { w, h } = fitDims(img, maxDim);
  const canvas = scratchCanvas(w, h);
  canvas.getContext("2d").drawImage(img, 0, 0, w, h);
  return canvas.toDataURL("image/jpeg", quality);
}

// Produces both sizes in one pass. Photos are no longer stuck at a 160px
// postage stamp: now that image bytes live outside the item documents, a 900px
// version costs nothing on the grid and is genuinely useful when a customer asks
// what a piece actually looks like.
// createObjectURL avoids base64-ing the original file into memory first, which
// matters when someone picks a 12 MP photo straight from a phone camera.
async function resizeImageFile(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url);
    return {
      thumb: renderToDataUrl(img, THUMB_DIM, 0.62),
      full: renderToDataUrl(img, FULL_DIM, 0.72),
      avgRgb: avgRgbOfImage(img),
    };
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function resizeDataUrl(dataUrl) {
  const img = await loadImage(dataUrl);
  return {
    thumb: renderToDataUrl(img, THUMB_DIM, 0.62),
    full: renderToDataUrl(img, FULL_DIM, 0.72),
    avgRgb: avgRgbOfImage(img),
  };
}

function thumbHtml(it, size) {
  size = size || 48;
  const fg = colorFor(it.category);
  const dim = `width:${size}px;height:${size}px`;

  // Already fetched this session: render it immediately, no flicker on re-render.
  const cached = cachedPhoto(it.id);
  if (cached && cached.thumb) {
    return `<img class="thumb" decoding="async" alt="" style="${dim}" src="${esc(cached.thumb)}" />`;
  }
  // Legacy inline photo that hasn't been migrated yet.
  if (it.image) {
    return `<img class="thumb" loading="lazy" decoding="async" alt="" style="${dim}" src="${esc(it.image)}" />`;
  }
  if (it.hasPhoto === false) {
    return `<div class="thumb-fallback" style="${dim};background:${fg}22"><div class="dot" style="background:${fg};width:${Math.round(size * 0.4)}px;height:${Math.round(size * 0.4)}px"></div></div>`;
  }
  // Placeholder tinted with the item's average colour, swapped for the real
  // photo when it scrolls into view. Using avgRgb means the placeholder already
  // roughly resembles the product instead of flashing grey.
  const rgb = ensureAvgRgb(it);
  return `<div class="thumb-fallback" data-photo-id="${esc(it.id)}" style="${dim};background:rgb(${rgb[0]},${rgb[1]},${rgb[2]});opacity:.35"></div>`;
}

// ============================================================================
// QR code generation
// ============================================================================
const qrCache = new Map();
async function getQRDataUrl(text, size) {
  size = size || 100;
  const cacheKey = text + ":" + size;
  if (qrCache.has(cacheKey)) return qrCache.get(cacheKey);
  await loadScript(LIB.qrcode);

  const holder = document.createElement("div");
  holder.style.cssText = "position:absolute;left:-9999px;top:0;";
  document.body.appendChild(holder);
  try {
    new QRCode(holder, { text: String(text), width: size, height: size, correctLevel: QRCode.CorrectLevel.M });
    // The library renders synchronously to <canvas> but asynchronously to <img>
    // on older browsers. Poll briefly instead of guessing a fixed 60 ms.
    for (let attempt = 0; attempt < 40; attempt++) {
      const canvas = holder.querySelector("canvas");
      if (canvas) { const u = canvas.toDataURL("image/png"); qrCache.set(cacheKey, u); return u; }
      const img = holder.querySelector("img");
      if (img && img.src && img.src.startsWith("data:")) { qrCache.set(cacheKey, img.src); return img.src; }
      await new Promise((r) => setTimeout(r, 15));
    }
    return null;
  } catch (e) {
    console.error(e);
    return null;
  } finally {
    holder.remove();
  }
}

// ============================================================================
// Firestore mutations. Every write requires isManager (enforced both in the UI,
// by hiding these controls, and for real by the security rules on the server —
// see firestore.rules) so someone can't just open dev tools and edit stock
// without signing in.
// ============================================================================
function itemDocRef(id) { return doc(db, "items", id); }

async function saveItemRemote(item) {
  try {
    await setDoc(itemDocRef(item.id), item);
    return true;
  } catch (e) {
    console.error(e);
    if (e.code === "permission-denied") {
      // Item writes and photo writes are governed by SEPARATE rules blocks, so a
      // setup where one works and the other doesn't is normal and worth naming
      // precisely — otherwise it reads as a login problem, which it isn't.
      toast(photoStep
        ? "Photo rejected — the itemPhotos rules block isn't published"
        : "Rejected by the database — your account isn't on the manager list");
    } else {
      toast("Couldn't save — check your connection");
    }
    return false;
  }
}

// Stock adjustment used to rewrite the ENTIRE document — including the base64
// photo — on every single tap of + or −. Tapping "+" ten times pushed roughly
// 200 KB up to Firestore and burned ten writes. Two problems fixed here:
//
//   1. Only the `stock` field is sent now, using Firestore's atomic increment.
//      Atomic matters with several staff on the floor at once: two people each
//      selling one sofa now correctly lands on −2, where the old read-modify-write
//      could lose one of them.
//   2. Taps are coalesced. Rapid taps merge into a single write ~600 ms after the
//      last one, so ten taps cost one write instead of ten.
const pendingStockDeltas = new Map(); // id -> net delta not yet written
let stockFlushTimer = null;

function flushStockWrites() {
  const work = Array.from(pendingStockDeltas.entries());
  pendingStockDeltas.clear();
  work.forEach(([id, delta]) => {
    if (delta === 0) return;
    updateDoc(itemDocRef(id), { stock: increment(delta) }).catch((e) => {
      console.error(e);
      toast(e.code === "permission-denied"
        ? "Rejected by the database — your account isn't on the manager list"
        : "Couldn't save stock change");
    });
  });
}

function adjustStock(id, delta) {
  const idx = ITEMS.findIndex((it) => it.id === id);
  if (idx === -1) return;
  const current = Number(ITEMS[idx].stock) || 0;
  const next = Math.max(0, current + delta);
  const applied = next - current;
  if (applied === 0) return; // already at zero, nothing to send

  ITEMS[idx] = { ...ITEMS[idx], stock: next };
  patchCardStock(ITEMS[idx]); // touch just this card, not the whole grid

  pendingStockDeltas.set(id, (pendingStockDeltas.get(id) || 0) + applied);
  clearTimeout(stockFlushTimer);
  stockFlushTimer = setTimeout(flushStockWrites, 600);
}

// Anything queued should still go out if the tab is closed or backgrounded.
window.addEventListener("pagehide", flushStockWrites);
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") flushStockWrites();
});

async function deleteItemLocal(id) {
  try {
    await deleteDoc(itemDocRef(id));
    // Deleting the item but leaving its photo behind would quietly accumulate
    // orphaned documents that nothing ever reads or cleans up.
    await deletePhoto(db, id);
    confirmDeleteId = null;
    toast("Item deleted");
    // ITEMS/render update arrives via the onSnapshot listener
  } catch (e) {
    console.error(e);
    toast(e.code === "permission-denied" ? "Sign in required to delete" : "Couldn't delete item");
  }
}

// ============================================================================
// Rendering
// ============================================================================
function render() {
  document.getElementById("pageTitle").textContent = mode === "showroom" ? "Browse Inventory" : "Manage Inventory";
  document.getElementById("btnShowroom").classList.toggle("active", mode === "showroom");
  document.getElementById("btnManager").classList.toggle("active", mode === "manager");
  document.getElementById("managerActions").classList.toggle("hidden", mode !== "manager");
  document.getElementById("statsBar").classList.toggle("hidden", mode !== "manager");
  document.getElementById("btnLoadSample").classList.toggle("hidden", !(mode === "manager" && ITEMS.length === 0));
  if (typeof refreshMigrateButton === "function") refreshMigrateButton();
  if (mode === "manager") renderStats();
  renderCategoryChips();
  renderGrid();
}

function renderStats() {
  let totalValue = 0, lowCount = 0;
  for (const it of ITEMS) {
    const stock = Number(it.stock) || 0;
    totalValue += (Number(it.cost) || 0) * stock;
    if (stock <= (Number(it.reorder) || 0)) lowCount++;
  }
  document.getElementById("statsBar").innerHTML = `
    <div class="stat-card"><div class="label">Stock value</div><div class="value">${fmtINR(totalValue)}</div></div>
    <div class="stat-card"><div class="label">Items</div><div class="value">${ITEMS.length}</div></div>
    <div class="stat-card" style="${lowCount ? "border-color:#B8863B" : ""}"><div class="label">Reorder</div><div class="value" style="${lowCount ? "color:#B8863B" : ""}">${lowCount}</div></div>
  `;
}

function renderCategoryChips() {
  const cats = new Set(Object.keys(CATS));
  ITEMS.forEach((it) => cats.add(it.category));
  const all = ["All", ...Array.from(cats)];
  document.getElementById("catChips").innerHTML = all.map((c) =>
    `<button class="chip ${catFilter === c ? "active" : ""}" data-cat="${esc(c)}">${esc(c)}</button>`
  ).join("");
}

function getFiltered() {
  const q = query.trim().toLowerCase();
  return ITEMS.filter((it) => {
    const matchQ = !q || [it.name, it.sku, it.supplier, it.material].some((v) => (v || "").toLowerCase().includes(q));
    const matchCat = catFilter === "All" || it.category === catFilter;
    return matchQ && matchCat;
  });
}

function cardHtml(it) {
  const s = stockInfo(it);
  const managerRow = mode === "manager" ? `
      <div class="manager-row sans">
        <div class="stock-adjust">
          <button class="icon-btn" data-action="dec" data-id="${esc(it.id)}" aria-label="Decrease stock">−</button>
          <span class="stock-count" style="font-size:13px;font-weight:700;min-width:18px;text-align:center;">${Number(it.stock) || 0}</span>
          <button class="icon-btn" data-action="inc" data-id="${esc(it.id)}" aria-label="Increase stock">+</button>
        </div>
        <div style="display:flex;gap:5px;">
          <button class="icon-btn" data-action="edit" data-id="${esc(it.id)}" aria-label="Edit item">✎</button>
          ${confirmDeleteId === it.id
            ? `<button class="icon-btn danger" data-action="delete-yes" data-id="${esc(it.id)}">Yes</button><button class="icon-btn" data-action="delete-no">No</button>`
            : `<button class="icon-btn" data-action="delete-ask" data-id="${esc(it.id)}" aria-label="Delete item">🗑</button>`}
        </div>
      </div>` : "";
  return `
      <div class="card" data-card-id="${esc(it.id)}">
        <div class="card-top">
          ${thumbHtml(it)}
          <div style="min-width:0;flex:1;">
            <div class="item-name">${esc(it.name)}</div>
            <div class="item-meta sans">${esc(it.sku)}</div>
          </div>
        </div>
        <div class="card-bottom sans">
          <span class="price">${fmtINR(it.price)}</span>
          <span class="badge" style="color:${s.color};background:${s.color}1A;">${s.text}</span>
        </div>
        ${managerRow}
      </div>`;
}

// A snapshot arrives every time ANY device touches ANY item, and the old code
// rebuilt the whole grid's innerHTML each time — throwing away and re-decoding
// every photo, killing scroll position and any in-progress tap. This signature
// lets us skip the rebuild entirely when nothing visible actually changed.
let lastGridSignature = null;
function gridSignature(list) {
  return mode + "|" + confirmDeleteId + "|" + list.map((it) =>
    `${it.id}~${it.stock}~${it.price}~${it.name}~${it.reorder}~${it.image ? it.image.length : 0}`
  ).join("|");
}

function renderGrid(force) {
  const list = getFiltered();
  const sig = gridSignature(list);
  if (!force && sig === lastGridSignature) return;
  lastGridSignature = sig;

  const grid = document.getElementById("grid");
  if (list.length === 0) {
    grid.innerHTML = `<div class="empty-state sans">No items match that search.</div>`;
    return;
  }
  grid.innerHTML = list.map(cardHtml).join("");
  observeAll(grid); // hand the new placeholders to the lazy-photo observer
}

// Updates one card in place after a +/− tap. Avoids a full grid rebuild, so the
// photo never flickers and the list doesn't jump under your thumb.
function patchCardStock(it) {
  const card = document.querySelector(`.card[data-card-id="${CSS.escape(it.id)}"]`);
  if (!card) { renderGrid(true); return; }
  const count = card.querySelector(".stock-count");
  if (count) count.textContent = Number(it.stock) || 0;
  const badge = card.querySelector(".badge");
  if (badge) {
    const s = stockInfo(it);
    badge.textContent = s.text;
    badge.style.color = s.color;
    badge.style.background = s.color + "1A";
  }
  // Re-sync the signature to the state we just painted. Our own write echoes back
  // through onSnapshot a moment later; without this the echo would look like a
  // change and trigger a full rebuild, undoing the point of patching in place.
  lastGridSignature = gridSignature(getFiltered());
}

// ============================================================================
// Photo search (color/tone matching)
// ============================================================================
function analyzeUploadedPhoto(dataUrl) {
  document.getElementById("uploadResultWrap").classList.remove("hidden");
  document.getElementById("uploadPreviewImg").src = dataUrl;
  document.getElementById("analysisStatus").textContent = "Analyzing photo…";
  document.getElementById("matchList").innerHTML = "";

  const img = new Image();
  img.onerror = () => { document.getElementById("analysisStatus").textContent = "Couldn't read that photo — try another."; };
  img.onload = () => {
    const [r, g, b] = avgRgbOfImage(img);
    const maxDist = Math.sqrt(255 * 255 * 3);
    const ranked = ITEMS.map((it) => {
      const c = ensureAvgRgb(it);
      const dist = Math.sqrt((r - c[0]) ** 2 + (g - c[1]) ** 2 + (b - c[2]) ** 2);
      return Object.assign({}, it, { similarity: Math.round((1 - dist / maxDist) * 100) });
    }).sort((a, b2) => b2.similarity - a.similarity);

    document.getElementById("analysisStatus").textContent = "Ranked by closest visual match";
    document.getElementById("matchList").innerHTML = ranked.slice(0, 6).map((it, idx) => {
      const s = stockInfo(it);
      const barColor = idx === 0 ? "#A8562E" : colorFor(it.category);
      return `
        <div class="match-row" style="${idx === 0 ? "border-color:#A8562E" : ""}">
          ${thumbHtml(it, 44)}
          <div style="flex:1;min-width:0;">
            <div style="font-size:13.5px;">${esc(it.name)}</div>
            <div class="sans" style="font-size:11px;color:var(--muted);margin-top:2px;">${esc(it.sku)} · ${fmtINR(it.price)} · <span style="color:${s.color};font-weight:700;">${s.text}</span></div>
            <div class="match-bar-track"><div class="match-bar-fill" style="width:${Math.max(it.similarity, 4)}%;background:${barColor};"></div></div>
          </div>
          <div class="sans" style="font-size:14px;font-weight:700;color:${idx === 0 ? "#A8562E" : "#2B2420"};min-width:40px;text-align:right;">${it.similarity}%</div>
        </div>`;
    }).join("");
  };
  img.src = dataUrl;
}

// ============================================================================
// QR scan
// ============================================================================
function handleQRSearchFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => decodeQRFromDataUrl(e.target.result);
  reader.readAsDataURL(file);
}

async function decodeQRFromDataUrl(dataUrl) {
  const wrap = document.getElementById("qrResultWrap");
  wrap.classList.remove("hidden");
  wrap.innerHTML = `<div class="qr-not-found sans">Reading QR code…</div>`;

  let img;
  try {
    await loadScript(LIB.jsqr);
    img = await loadImage(dataUrl);
  } catch (e) {
    console.error(e);
    wrap.innerHTML = `<div class="qr-not-found sans">Couldn't read that image — check your connection and try again.</div>`;
    return;
  }

  {
    const { w, h } = fitDims(img, 900);
    const canvas = scratchCanvas(w, h);
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(img, 0, 0, w, h);
    const imageData = ctx.getImageData(0, 0, w, h);
    const result = jsQR(imageData.data, w, h);

    if (!result) {
      wrap.innerHTML = `
        <div class="qr-not-found sans">
          <div style="font-size:24px;margin-bottom:8px;">🔍</div>
          <div style="font-weight:600;">No QR code found in that photo</div>
          <div style="font-size:12px;margin-top:4px;">Try again with better lighting and the tag filling more of the frame.</div>
          <button id="btnQRRetry" class="btn" style="margin:14px auto 0;">↺ Try again</button>
        </div>`;
      document.getElementById("btnQRRetry").onclick = resetQRSearch;
      return;
    }

    const decoded = (result.data || "").trim();
    const needle = decoded.toLowerCase();
    // String(...) guards against a numeric SKU coming back from Excel, which
    // previously threw "it.sku.toLowerCase is not a function" and left the panel stuck.
    const found = ITEMS.find((it) =>
      String(it.sku || "").toLowerCase() === needle || String(it.id || "").toLowerCase() === needle);

    if (!found) {
      wrap.innerHTML = `
        <div class="qr-not-found sans">
          <div style="font-size:24px;margin-bottom:8px;">❓</div>
          <div style="font-weight:600;">Scanned "${esc(decoded)}" but no matching item</div>
          <button id="btnQRRetry" class="btn" style="margin:14px auto 0;">↺ Scan another</button>
        </div>`;
      document.getElementById("btnQRRetry").onclick = resetQRSearch;
      return;
    }

    const s = stockInfo(found);
    wrap.innerHTML = `
      <div class="qr-found-card">
        ${thumbHtml(found, 60)}
        <div style="flex:1;min-width:0;" class="sans">
          <div style="font-size:14.5px;font-family:Georgia,serif;color:var(--ink);">${esc(found.name)}</div>
          <div style="font-size:11.5px;color:var(--muted);margin-top:3px;">${esc(found.sku)} · ${esc(found.category)}</div>
          <div style="display:flex;align-items:center;gap:10px;margin-top:7px;">
            <span style="font-size:15px;font-weight:700;">${fmtINR(found.price)}</span>
            <span class="badge" style="color:${s.color};background:${s.color}1A;">${s.text}</span>
          </div>
        </div>
      </div>
      <button id="btnQRRetry" class="btn sans" style="margin-top:12px;">↺ Scan another</button>`;
    document.getElementById("btnQRRetry").onclick = resetQRSearch;
  }
}

function resetQRSearch() {
  document.getElementById("qrResultWrap").classList.add("hidden");
  document.getElementById("qrResultWrap").innerHTML = "";
  document.getElementById("qrSearchInput").value = "";
}

// ============================================================================
// Print QR tags
// ============================================================================
async function printQRTags() {
  const list = getFiltered();
  if (list.length === 0) { toast("No items to print"); return; }
  toast(`Generating ${list.length} tag${list.length === 1 ? "" : "s"}…`);
  // Was Promise.all over every item, which spawned one hidden DOM node per item
  // simultaneously — visibly locked up the page on a large catalogue.
  const withQR = await mapLimit(list, 6, async (it) => ({ it, qr: await getQRDataUrl(it.sku, 160) }));
  const tagsHtml = withQR.map(({ it, qr }) => `
    <div class="tag">
      <img src="${qr || ""}" />
      <div class="tag-name">${esc(it.name)}</div>
      <div class="tag-sku">${esc(it.sku)}</div>
      <div class="tag-price">${fmtINR(it.price)}</div>
    </div>`).join("");
  const win = window.open("", "_blank");
  if (!win) { toast("Allow pop-ups to print tags"); return; }
  win.document.write(`
    <!DOCTYPE html><html><head><title>QR Tags — Print</title>
    <style>
      body { font-family: system-ui, sans-serif; margin: 20px; }
      .tags { display: flex; flex-wrap: wrap; gap: 14px; }
      .tag { width: 150px; border: 1px solid #ccc; border-radius: 8px; padding: 10px; text-align: center; page-break-inside: avoid; }
      .tag img { width: 110px; height: 110px; }
      .tag-name { font-size: 12px; font-weight: 600; margin-top: 6px; min-height: 30px; }
      .tag-sku { font-size: 11px; color: #666; }
      .tag-price { font-size: 13px; font-weight: 700; margin-top: 4px; }
      @media print { body { margin: 0; } }
    </style></head>
    <body><div class="tags">${tagsHtml}</div>
      <script>window.onload = () => setTimeout(() => window.print(), 300);</script>
    </body></html>`);
  win.document.close();
}

// ============================================================================
// Add / edit modal
// ============================================================================
function openAddModal() {
  editingId = null;
  formPhoto = null;
  formPhotoOrigin = "none";
  formPhotoRemoved = false;
  document.getElementById("btnRemovePhoto").classList.add("hidden");
  document.getElementById("modalTitle").textContent = "Add item";
  document.getElementById("btnSubmitForm").textContent = "Add item";
  ["sku", "name", "category", "material", "supplier", "warehouse", "cost", "price", "stock", "reorder"].forEach((k) => {
    document.getElementById("f_" + k).value = "";
  });
  document.getElementById("formPhotoPreview").classList.add("hidden");
  document.getElementById("formPhotoPlaceholder").classList.remove("hidden");
  document.getElementById("photoPickerLabel").textContent = "Add photo";
  document.getElementById("formModal").classList.remove("hidden");
}

function openEditModal(it) {
  editingId = it.id;
  formPhotoRemoved = false;
  const existing = cachedPhoto(it.id);
  if (existing) {
    formPhoto = { ...existing };
    formPhotoOrigin = "existing";
  } else if (it.image) {
    // Legacy inline photo. It is NOT yet in itemPhotos, so saving must write it.
    formPhoto = { thumb: it.image, full: it.image, avgRgb: it.avgRgb };
    formPhotoOrigin = "legacy";
  } else if (it.hasPhoto) {
    // Exists in itemPhotos but hasn't loaded yet. Marking this "loading" rather
    // than "none" is what stops a quick save from deleting a photo we simply
    // hadn't finished fetching.
    formPhoto = null;
    formPhotoOrigin = "loading";
  } else {
    formPhoto = null;
    formPhotoOrigin = "none";
  }
  document.getElementById("btnRemovePhoto").classList.toggle("hidden", formPhotoOrigin === "none");

  if (formPhotoOrigin === "loading") {
    loadPhoto(db, it.id).then((ph) => {
      if (!ph || editingId !== it.id || formPhotoRemoved || formPhotoOrigin === "new") return;
      formPhoto = { ...ph };
      formPhotoOrigin = "existing";
      const prev = document.getElementById("formPhotoPreview");
      prev.src = ph.thumb;
      prev.classList.remove("hidden");
      document.getElementById("formPhotoPlaceholder").classList.add("hidden");
      document.getElementById("photoPickerLabel").textContent = "Change photo";
    });
  }
  document.getElementById("modalTitle").textContent = "Edit item";
  document.getElementById("btnSubmitForm").textContent = "Save changes";
  document.getElementById("f_sku").value = it.sku;
  document.getElementById("f_name").value = it.name;
  document.getElementById("f_category").value = it.category;
  document.getElementById("f_material").value = it.material;
  document.getElementById("f_supplier").value = it.supplier;
  document.getElementById("f_warehouse").value = it.warehouse;
  document.getElementById("f_cost").value = it.cost;
  document.getElementById("f_price").value = it.price;
  document.getElementById("f_stock").value = it.stock;
  document.getElementById("f_reorder").value = it.reorder;
  if (it.image) {
    document.getElementById("formPhotoPreview").src = it.image;
    document.getElementById("formPhotoPreview").classList.remove("hidden");
    document.getElementById("formPhotoPlaceholder").classList.add("hidden");
    document.getElementById("photoPickerLabel").textContent = "Change photo";
  } else {
    document.getElementById("formPhotoPreview").classList.add("hidden");
    document.getElementById("formPhotoPlaceholder").classList.remove("hidden");
    document.getElementById("photoPickerLabel").textContent = "Add photo";
  }
  document.getElementById("formModal").classList.remove("hidden");
}

async function submitForm() {
  const sku = document.getElementById("f_sku").value.trim();
  const name = document.getElementById("f_name").value.trim();
  if (!sku || !name) { toast("SKU and item name are required"); return; }
  if (!editingId && ITEMS.some((it) => it.sku.toLowerCase() === sku.toLowerCase())) {
    toast("An item with that SKU already exists");
    return;
  }
  if (editingId) {
    const clash = ITEMS.find((it) => it.id !== editingId && it.sku.toLowerCase() === sku.toLowerCase());
    if (clash) { toast("Another item already uses that SKU"); return; }
  }
  const category = document.getElementById("f_category").value.trim() || "Other";
  // The document ID is the SKU. Editing an item's SKU used to change only the
  // `sku` field, leaving the ID as the old value — so IDs and SKUs silently drifted
  // apart and printed tags stopped matching. A rename now moves the document.
  const renamedFrom = editingId && editingId !== sku ? editingId : null;
  const item = {
    id: sku,
    sku, name, category,
    material: document.getElementById("f_material").value.trim(),
    supplier: document.getElementById("f_supplier").value.trim(),
    warehouse: document.getElementById("f_warehouse").value.trim(),
    cost: Number(document.getElementById("f_cost").value) || 0,
    price: Number(document.getElementById("f_price").value) || 0,
    stock: Math.max(0, Number(document.getElementById("f_stock").value) || 0),
    reorder: Math.max(0, Number(document.getElementById("f_reorder").value) || 0),
    // Image bytes now live in the itemPhotos collection. The item document keeps
    // only avgRgb (needed by photo search to rank items without loading photos)
    // and a flag so the grid knows whether to expect one.
    image: null,
    // "loading" still means the item HAS a photo — we just hadn't fetched it.
    hasPhoto: !formPhotoRemoved && formPhotoOrigin !== "none",
    avgRgb: formPhoto ? formPhoto.avgRgb : mixColor(colorFor(category), "#FFFFFF", 0.55),
  };
  const btn = document.getElementById("btnSubmitForm");
  btn.disabled = true;
  try {
    // Write the photo BEFORE the item document. If the connection drops between
    // the two, an orphaned photo is harmless; an item flagged hasPhoto with no
    // photo document would show a permanent empty placeholder.
    let photoStep = false;
    if (formPhoto && (formPhotoOrigin === "new" || formPhotoOrigin === "legacy")) {
      photoStep = true;
      // "legacy" is the case that used to lose photos: the image lived inline on
      // the item document, the save cleared that field, and nothing ever wrote it
      // to itemPhotos. Migrating it here is what makes editing a legacy item safe.
      let payload = { thumb: formPhoto.thumb, full: formPhoto.full, avgRgb: formPhoto.avgRgb };
      if (formPhotoOrigin === "legacy") {
        try { payload = await resizeDataUrl(formPhoto.thumb); } catch { /* keep the original */ }
      }
      await savePhoto(db, item.id, payload);
      photoStep = false;
    } else if (formPhoto && renamedFrom) {
      // SKU rename: the photo document is keyed by item id, so it has to move too.
      const existing = cachedPhoto(renamedFrom) || await loadPhoto(db, renamedFrom);
      if (existing) await savePhoto(db, item.id, existing);
    }

    if (renamedFrom) {
      const batch = writeBatch(db);
      batch.set(itemDocRef(item.id), item);
      batch.delete(itemDocRef(renamedFrom));
      await batch.commit();
      await deletePhoto(db, renamedFrom);
    } else {
      const ok = await saveItemRemote(item);
      if (!ok) return;
      // Deletion happens ONLY on an explicit Remove photo tap. Previously any
      // falsy formPhoto triggered it, so saving before the photo finished
      // loading silently deleted it.
      if (formPhotoRemoved) await deletePhoto(db, item.id);
    }
    forgetPhoto(item.id); // drop the stale cache entry so the grid refetches
    document.getElementById("formModal").classList.add("hidden");
    document.getElementById("formPhotoInput").value = "";
    toast(editingId ? "Item updated" : "Item added");
  } catch (e) {
    console.error(e);
    if (e.code === "permission-denied") {
      // Item writes and photo writes are governed by SEPARATE rules blocks, so a
      // setup where one works and the other doesn't is normal and worth naming
      // precisely — otherwise it reads as a login problem, which it isn't.
      toast(photoStep
        ? "Photo rejected — the itemPhotos rules block isn't published"
        : "Rejected by the database — your account isn't on the manager list");
    } else {
      toast("Couldn't save — check your connection");
    }
  } finally {
    btn.disabled = false;
  }
}

// ============================================================================
// Excel import. Reads an .xlsx (including embedded photos), shows a confirmation
// preview, then MERGES into Firestore — existing SKUs are updated, new ones added,
// and anything already in the database that's NOT in the file is left untouched.
// It never wipes the collection, so a stray import can't destroy everyone's data.
// ============================================================================
let pendingImport = null; // holds parsed items awaiting confirmation

function resolveZipPath(baseDir, relTarget) {
  if (relTarget.startsWith("/")) return relTarget.replace(/^\//, "");
  const stack = baseDir ? baseDir.split("/") : [];
  relTarget.split("/").forEach((p) => {
    if (p === "..") stack.pop();
    else if (p === "." || p === "") { /* skip */ }
    else stack.push(p);
  });
  return stack.join("/");
}
function bytesToBase64(bytes) {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return btoa(binary);
}
function extractImagesFromWorkbook(wb, sheetName) {
  const result = new Map();
  try {
    const files = wb.files;
    if (!files) return result;
    const getText = (path) => {
      const f = files[path];
      if (!f || f.content == null) return null;
      return typeof f.content === "string" ? f.content : new TextDecoder("utf-8").decode(f.content);
    };
    const getBytes = (path) => {
      const f = files[path];
      if (!f || f.content == null) return null;
      if (typeof f.content === "string") return new TextEncoder().encode(f.content);
      return f.content instanceof Uint8Array ? f.content : new Uint8Array(f.content);
    };
    const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const workbookXml = getText("xl/workbook.xml");
    if (!workbookXml) return result;
    let sheetM = workbookXml.match(new RegExp(`<sheet[^>]*name="${escRe(sheetName)}"[^>]*r:id="(rId\\d+)"`));
    if (!sheetM) sheetM = workbookXml.match(new RegExp(`<sheet[^>]*r:id="(rId\\d+)"[^>]*name="${escRe(sheetName)}"`));
    if (!sheetM) return result;
    const wbRels = getText("xl/_rels/workbook.xml.rels");
    if (!wbRels) return result;
    const relM = wbRels.match(new RegExp(`Id="${sheetM[1]}"[^>]*Target="([^"]+)"`));
    if (!relM) return result;
    let sheetPath = relM[1].replace(/^\//, "");
    if (!sheetPath.startsWith("xl/")) sheetPath = "xl/" + sheetPath;
    const sheetDir = sheetPath.substring(0, sheetPath.lastIndexOf("/"));
    const sheetFile = sheetPath.substring(sheetPath.lastIndexOf("/") + 1);
    const sheetRels = getText(`${sheetDir}/_rels/${sheetFile}.rels`);
    if (!sheetRels) return result;
    const drawRelM = sheetRels.match(/Id="(rId\d+)"[^>]*Target="([^"]*drawing[^"]*)"/i);
    if (!drawRelM) return result;
    const drawingPath = resolveZipPath(sheetDir, drawRelM[2]);
    const drawingXml = getText(drawingPath);
    if (!drawingXml) return result;
    const anchorBlocks = drawingXml.match(/<xdr:twoCellAnchor[\s\S]*?<\/xdr:twoCellAnchor>/g) || [];
    const rowToRid = [];
    anchorBlocks.forEach((block) => {
      const rowM = block.match(/<xdr:from>\s*<xdr:col>\d+<\/xdr:col>\s*<xdr:colOff>\d+<\/xdr:colOff>\s*<xdr:row>(\d+)<\/xdr:row>/);
      const ridM = block.match(/r:embed="(rId\d+)"/);
      if (rowM && ridM) rowToRid.push({ row: parseInt(rowM[1], 10), rid: ridM[1] });
    });
    if (rowToRid.length === 0) return result;
    const drawingDir = drawingPath.substring(0, drawingPath.lastIndexOf("/"));
    const drawingFile = drawingPath.substring(drawingPath.lastIndexOf("/") + 1);
    const drawingRels = getText(`${drawingDir}/_rels/${drawingFile}.rels`);
    if (!drawingRels) return result;
    const ridToMedia = {};
    Array.from(drawingRels.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)).forEach((rm) => {
      ridToMedia[rm[1]] = resolveZipPath(drawingDir, rm[2]);
    });
    rowToRid.forEach(({ row, rid }) => {
      const mediaPath = ridToMedia[rid];
      if (!mediaPath) return;
      const bytes = getBytes(mediaPath);
      if (!bytes) return;
      const ext = mediaPath.split(".").pop().toLowerCase();
      const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : "image/png";
      result.set(row, `data:${mime};base64,${bytesToBase64(bytes)}`);
    });
  } catch (e) { console.error("Image extraction failed:", e); }
  return result;
}
// Excel embeds photos at full resolution, which can be several MB each — far over
// Firestore's ~1MB-per-field limit. Shrink every imported photo to the same small
// size the in-app camera uses (~160px, JPEG) so it fits comfortably and stays fast.
//
// Both the shrink and the colour sample now run on their own canvas. Previously
// they shared the two hidden canvases in index.html while the importer processed
// every row in parallel, so concurrent rows overwrote each other's pixels and
// items could end up wearing another item's photo and colour.
async function shrinkAndSample(dataUrl, maxDim = 160, quality = 0.6) {
  try {
    const img = await loadImage(dataUrl);
    const { w, h } = fitDims(img, maxDim);
    const canvas = scratchCanvas(w, h);
    canvas.getContext("2d").drawImage(img, 0, 0, w, h);
    return { image: canvas.toDataURL("image/jpeg", quality), avgRgb: avgRgbOfImage(img) };
  } catch (e) {
    return { image: null, avgRgb: null }; // drop a bad photo rather than fail the import
  }
}

async function handleExcelImport(file) {
  if (!file) return;
  const errEl = document.getElementById("importError");
  errEl.style.display = "none";
  try {
    toast("Reading spreadsheet…");
    await loadScript(LIB.xlsx);
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", bookFiles: true });
    const sheetName = wb.SheetNames.includes("Inventory") ? "Inventory" : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const headerIdx = aoa.findIndex((r) => r && r[0] === "SKU");
    if (headerIdx === -1) {
      showImportModal("Couldn't find a SKU column in that file. Make sure you're using the inventory workbook with a header row that starts with 'SKU'.", true);
      return;
    }
    const headers = aoa[headerIdx];
    const col = (name) => headers.indexOf(name);
    const imageMap = extractImagesFromWorkbook(wb, sheetName);

    const rows = [];
    aoa.forEach((r, idx) => {
      if (idx <= headerIdx) return;
      if (r && r[0] && r[0] !== "TOTALS") rows.push({ r, aoaIndex: idx });
    });
    const rowIndexSet = new Set(rows.map((x) => x.aoaIndex));
    const resolvedImageMap = new Map();
    imageMap.forEach((dataUrl, anchorRow) => {
      if (rowIndexSet.has(anchorRow)) { resolvedImageMap.set(anchorRow, dataUrl); return; }
      for (const delta of [1, -1, 2, -2]) {
        if (rowIndexSet.has(anchorRow + delta)) { resolvedImageMap.set(anchorRow + delta, dataUrl); return; }
      }
    });

    const existingById = new Map(ITEMS.map((it) => [it.id, it]));

    // Four at a time rather than every row at once: importing a few hundred
    // full-size photos in parallel was enough to crash the tab on a mid-range phone.
    const parsed = await mapLimit(rows, 4, async ({ r, aoaIndex }) => {
      const category = String(r[col("Category")] || "Other");
      // Firestore document IDs can't contain "/" and can't be empty.
      const sku = String(r[col("SKU")]).trim().replace(/\//g, "-");
      const rawImage = resolvedImageMap.get(aoaIndex) || null;
      let { image, avgRgb } = rawImage ? await shrinkAndSample(rawImage) : { image: null, avgRgb: null };

      // A spreadsheet row with no embedded photo used to overwrite the item with
      // image: null — silently deleting a photo that had been added in the app.
      // Keep whatever is already stored when the file doesn't supply one.
      if (!image) {
        const prev = existingById.get(sku);
        if (prev && prev.image) { image = prev.image; avgRgb = prev.avgRgb || null; }
      }

      return {
        id: sku, sku, name: String(r[col("Item Name")] || ""),
        category, material: String(r[col("Material")] || ""), supplier: String(r[col("Supplier")] || ""),
        warehouse: String(r[col("Warehouse/Location")] || ""),
        cost: Number(r[col("Cost Price")]) || 0, price: Number(r[col("Selling Price")]) || 0,
        stock: Number(r[col("Stock Qty")]) || 0, reorder: Number(r[col("Reorder Level")]) || 0,
        image, avgRgb: avgRgb || mixColor(colorFor(category), "#FFFFFF", 0.55),
      };
    });

    if (parsed.length === 0) { showImportModal("No item rows found in that file.", true); return; }

    pendingImport = parsed;
    const existingIds = new Set(ITEMS.map((it) => it.id));
    const updating = parsed.filter((it) => existingIds.has(it.id)).length;
    const adding = parsed.length - updating;
    const withPhotos = parsed.filter((it) => it.image).length;

    // Catches the most natural mistake with this workbook: editing "Stock Status"
    // from OUT OF STOCK to IN STOCK. That column is calculated from Stock Qty and
    // is ignored on import, so without this check the import "succeeds", changes
    // nothing, and reports no problem at all.
    const statusCol = headers.findIndex((h) => String(h || "").startsWith("Stock Status"));
    const editedAutoCols = [];
    if (statusCol !== -1) {
      rows.forEach(({ r }) => {
        const claimed = String(r[statusCol] || "").trim().toUpperCase();
        if (!claimed) return;
        const qty = Number(r[col("Stock Qty")]) || 0;
        const reorder = Number(r[col("Reorder Level")]) || 0;
        const actual = qty <= 0 ? "OUT OF STOCK" : qty <= reorder ? "LOW STOCK" : "IN STOCK";
        if (claimed !== actual) {
          editedAutoCols.push({ sku: String(r[col("SKU")] || ""), claimed, qty });
        }
      });
    }

    let warning = "";
    if (editedAutoCols.length) {
      const sample = editedAutoCols.slice(0, 5).map((x) => `${esc(x.sku)} (says ${esc(x.claimed)}, qty ${x.qty})`).join(", ");
      warning =
        `<div class="import-warning sans"><b>⚠ ${editedAutoCols.length} row${editedAutoCols.length === 1 ? "" : "s"} ` +
        `changed the “Stock Status” column.</b><br>` +
        `That column is calculated from <b>Stock Qty</b> and is ignored on import, so those edits will have no effect. ` +
        `To put an item back in stock, set its <b>Stock Qty</b> to the number you actually have.<br>` +
        `<span style="opacity:.75">${sample}${editedAutoCols.length > 5 ? " …" : ""}</span></div>`;
    }

    showImportModal(
      warning +
      `<b>${parsed.length}</b> items found in the file (${withPhotos} with photos).<br><br>` +
      `• <b>${adding}</b> new item${adding === 1 ? "" : "s"} will be added<br>` +
      `• <b>${updating}</b> existing item${updating === 1 ? "" : "s"} (same SKU) will be updated<br><br>` +
      `Nothing else in your inventory is touched — items already in the app but not in this file stay as they are.`,
      false
    );
  } catch (e) {
    console.error(e);
    showImportModal("Couldn't read that file — check it's a valid .xlsx spreadsheet.", true);
  } finally {
    // Without this, picking the SAME file again after an error did nothing at all,
    // because the input's value hadn't changed so no change event fired.
    document.getElementById("importExcelInput").value = "";
  }
}

function showImportModal(html, isError) {
  document.getElementById("importSummary").innerHTML = isError ? "" : html;
  const errEl = document.getElementById("importError");
  if (isError) { errEl.textContent = html.replace(/<[^>]+>/g, ""); errEl.style.display = "block"; }
  else errEl.style.display = "none";
  document.getElementById("btnConfirmImport").style.display = isError ? "none" : "";
  document.getElementById("importModal").classList.remove("hidden");
}
function closeImportModal() {
  document.getElementById("importModal").classList.add("hidden");
  pendingImport = null;
  document.getElementById("importExcelInput").value = "";
}

async function confirmImport() {
  if (!pendingImport) return;
  const btn = document.getElementById("btnConfirmImport");
  btn.textContent = "Importing…"; btn.disabled = true;
  try {
    // Final safety guard: if any single image is still over ~0.9MB (Firestore's hard
    // limit is ~1MB per field), drop just that photo rather than let it fail the whole
    // import. The item itself still imports, just without its picture.
    let droppedPhotos = 0;
    pendingImport.forEach((it) => {
      if (it.image && it.image.length > 900000) { it.image = null; droppedPhotos++; }
    });

    // Two writes to the same document inside one batch are rejected outright, so a
    // file with a repeated SKU used to fail the whole import. Last row wins.
    const deduped = Array.from(new Map(pendingImport.map((it) => [it.id, it])).values());

    // The old code batched purely by count (400 docs). A batch also has a ~10 MB
    // payload ceiling, and 400 items carrying photos comfortably blew past it —
    // failing with an opaque "transaction too big". Now a batch closes when either
    // limit is reached, whichever comes first.
    const MAX_OPS = 400;
    const MAX_BYTES = 8 * 1024 * 1024;
    let batch = writeBatch(db), ops = 0, bytes = 0;
    for (const it of deduped) {
      const size = (it.image ? it.image.length : 0) + 600;
      if (ops > 0 && (ops >= MAX_OPS || bytes + size > MAX_BYTES)) {
        await batch.commit();
        batch = writeBatch(db); ops = 0; bytes = 0;
      }
      batch.set(itemDocRef(it.id), it);
      ops++; bytes += size;
    }
    if (ops > 0) await batch.commit();

    toast(droppedPhotos
      ? `Imported ${pendingImport.length} items (${droppedPhotos} photo${droppedPhotos === 1 ? "" : "s"} too large, skipped)`
      : `Imported ${pendingImport.length} items`);
    closeImportModal();
  } catch (e) {
    console.error(e);
    toast(e.code === "permission-denied" ? "Sign in required to import" : "Import failed — check your connection");
  } finally {
    btn.textContent = "Import"; btn.disabled = false;
  }
}


function todayStamp() { return new Date().toISOString().slice(0, 10); }

// The anchor now goes into the document and the URL is revoked on a delay.
// Revoking immediately after .click() cancelled the download in Firefox and Safari,
// which is why exports sometimes produced nothing at all.
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 10000);
}

function handleJsonExport() {
  const blob = new Blob([JSON.stringify(ITEMS, null, 2)], { type: "application/json" });
  downloadBlob(blob, `inventory_backup_${todayStamp()}.json`);
  toast("Backup ready — save it somewhere safe");
}

// Exports current inventory to a real .xlsx with embedded product photos, using
// ExcelJS (SheetJS's free build can't embed images; ExcelJS can). Same column layout
// the importer expects, so an exported file round-trips: export → edit → re-import.
async function handleExcelExport() {
  toast("Building Excel file…");
  try {
    await loadScript(LIB.exceljs);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Inventory");

    // Margin %, Stock Value and Stock Status are CALCULATED from the other columns.
    // They are written for reading, and ignored on import. Labelling them "(auto)"
    // matters: editing "Stock Status" from OUT OF STOCK to IN STOCK looks like it
    // should work, changes nothing on import, and gives no error — the status is
    // derived from Stock Qty, so Stock Qty is the cell to edit.
    const AUTO_COLS = new Set([10, 13, 14]); // 1-based: Margin %, Stock Value, Stock Status
    const headers = ["SKU", "Image", "Item Name", "Category", "Material", "Supplier",
      "Warehouse/Location", "Cost Price", "Selling Price", "Margin % (auto)", "Stock Qty",
      "Reorder Level", "Stock Value (auto)", "Stock Status (auto)"];
    ws.addRow(headers);
    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      // Calculated columns get a muted header so they read as output, not input.
      cell.fill = { type: "pattern", pattern: "solid",
        fgColor: { argb: AUTO_COLS.has(colNumber) ? "FF6B7280" : "FF1F2937" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });
    ws.columns = [
      { width: 12 }, { width: 10 }, { width: 26 }, { width: 12 }, { width: 16 }, { width: 20 },
      { width: 16 }, { width: 11 }, { width: 12 }, { width: 9 }, { width: 10 }, { width: 12 }, { width: 12 }, { width: 13 },
    ];

    // Photos no longer live on the item documents, so the export has to fetch
    // them from the itemPhotos collection first. Without this the workbook came
    // out with an empty Image column even though the app showed photos fine.
    const needPhotos = ITEMS.filter((it) => it.hasPhoto || it.image);
    if (needPhotos.length) toast(`Fetching ${needPhotos.length} photo(s)…`);
    const photoFor = new Map();
    await mapLimit(needPhotos, 6, async (it) => {
      // Legacy inline photos are still honoured for anything not yet migrated.
      if (it.image && it.image.startsWith("data:image")) { photoFor.set(it.id, it.image); return; }
      const ph = cachedPhoto(it.id) || await loadPhoto(db, it.id);
      // The thumbnail is used deliberately: rows are only ~56px tall, so the
      // 900px version would bloat the file for no visible gain.
      if (ph && (ph.thumb || ph.full)) photoFor.set(it.id, ph.thumb || ph.full);
    });

    let rowNum = 2;
    for (const it of ITEMS) {
      const margin = it.price ? Math.round(((it.price - it.cost) / it.price) * 1000) / 10 : 0;
      const row = ws.addRow([
        it.sku, "", it.name, it.category, it.material, it.supplier, it.warehouse,
        it.cost, it.price, margin, it.stock, it.reorder, (it.cost || 0) * (it.stock || 0),
        stockInfo(it).text,
      ]);
      row.height = 60;
      row.getCell(8).numFmt = "#,##0";
      row.getCell(9).numFmt = "#,##0";
      row.getCell(10).numFmt = "0.0";
      row.getCell(13).numFmt = "#,##0";
      // Grey italics on every calculated cell, reinforcing that edits here do nothing.
      for (const c of [10, 13, 14]) {
        row.getCell(c).font = { italic: true, color: { argb: "FF9CA3AF" } };
      }

      const photoData = photoFor.get(it.id);
      if (photoData && photoData.startsWith("data:image")) {
        try {
          const ext = photoData.substring(11, photoData.indexOf(";")); // "jpeg" or "png"
          const imgId = wb.addImage({ base64: photoData, extension: ext === "jpeg" ? "jpeg" : "png" });
          // place in column B (index 1, zero-based) at this row, sized to fit the cell
          ws.addImage(imgId, {
            tl: { col: 1, row: rowNum - 1 },
            ext: { width: 56, height: 56 },
            editAs: "oneCell",
          });
        } catch (e) { /* skip a bad image rather than fail the whole export */ }
      }
      rowNum++;
    }

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    downloadBlob(blob, `inventory_export_${todayStamp()}.xlsx`);
    const withPhotos = photoFor.size;
    toast(`Excel exported with ${withPhotos} photo${withPhotos === 1 ? "" : "s"}`);
  } catch (e) {
    console.error(e);
    toast("Export failed — try again");
  }
}

// ============================================================================
// Manager login (Firebase Authentication — Email/Password)
// ============================================================================
function openLoginModal(message) {
  document.getElementById("loginMessage").textContent = message ||
    "Staff can browse and search without signing in — this is only needed to add, edit, or adjust stock.";
  document.getElementById("loginError").style.display = "none";
  document.getElementById("loginModal").classList.remove("hidden");
  document.getElementById("loginEmail").focus();
}
function closeLoginModal() { document.getElementById("loginModal").classList.add("hidden"); }

async function tryLogin() {
  const email = document.getElementById("loginEmail").value.trim();
  const password = document.getElementById("loginPassword").value;
  const err = document.getElementById("loginError");
  if (!email || !password) { err.textContent = "Enter your email and password"; err.style.display = "block"; return; }
  err.style.display = "none";
  try {
    await signInWithEmailAndPassword(auth, email, password);
    closeLoginModal();
    mode = "manager";
    render();
    toast("Signed in as manager");
  } catch (e) {
    console.error(e);
    // "Wrong email or password" for a network outage sent people hunting for a
    // password problem that didn't exist.
    err.textContent = e.code === "auth/network-request-failed"
      ? "No connection — check your internet and try again"
      : e.code === "auth/too-many-requests"
        ? "Too many attempts. Wait a minute and try again."
        : "Wrong email or password";
    err.style.display = "block";
  }
}

// ============================================================================
// Event wiring
// ============================================================================
document.getElementById("btnShowroom").onclick = () => { mode = "showroom"; render(); };
document.getElementById("btnManager").onclick = () => {
  if (isManager) { mode = "manager"; render(); } else { openLoginModal(); }
};
document.getElementById("btnSignOut").onclick = () => signOut(auth);

document.getElementById("btnCloseLogin").onclick = closeLoginModal;
document.getElementById("loginModal").onclick = (e) => { if (e.target.id === "loginModal") closeLoginModal(); };
document.getElementById("btnLoginSubmit").onclick = tryLogin;
document.getElementById("loginPassword").addEventListener("keydown", (e) => { if (e.key === "Enter") tryLogin(); });

document.getElementById("btnTextSearch").onclick = () => setSearchType("text");
document.getElementById("btnPhotoSearch").onclick = () => setSearchType("image");
document.getElementById("btnQRSearch").onclick = () => {
  setSearchType("qr");
  // Start fetching the QR reader the moment the tab is opened, so it's ready
  // by the time a photo is actually taken.
  loadScript(LIB.jsqr).catch(() => {});
};
function setSearchType(type) {
  searchType = type;
  document.getElementById("btnTextSearch").classList.toggle("active", type === "text");
  document.getElementById("btnPhotoSearch").classList.toggle("active", type === "image");
  document.getElementById("btnQRSearch").classList.toggle("active", type === "qr");
  document.getElementById("textSearchPanel").classList.toggle("hidden", type !== "text");
  document.getElementById("photoSearchPanel").classList.toggle("hidden", type !== "image");
  document.getElementById("qrSearchPanel").classList.toggle("hidden", type !== "qr");
}

let searchDebounceTimer = null;
document.getElementById("searchInput").oninput = (e) => {
  query = e.target.value;
  clearTimeout(searchDebounceTimer);
  searchDebounceTimer = setTimeout(() => renderGrid(true), 150);
};
document.getElementById("catChips").onclick = (e) => {
  const btn = e.target.closest("[data-cat]");
  if (!btn) return;
  catFilter = btn.dataset.cat;
  renderCategoryChips(); renderGrid(true);
};

document.getElementById("grid").onclick = async (e) => {
  const btn = e.target.closest("[data-action]");
  if (!btn) return;
  const id = btn.dataset.id;
  const action = btn.dataset.action;
  const top = e.target.closest(".card-top");
  if (top && !action) {
    const card = top.closest("[data-card-id]");
    if (card) openLightbox(card.dataset.cardId);
    return;
  }

  if (action === "inc" || action === "dec") { adjustStock(id, action === "inc" ? 1 : -1); return; }

  if (action === "edit") {
    const it = ITEMS.find((x) => x.id === id);
    // Another device may have deleted this item between render and tap. Previously
    // that threw on `it.image` and the modal opened half-populated.
    if (!it) { toast("That item no longer exists"); renderGrid(true); return; }
    openEditModal(it);
  } else if (action === "delete-ask") { confirmDeleteId = id; renderGrid(); }
  else if (action === "delete-no") { confirmDeleteId = null; renderGrid(); }
  else if (action === "delete-yes") await deleteItemLocal(id);
};

document.getElementById("btnAdd").onclick = openAddModal;
document.getElementById("btnCloseModal").onclick = () => document.getElementById("formModal").classList.add("hidden");
document.getElementById("formModal").onclick = (e) => { if (e.target.id === "formModal") e.currentTarget.classList.add("hidden"); };
document.getElementById("btnRemovePhoto").onclick = () => {
  formPhoto = null;
  formPhotoOrigin = "none";
  formPhotoRemoved = true; // the only thing that authorises deleting a photo
  document.getElementById("formPhotoPreview").classList.add("hidden");
  document.getElementById("formPhotoPlaceholder").classList.remove("hidden");
  document.getElementById("photoPickerLabel").textContent = "Add photo";
  document.getElementById("btnRemovePhoto").classList.add("hidden");
};

document.getElementById("btnSubmitForm").onclick = submitForm;

const formImgRef = document.getElementById("formPhotoInput");
document.getElementById("photoPicker").onclick = () => formImgRef.click();
formImgRef.onchange = async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    // A failure here used to be an unhandled rejection: the picker just did
    // nothing and the user had no idea why.
    formPhoto = await resizeImageFile(file);
    formPhotoOrigin = "new";
    formPhotoRemoved = false;
    document.getElementById("btnRemovePhoto").classList.remove("hidden");
    document.getElementById("formPhotoPreview").src = formPhoto.thumb;
    document.getElementById("formPhotoPreview").classList.remove("hidden");
    document.getElementById("formPhotoPlaceholder").classList.add("hidden");
    document.getElementById("photoPickerLabel").textContent = "Change photo";
  } catch (err) {
    console.error(err);
    toast("Couldn't read that photo — try another");
  } finally {
    e.target.value = ""; // lets you pick the exact same file again
  }
};

document.getElementById("btnExportJson").onclick = handleJsonExport;
document.getElementById("btnExportExcel").onclick = handleExcelExport;
document.getElementById("btnPrintTags").onclick = printQRTags;

document.getElementById("btnImportExcel").onclick = () => document.getElementById("importExcelInput").click();
document.getElementById("importExcelInput").onchange = (e) => handleExcelImport(e.target.files[0]);
document.getElementById("btnCloseImport").onclick = closeImportModal;
document.getElementById("btnCancelImport").onclick = closeImportModal;
document.getElementById("btnConfirmImport").onclick = confirmImport;
document.getElementById("importModal").onclick = (e) => { if (e.target.id === "importModal") closeImportModal(); };

const dropzone = document.getElementById("dropzone");
dropzone.onclick = () => document.getElementById("photoSearchInput").click();
dropzone.ondragover = (e) => e.preventDefault();
dropzone.ondrop = (e) => { e.preventDefault(); handlePhotoSearchFile(e.dataTransfer.files[0]); };
document.getElementById("photoSearchInput").onchange = (e) => handlePhotoSearchFile(e.target.files[0]);
function handlePhotoSearchFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => analyzeUploadedPhoto(e.target.result);
  reader.readAsDataURL(file);
}
document.getElementById("btnNewPhoto").onclick = () => {
  document.getElementById("uploadResultWrap").classList.add("hidden");
  document.getElementById("photoSearchInput").value = "";
};

const qrDropzone = document.getElementById("qrDropzone");
qrDropzone.onclick = () => document.getElementById("qrSearchInput").click();
document.getElementById("qrSearchInput").onchange = (e) => handleQRSearchFile(e.target.files[0]);

// ============================================================================
// PWA install prompt + online/offline indicator
// ============================================================================
let deferredInstallPrompt = null;
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  document.getElementById("installBanner").classList.remove("hidden");
});
document.getElementById("btnInstall").onclick = async () => {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  await deferredInstallPrompt.userChoice;
  deferredInstallPrompt = null;
  document.getElementById("installBanner").classList.add("hidden");
};
window.addEventListener("appinstalled", () => document.getElementById("installBanner").classList.add("hidden"));

const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
const isStandalone = window.matchMedia("(display-mode: standalone)").matches || window.navigator.standalone;
if (isIos && !isStandalone && !localStorage.getItem("iosInstallHintDismissed")) {
  const banner = document.getElementById("installBanner");
  banner.classList.remove("hidden");
  banner.innerHTML = `<span>On iPhone/iPad: tap the Share button, then "Add to Home Screen".</span>
    <button id="btnDismissIosHint" class="btn" style="flex-shrink:0;">Got it</button>`;
  document.getElementById("btnDismissIosHint").onclick = () => {
    localStorage.setItem("iosInstallHintDismissed", "1");
    banner.classList.add("hidden");
  };
}

function updateOnlineStatus() {
  document.getElementById("offlinePill").classList.toggle("hidden", navigator.onLine);
}
window.addEventListener("online", updateOnlineStatus);
window.addEventListener("offline", updateOnlineStatus);

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch((e) => console.error("Service worker registration failed:", e));
  });
}

// ============================================================================
// Live data subscription + auth state + init
// ============================================================================

// ============================================================================
// Google Sheets two-way sync UI
//
// Every sync is a two-step action: build a plan, show it, then apply only if the
// manager confirms. Two-way sync fails silently or not at all, so it must never
// run without being seen first.
// ============================================================================
const GOOGLE_CLIENT_ID = firebaseConfig.googleClientId || "";

function syncBody(html) { document.getElementById("syncBody").innerHTML = html; }
function openSyncModal() { document.getElementById("syncModal").classList.remove("hidden"); }
function closeSyncModal() { document.getElementById("syncModal").classList.add("hidden"); }

async function startSync() {
  if (!isManager) { openLoginModal("Sign in to sync."); return; }
  if (!GOOGLE_CLIENT_ID) {
    openSyncModal();
    syncBody(`<p>No Google client ID configured. Add <code>googleClientId</code> to
      <code>firebase-config.js</code> — see the README, Part 5.</p>`);
    document.getElementById("btnSyncApply").disabled = true;
    return;
  }

  openSyncModal();
  document.getElementById("btnSyncApply").disabled = true;
  syncBody(`<p>Reading the sheet…</p>`);

  try {
    const result = await planSync(db, GOOGLE_CLIENT_ID, ITEMS);

    if (result.blocked) {
      // Structural damage to the sheet stops everything before any comparison.
      // Comparing against a sheet with renamed columns would read every field as
      // blank and look identical to "the whole catalogue was cleared".
      syncBody(`<p><strong>Sync stopped — the sheet's structure has changed.</strong></p>
        <ul>${result.blocked.map((i) => `<li>${esc(i)}</li>`).join("")}</ul>
        <p>Nothing was read or written. Fix the sheet and try again.</p>`);
      return;
    }

    const { plan, summary } = result;
    if (!summary.fromApp && !summary.fromSheet && !summary.conflicts && !summary.needsDeleteConfirm) {
      syncBody(`<p>Everything is already in sync — ${summary.unchanged} item(s), nothing to do.</p>
        <p><a href="${esc(result.cfg.url || "#")}" target="_blank" rel="noopener">Open the sheet</a></p>`);
      return;
    }

    syncBody(renderPlanHtml(plan, esc));
    refreshApplyButton();
  } catch (e) {
    console.error(e);
    if (e.needsConsent) {
      syncBody(`<p>Google access is needed to read the sheet.</p>
        <button id="btnGrant" class="btn primary">Connect Google</button>`);
      document.getElementById("btnGrant").onclick = async () => {
        try { await planSync(db, GOOGLE_CLIENT_ID, ITEMS); startSync(); }
        catch (err) { syncBody(`<p>${esc(err.message)}</p>`); }
      };
      return;
    }
    syncBody(`<p>Couldn't sync: ${esc(e.message)}</p><p>Nothing was changed.</p>`);
  }
}

function refreshApplyButton() {
  const btn = document.getElementById("btnSyncApply");
  const left = unresolvedCount();
  btn.disabled = left > 0;
  btn.textContent = left > 0 ? `Resolve ${left} conflict(s) first` : "Apply changes";
}

document.getElementById("syncBody").onclick = (e) => {
  const opt = e.target.closest(".conflict-opt");
  if (!opt) return;
  const [id, field] = opt.dataset.key.split("::");
  setConflictChoice(id, field, opt.dataset.pick);
  opt.parentElement.querySelectorAll(".conflict-opt").forEach((b) => b.classList.remove("chosen"));
  opt.classList.add("chosen");
  refreshApplyButton();
};

document.getElementById("btnSyncSheet").onclick = startSync;
document.getElementById("btnSyncCancel").onclick = () => { disconnectSheets(); closeSyncModal(); };

document.getElementById("btnSyncApply").onclick = async () => {
  const btn = document.getElementById("btnSyncApply");
  btn.disabled = true;
  const confirmed = new Set(
    [...document.querySelectorAll(".del-check:checked")].map((c) => c.dataset.id)
  );
  try {
    const res = await applyPlan(db, GOOGLE_CLIENT_ID, {
      confirmedDeletions: confirmed,
      onProgress: (msg) => syncBody(`<p>${esc(msg)}</p>`),
    });
    closeSyncModal();
    toast(`Synced — ${res.fromSheet} in, ${res.fromApp} out${res.deleted ? `, ${res.deleted} deleted` : ""}`);
  } catch (e) {
    console.error(e);
    // The baseline is written last, so a failure here leaves the previous sync
    // state intact and the next run simply re-proposes the same work.
    syncBody(`<p>Sync failed partway: ${esc(e.message)}</p>
      <p>Your sync state was not advanced — run the sync again to retry safely.</p>`);
    btn.disabled = false;
  }
};

// ============================================================================
// Photo lightbox — the payoff for storing a large version
// ============================================================================
async function openLightbox(id) {
  const it = ITEMS.find((x) => x.id === id);
  if (!it) return;
  const box = document.getElementById("lightbox");
  const img = document.getElementById("lightboxImg");
  document.getElementById("lightboxCaption").textContent = `${it.name} · ${it.sku}`;
  const cached = cachedPhoto(id);
  img.src = (cached && (cached.full || cached.thumb)) || it.image || "";
  box.classList.remove("hidden");
  if (!cached) {
    const ph = await loadPhoto(db, id);
    if (ph && !box.classList.contains("hidden")) img.src = ph.full || ph.thumb;
  }
}
document.getElementById("btnCloseLightbox").onclick = () =>
  document.getElementById("lightbox").classList.add("hidden");
document.getElementById("lightbox").onclick = (e) => {
  if (e.target.id === "lightbox") document.getElementById("lightbox").classList.add("hidden");
};

// ============================================================================
// One-time migration of inline photos into the itemPhotos collection
// ============================================================================
async function refreshMigrateButton() {
  const btn = document.getElementById("btnMigratePhotos");
  if (!isManager) { btn.classList.add("hidden"); return; }
  const n = ITEMS.filter((it) => typeof it.image === "string" && it.image.startsWith("data:")).length;
  btn.classList.toggle("hidden", n === 0);
  btn.textContent = `📦 Move ${n} photo(s) out of items`;
}

document.getElementById("btnMigratePhotos").onclick = async () => {
  const btn = document.getElementById("btnMigratePhotos");
  btn.disabled = true;
  try {
    const res = await migrateLegacyPhotos(db, {
      makeSizes: (dataUrl) => resizeDataUrl(dataUrl),
      onProgress: (done, total) => { btn.textContent = `Moving ${done}/${total}…`; },
    });
    toast(`Moved ${res.migrated} photo(s) out of the item records`);
  } catch (e) {
    console.error(e);
    toast("Migration failed — nothing was lost, try again");
  } finally {
    btn.disabled = false;
    refreshMigrateButton();
  }
};

initLazyPhotos(db);

// Signing in and being ALLOWED TO WRITE are two different things. Firestore only
// accepts writes from accounts listed in the `managers` collection, so a valid
// login with no managers entry produced a baffling "sign in required to save"
// on every action. This checks up front and says exactly what's wrong, including
// the UID to paste into the console.
async function checkManagerAllowlist(user) {
  const banner = document.getElementById("managerWarning");
  if (!user || user.isAnonymous) { banner.classList.add("hidden"); return; }
  try {
    const snap = await getDoc(doc(db, "managers", user.uid));
    if (snap.exists()) { banner.classList.add("hidden"); return; }
    banner.innerHTML =
      `<b>You're signed in, but not on the manager list — saving will fail.</b><br>` +
      `In Firebase Console → Firestore Database → Data, create a collection named ` +
      `<code>managers</code> with a document whose ID is exactly:<br>` +
      `<code class="uid">${esc(user.uid)}</code><br>` +
      `Give it any single field (for example <code>role: manager</code>) so the document actually saves.`;
    banner.classList.remove("hidden");
  } catch (e) {
    console.error("manager check failed", e);
    banner.classList.add("hidden"); // never block the UI on a diagnostic
  }
}

// Seeding used to happen automatically whenever the collection looked empty. Three
// things went wrong with that:
//   • Delete all ten demo items and they instantly reappeared on the next load.
//   • Staff without a login triggered it, and the writes failed silently.
//   • Two devices opening at once both seeded, racing each other.
// It's now an explicit button that only appears when the inventory is genuinely
// empty and you're signed in as manager.
async function loadSampleData() {
  if (!isManager) { openLoginModal("Sign in to load the sample catalogue."); return; }
  if (ITEMS.length > 0) { toast("Inventory isn't empty — sample data not loaded"); return; }
  try {
    const batch = writeBatch(db);
    SEED_ITEMS.forEach((it) => batch.set(itemDocRef(it.id), it));
    await batch.commit();
    toast("Sample catalogue loaded");
  } catch (e) {
    console.error(e);
    toast("Couldn't load sample data");
  }
}
document.getElementById("btnLoadSample").onclick = loadSampleData;

onSnapshot(collection(db, "items"),
  (snapshot) => {
    ITEMS = snapshot.docs.map((d) => d.data());
    ITEMS.forEach(ensureAvgRgb);
    render();
  },
  (error) => {
    console.error(error);
    toast(error.code === "permission-denied"
      ? "No access to the inventory — check the Firestore rules"
      : "Couldn't connect — check your internet connection");
  }
);

// Reads are public by your decision, so browsing needs no sign-in at all and no
// anonymous auth setup. Only manager writes are gated, by the `managers`
// allowlist enforced in firestore.rules.
onAuthStateChanged(auth, (user) => {
  if (!user) {
    isManager = false;
    if (mode === "manager") mode = "showroom";
    render();
    return;
  }
  // isAnonymous is still checked so that turning anonymous auth on later for any
  // reason can't accidentally hand every browsing device the manager interface.
  isManager = !user.isAnonymous;
  checkManagerAllowlist(user);
  if (!isManager && mode === "manager") mode = "showroom";
  render();
});

updateOnlineStatus();
render();
