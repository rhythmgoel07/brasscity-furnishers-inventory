// Bump this version any time you change a cached file, so phones pick up updates.
const CACHE_VERSION = "furniture-inventory-firebase-v15";

// App SHELL only — not the live inventory data. The data comes from Firestore,
// which keeps its own separate offline cache (configured in app.js), independent
// of anything here.
//
// The CDN libraries are deliberately NOT precached any more. app.js loads them on
// demand, and the runtime cache below picks them up the first time they're really
// used. Precaching them meant every install pulled ~2 MB up front, and a single
// failed URL aborted the entire install.
const PRECACHE_URLS = [
  "./index.html",
  "./app.js",
  "./firebase-config.js",
  "./sync-engine.js",
  "./sheets-sync.js",
  "./sync-ui.js",
  "./photos.js",
  "./tags.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png",
];

// Files that should always try the network first, so a deployed fix reaches phones
// on the next load instead of sitting behind a stale cache entry.
const NETWORK_FIRST = [
  "/index.html", "/app.js", "/firebase-config.js",
  "/sync-engine.js", "/sheets-sync.js", "/sync-ui.js", "/photos.js", "/tags.js",
];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_VERSION);
    // cache.addAll() is all-or-nothing: one 404 or flaky request used to abort the
    // whole install and leave the app with no service worker at all. Cache each
    // file independently and tolerate individual failures.
    await Promise.all(PRECACHE_URLS.map((url) =>
      cache.add(new Request(url, { cache: "reload" })).catch((e) => {
        console.warn("Precache skipped:", url, e);
      })
    ));
    await self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter((n) => n !== CACHE_VERSION).map((n) => caches.delete(n)));
    if (self.registration.navigationPreload) {
      await self.registration.navigationPreload.enable();
    }
    await self.clients.claim();
  })());
});

function isFirebaseTraffic(url) {
  return url.includes("googleapis.com")
    || url.includes("accounts.google.com")
    || url.includes("gstatic.com/firebasejs")
    || url.includes("firebaseio.com")
    || url.includes("firebaseapp.com");
}

self.addEventListener("fetch", (event) => {
  const req = event.request;

  // Only GET is cacheable; POSTs to Firestore and friends must pass straight through.
  if (req.method !== "GET") return;

  // The Cache API only accepts http(s). Browser extensions issue requests with
  // schemes like chrome-extension: and moz-extension:, and trying to cache one
  // throws "Request scheme 'chrome-extension' is unsupported" — noise that has
  // nothing to do with the app but fills the console and hides real errors.
  if (!req.url.startsWith("http")) return;

  // Never intercept Firebase traffic — real-time sync and auth need the network
  // directly, and should fail honestly rather than serve stale responses.
  if (isFirebaseTraffic(req.url)) return;

  const url = new URL(req.url);

  // HTML navigations: network first, cached shell as the fallback.
  if (req.mode === "navigate") {
    event.respondWith((async () => {
      try {
        const preload = await event.preloadResponse;
        if (preload) return preload;
        return await fetch(req);
      } catch (e) {
        return (await caches.match("./index.html")) || Response.error();
      }
    })());
    return;
  }

  // app.js and config: network first. Under the old cache-first rule you could end
  // up running yesterday's app.js against today's index.html.
  if (url.origin === self.location.origin && NETWORK_FIRST.some((p) => url.pathname.endsWith(p))) {
    event.respondWith((async () => {
      try {
        const res = await fetch(req);
        if (res && res.status === 200) {
          const cache = await caches.open(CACHE_VERSION);
          cache.put(req, res.clone());
        }
        return res;
      } catch (e) {
        const cached = await caches.match(req);
        if (cached) return cached;
        throw e;
      }
    })());
    return;
  }

  // Everything else (icons, the lazily-loaded CDN libraries): stale-while-revalidate.
  event.respondWith((async () => {
    const cached = await caches.match(req);
    const network = fetch(req).then(async (res) => {
      // Opaque cross-origin responses report status 0. Caching them is still useful
      // for the CDN libraries, which is why status 0 is allowed through here.
      if (res && (res.status === 200 || res.type === "opaque")) {
        const cache = await caches.open(CACHE_VERSION);
        cache.put(req, res.clone());
      }
      return res;
    }).catch(() => cached);
    return cached || network;
  })());
});
