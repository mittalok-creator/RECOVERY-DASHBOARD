const CACHE_NAME = 'recovery-dashboard-shell-v10';
// Recovery Dashboard (branch portal) -- forked from npadashboard.alokmittal.net's
// own sw.js. Same stale-while-revalidate shell + network-first data pattern;
// only the SHELL_ASSETS list changed (js/login.js instead of js/auth.js/
// js/publish.js/js/splash.js -- this portal has no Admin/OAuth/Publish flow
// at all). Keep this list in sync with index.html's actual ?v= query params
// on every future version bump, same discipline the production repo uses.
const SHELL_ASSETS = [
  './',
  './index.html',
  './css/styles.css?v=20260925d',
  './js/app.js?v=20260925i',
  './js/login.js?v=20260925a',
  './manifest.webmanifest',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_ASSETS)).catch(() => {})
  );
  self.skipWaiting();
});

// Holds the last successfully-fetched data/latest.json, separately from
// CACHE_NAME, so "Download for Offline" survives an app-shell update (same
// reasoning as the production repo's own sw.js -- see its comment).
const DATA_CACHE_NAME = 'recovery-dashboard-data';

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME && n !== DATA_CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

// data/latest.json is fetched CROSS-ORIGIN (https://npadashboard.alokmittal.net/...,
// see DATA_ORIGIN in js/app.js -- this portal has no data of its own) with a
// `?t=<timestamp>` cache-buster that's a different URL every load. A fetch
// event still fires for cross-origin requests made by pages in this SW's
// scope, and the response is a normal (non-opaque) CORS response since the
// production site serves these files with access-control-allow-origin:* --
// so this caches exactly the same way the production site's own sw.js
// caches its own same-origin data/latest.json. url.pathname is
// origin-independent, so the pattern below matches regardless. Only
// data/latest.json gets an offline fallback (not data/kcc-overdue.json),
// matching the production site's own "Download for Offline" button, which
// only ever fetches data/latest.json.
const DATA_URL_PATTERN = /\/data\/latest\.json$/;
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (DATA_URL_PATTERN.test(url.pathname)) {
    const cacheKey = url.pathname;
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const copy = response.clone();
          caches.open(DATA_CACHE_NAME).then((cache) => cache.put(cacheKey, copy)).catch(() => {});
          return response;
        })
        .catch(() => caches.open(DATA_CACHE_NAME).then((cache) => cache.match(cacheKey)))
    );
    return;
  }
  // KCC Overdue's cross-origin fetch (data/kcc-overdue.json) and any other
  // cross-origin request just passes straight through, uncached -- same
  // "no special handling" default the production sw.js applies to anything
  // that isn't its own shell or data/latest.json.
  if (url.origin !== self.location.origin) {
    return;
  }
  // Stale-while-revalidate for this portal's own shell (HTML/CSS/JS/manifest),
  // same as the production site's own sw.js.
  event.respondWith(
    caches.open(CACHE_NAME).then((cache) =>
      cache.match(event.request).then((cached) => {
        const network = fetch(event.request)
          .then((response) => { cache.put(event.request, response.clone()).catch(() => {}); return response; })
          .catch(() => cached);
        return cached || network;
      })
    )
  );
});
