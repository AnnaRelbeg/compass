// Offline support: the app files are cached, and refreshed in the background whenever you're online.
// Firestore keeps its own offline copy of your data, so this only covers the app itself.
const CACHE = "money-v3";
const SHELL = ["./", "index.html", "app.css", "app.js", "calc.js", "importer.js", "store.js", "firebase-config.js", "manifest.webmanifest", "icons/icon-192.png"];

self.addEventListener("install", (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET") return;
  // never cache Firebase API traffic; do cache the app, fonts, SheetJS and the Firebase SDK files
  if (/googleapis\.com$/.test(u.hostname) && !u.hostname.startsWith("fonts")) return;
  if (u.hostname.includes("firestore") || u.hostname.includes("identitytoolkit") || u.hostname.includes("securetoken")) return;
  e.respondWith(caches.open(CACHE).then(async (c) => {
    const hit = await c.match(e.request);
    const net = fetch(e.request).then((r) => { if (r.ok) c.put(e.request, r.clone()); return r; }).catch(() => hit);
    return hit || net;
  }));
});
