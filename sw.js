// Devis Maquette : garde une copie de l'appli pour l'ouvrir sans réseau.
// Réseau d'abord (les mises à jour publiées arrivent tout de suite), copie locale si pas de réseau.
const CACHE = "devis-maquette-v29";
const FILES = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png", "./logo-driv3d.png"];
// Bibliothèques et polices externes utilisées par l'appli (mises en cache aussi).
const EXTERNES = ["cdnjs.cloudflare.com", "fonts.googleapis.com", "fonts.gstatic.com"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== CACHE).map(x => caches.delete(x))))); self.clients.claim(); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET") return;
  if (u.origin !== location.origin && !EXTERNES.includes(u.hostname)) return; // jamais le script Google Drive
  // Pour l'appli elle-même : on court-circuite le cache du navigateur afin d'avoir toujours la dernière version publiée.
  const req = u.origin === location.origin ? fetch(e.request.url, { cache: "no-cache" }) : fetch(e.request);
  e.respondWith(req.then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r; })
    .catch(() => caches.match(e.request).then(r => r || (e.request.mode === "navigate" ? caches.match("./index.html") : undefined))));
});
