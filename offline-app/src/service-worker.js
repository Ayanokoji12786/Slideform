const cacheName = "slideform-v5";
const assets = [
  "./", "./index.html", "./styles.css", "./app.js", "./demo.js", "./manifest.webmanifest",
  "./vendor/jszip.min.js", "./vendor/exceljs.min.js",
  "./vendor/pdf.min.mjs", "./vendor/pdf.worker.min.mjs",
  "./assets/guide-character.png", "./assets/guide-reviewing.png", "./assets/guide-complete.png"
];

self.addEventListener("install", (event) => event.waitUntil(caches.open(cacheName).then((cache) => cache.addAll(assets))));
self.addEventListener("activate", (event) => event.waitUntil(
  caches.keys()
    .then((names) => Promise.all(names.filter((name) => name !== cacheName).map((name) => caches.delete(name))))
    .then(() => self.clients.claim())
));
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request)));
});
