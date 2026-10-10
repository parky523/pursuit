// Pursuit service worker: makes the phone page installable as an app.
// Network-first passthrough; the page itself keeps its own offline copy in localStorage.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(fetch(e.request).catch(() => new Response("Pursuit: PC offline", { status: 503 })));
});
