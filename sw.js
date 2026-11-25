const CACHE = 'campo-eeq-v1';
const CORE = ['./', 'index.html', 'style.css', 'app.js', 'manifest.json', 'icon-192.png', 'icon-512.png', 'lib/jszip.min.js', 'lib/leaflet.js', 'lib/leaflet.css'];
const OPCIONAL = ['calles.json'];   // las calles: si fallan, la app igual instala y lo reintenta después

async function bajar(cache, url, intentos) {
  for (let i = 0; i < intentos; i++) {
    try {
      const r = await fetch(new Request(url, { cache: 'reload' }));
      if (r && r.ok) { await cache.put(url, r.clone()); return true; }
    } catch (_) { /* reintenta */ }
    await new Promise((res) => setTimeout(res, 700 * (i + 1)));
  }
  return false;
}
self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    const res = await Promise.all(CORE.map((u) => bajar(c, u, 4)));
    if (res.includes(false)) throw new Error('No se pudieron guardar todos los archivos; se reintentará.');
    await Promise.all(OPCIONAL.map((u) => bajar(c, u, 2)));
    await self.skipWaiting();
  })());
});
self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    const k = await caches.keys(); await Promise.all(k.filter((x) => x !== CACHE).map((x) => caches.delete(x)));
    // si faltó algo opcional en la instalación, completarlo ahora
    const c = await caches.open(CACHE); for (const u of OPCIONAL) { if (!(await c.match(u))) await bajar(c, u, 2); }
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', (e) => {
  const req = e.request; if (req.method !== 'GET') return;
  e.respondWith((async () => {
    const hit = await caches.match(req, { ignoreSearch: true });
    if (hit) return hit;
    try { return await fetch(req); }
    catch (_) {
      if (req.mode === 'navigate') { const i = await caches.match('index.html'); if (i) return i; }
      return new Response('Sin conexión', { status: 503, statusText: 'Sin conexión' });
    }
  })());
});
