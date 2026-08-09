// ============================================================
//  Service Worker — Aplikasi Jadwal KBM SMANSABA
//  v3 — dengan halaman offline ramah
// ============================================================

const CACHE_NAME   = 'jadwal-smansaba-v3';
const CACHE_STATIC = 'jadwal-static-v3';
const OFFLINE_URL  = '/offline.html';

const STATIC_ASSETS = [
  '/',
  '/index.html',
  '/offline.html',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
  'https://cdn.tailwindcss.com',
];

// ── INSTALL ──────────────────────────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_STATIC)
      .then(cache => Promise.allSettled(
        STATIC_ASSETS.map(url => cache.add(url).catch(e => console.warn('[SW] Skip:', url, e)))
      ))
      .then(() => self.skipWaiting())
  );
});

// ── ACTIVATE ─────────────────────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME && k !== CACHE_STATIC)
            .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// ── FETCH ─────────────────────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // 1. Apps Script → Network only, offline fallback JSON
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('script.googleusercontent.com')) {
    event.respondWith(
      fetch(event.request.clone()).catch(() =>
        new Response(
          JSON.stringify({ ok:false, msg:'Anda sedang offline. Koneksi internet tidak tersedia.' }),
          { headers: { 'Content-Type':'application/json' } }
        )
      )
    );
    return;
  }

  // 2. Navigasi (buka halaman) → Network First, fallback ke cache lalu offline.html
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request).catch(async () => {
        const cached = await caches.match(event.request);
        return cached || caches.match(OFFLINE_URL);
      })
    );
    return;
  }

  // 3. Tailwind CDN → Cache First
  if (url.hostname === 'cdn.tailwindcss.com') {
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(res => {
          if (res.ok) caches.open(CACHE_STATIC).then(c => c.put(event.request, res.clone()));
          return res;
        });
      })
    );
    return;
  }

  // 4. Aset lokal → Stale While Revalidate
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(event.request).then(cached => {
        const network = fetch(event.request).then(res => {
          if (res.ok) caches.open(CACHE_STATIC).then(c => c.put(event.request, res.clone()));
          return res;
        }).catch(() => cached);
        return cached || network;
      })
    );
    return;
  }

  event.respondWith(fetch(event.request));
});

self.addEventListener('message', event => {
  if (event.data?.type === 'SKIP_WAITING') self.skipWaiting();
});
