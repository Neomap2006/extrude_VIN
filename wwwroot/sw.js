const CACHE_NAME = 'velasto-pwa-cache-v19';
const NAV_TIMEOUT_MS = 500; // 2 detik — jika server lambat/mati, cepat beralih ke cache offline

const ASSETS_TO_CACHE = [
  // Halaman Fallback Offline & Halaman PWA
  '/offline.html',

  // Lokal aset — CSS
  '/lib/offline-cdn/bootstrap.min.css',
  '/lib/offline-cdn/bootstrap-icons.css',
  '/lib/offline-cdn/sweetalert2.min.css',
  '/lib/offline-cdn/caveat.css',
  '/css/site.css',

  // Lokal aset — JS
  '/lib/offline-cdn/bootstrap.bundle.min.js',
  '/lib/offline-cdn/jquery-3.6.0.min.js',
  '/lib/offline-cdn/sweetalert2.all.min.js',
  '/lib/offline-cdn/html5-qrcode.min.js',
  '/lib/offline-cdn/chart.umd.min.js',
  '/lib/offline-cdn/chartjs-plugin-zoom.min.js',
  '/lib/offline-cdn/hammer.min.js',
  '/lib/offline-cdn/signalr.min.js',
  '/lib/offline-cdn/html2pdf.bundle.min.js',
  '/lib/offline-cdn/jquery.validate.min.js',
  '/lib/offline-cdn/jquery.validate.unobtrusive.min.js',
  '/lib/offline-cdn/dexie.min.js',
  '/js/site.js',
  '/js/offline-sync.js',

  // Font
  '/lib/offline-cdn/fonts/bootstrap-icons.woff2',
  '/lib/offline-cdn/fonts/bootstrap-icons.woff',
  '/lib/offline-cdn/fonts/caveat.ttf',
  '/favicon.svg',
  '/images/icon-192.png',
  '/images/icon-512.png',
];

// ─── Install: cache semua aset statis ─────────────────────────────────────────
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return Promise.all(
        ASSETS_TO_CACHE.map((url) => {
          return cache.add(url).catch((err) => {
            console.warn('[SW] Failed to cache asset:', url, err);
          });
        })
      );
    })
  );
});

// ─── Activate: hapus cache versi lama ─────────────────────────────────────────
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) =>
      Promise.all(
        cacheNames.map((name) => {
          if (name !== CACHE_NAME && name !== 'velasto-dynamic-pages') return caches.delete(name);
        })
      )
    )
  );
  self.clients.claim();
});

// ─── Helper: fetch dengan batas waktu (AbortController) ───────────────────────
function fetchWithTimeout(request, ms) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), ms);
  return fetch(request, { signal: controller.signal })
    .then((res) => { clearTimeout(timeoutId); return res; })
    .catch((err) => { clearTimeout(timeoutId); throw err; });
}

// ─── Fetch: intercept semua request GET ───────────────────────────────────────
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  // Biarkan API calls & data diurus langsung (IndexedDB menangani offline)
  const url = event.request.url;
  if (url.includes('/ProductionReport/Get') ||
      url.includes('/Dimensi/Get') ||
      url.includes('/Dimensi/SaveData') ||
      url.includes('/OfflineSync/') ||
      url.includes('/api/')) {
    return;
  }

  const isNavigate = event.request.mode === 'navigate';

  if (isNavigate) {
    // NAVIGASI: Network-first DENGAN timeout 4 detik
    // Jika server tidak merespons dalam 4 detik → langsung pakai cache
    // Ini mencegah loading lama saat WiFi mati
    event.respondWith(
      fetchWithTimeout(event.request, NAV_TIMEOUT_MS)
        .then((networkResponse) => {
          // Berhasil dari network → update cache untuk offline berikutnya (JANGAN cache halaman login sebagai halaman lain)
          if (networkResponse.ok && !networkResponse.url.includes('/Account/Login')) {
            const clone = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return networkResponse;
        })
        .catch(() =>
          // Network gagal / timeout → ambil dari cache
          caches.match(event.request, { ignoreSearch: true }).then((cached) => {
            if (cached) {
              console.log('[SW] Offline: serving cached page:', event.request.url);
              return cached;
            }
            // Halaman tidak ada di cache → coba offline.html
            return caches.match('/offline.html').then((offlinePage) => {
              if (offlinePage) return offlinePage;
              // Emergency: offline.html juga tidak ada → jangan return undefined (ERR_FAILED!)
              return new Response(
                `<!DOCTYPE html><html lang="id">
                 <head><meta charset="UTF-8"><title>Mode Offline</title>
                 <style>body{font-family:Arial,sans-serif;background:#0a1628;color:#fff;
                 display:flex;flex-direction:column;align-items:center;justify-content:center;
                 height:100vh;margin:0;text-align:center;gap:16px;}
                 h1{font-size:2rem;margin:0;}p{opacity:.7;margin:0;}
                 button{background:#00c9d4;border:none;color:#000;padding:12px 28px;
                 border-radius:8px;font-size:1rem;cursor:pointer;margin-top:8px;}</style>
                 </head><body>
                 <div style="font-size:3rem">⚡</div>
                 <h1>Mode Offline</h1>
                 <p>Server tidak dapat dijangkau.</p>
                 <p>Buka halaman login untuk masuk secara offline.</p>
                 <button onclick="window.location.href='/Account/Login'">Ke Halaman Login</button>
                 <button onclick="window.location.reload()" style="background:#334;color:#fff;margin-left:8px">Coba Lagi</button>
                 </body></html>`,
                { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
              );
            });
          })
        )
    );
  } else {
    // ASET STATIS (CSS, JS, font, gambar): Cache-first
    event.respondWith(
      caches.match(event.request, { ignoreSearch: true }).then((cached) => {
        if (cached) return cached;
        return fetch(event.request).then((networkResponse) => {
          if (url.startsWith(self.location.origin)) {
            const clone = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return networkResponse;
        }).catch(() => { 
          // Jika aset tidak tersedia saat offline, return Response kosong agar tidak TypeError
          return new Response('', { status: 404, statusText: 'Not Found Offline' });
        });
      })
    );
  }
});
