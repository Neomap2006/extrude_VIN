// wwwroot/sw.js
// SERVICE WORKER DINONAKTIFKAN — tidak ada caching halaman, tidak ada intercept request.
// File ini sengaja dikosongkan agar browser tidak memunculkan halaman Offline/Server Tidak Terjangkau.

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => {
    // Hapus semua cache lama agar tidak ada halaman cached yang muncul
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(keys.map(k => caches.delete(k)))
        ).then(() => self.clients.claim())
    );
});
// Tidak ada fetch handler — semua request langsung ke server seperti biasa.
