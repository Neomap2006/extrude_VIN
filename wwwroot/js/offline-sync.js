// wwwroot/js/offline-sync.js
// FITUR OFFLINE DINONAKTIFKAN — tidak ada IndexedDB, tidak ada Service Worker sync.
// File ini sengaja dikosongkan agar tidak ada error saat skrip lain memanggil fungsi ini.

// Stub window.db agar guard (if window.db && window.db.apiCache) tidak throw error
// Semua operasi akan selalu return null/false sehingga fallback ke fetch normal.
window.db = {
    apiCache: {
        get:     async function() { return null; },
        put:     async function() {},
        bulkPut: async function() {},
        count:   async function() { return 0; },
        delete:  async function() {}
    },
    syncQueue: {
        add:     async function() {},
        count:   async function() { return 0; },
        orderBy: function() { return { toArray: async function() { return []; } }; },
        delete:  async function() {},
        update:  async function() {}
    }
};

// Stub fungsi agar tidak ada ReferenceError di halaman lain
window.enqueueOfflineRequest     = async function() { console.info('[OfflineSync DISABLED]'); };
window.syncOfflineRequests       = async function() {};
window.autoSyncMasterData        = async function() {};
window.downloadOfflineMasterData = async function() {};
window.updateOfflineIndicator    = async function() {};
window.showSyncingIndicator      = function() {};
window.checkNetworkStatus        = async function() { return true; };
window.isAppOnline               = true;
