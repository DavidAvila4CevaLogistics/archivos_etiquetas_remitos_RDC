// Service Worker - Control Etiquetas/Remitos Despacho RDC Troncal Rosario
// Estrategia: cache-first con precarga de todos los recursos para uso offline.
// Subí CACHE_NAME (v2 -> v3...) cada vez que cambien los archivos cacheados.

const CACHE_NAME = 'control-etiquetas-cache-v3';

const OFFLINE_URLS = [
    './',
    './index.html',
    './manifest.webmanifest',
    './css/styles.css',
    './js/app.js',
    './libs/xlsx.full.min.js',
    './libs/jspdf.umd.min.js',
    './icons/icon-192.png',
    './icons/icon-512.png',
    './icons/icon-512-maskable.png'
];

self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(OFFLINE_URLS))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(
                keys
                    .filter(key => key !== CACHE_NAME)
                    .map(key => caches.delete(key))
            )
        ).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', event => {
    if (event.request.method !== 'GET') return;

    event.respondWith(
        caches.match(event.request).then(cached => {
            if (cached) return cached;

            return fetch(event.request)
                .then(response => {
                    // Guarda en caché las respuestas válidas del mismo origen.
                    if (response.ok && new URL(event.request.url).origin === self.location.origin) {
                        const copy = response.clone();
                        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
                    }
                    return response;
                })
                .catch(() => caches.match('./index.html'));
        })
    );
});
