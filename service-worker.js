const CACHE_NAME = 'control-etiquetas-cache-v1';
const OFFLINE_URLS = [
    './',
    './index.html',
    './manifest.webmanifest'
    // agrega aquí /libs/xlsx.full.min.js, /libs/jspdf.umd.min.js si no usas CDN
];

self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(OFFLINE_URLS))
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
        )
    );
});

self.addEventListener('fetch', event => {
    event.respondWith(
        caches.match(event.request).then(resp => resp || fetch(event.request))
    );
});