const CACHE = 'treasures-v69';
const ASSETS = ['./', './index.html', './owner.html', './styles.css', './app.js', './manifest.webmanifest', './assets/treasures-icon.svg', './assets/reference/hero-products.png', './assets/reference/pantry.png', './assets/reference/grains.png', './assets/reference/household.png', './assets/reference/fresh.png', './assets/reference/package-products.png', './assets/premium-package.jpeg', './assets/healthy-package.jpeg', './assets/all-inclusive-package.jpeg'];
self.addEventListener('install', event => event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS))));
self.addEventListener('activate', event => event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE).map(key => caches.delete(key)))).then(() => self.clients.claim())));
self.addEventListener('fetch', event => event.respondWith(fetch(event.request).catch(() => caches.match(event.request, { ignoreSearch: true }))));

