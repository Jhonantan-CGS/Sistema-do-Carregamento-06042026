const CACHE_NAME = 'cysy-log360-v20260406r02';
const TILE_CACHE = 'cysy-map-tiles-v20260406r02';
const TILE_CACHE_LIMIT = 5000;
const APP_SHELL = [
  './',
  './index.html',
  './index.html?v=20260406r02',
  './launch.html',
  './launch.html?v=20260406r02',
  './app-shell.html',
  './app-shell.html?v=20260406r02',
  './offline.html',
  './manifest.webmanifest',
  './manifest.webmanifest?v=20260406r02',
  './assets/css/modern.css',
  './assets/css/modern.css?v=20260406r02',
  './assets/logo.svg',
  './assets/icons/icon-192.png',
  './assets/icons/icon-192.png?v=20260406r02',
  './assets/icons/icon-512.png',
  './assets/icons/icon-512.png?v=20260406r02',
  './assets/js/app-core-1.js',
  './assets/js/app-core-1.js?v=20260406r02',
  './assets/js/app-core-2.js',
  './assets/js/app-core-2.js?v=20260406r02',
  './assets/js/app-core-3.js',
  './assets/js/app-core-3.js?v=20260406r02',
  './assets/js/app-core-4.js',
  './assets/js/app-core-4.js?v=20260406r02',
  './assets/js/app-core-5.js',
  './assets/js/app-core-5.js?v=20260406r02',
  './assets/js/app-core-6.js',
  './assets/js/app-core-6.js?v=20260406r02',
  './assets/js/map-operacional-v2.js',
  './assets/js/map-operacional-v2.js?v=20260406r02',
  './assets/vendor/leaflet.css',
  './assets/vendor/leaflet.css?v=20260406r02',
  './assets/vendor/leaflet.js',
  './assets/vendor/leaflet.js?v=20260406r02'
];

const TILE_DOMAINS = [
  'server.arcgisonline.com',
  'tile.openstreetmap.org',
  'tiles.stadiamaps.com'
];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await Promise.allSettled(APP_SHELL.map((url) => cache.add(url).catch(() => {})));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => ![CACHE_NAME, TILE_CACHE].includes(k)).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  if (url.origin !== self.location.origin) {
    if (TILE_DOMAINS.some(domain => url.hostname.includes(domain))) {
      event.respondWith(handleTileRequest(req));
      return;
    }
    event.respondWith(handleExternalRequest(req));
    return;
  }

  event.respondWith(handleAppRequest(req));
});

async function handleTileRequest(req) {
  const tileCache = await caches.open(TILE_CACHE);
  const cached = await tileCache.match(req);

  fetch(req).then(response => {
    if (response && response.ok) {
      tileCache.put(req, response).catch(() => {});
    }
  }).catch(() => {});

  if (cached) return cached;

  try {
    const network = await fetch(req);
    if (network && network.ok) {
      tileCache.put(req, network.clone()).catch(() => {});
      trimTileCache(tileCache);
    }
    return network;
  } catch (_) {
    return Response.error();
  }
}

async function trimTileCache(cache) {
  try {
    const keys = await cache.keys();
    if (keys.length > TILE_CACHE_LIMIT) {
      const toDelete = keys.slice(0, keys.length - TILE_CACHE_LIMIT);
      await Promise.allSettled(toDelete.map(k => cache.delete(k)));
    }
  } catch (_) {}
}

async function handleExternalRequest(req) {
  try {
    return await fetch(req);
  } catch (_) {
    return caches.match(req) || Response.error();
  }
}

async function handleAppRequest(req) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(req);

  const isNavigation = req.mode === 'navigate';
  const reqUrl = new URL(req.url);
  const isVersionedStatic = reqUrl.searchParams.has('v') || /\.(css|js|svg|png|webmanifest)$/i.test(reqUrl.pathname);

  if (cached && (isNavigation || isVersionedStatic)) {
    fetch(req).then((network) => {
      if (network && network.status === 200) {
        cache.put(req, network.clone()).catch(() => {});
      }
    }).catch(() => {});
    return cached;
  }

  try {
    const network = await fetch(req);
    if (network && network.status === 200) {
      cache.put(req, network.clone()).catch(() => {});
    }
    return network;
  } catch (_) {
    if (isNavigation) {
      return cached || await caches.match('./launch.html?v=20260406r02') ||
        await caches.match('./launch.html') ||
        await caches.match('./app-shell.html?v=20260406r02') ||
        await caches.match('./app-shell.html') ||
        await caches.match('./index.html?v=20260406r02') ||
        await caches.match('./index.html') ||
        await caches.match('./offline.html');
    }
    return cached || Response.error();
  }
}

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil((async () => {
    const allClients = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    const target = allClients.find((client) => client.url.includes('/app-shell.html') || client.url.includes('/launch.html'));
    if (target) {
      await target.focus();
      return;
    }
    if (self.clients.openWindow) {
      await self.clients.openWindow('./app-shell.html?v=20260406r02');
    }
  })());
});
