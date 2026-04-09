const CACHE_NAME = 'cysy-log360-v20260409r05';
const CACHE_TTL_MS = 12 * 60 * 60 * 1000;
const PWA_START_URL = './app-shell.html?source=pwa';
const APP_SHELL = [
  './',
  './index.html',
  './index.html?v=20260409r05',
  './launch.html',
  './launch.html?v=20260409r05',
  './app-shell.html',
  './app-shell.html?v=20260409r05',
  PWA_START_URL,
  './offline.html',
  './manifest.webmanifest',
  './manifest.webmanifest?v=20260409r05',
  './assets/css/modern.css',
  './assets/css/modern.css?v=20260409r05',
  './assets/css/theme-3d.css',
  './assets/css/theme-3d.css?v=20260409r05',
  './assets/logo.svg',
  './assets/icons/favicon.png',
  './assets/icons/favicon.png?v=20260409r05',
  './assets/icons/apple-touch-icon.png',
  './assets/icons/apple-touch-icon.png?v=20260409r05',
  './assets/icons/icon-192.png',
  './assets/icons/icon-192.png?v=20260409r05',
  './assets/icons/icon-512.png',
  './assets/icons/icon-512.png?v=20260409r05',
  './assets/icons/icon-512-maskable.png',
  './assets/icons/icon-512-maskable.png?v=20260409r05',
  './assets/js/app-core-1.js',
  './assets/js/app-core-1.js?v=20260409r05',
  './assets/js/app-core-2.js',
  './assets/js/app-core-2.js?v=20260409r05',
  './assets/js/app-core-3.js',
  './assets/js/app-core-3.js?v=20260409r05',
  './assets/js/app-core-5.js',
  './assets/js/app-core-5.js?v=20260409r05',
  './assets/js/app-core-6.js',
  './assets/js/app-core-6.js?v=20260409r05'
];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await Promise.allSettled(APP_SHELL.map((url) => cache.add(url).catch(() => {})));
    await cache.put('./__cache_meta__', new Response(JSON.stringify({ createdAt: Date.now() })));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(
      keys.filter((key) => key.startsWith('cysy-') && key !== CACHE_NAME).map((key) => caches.delete(key))
    );
    await ensureFreshCache();
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  if (url.origin !== self.location.origin) {
    event.respondWith(handleExternalRequest(req));
    return;
  }

  event.respondWith(handleAppRequest(req));
});

async function handleExternalRequest(req) {
  try {
    return await fetch(req);
  } catch (_) {
    return caches.match(req) || Response.error();
  }
}

async function handleAppRequest(req) {
  await ensureFreshCache();
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
      return cached || await caches.match(PWA_START_URL) ||
      await caches.match('./app-shell.html?v=20260409r05') ||
        await caches.match('./app-shell.html') ||
      await caches.match('./launch.html?v=20260409r05') ||
        await caches.match('./launch.html') ||
      await caches.match('./index.html?v=20260409r05') ||
        await caches.match('./index.html') ||
        await caches.match('./offline.html');
    }
    return cached || Response.error();
  }
}

async function ensureFreshCache() {
  const cache = await caches.open(CACHE_NAME);
  const metaRes = await cache.match('./__cache_meta__');
  let createdAt = 0;
  try {
    const meta = metaRes ? await metaRes.json() : null;
    createdAt = Number(meta?.createdAt || 0);
  } catch (_) {}
  if (createdAt && (Date.now() - createdAt) < CACHE_TTL_MS) return;
  const keys = await cache.keys();
  await Promise.allSettled(keys.map((key) => cache.delete(key)));
  await Promise.allSettled(APP_SHELL.map((url) => cache.add(url).catch(() => {})));
  await cache.put('./__cache_meta__', new Response(JSON.stringify({ createdAt: Date.now() })));
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
      await self.clients.openWindow(PWA_START_URL);
    }
  })());
});
