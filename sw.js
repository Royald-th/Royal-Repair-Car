/**
 * ============================================================
 *  Service Worker — ระบบแจ้งซ่อมรถยนต์ (Royal Repair Car PWA)
 * ============================================================
 *  กลยุทธ์:
 *  - App shell (html/css/js/icons)  → Cache First  (เร็ว + ใช้ได้ offline)
 *  - CDN libraries (bootstrap ฯลฯ) → Cache First  (URL มีเลขเวอร์ชันอยู่แล้ว)
 *  - Google Apps Script API calls  → Network Only (ห้าม cache ข้อมูลงานซ่อม)
 *  - Google Drive รูปภาพ           → Network Only (รูปเปลี่ยนได้ตลอด)
 *  - LINE LIFF SDK                  → Network First (ต้องได้เวอร์ชันล่าสุดเพื่อ auth)
 *
 *  เพิ่มเลขเวอร์ชัน CACHE_VERSION ทุกครั้งที่แก้ไฟล์ app shell (index.html/app.js/style.css)
 *  เพื่อบังคับให้ผู้ใช้ได้โค้ดใหม่ ไม่ใช่ไฟล์เก่าที่ค้างอยู่ใน cache
 * ============================================================ */

const CACHE_VERSION   = 'v2';
const APP_SHELL_CACHE = `repair-app-shell-${CACHE_VERSION}`;
const RUNTIME_CACHE    = `repair-runtime-${CACHE_VERSION}`;

// ไฟล์หลักของแอป — ใช้ path แบบ relative เพื่อรองรับ GitHub Pages subpath
const APP_SHELL_FILES = [
  './',
  './index.html',
  './app.js',
  './style.css',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './favicon.png',
];

// โดเมนที่ "ห้าม" cache เด็ดขาด — ข้อมูลสด/authentication เท่านั้น
const NETWORK_ONLY_HOSTS = [
  'script.google.com',       // Google Apps Script API (ข้อมูลงานซ่อมทั้งหมด)
  'script.googleusercontent.com',
  'drive.google.com',        // รูปภาพที่อัปโหลด (รูปแจ้งซ่อม/บิล/ใบประเมิน)
  'lh3.googleusercontent.com',
];

// โดเมนที่ต้อง network-first (พยายามโหลดสดก่อน ถ้าล้มเหลวค่อย fallback cache)
const NETWORK_FIRST_HOSTS = [
  'static.line-scdn.net',    // LINE LIFF SDK
  'access.line.me',
];

/* ============================================================
   INSTALL — precache app shell
   ============================================================ */
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(APP_SHELL_CACHE)
      .then((cache) => cache.addAll(APP_SHELL_FILES))
      .catch((err) => console.warn('[SW] precache failed (ไม่ใช่ปัญหาร้ายแรง):', err))
      .then(() => self.skipWaiting())
  );
});

/* ============================================================
   ACTIVATE — ลบ cache เวอร์ชันเก่าทิ้ง
   ============================================================ */
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== APP_SHELL_CACHE && key !== RUNTIME_CACHE)
          .map((key) => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

/* ============================================================
   FETCH — เลือกกลยุทธ์ตามประเภท request
   ============================================================ */
self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // ข้ามทุก request ที่ไม่ใช่ GET (POST ไป GAS ต้องผ่านตรงเสมอ)
  if (req.method !== 'GET') return;

  // 1) ห้าม cache เด็ดขาด — ข้อมูลสด / รูปภาพ Drive
  if (NETWORK_ONLY_HOSTS.some((h) => url.hostname.includes(h))) {
    event.respondWith(fetch(req).catch(() => networkErrorResponse()));
    return;
  }

  // 2) Network-first — LINE LIFF SDK
  if (NETWORK_FIRST_HOSTS.some((h) => url.hostname.includes(h))) {
    event.respondWith(networkFirst(req));
    return;
  }

  // 3) Navigation requests (เปิดหน้าเว็บ) — network-first พร้อม fallback เป็น cached index.html
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).catch(() => caches.match('./index.html'))
    );
    return;
  }

  // 4) ไฟล์ในโดเมนตัวเอง (app shell) — cache-first
  if (url.origin === self.location.origin) {
    event.respondWith(cacheFirst(req, APP_SHELL_CACHE));
    return;
  }

  // 5) CDN ภายนอกอื่นๆ (bootstrap, fonts, chart.js ฯลฯ) — cache-first + เติม runtime cache
  event.respondWith(cacheFirst(req, RUNTIME_CACHE));
});

/* ============================================================
   STRATEGIES
   ============================================================ */
async function cacheFirst(req, cacheName) {
  const cached = await caches.match(req);
  if (cached) return cached;
  try {
    const res = await fetch(req);
    // cache เฉพาะ response ที่สำเร็จ (status 200)
    if (res && res.status === 200) {
      const cache = await caches.open(cacheName);
      cache.put(req, res.clone());
    }
    return res;
  } catch (err) {
    return networkErrorResponse();
  }
}

async function networkFirst(req) {
  try {
    const res = await fetch(req);
    if (res && res.status === 200) {
      const cache = await caches.open(RUNTIME_CACHE);
      cache.put(req, res.clone());
    }
    return res;
  } catch (err) {
    const cached = await caches.match(req);
    return cached || networkErrorResponse();
  }
}

function networkErrorResponse() {
  return new Response(
    JSON.stringify({ status: 'error', message: 'ไม่มีการเชื่อมต่ออินเทอร์เน็ต' }),
    { status: 503, headers: { 'Content-Type': 'application/json' } }
  );
}

/* ============================================================
   MESSAGE — รองรับการสั่ง skipWaiting จากหน้าเว็บ (เมื่อกด "อัปเดตแอป")
   ============================================================ */
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
