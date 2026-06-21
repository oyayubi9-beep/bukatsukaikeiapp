// シンプルなアプリシェルキャッシュ。
// データはGoogle API経由なのでキャッシュせず、見た目（HTML/CSS/JS）だけ
// 高速表示・最低限のオフライン起動に対応させる。
const CACHE_NAME = "bukatsu-kaikei-v1";
const SHELL_FILES = [
  "./index.html",
  "./style.css",
  "./app.js",
  "./config.js",
  "./manifest.json",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  // Google APIへのリクエストは素通しする（キャッシュしない）
  if (url.origin.includes("googleapis.com") || url.origin.includes("google.com")) {
    return;
  }
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
