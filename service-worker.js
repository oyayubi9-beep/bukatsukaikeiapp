// シンプルなアプリシェルキャッシュ。
// データはGoogle API経由なのでキャッシュせず、見た目（HTML/CSS/JS）だけ
// 高速表示・最低限のオフライン起動に対応させる。
//
// 【重要】index.html / app.js / config.js などを更新するたびに、
// このCACHE_NAMEの数字を1つ増やしてください（v2 → v3 など）。
// 増やさないと、スマホ側で古い内容がキャッシュされ続けて反映されません。
const CACHE_NAME = "bukatsu-kaikei-v3";
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

// ネットワーク優先：まず最新を取りに行き、取れたらキャッシュも更新する。
// オフラインなど取得に失敗したときだけ、キャッシュ済みの内容を表示する。
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  // Google APIへのリクエストは素通しする（キャッシュしない）
  if (url.origin.includes("googleapis.com") || url.origin.includes("google.com")) {
    return;
  }
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
