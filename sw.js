// ごはんナビ：保存済みの画面をすぐ表示し、新しい版は裏で取り込む（データ通信はキャッシュしない）
const CACHE = "gohan-navi-v10";
const SHELL = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png", "./icons/apple-touch-icon.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
let pageCheck = Promise.resolve(false);
self.addEventListener("message", e => {
  if (!e.data || e.data.type !== "gn-check" || !e.source) return;
  e.waitUntil(pageCheck.then(ch => { if (ch) e.source.postMessage({ type: "gn-update" }); }));
});
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return; // 外部（店舗データ・同期）は素通し
  const page = e.request.mode === "navigate" || /\/(index\.html)?$/.test(u.pathname);
  const key = page ? "./index.html" : e.request;
  e.respondWith(caches.open(CACHE).then(async c => {
    const hit = await c.match(key);
    const prev = page && hit ? hit.clone() : null;   // 比べる用の写し（本体は画面に渡すので、先に取っておく）
    // 裏で最新を取りに行く（ブラウザの一時保存は使わず、サーバーに確認する）
    const net = fetch(page ? u.pathname : e.request, { cache: "no-cache" }).then(async r => {
      if (!r.ok) return page ? { r, changed: false } : r;
      let changed = false;
      if (prev) { const [a, b] = await Promise.all([prev.text(), r.clone().text()]); changed = a !== b; }
      await c.put(key, r.clone());
      return page ? { r, changed } : r;
    });
    // 画面を開いた直後に「新しい版が入ったか」を聞かれるので、結果を覚えておく
    if (page) pageCheck = net.then(x => x.changed === true, () => false);
    const res = net.then(x => (page && x.r) || x);
    if (hit) { e.waitUntil(res.catch(() => {})); return hit; }   // 保存済みがあれば、待たずにそれを出す
    return res.catch(() => c.match("./index.html"));
  }));
});
