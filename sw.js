/* Launchpad Suite: lets the Suite be installed as an app and work offline.
   Pages and app files are kept so the apps open with no internet; files
   you make while offline are saved in the browser and move to your account
   when you're back online. Nothing personal is ever cached: sign-in and
   the database (hehs-live.duckdns.org) always go straight to the network. */
var VERSION = "suite-20261109b";
var SHELL = [
  "/", "/write/", "/grid/", "/present/", "/notebook/", "/pdf/", "/draw/", "/forms/", "/design/", "/vids/",
  "/office/office.css", "/office/office.js", "/office/live.js", "/office/jszip.min.js", "/office/odf.js", "/auth-store.js", "/favicon.png", "/manifest.webmanifest",
  "/write/write.js", "/write/docx.js", "/grid/grid.js", "/grid/formula.js", "/grid/xlsx.js",
  "/present/present.js", "/present/pptx.js", "/notebook/notebook.js", "/draw/draw.js", "/forms/forms.js", "/design/design.js", "/vids/vids.js", "/office/scene.js",
  "/pdf/pdf.js", "/pdf/lib/pdf.min.mjs", "/pdf/lib/pdf.worker.min.mjs", "/pdf/lib/pdf-lib.min.js",
  "/icons/suite-192.png", "/icons/suite-512.png", "/icons/apple-touch-icon.png"
];
var CDN = [
  "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.117.0/dist/umd/supabase.js",
  "https://civilinitiatives.org/glass.css"
];
self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(VERSION).then(function (c) {
    // one missing file must not stop the rest being kept
    return Promise.all(SHELL.concat(CDN).map(function (u) { return c.add(new Request(u, { cache: "reload" })).catch(function () {}); }));
  }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== VERSION; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
function cacheable(url) {
  if (url.origin === location.origin) return true;
  return url.href.indexOf("https://cdn.jsdelivr.net/npm/") === 0 || url.href.indexOf("https://cdnjs.cloudflare.com/") === 0 ||
    url.href.indexOf("https://civilinitiatives.org/glass.css") === 0;
}
self.addEventListener("fetch", function (e) {
  var req = e.request;
  if (req.method !== "GET") return;
  var url = new URL(req.url);
  if (!cacheable(url)) return;
  if (req.mode === "navigate") {
    // pages: the newest from the network, the kept copy when offline
    e.respondWith(fetch(req).then(function (res) {
      if (res.ok) { var copy = res.clone(); caches.open(VERSION).then(function (c) { c.put(url.pathname, copy); }); }
      return res;
    }).catch(function () {
      return caches.match(url.pathname, { ignoreSearch: true }).then(function (r) { return r || caches.match("/", { ignoreSearch: true }); });
    }));
    return;
  }
  // files: the newest from the network (so a new version never mixes
  // with old files), the kept copy when offline
  e.respondWith(fetch(req).then(function (res) {
    if (res.ok && (res.type === "basic" || res.type === "cors")) { var copy = res.clone(); caches.open(VERSION).then(function (c) { c.put(req, copy); }); }
    return res;
  }).catch(function () { return caches.match(req, { ignoreSearch: true }); }));
});
