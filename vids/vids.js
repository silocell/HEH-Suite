/* Launchpad Vids - make a video from scenes, like Google Vids.
   Each scene has text, stickers, shapes, photos or video clips that come in
   with animations, a caption, and maybe your voice. Add music, record
   yourself, then make an .mp4 / .webm video.

   Everything stays on this device (in the browser's own storage):
   videos and recordings are too big to keep in the Suite's shared space. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var S = window.Scene, esc = Office.esc;
  var cv = $("cv"), ctx = cv.getContext("2d"), ov = $("ov"), screenEl = $("screen"), stage = $("stage");
  var proj = null, cur = 0, sel = null, t = 0, playing = false, mode = "edit", zoom = 1, panel = "scene", editing = null;
  var undo = [], redo = [];
  var meas = document.createElement("canvas").getContext("2d");
  var TD = 0.5; // how long a transition takes
  function uid() { return Math.random().toString(36).slice(2, 10); }
  function SC() { return proj.scenes[cur]; }
  function find(id) { var e = null; (SC().els || []).forEach(function (x) { if (x.id === id) e = x; }); return e; }
  function selEl() { return sel ? find(sel) : null; }
  function fmt(s) { s = Math.max(0, s); var m = Math.floor(s / 60), r = Math.floor(s % 60); return m + ":" + (r < 10 ? "0" : "") + r; }
  function total() { return proj.scenes.reduce(function (n, s) { return n + s.dur; }, 0); }
  function startOf(i) { var n = 0; for (var k = 0; k < i; k++) n += proj.scenes[k].dur; return n; }
  function sceneAt(tt) { var n = 0; for (var i = 0; i < proj.scenes.length; i++) { if (tt < n + proj.scenes[i].dur || i === proj.scenes.length - 1) return { i: i, lt: Math.max(0, tt - n) }; n += proj.scenes[i].dur; } return { i: 0, lt: 0 }; }

  /* ---------------- keeping things on this device (IndexedDB) ---------------- */
  var DB = null, memory = { projects: {}, media: {} };
  function db() {
    if (DB) return DB;
    DB = new Promise(function (ok) {
      if (!window.indexedDB) { ok(null); return; }
      var r; try { r = indexedDB.open("hl-vids", 1); } catch (e) { ok(null); return; }
      r.onupgradeneeded = function () { var d = r.result; d.createObjectStore("projects", { keyPath: "id" }); var m = d.createObjectStore("media", { keyPath: "id" }); m.createIndex("pid", "pid"); };
      r.onsuccess = function () { ok(r.result); };
      r.onerror = r.onblocked = function () { ok(null); };
    });
    return DB;
  }
  function tx(store, mode, fn) {
    return db().then(function (d) {
      if (!d) return fn(null);
      return new Promise(function (ok, no) {
        var x = d.transaction(store, mode), st = x.objectStore(store), res = fn(st);
        x.oncomplete = function () { ok(res && res.result !== undefined ? res.result : res); };
        x.onerror = x.onabort = function () { no(x.error || new Error("Storage error")); };
      });
    });
  }
  function put(store, v) { return tx(store, "readwrite", function (st) { if (!st) { memory[store][v.id] = v; return; } return st.put(v); }); }
  function getOne(store, id) { return tx(store, "readonly", function (st) { if (!st) return { result: memory[store][id] }; return st.get(id); }); }
  function all(store) { return tx(store, "readonly", function (st) { if (!st) return { result: Object.keys(memory[store]).map(function (k) { return memory[store][k]; }) }; return st.getAll(); }); }
  function delOne(store, id) { return tx(store, "readwrite", function (st) { if (!st) { delete memory[store][id]; return; } return st.delete(id); }); }
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () {});

  /* ---------------- media: photos, clips, sounds ---------------- */
  var MEDIA = {}; // id -> { id, kind, name, url, el, dur }
  function mediaOf(src) { return src && src.indexOf("idb:") === 0 ? MEDIA[src.slice(4)] : null; }
  function loadMedia(rec) {
    if (MEDIA[rec.id]) return Promise.resolve(MEDIA[rec.id]);
    var m = { id: rec.id, kind: rec.kind, name: rec.name || rec.kind, url: URL.createObjectURL(rec.blob), dur: rec.dur || 0, pid: rec.pid };
    MEDIA[rec.id] = m;
    return new Promise(function (ok) {
      if (m.kind === "image") { var im = new Image(); im.onload = im.onerror = function () { ok(m); }; im.src = m.url; m.el = im; }
      else if (m.kind === "video" || m.kind === "audio") {
        var v = document.createElement(m.kind); v.preload = "auto"; v.playsInline = true; v.setAttribute("playsinline", ""); v.src = m.url; m.el = v;
        if (m.kind === "video") v.muted = false;
        var done = false; function fin() { if (done) return; done = true; if (isFinite(v.duration) && v.duration > 0) m.dur = v.duration; ok(m); }
        v.onloadedmetadata = function () {
          // some recordings say their length is Infinity until you look at the end
          if (!isFinite(v.duration)) { v.currentTime = 1e7; v.ontimeupdate = function () { v.ontimeupdate = null; v.currentTime = 0; fin(); }; }
          else fin();
        };
        v.onerror = fin; setTimeout(fin, 6000);
      } else ok(m);
    });
  }
  function storeMedia(blob, kind, name, dur) {
    var rec = { id: uid() + uid().slice(0, 4), pid: proj.id, blob: blob, kind: kind, name: name || kind, dur: dur || 0, created: Date.now() };
    return put("media", rec).then(function () { return loadMedia(rec); });
  }
  function loadProjectMedia(pid) {
    return all("media").then(function (rows) {
      return Promise.all((rows || []).filter(function (r) { return r.pid === pid; }).map(loadMedia));
    });
  }
  /* each video clip on a scene gets its own player, so two clips can play together */
  var CLIPS = {};
  function clipFor(el) {
    var m = mediaOf(el.src); if (!m || m.kind !== "video") return null;
    var k = el.id + "|" + m.id, v = CLIPS[k];
    if (!v) { v = document.createElement("video"); v.src = m.url; v.preload = "auto"; v.playsInline = true; v.setAttribute("playsinline", ""); v.muted = !!el.mute; CLIPS[k] = v; v.onseeked = v.onloadeddata = function () { if (!playing) draw(); }; }
    return v;
  }

  /* ---------------- animations ---------------- */
  var ANIMS = [["none", "None"], ["fade", "Fade in"], ["rise", "Rise up"], ["drop", "Drop in"], ["left", "From the left"], ["right", "From the right"],
    ["zoom", "Zoom in"], ["pop", "Pop"], ["spin", "Spin in"], ["type", "Typewriter"], ["kb", "Slow zoom (photos)"], ["pulse", "Pulse"]];
  function animState(el, lt, dur) {
    var at = el.at || 0;
    if (lt < at) return { o: 0 };
    var k = el.anim || "none", p = Math.min(1, (lt - at) / 0.6), e = 1 - Math.pow(1 - p, 3), W = proj.W, H = proj.H;
    switch (k) {
      case "fade": return { o: e };
      case "rise": return { o: e, dy: (1 - e) * H * 0.08 };
      case "drop": return { o: e, dy: -(1 - e) * H * 0.08 };
      case "left": return { o: e, dx: -(1 - e) * W * 0.25 };
      case "right": return { o: e, dx: (1 - e) * W * 0.25 };
      case "zoom": return { o: e, s: 0.5 + 0.5 * e };
      case "pop": return { o: Math.min(1, p * 3), s: (0.3 + 0.7 * e) * (1 + Math.sin(p * Math.PI) * 0.18) };
      case "spin": return { o: e, r: (1 - e) * -180, s: 0.4 + 0.6 * e };
      case "type": return el.t === "text" ? { chars: Math.floor((lt - at) * 26) } : { o: e };
      case "kb": return { s: 1 + 0.12 * Math.min(1, lt / Math.max(1, dur)) };
      case "pulse": return { s: 1 + 0.05 * Math.sin((lt - at) * Math.PI * 2) };
    }
    return null;
  }

  /* ---------------- drawing a moment of the video ---------------- */
  function bgFor(bg) { if (bg && bg.img && bg.img.indexOf("idb:") === 0) { var m = mediaOf(bg.img); return m ? { img: m.url, dim: bg.dim } : { c: "#000000" }; } return bg; }
  function drawScene(g, sc, lt, animate, skip) {
    var W = proj.W, H = proj.H;
    S.drawPage(g, { bg: bgFor(sc.bg), els: sc.els }, W, H, {
      onload: draw, skip: skip, clean: animate,
      anim: animate ? function (el) { return animState(el, lt, sc.dur); } : null,
      media: function (el) {
        if (el.t === "video") return clipFor(el);
        if (el.t === "img") { var m = mediaOf(el.src); return m ? m.el : null; }
        return null;
      }
    });
    if (sc.cap) caption(g, sc.cap, W, H);
  }
  function caption(g, text, W, H) {
    var size = Math.round(Math.min(W, H) * 0.045), e = { t: "text", text: text, font: "Inter", bold: true, size: size, w: W * 0.84, lh: 1.25, align: "center", color: "#ffffff", x: 0, y: 0, h: 10 };
    var lines = S.wrap(g, e), lh = size * 1.25, bh = lines.length * lh + size * 0.6, y = H - bh - H * 0.05;
    g.save(); g.fillStyle = "rgba(0,0,0,.62)";
    var bw = 0; g.font = S.fontOf(e); lines.forEach(function (l) { bw = Math.max(bw, g.measureText(l).width); });
    bw += size * 1.2;
    var x = (W - bw) / 2; g.translate(x, y); S.shapePath(g, "rect", bw, bh, size * 0.4); g.fill(); g.translate(-x, -y);
    e.x = (W - e.w) / 2; e.y = y + size * 0.3; e.h = bh; S.drawEl(g, e);
    g.restore();
  }
  var offA = document.createElement("canvas"), offB = document.createElement("canvas");
  function frame(g, tt, scale) {
    var W = proj.W, H = proj.H, at = sceneAt(tt), sc = proj.scenes[at.i];
    g.save(); g.fillStyle = "#000"; g.fillRect(0, 0, W, H);
    var tr = sc.trans || "none";
    if (at.i > 0 && tr !== "none" && at.lt < TD) {
      // a transition: the last moment of the scene before, and this one coming in
      var prev = proj.scenes[at.i - 1], p = at.lt / TD, e = p * p * (3 - 2 * p);
      [offA, offB].forEach(function (c) { if (c.width !== Math.round(W * scale) || c.height !== Math.round(H * scale)) { c.width = Math.round(W * scale); c.height = Math.round(H * scale); } });
      var ga = offA.getContext("2d"), gb = offB.getContext("2d");
      ga.setTransform(scale, 0, 0, scale, 0, 0); gb.setTransform(scale, 0, 0, scale, 0, 0);
      drawScene(ga, prev, prev.dur, true); drawScene(gb, sc, at.lt, true);
      var sw = W, sh = H;
      g.drawImage(offA, 0, 0, sw, sh);
      g.save();
      if (tr === "fade") { g.globalAlpha = e; g.drawImage(offB, 0, 0, sw, sh); }
      else if (tr === "slide") { g.drawImage(offB, (1 - e) * W, 0, sw, sh); }
      else if (tr === "up") { g.drawImage(offB, 0, (1 - e) * H, sw, sh); }
      else if (tr === "zoom") { var z = 0.6 + 0.4 * e; g.globalAlpha = e; g.drawImage(offB, W * (1 - z) / 2, H * (1 - z) / 2, W * z, H * z); }
      else if (tr === "wipe") { g.beginPath(); g.rect(0, 0, W * e, H); g.clip(); g.drawImage(offB, 0, 0, sw, sh); }
      else if (tr === "circle") { g.beginPath(); g.arc(W / 2, H / 2, Math.hypot(W, H) / 2 * e, 0, Math.PI * 2); g.clip(); g.drawImage(offB, 0, 0, sw, sh); }
      else g.drawImage(offB, 0, 0, sw, sh);
      g.restore();
    } else drawScene(g, sc, at.lt, true);
    g.restore();
  }
  var dpr = Math.min(2, window.devicePixelRatio || 1), rq = 0;
  function draw() { if (rq) return; rq = requestAnimationFrame(function () { rq = 0; drawNow(); }); }
  function drawNow() {
    if (!proj) return;
    var W = proj.W, H = proj.H, sw = Math.round(W * zoom), sh = Math.round(H * zoom);
    if (cv.width !== Math.round(sw * dpr) || cv.height !== Math.round(sh * dpr)) { cv.width = Math.round(sw * dpr); cv.height = Math.round(sh * dpr); }
    cv.style.width = sw + "px"; cv.style.height = sh + "px"; screenEl.style.width = sw + "px"; screenEl.style.height = sh + "px";
    ctx.setTransform(zoom * dpr, 0, 0, zoom * dpr, 0, 0);
    if (mode === "edit") {
      ctx.fillStyle = "#000"; ctx.fillRect(0, 0, W, H);
      drawScene(ctx, SC(), SC().dur, false, editing);
      // in edit mode clips show their first frame
      SC().els.forEach(function (el) { if (el.t === "video") { var v = clipFor(el); if (v && !playing && Math.abs(v.currentTime - (el.from || 0)) > 0.05 && v.readyState >= 1) try { v.currentTime = el.from || 0; } catch (e) {} } });
    } else frame(ctx, t, zoom * dpr);
    selBox();
  }
  function fitZoom() {
    if (!proj) return;
    var r = stage.getBoundingClientRect(), pad = window.innerWidth < 560 ? 24 : 48;
    zoom = Math.max(0.05, Math.min((r.width - pad) / proj.W, (r.height - pad) / proj.H, 1.5));
  }

  /* ---------------- the project ---------------- */
  var RATIOS = { "16:9": [1280, 720], "9:16": [720, 1280], "1:1": [1080, 1080] };
  function blankScene() { return { id: uid(), dur: 4, bg: { c: "#111827" }, trans: "fade", cap: "", els: [] }; }
  function clean(b) {
    var r = RATIOS[b && b.ratio] ? b.ratio : "16:9", d = { ratio: r, W: RATIOS[r][0], H: RATIOS[r][1], scenes: [], music: null };
    var seen = {};
    ((b && b.scenes) || []).slice(0, 80).forEach(function (s, i) {
      if (!s || typeof s !== "object") return;
      var id = String(s.id || "").replace(/[^a-z0-9]/gi, "").slice(0, 16); if (!id || seen[id]) id = "s" + i + uid().slice(0, 4); seen[id] = 1;
      var o = { id: id, dur: S.num(s.dur, 4, 0.5, 120), bg: S.cleanBg(s.bg), trans: /^(none|fade|slide|up|zoom|wipe|circle)$/.test(s.trans) ? s.trans : "none",
        cap: String(s.cap || "").slice(0, 400), els: (Array.isArray(s.els) ? s.els : []).slice(0, 120).map(S.cleanEl).filter(Boolean) };
      if (s.voice && /^idb:[a-z0-9]+$/i.test(s.voice)) { o.voice = s.voice; o.vvol = S.num(s.vvol, 1, 0, 1); }
      d.scenes.push(o);
    });
    if (!d.scenes.length) d.scenes.push(blankScene());
    if (b && b.music && /^idb:[a-z0-9]+$/i.test(b.music.src)) d.music = { src: b.music.src, vol: S.num(b.music.vol, 0.4, 0, 1), name: String(b.music.name || "Music").slice(0, 80) };
    return d;
  }
  var saveT = null, meta = { id: null, title: "Untitled video" };
  function changed() {
    clearTimeout(saveT); $("status").textContent = "Editing…";
    saveT = setTimeout(saveNow, 700); tlSoon();
  }
  function saveNow() {
    clearTimeout(saveT); if (!proj || !meta.id) return Promise.resolve();
    meta.title = $("title").value.trim() || "Untitled video";
    var rec = { id: meta.id, title: meta.title, updated: Date.now(), body: proj, thumb: thumbURL() };
    return put("projects", rec).then(function () { $("status").textContent = "✓ Saved on this device"; }, function (e) { $("status").textContent = "⚠ Not saved"; Office.toast("Couldn't save: " + (e.message || e), true); });
  }
  function thumbURL() {
    try {
      var c = document.createElement("canvas"), w = 192, h = Math.round(192 * proj.H / proj.W); c.width = w; c.height = h;
      var g = c.getContext("2d"); g.scale(w / proj.W, h / proj.H); drawScene(g, proj.scenes[0], proj.scenes[0].dur, false); return c.toDataURL("image/jpeg", 0.7);
    } catch (e) { return ""; }
  }
  function snap() { undo.push(JSON.stringify(proj)); if (undo.length > 60) undo.shift(); redo = []; }
  function doUndo() { if (!undo.length) return; redo.push(JSON.stringify(proj)); proj = clean(JSON.parse(undo.pop())); after(); }
  function doRedo() { if (!redo.length) return; undo.push(JSON.stringify(proj)); proj = clean(JSON.parse(redo.pop())); after(); }
  function after() { cur = Math.min(cur, proj.scenes.length - 1); if (sel && !find(sel)) sel = null; stopEdit(); fitZoom(); draw(); timeline(); ctxBar(); showPanel(); changed(); }

  /* ---------------- templates ---------------- */
  /* x, y, w in fractions of the screen; s = text size as a fraction of the shorter side */
  var TPLS = [
    { n: "Book trailer", d: "Build suspense for a book you love", sc: [
      { dur: 3.5, bg: { c: "#000000" }, trans: "fade", els: [{ t: "text", text: "In a world where nothing is as it seems…", x: .1, y: .4, w: .8, s: .07, font: "Playfair Display", italic: 1, color: "#f5f5f4", align: "center", anim: "type" }] },
      { dur: 3, bg: { g: ["#1c1917", "#7c2d12", 160] }, trans: "zoom", els: [{ t: "emoji", ch: "📖", x: .4, y: .2, h: .3, anim: "pop" }, { t: "text", text: "One book will change everything", x: .1, y: .58, w: .8, s: .07, font: "Bebas Neue", color: "#fed7aa", align: "center", anim: "rise", at: .4 }] },
      { dur: 3.5, bg: { c: "#0c0a09" }, trans: "fade", cap: "Say the title and the author here", els: [{ t: "text", text: "TITLE OF THE BOOK", x: .08, y: .34, w: .84, s: .12, font: "Anton", color: "#fbbf24", align: "center", anim: "zoom", sh: 1 }, { t: "text", text: "by Author Name", x: .1, y: .6, w: .8, s: .05, font: "Inter", color: "#e7e5e4", align: "center", anim: "fade", at: .8 }] },
      { dur: 3, bg: { g: ["#7c2d12", "#000000", 180] }, trans: "circle", els: [{ t: "text", text: "Read it at your library", x: .1, y: .42, w: .8, s: .07, font: "Poppins", bold: 1, color: "#ffffff", align: "center", anim: "rise" }, { t: "emoji", ch: "⭐", x: .45, y: .64, h: .12, anim: "spin", at: .5 }] }] },
    { n: "Class news", d: "A news report with headlines and weather", sc: [
      { dur: 3, bg: { g: ["#1d4ed8", "#0f172a", 160] }, trans: "none", els: [{ t: "shape", shape: "rect", x: 0, y: .62, w: 1, h: .16, fill: "#dc2626", anim: "left" }, { t: "text", text: "7B NEWS", x: .06, y: .64, w: .6, s: .1, font: "Anton", color: "#ffffff", anim: "left", at: .2 }, { t: "emoji", ch: "📺", x: .7, y: .18, h: .3, anim: "pop" }] },
      { dur: 4, bg: { c: "#f8fafc" }, trans: "slide", cap: "Read out your top story", els: [{ t: "text", text: "TOP STORY", x: .06, y: .1, w: .5, s: .05, font: "Inter", bold: 1, color: "#dc2626", ls: 4 }, { t: "text", text: "Our class won the reading challenge!", x: .06, y: .2, w: .6, s: .08, font: "Poppins", bold: 1, color: "#0f172a", anim: "rise" }, { t: "emoji", ch: "🏆", x: .72, y: .3, h: .35, anim: "pop", at: .5 }] },
      { dur: 4, bg: { c: "#ecfccb" }, trans: "slide", cap: "What happened in sport this week", els: [{ t: "text", text: "SPORT", x: .06, y: .1, w: .5, s: .05, font: "Inter", bold: 1, color: "#15803d", ls: 4 }, { t: "text", text: "Football team scores 3 goals", x: .06, y: .2, w: .6, s: .08, font: "Poppins", bold: 1, color: "#14532d", anim: "rise" }, { t: "emoji", ch: "⚽", x: .72, y: .3, h: .3, anim: "spin", at: .4 }] },
      { dur: 3.5, bg: { g: ["#38bdf8", "#e0f2fe", 180] }, trans: "slide", cap: "And now the weather", els: [{ t: "text", text: "WEATHER", x: .06, y: .1, w: .5, s: .05, font: "Inter", bold: 1, color: "#075985", ls: 4 }, { t: "emoji", ch: "☀️", x: .12, y: .3, h: .3, anim: "pulse" }, { t: "text", text: "Sunny, 18°", x: .45, y: .4, w: .5, s: .1, font: "Poppins", bold: 1, color: "#0c4a6e", anim: "fade", at: .3 }] },
      { dur: 3, bg: { g: ["#1d4ed8", "#0f172a", 160] }, trans: "fade", els: [{ t: "text", text: "That's all for today!", x: .1, y: .4, w: .8, s: .08, font: "Poppins", bold: 1, color: "#ffffff", align: "center", anim: "pop" }] }] },
    { n: "How-to", d: "Explain how to do something, step by step", sc: [
      { dur: 3, bg: { c: "#fef3c7" }, trans: "none", els: [{ t: "text", text: "How to make a paper plane", x: .08, y: .3, w: .6, s: .09, font: "Fredoka", bold: 1, color: "#92400e", anim: "rise" }, { t: "emoji", ch: "✈️", x: .7, y: .3, h: .3, anim: "right", at: .3 }] },
      { dur: 4, bg: { c: "#ffffff" }, trans: "slide", cap: "Fold the paper in half, long way", els: [{ t: "shape", shape: "circle", x: .06, y: .12, w: .1, h: .1 * 16 / 9, fill: "#f59e0b", anim: "pop" }, { t: "text", text: "1", x: .06, y: .15, w: .1, s: .08, font: "Poppins", bold: 1, color: "#ffffff", align: "center", anim: "pop" }, { t: "text", text: "Fold it in half", x: .2, y: .14, w: .7, s: .08, font: "Fredoka", bold: 1, color: "#1c1917", anim: "left", at: .2 }, { t: "emoji", ch: "📄", x: .4, y: .4, h: .4, anim: "zoom", at: .5 }] },
      { dur: 4, bg: { c: "#ffffff" }, trans: "slide", cap: "Fold the top corners into the middle", els: [{ t: "shape", shape: "circle", x: .06, y: .12, w: .1, h: .1 * 16 / 9, fill: "#f59e0b", anim: "pop" }, { t: "text", text: "2", x: .06, y: .15, w: .1, s: .08, font: "Poppins", bold: 1, color: "#ffffff", align: "center", anim: "pop" }, { t: "text", text: "Fold the corners in", x: .2, y: .14, w: .7, s: .08, font: "Fredoka", bold: 1, color: "#1c1917", anim: "left", at: .2 }, { t: "emoji", ch: "📐", x: .4, y: .4, h: .4, anim: "zoom", at: .5 }] },
      { dur: 4, bg: { c: "#ffffff" }, trans: "slide", cap: "Fold the wings down and throw it!", els: [{ t: "shape", shape: "circle", x: .06, y: .12, w: .1, h: .1 * 16 / 9, fill: "#f59e0b", anim: "pop" }, { t: "text", text: "3", x: .06, y: .15, w: .1, s: .08, font: "Poppins", bold: 1, color: "#ffffff", align: "center", anim: "pop" }, { t: "text", text: "Fold the wings, then fly!", x: .2, y: .14, w: .7, s: .08, font: "Fredoka", bold: 1, color: "#1c1917", anim: "left", at: .2 }, { t: "emoji", ch: "🛫", x: .4, y: .4, h: .4, anim: "right", at: .5 }] },
      { dur: 3, bg: { g: ["#f59e0b", "#ef4444", 135] }, trans: "zoom", els: [{ t: "text", text: "Your turn!", x: .1, y: .38, w: .8, s: .14, font: "Fredoka", bold: 1, color: "#ffffff", align: "center", anim: "pop", sh: 1 }] }] },
    { n: "Science explainer", d: "Answer a big question in a few scenes", sc: [
      { dur: 3.5, bg: { g: ["#0f172a", "#1e3a8a", 160] }, trans: "none", els: [{ t: "text", text: "Why is the sky blue?", x: .1, y: .36, w: .8, s: .1, font: "Poppins", bold: 1, color: "#e0f2fe", align: "center", anim: "type" }, { t: "emoji", ch: "🤔", x: .44, y: .62, h: .18, anim: "pop", at: 1.2 }] },
      { dur: 4, bg: { c: "#fffbeb" }, trans: "fade", cap: "Sunlight is made of every colour of the rainbow", els: [{ t: "emoji", ch: "☀️", x: .12, y: .25, h: .4, anim: "pulse" }, { t: "emoji", ch: "🌈", x: .55, y: .25, h: .4, anim: "right", at: .5 }] },
      { dur: 4, bg: { c: "#dbeafe" }, trans: "fade", cap: "Blue light bounces off the air the most, so it spreads across the sky", els: [{ t: "text", text: "Blue light scatters the most", x: .1, y: .14, w: .8, s: .075, font: "Poppins", bold: 1, color: "#1e3a8a", align: "center", anim: "rise" }, { t: "shape", shape: "circle", x: .3, y: .45, w: .08, h: .08 * 16 / 9, fill: "#3b82f6", anim: "pop", at: .5 }, { t: "shape", shape: "circle", x: .46, y: .55, w: .06, h: .06 * 16 / 9, fill: "#60a5fa", anim: "pop", at: .7 }, { t: "shape", shape: "circle", x: .6, y: .42, w: .07, h: .07 * 16 / 9, fill: "#2563eb", anim: "pop", at: .9 }] },
      { dur: 3, bg: { g: ["#38bdf8", "#0284c7", 180] }, trans: "zoom", els: [{ t: "text", text: "So the sky looks blue!", x: .1, y: .4, w: .8, s: .1, font: "Poppins", bold: 1, color: "#ffffff", align: "center", anim: "pop", sh: 1 }] }] },
    { n: "All about me", d: "Introduce yourself in four scenes", sc: [
      { dur: 3, bg: { g: ["#a855f7", "#ec4899", 135] }, trans: "none", els: [{ t: "text", text: "Hi, I'm …", x: .1, y: .36, w: .8, s: .12, font: "Pacifico", color: "#ffffff", align: "center", anim: "pop" }, { t: "emoji", ch: "👋", x: .45, y: .64, h: .18, anim: "pulse", at: .5 }] },
      { dur: 3.5, bg: { c: "#fdf4ff" }, trans: "slide", cap: "Say what you love doing", els: [{ t: "text", text: "I love …", x: .08, y: .3, w: .55, s: .1, font: "Fredoka", bold: 1, color: "#86198f", anim: "left" }, { t: "emoji", ch: "⚽", x: .68, y: .3, h: .3, anim: "spin", at: .3 }] },
      { dur: 3.5, bg: { c: "#ecfeff" }, trans: "slide", cap: "Share something you want to do one day", els: [{ t: "text", text: "One day I want to …", x: .08, y: .3, w: .55, s: .09, font: "Fredoka", bold: 1, color: "#155e75", anim: "left" }, { t: "emoji", ch: "🚀", x: .68, y: .3, h: .3, anim: "rise", at: .3 }] },
      { dur: 3, bg: { g: ["#ec4899", "#a855f7", 135] }, trans: "circle", els: [{ t: "text", text: "Thanks for watching!", x: .1, y: .4, w: .8, s: .09, font: "Poppins", bold: 1, color: "#ffffff", align: "center", anim: "rise" }] }] }
  ];
  function fromTpl(tp) {
    /* templates are laid out for a wide screen; on a tall or square one they
       sit in a wide band across the middle (the background fills it all) */
    var W = proj.W, H = proj.H, bw = W, bh = Math.min(H, W * 9 / 16), y0 = (H - bh) / 2, big = H > W ? 1.15 : 1, m = Math.min(bw, bh) * big;
    return tp.sc.map(function (s) {
      return { id: uid(), dur: s.dur, bg: JSON.parse(JSON.stringify(s.bg)), trans: s.trans, cap: s.cap || "", els: s.els.map(function (e) {
        var o = { id: uid(), t: e.t, x: e.x * bw, y: y0 + e.y * bh, w: (e.w || .2) * bw, h: (e.h || .1) * bh, r: 0, o: 1 };
        ["shape", "fill", "font", "color", "align", "text", "ch", "anim", "at"].forEach(function (k) { if (e[k] != null) o[k] = e[k]; });
        ["bold", "italic", "sh"].forEach(function (k) { if (e[k]) o[k] = true; }); if (e.ls) o.ls = e.ls * m / 720;
        if (e.t === "text") { o.size = Math.round(e.s * m); o.h = S.textHeight(meas, o); if (big > 1 && e.align === "center") { o.y -= (o.h - (e.s * Math.min(bw, bh) * 1.2)) / 2; } }
        if (e.t === "emoji") { o.h = e.h * bh * big; o.w = o.h; o.x = Math.min(o.x, W - o.w); }
        if (e.t === "shape" && e.shape === "circle") { var c = Math.min(o.w, o.h); o.w = o.h = c; }
        return o;
      }) };
    });
  }

  /* ---------------- the box round what's picked ---------------- */
  var box = document.createElement("div"); box.id = "selbox"; box.hidden = true;
  ["nw", "ne", "se", "sw"].forEach(function (h, i) { var d = document.createElement("div"); d.className = "hd"; d.dataset.h = h; d.style.left = (i === 1 || i === 2 ? 100 : 0) + "%"; d.style.top = (i >= 2 ? 100 : 0) + "%"; box.appendChild(d); });
  screenEl.appendChild(box);
  function selBox() {
    var e = mode === "edit" ? selEl() : null;
    if (!e) { box.hidden = true; return; }
    box.hidden = false;
    box.style.left = e.x * zoom + "px"; box.style.top = e.y * zoom + "px"; box.style.width = e.w * zoom + "px"; box.style.height = e.h * zoom + "px";
    box.style.transform = "rotate(" + (e.r || 0) + "deg)";
  }
  function toPage(ev) { var r = cv.getBoundingClientRect(); return { x: (ev.clientX - r.left) / zoom, y: (ev.clientY - r.top) / zoom }; }
  function topAt(p) { var els = SC().els; for (var i = els.length - 1; i >= 0; i--) if (S.hit(els[i], p.x, p.y, 4 / zoom)) return els[i]; return null; }
  function select(id) { if (editing && editing !== id) stopEdit(); sel = id; selBox(); ctxBar(); }
  var act = null;
  ov.addEventListener("pointerdown", function (ev) {
    if (ev.button > 0) return;
    if (playing) pause();
    if (mode !== "edit") { mode = "edit"; var at = sceneAt(t); cur = at.i; timeline(); }
    ov.focus({ preventScroll: true });
    var p = toPage(ev), e = topAt(p);
    if (editing && e && e.id === editing) return;
    if (!e) { select(null); stopEdit(); draw(); return; }
    if (sel !== e.id) select(e.id);
    act = { kind: "move", id: e.id, sx: p.x, sy: p.y, ox: e.x, oy: e.y, moved: false };
    ov.setPointerCapture(ev.pointerId); draw();
  });
  box.addEventListener("pointerdown", function (ev) {
    var h = ev.target.dataset && ev.target.dataset.h, e = selEl(); if (!h || !e) return;
    ev.preventDefault(); ev.stopPropagation();
    var p = toPage(ev); act = { kind: "size", h: h, id: e.id, sx: p.x, sy: p.y, o: JSON.parse(JSON.stringify(e)), moved: false };
    ev.target.setPointerCapture(ev.pointerId);
    ev.target.onpointermove = move; ev.target.onpointerup = ev.target.onpointercancel = function (e2) { ev.target.onpointermove = null; end(); };
  });
  ov.addEventListener("pointermove", function (ev) { if (act) move(ev); });
  ov.addEventListener("pointerup", end); ov.addEventListener("pointercancel", end);
  function move(ev) {
    var e = act && find(act.id); if (!e) return;
    var p = toPage(ev), dx = p.x - act.sx, dy = p.y - act.sy;
    if (!act.moved) { if (Math.abs(dx) * zoom < 2 && Math.abs(dy) * zoom < 2) return; act.moved = true; snap(); }
    if (act.kind === "move") {
      var nx = act.ox + dx, ny = act.oy + dy, T = 6 / zoom;
      if (Math.abs(nx + e.w / 2 - proj.W / 2) < T) nx = proj.W / 2 - e.w / 2;
      if (Math.abs(ny + e.h / 2 - proj.H / 2) < T) ny = proj.H / 2 - e.h / 2;
      e.x = nx; e.y = ny;
    } else {
      var o = act.o, h = act.h, w = o.w + (h.indexOf("e") >= 0 ? dx : -dx), ht = o.h + (h.indexOf("s") >= 0 ? dy : -dy);
      var keep = e.t !== "shape" || ev.shiftKey;
      if (keep) { var k = Math.max(w / o.w, ht / o.h, 12 / Math.min(o.w, o.h)); w = o.w * k; ht = o.h * k; }
      w = Math.max(12, w); ht = Math.max(12, ht);
      if (e.t === "text") { e.size = Math.max(6, Math.round(o.size * w / o.w)); if (o.sw) e.sw = o.sw * w / o.w; e.w = w; ht = S.textHeight(meas, e); }
      e.w = w; e.h = ht;
      e.x = h.indexOf("w") >= 0 ? o.x + o.w - w : o.x; e.y = h.indexOf("n") >= 0 ? o.y + o.h - ht : o.y;
    }
    draw(); if (editing === e.id) placeEditor();
  }
  function end() { if (!act) return; var a = act; act = null; if (a.moved) { changed(); ctxBar(); } }
  ov.addEventListener("dblclick", function (ev) { var e = topAt(toPage(ev)); if (e && e.t === "text") { select(e.id); startEdit(e); } });

  /* typing in a text box */
  function startEdit(e, all) {
    stopEdit(); editing = e.id;
    var ta = document.createElement("textarea"); ta.id = "ted"; ta.value = e.text; screenEl.appendChild(ta); placeEditor(); draw();
    ta.focus(); if (all) ta.select();
    var snapped = false;
    ta.addEventListener("input", function () { var x = find(editing); if (!x) return; if (!snapped) { snap(); snapped = true; } x.text = ta.value.slice(0, 3000); x.h = S.textHeight(meas, x); placeEditor(); selBox(); changed(); });
    ta.addEventListener("keydown", function (k) { if (k.key === "Escape") { k.preventDefault(); stopEdit(); } k.stopPropagation(); });
    ta.addEventListener("blur", function () { setTimeout(function () { if (document.activeElement !== ta) stopEdit(); }, 120); });
  }
  function placeEditor() {
    var ta = $("ted"), e = editing && find(editing); if (!ta || !e) return;
    ta.style.left = e.x * zoom + "px"; ta.style.top = e.y * zoom + "px"; ta.style.width = e.w * zoom + 2 + "px"; ta.style.height = Math.max(e.h, e.size * (e.lh || 1.2)) * zoom + "px";
    ta.style.font = S.fontOf(e, zoom); ta.style.lineHeight = e.lh || 1.2; ta.style.textAlign = e.align || "left";
    ta.style.color = /^lin:/.test(e.color) ? e.color.slice(4).split(",")[0] : e.color; ta.style.transform = "rotate(" + (e.r || 0) + "deg)";
    ta.style.background = e.bgc || "transparent"; ta.style.textTransform = e.upper ? "uppercase" : "none";
  }
  function stopEdit() {
    var ta = $("ted"); if (!ta) { editing = null; return; }
    var id = editing; editing = null; ta.remove();
    var e = id && find(id); if (e && !e.text.trim()) { SC().els = SC().els.filter(function (x) { return x.id !== id; }); sel = null; changed(); }
    draw(); ctxBar();
  }

  /* ---------------- adding things ---------------- */
  function add(e, edit) {
    snap(); if (mode !== "edit") { mode = "edit"; }
    if (e.at == null) e.at = 0;
    SC().els.push(e); select(e.id); draw(); changed();
    if (edit && e.t === "text") startEdit(e, true);
  }
  function textEl(text, frac, more) {
    var m = Math.min(proj.W, proj.H), e = { id: uid(), t: "text", text: text, font: "Poppins", size: Math.round(frac * m), color: "#ffffff", align: "center", lh: 1.2, x: 0, y: 0, w: proj.W * 0.8, h: 10, r: 0, o: 1, anim: "rise", at: 0 };
    Object.keys(more || {}).forEach(function (k) { e[k] = more[k]; });
    e.h = S.textHeight(meas, e); e.x = (proj.W - e.w) / 2; e.y = (proj.H - e.h) / 2; return e;
  }
  var EMOJI = ("⭐ 🌟 ✨ 💥 🔥 ❤️ 🎉 🎈 🎁 🎂 🍕 🍦 🌈 ☀️ 🌙 ⚡ ❄️ 🌊 🌸 🌍 🐶 🐱 🦊 🐼 🦁 🦄 🐉 🚀 ✈️ 🚗 ⚽ 🏀 🏆 🥇 🎮 🎧 🎵 🎨 📚 ✏️ 🔬 🧪 🧠 💡 🏫 📣 📌 ✅ ❌ ❓ ❗ 👍 👏 🙌 👋 😀 😂 😍 😎 🤩 🥳 😮 🤔 👀 💯 ➡️ ⬇️ 🔔 ⏰ 📺 🎬 📷 🎤").split(" ");
  var COLOURS = ["#ffffff", "#000000", "#111827", "#334155", "#94a3b8", "#f8fafc", "#ef4444", "#f97316", "#f59e0b", "#fde047", "#22c55e", "#14b8a6", "#0ea5e9", "#2563eb", "#7c3aed", "#db2777", "#fecaca", "#fed7aa", "#fef08a", "#bbf7d0", "#bae6fd", "#ddd6fe", "#fbcfe8", "#fef3c7"];
  var GRADS = [["#7c3aed", "#db2777", 135], ["#06b6d4", "#3b82f6", 135], ["#f97316", "#facc15", 135], ["#22c55e", "#14b8a6", 135], ["#0f172a", "#334155", 180], ["#1e3a8a", "#0ea5e9", 200], ["#020617", "#4c1d95", 120], ["#7c2d12", "#000000", 180], ["#a855f7", "#ec4899", 135], ["#38bdf8", "#e0f2fe", 180], ["#fbcfe8", "#c4b5fd", 135], ["#111827", "#ef4444", 135]];
  function addMediaEl(m) {
    var W = proj.W, H = proj.H, iw = m.el && (m.el.naturalWidth || m.el.videoWidth) || 16, ih = m.el && (m.el.naturalHeight || m.el.videoHeight) || 9;
    if (m.kind === "video") {
      // a clip fills the screen, and the scene is as long as the clip
      var k = Math.min(W / iw, H / ih), full = Math.abs(iw / ih - W / H) < 0.15;
      var e = { id: uid(), t: "video", src: "idb:" + m.id, x: 0, y: 0, w: full ? W : iw * k, h: full ? H : ih * k, r: 0, o: 1, from: 0, vol: 1, at: 0 };
      e.x = (W - e.w) / 2; e.y = (H - e.h) / 2;
      snap(); SC().els.unshift(e); if (m.dur) SC().dur = Math.min(120, Math.max(SC().dur, Math.round(m.dur * 10) / 10));
      select(e.id); draw(); timeline(); changed(); return;
    }
    var k2 = Math.min(W * 0.6 / iw, H * 0.6 / ih);
    add({ id: uid(), t: "img", src: "idb:" + m.id, x: (W - iw * k2) / 2, y: (H - ih * k2) / 2, w: iw * k2, h: ih * k2, r: 0, o: 1, anim: "fade", at: 0 });
  }
  function importFiles(files) {
    var list = Array.prototype.slice.call(files || []).slice(0, 10); if (!list.length) return;
    Office.toast("Adding " + (list.length > 1 ? list.length + " things" : "it") + "…");
    list.reduce(function (pr, f) {
      return pr.then(function () {
        var kind = /^image\//.test(f.type) ? "image" : /^video\//.test(f.type) ? "video" : /^audio\//.test(f.type) ? "audio" : null;
        if (!kind) { Office.toast(f.name + " isn't a picture, video or sound.", true); return; }
        if (f.size > 500 * 1024 * 1024) { Office.toast(f.name + " is too big (over 500 MB).", true); return; }
        var p = kind === "image" ? smaller(f) : Promise.resolve(f);
        return p.then(function (bl) { return storeMedia(bl, kind, f.name); }).then(function (m) {
          if (kind === "audio") { setMusic(m); return; }
          addMediaEl(m); if (panel === "media") showPanel();
        });
      }).catch(function (e) { Office.toast("Couldn't add " + f.name + ": " + (e.message || e), true); });
    }, Promise.resolve());
  }
  function smaller(f) {
    return new Promise(function (ok) {
      var u = URL.createObjectURL(f), im = new Image();
      im.onload = function () {
        var k = Math.min(1, 2000 / Math.max(im.width, im.height)); if (k >= 1 && f.size < 3e6) { URL.revokeObjectURL(u); ok(f); return; }
        var c = document.createElement("canvas"); c.width = Math.round(im.width * k); c.height = Math.round(im.height * k);
        c.getContext("2d").drawImage(im, 0, 0, c.width, c.height); URL.revokeObjectURL(u);
        c.toBlob(function (b) { ok(b || f); }, /png|gif|webp/.test(f.type) ? "image/png" : "image/jpeg", 0.88);
      };
      im.onerror = function () { URL.revokeObjectURL(u); ok(f); };
      im.src = u;
    });
  }
  $("mediaIn").onchange = function () { importFiles(this.files); this.value = ""; };
  $("audioIn").onchange = function () { var f = this.files[0]; this.value = ""; if (f) storeMedia(f, "audio", f.name).then(setMusic).catch(function (e) { Office.toast(e.message || e, true); }); };
  stage.addEventListener("dragover", function (e) { e.preventDefault(); });
  stage.addEventListener("drop", function (e) { e.preventDefault(); if (e.dataTransfer && e.dataTransfer.files.length) importFiles(e.dataTransfer.files); });

  /* ---------------- music made here (free to use) ---------------- */
  var TUNES = [
    { n: "Happy", bpm: 112, chords: [[60, 64, 67], [55, 59, 62], [57, 60, 64], [53, 57, 60]], lead: [72, 74, 76, 79, 76, 74, 72, 67], wave: "triangle", drums: 1 },
    { n: "Calm", bpm: 76, chords: [[57, 60, 64], [53, 57, 60], [48, 52, 55], [55, 59, 62]], lead: [69, 0, 72, 0, 71, 0, 67, 0], wave: "sine", drums: 0 },
    { n: "Adventure", bpm: 124, chords: [[50, 57, 62], [46, 53, 58], [48, 55, 60], [45, 52, 57]], lead: [74, 74, 77, 79, 81, 79, 77, 74], wave: "sawtooth", drums: 2 },
    { n: "Lo-fi", bpm: 84, chords: [[53, 57, 60, 64], [52, 55, 59, 62], [50, 53, 57, 60], [48, 52, 55, 59]], lead: [0, 72, 0, 69, 0, 67, 0, 64], wave: "triangle", drums: 3 },
    { n: "Mystery", bpm: 90, chords: [[45, 48, 52], [44, 48, 51], [43, 46, 50], [44, 47, 51]], lead: [69, 0, 68, 0, 67, 0, 68, 0], wave: "sine", drums: 0 }
  ];
  function hz(n) { return 440 * Math.pow(2, (n - 69) / 12); }
  function makeTune(tu) {
    var Off = window.OfflineAudioContext || window.webkitOfflineAudioContext; if (!Off) return Promise.reject(new Error("This browser can't make music."));
    var beat = 60 / tu.bpm, bars = 8, len = bars * 4 * beat, rate = 44100, oc = new Off(2, Math.ceil(len * rate), rate), out = oc.createGain(); out.gain.value = 0.5; out.connect(oc.destination);
    function note(f, at, d, type, vol) {
      var o = oc.createOscillator(), g = oc.createGain(); o.type = type; o.frequency.value = f;
      g.gain.setValueAtTime(0, at); g.gain.linearRampToValueAtTime(vol, at + 0.02); g.gain.exponentialRampToValueAtTime(0.0008, at + d);
      o.connect(g); g.connect(out); o.start(at); o.stop(at + d + 0.05);
    }
    function noise(at, d, vol, hp) {
      var b = oc.createBuffer(1, Math.ceil(d * rate), rate), x = b.getChannelData(0); for (var i = 0; i < x.length; i++) x[i] = Math.random() * 2 - 1;
      var s = oc.createBufferSource(), f = oc.createBiquadFilter(), g = oc.createGain(); s.buffer = b; f.type = "highpass"; f.frequency.value = hp;
      g.gain.setValueAtTime(vol, at); g.gain.exponentialRampToValueAtTime(0.001, at + d); s.connect(f); f.connect(g); g.connect(out); s.start(at);
    }
    function kick(at) { var o = oc.createOscillator(), g = oc.createGain(); o.frequency.setValueAtTime(140, at); o.frequency.exponentialRampToValueAtTime(40, at + 0.18); g.gain.setValueAtTime(0.9, at); g.gain.exponentialRampToValueAtTime(0.001, at + 0.25); o.connect(g); g.connect(out); o.start(at); o.stop(at + 0.3); }
    for (var bar = 0; bar < bars; bar++) {
      var ch = tu.chords[bar % tu.chords.length], t0 = bar * 4 * beat;
      ch.forEach(function (n) { note(hz(n), t0, 4 * beat * 0.98, tu.wave === "sawtooth" ? "triangle" : "sine", 0.09); });
      note(hz(ch[0] - 12), t0, 2 * beat, "sine", 0.22); note(hz(ch[0] - 12), t0 + 2 * beat, 2 * beat, "sine", 0.18);
      for (var i = 0; i < 8; i++) { var ln = tu.lead[(bar * 8 + i) % tu.lead.length]; if (ln && bar % 2 === 1) note(hz(ln), t0 + i * beat / 2, beat * 0.45, tu.wave, 0.07); }
      for (var q = 0; q < 4; q++) {
        var tq = t0 + q * beat;
        if (tu.drums === 1) { if (q % 2 === 0) kick(tq); else noise(tq, 0.12, 0.25, 1500); noise(tq + beat / 2, 0.04, 0.12, 7000); }
        if (tu.drums === 2) { kick(tq); if (q % 2) noise(tq, 0.15, 0.3, 1200); noise(tq + beat / 2, 0.05, 0.14, 6000); }
        if (tu.drums === 3) { if (q === 0 || q === 2.5) kick(tq); if (q % 2) noise(tq, 0.1, 0.2, 2500); noise(tq + beat / 2, 0.03, 0.08, 8000); }
      }
    }
    return oc.startRendering().then(function (buf) { return wav(buf); });
  }
  function wav(buf) {
    var ch = buf.numberOfChannels, n = buf.length, b = new ArrayBuffer(44 + n * ch * 2), v = new DataView(b), o = 0;
    function s(str) { for (var i = 0; i < str.length; i++) v.setUint8(o++, str.charCodeAt(i)); } function u32(x) { v.setUint32(o, x, true); o += 4; } function u16(x) { v.setUint16(o, x, true); o += 2; }
    s("RIFF"); u32(36 + n * ch * 2); s("WAVE"); s("fmt "); u32(16); u16(1); u16(ch); u32(buf.sampleRate); u32(buf.sampleRate * ch * 2); u16(ch * 2); u16(16); s("data"); u32(n * ch * 2);
    var d = []; for (var c = 0; c < ch; c++) d.push(buf.getChannelData(c));
    for (var i = 0; i < n; i++) for (c = 0; c < ch; c++) { var x = Math.max(-1, Math.min(1, d[c][i])); v.setInt16(o, x < 0 ? x * 0x8000 : x * 0x7fff, true); o += 2; }
    return new Blob([b], { type: "audio/wav" });
  }
  function setMusic(m) { snap(); proj.music = { src: "idb:" + m.id, vol: proj.music ? proj.music.vol : 0.4, name: m.name }; changed(); timeline(); if (panel === "music") showPanel(); Office.toast("Music added."); }

  /* ---------------- playing it ---------------- */
  var AC = null, master = null, speaker = null, dest = null;
  function audioOn() {
    if (AC) { if (AC.state === "suspended") AC.resume(); return AC; }
    var A = window.AudioContext || window.webkitAudioContext; if (!A) return null;
    AC = new A(); master = AC.createGain(); speaker = AC.createGain(); master.connect(speaker); speaker.connect(AC.destination);
    if (AC.createMediaStreamDestination) { dest = AC.createMediaStreamDestination(); master.connect(dest); }
    return AC;
  }
  function route(el, vol) {
    if (!el) return;
    if (AC && !el._g) { try { var src = AC.createMediaElementSource(el); el._g = AC.createGain(); src.connect(el._g); el._g.connect(master); } catch (e) { el._g = null; el._nog = true; } }
    if (el._g) el._g.gain.value = vol; else el.volume = Math.max(0, Math.min(1, vol));
  }
  var clock = { at: 0, t0: 0 }, loopId = 0;
  function play() {
    if (!proj) return; stopEdit();
    if (t >= total() - 0.05) t = 0;
    audioOn(); playing = true; mode = "play"; $("playBtn").textContent = "⏸"; $("playBtn").setAttribute("aria-label", "Pause");
    clock = { at: performance.now(), t0: t };
    var id = ++loopId;
    (function tick() {
      if (id !== loopId || !playing) return;
      t = clock.t0 + (performance.now() - clock.at) / 1000;
      if (t >= total()) { t = total(); sync(t, true); drawNow(); pause(); onEnd(); return; }
      sync(t, false); drawNow(); progress();
      requestAnimationFrame(tick);
    })();
  }
  var onEnd = function () {};
  function pause() {
    playing = false; loopId++; $("playBtn").textContent = "▶"; $("playBtn").setAttribute("aria-label", "Play");
    stopAll(); progress();
  }
  function stopAll() {
    if (!proj) return;
    Object.keys(CLIPS).forEach(function (k) { try { CLIPS[k].pause(); } catch (e) {} });
    Object.keys(MEDIA).forEach(function (k) { var m = MEDIA[k]; if (m.el && m.el.pause && (m.kind === "audio")) try { m.el.pause(); } catch (e) {} });
  }
  /* every clip and sound at the right place for the moment tt */
  function sync(tt, stop) {
    var at = sceneAt(tt), sc = proj.scenes[at.i];
    // clips on this scene (and the one before, during a transition)
    var live = {};
    [at.i, at.lt < TD ? at.i - 1 : -1].forEach(function (i) {
      if (i < 0) return;
      var s = proj.scenes[i], lt = i === at.i ? at.lt : s.dur;
      s.els.forEach(function (el) {
        if (el.t !== "video") return;
        var v = clipFor(el); if (!v) return; live[el.id + "|" + el.src.slice(4)] = 1;
        var want = (el.from || 0) + Math.max(0, lt - (el.at || 0));
        if (i !== at.i || lt < (el.at || 0) || stop) { if (!v.paused) v.pause(); return; }
        if (v.readyState >= 1 && Math.abs(v.currentTime - want) > 0.35) try { v.currentTime = want; } catch (e) {}
        route(v, el.mute ? 0 : (el.vol == null ? 1 : el.vol));
        if (v.paused && !v.ended) { var pr = v.play(); if (pr && pr.catch) pr.catch(function () {}); }
      });
    });
    Object.keys(CLIPS).forEach(function (k) { if (!live[k] && !CLIPS[k].paused) CLIPS[k].pause(); });
    // the voice for this scene
    proj.scenes.forEach(function (s, i) {
      var m = mediaOf(s.voice); if (!m || !m.el) return;
      var v = m.el, lt = tt - startOf(i);
      if (i === at.i && !stop && lt < (m.dur || 999)) {
        if (Math.abs(v.currentTime - lt) > 0.35) try { v.currentTime = Math.max(0, lt); } catch (e) {}
        route(v, s.vvol == null ? 1 : s.vvol); if (v.paused) { var pr = v.play(); if (pr && pr.catch) pr.catch(function () {}); }
      } else if (!v.paused) v.pause();
    });
    // music all the way through, quieter at the very end
    var mm = proj.music && mediaOf(proj.music.src);
    if (mm && mm.el) {
      var a = mm.el, T = total(), md = mm.dur || a.duration || 0;
      if (stop || !md) { if (!a.paused) a.pause(); }
      else {
        var want = tt % md, fade = Math.min(1, Math.max(0, (T - tt) / 1.5));
        if (Math.abs(a.currentTime - want) > 0.4) try { a.currentTime = want; } catch (e) {}
        a.loop = true; route(a, proj.music.vol * fade); if (a.paused) { var p2 = a.play(); if (p2 && p2.catch) p2.catch(function () {}); }
      }
    }
  }
  function progress() {
    if (!proj) return;
    var T = total(); $("tm").textContent = fmt(t) + " / " + fmt(T);
    if (!scrubbing) $("scrub").value = T ? Math.round(t / T * 1000) : 0;
    var ph = $("phead"); if (ph) ph.style.left = (12 + t * PPS + gapsBefore(t)) + "px";
    if (playing) { var at = sceneAt(t); if (at.i !== cur) { cur = at.i; markScene(); } }
  }
  $("playBtn").onclick = function () { if (playing) pause(); else play(); };
  $("toStart").onclick = function () { pause(); t = 0; cur = 0; mode = "edit"; markScene(); progress(); draw(); ctxBar(); };
  var scrubbing = false;
  $("scrub").addEventListener("input", function () {
    scrubbing = true; if (playing) pause();
    t = +this.value / 1000 * total(); mode = "peek"; var at = sceneAt(t); if (at.i !== cur) { cur = at.i; sel = null; markScene(); ctxBar(); }
    sync(t, true); draw(); progress();
  });
  $("scrub").addEventListener("change", function () { scrubbing = false; });

  /* ---------------- the timeline ---------------- */
  var PPS = 26; // pixels for each second
  function gapsBefore(tt) { return sceneAt(tt).i * 6; }
  var tlT = null; function tlSoon() { clearTimeout(tlT); tlT = setTimeout(timeline, 250); }
  function timeline() {
    if (!proj) return;
    var inn = $("tlin"), h = '<div class="lane" id="lane">';
    proj.scenes.forEach(function (s, i) {
      var w = Math.max(40, s.dur * PPS);
      h += '<button type="button" class="scn" data-sc="' + i + '" aria-current="' + (i === cur) + '" style="width:' + w + 'px" aria-label="Scene ' + (i + 1) + ', ' + s.dur + ' seconds"><canvas></canvas>' +
        '<span class="n">' + (i + 1) + (s.trans !== "none" && i ? " ⇢" : "") + '</span><span class="d">' + (Math.round(s.dur * 10) / 10) + "s</span>" + (s.voice ? '<span class="mic">🎙</span>' : "") + '<span class="grab" data-grab="' + i + '" title="Drag to make it longer or shorter"></span></button>';
    });
    h += '<button type="button" class="addscn" id="addScn" aria-label="Add a scene" title="Add a scene">＋</button></div>';
    var mu = proj.music;
    h += '<div class="mlane' + (mu ? "" : " none") + '" id="mlane" style="width:' + Math.max(120, total() * PPS + proj.scenes.length * 6) + 'px">' + (mu ? "🎵 " + esc(mu.name) + " · " + Math.round(mu.vol * 100) + "%" : "🎵 Add music") + "</div>";
    h += '<div class="phead" id="phead"></div>';
    inn.innerHTML = h; inn.style.width = (total() * PPS + proj.scenes.length * 6 + 120) + "px";
    inn.querySelectorAll(".scn canvas").forEach(function (c, i) {
      var s = proj.scenes[i], bh = 70, bw = Math.max(40, s.dur * PPS);
      c.width = bw * 2; c.height = bh * 2;
      var g = c.getContext("2d"), k = Math.max(bw * 2 / proj.W, bh * 2 / proj.H);
      g.translate((bw * 2 - proj.W * k) / 2, (bh * 2 - proj.H * k) / 2); g.scale(k, k);
      drawScene(g, s, s.dur, false);
    });
    progress();
  }
  function markScene() { $("tlin").querySelectorAll(".scn").forEach(function (b) { b.setAttribute("aria-current", String(+b.dataset.sc === cur)); }); }
  $("tl").addEventListener("click", function (ev) {
    if (ev.target.closest("[data-grab]")) return;
    var b = ev.target.closest("[data-sc]");
    if (b) { if (playing) pause(); stopEdit(); cur = +b.dataset.sc; sel = null; mode = "edit"; t = startOf(cur); markScene(); progress(); draw(); ctxBar(); if (panel === "scene") showPanel(); return; }
    if (ev.target.closest("#addScn")) { addScene(); return; }
    if (ev.target.closest("#mlane")) { panel = "music"; showPanel(); }
  });
  $("tl").addEventListener("pointerdown", function (ev) {
    var g = ev.target.closest("[data-grab]"); if (!g) return;
    ev.preventDefault(); var i = +g.dataset.grab, s = proj.scenes[i], x0 = ev.clientX, d0 = s.dur, moved = false;
    g.setPointerCapture(ev.pointerId);
    g.onpointermove = function (e) {
      var nd = Math.max(0.5, Math.min(120, Math.round((d0 + (e.clientX - x0) / PPS) * 10) / 10));
      if (!moved && nd !== d0) { snap(); moved = true; }
      s.dur = nd; var bb = g.parentNode; bb.style.width = Math.max(40, nd * PPS) + "px"; bb.querySelector(".d").textContent = nd + "s"; progress();
    };
    g.onpointerup = g.onpointercancel = function () { g.onpointermove = null; if (moved) { timeline(); changed(); if (panel === "scene") showPanel(); } };
  });
  function addScene(copy) {
    snap(); stopEdit();
    var s = copy ? JSON.parse(JSON.stringify(SC())) : blankScene();
    if (copy) { s.id = uid(); s.els.forEach(function (e) { e.id = uid(); }); delete s.voice; }
    else { s.bg = JSON.parse(JSON.stringify(SC().bg.img ? { c: "#111827" } : SC().bg)); }
    proj.scenes.splice(cur + 1, 0, s); cur++; sel = null; mode = "edit"; t = startOf(cur);
    draw(); timeline(); ctxBar(); showPanel(); changed();
  }

  /* ---------------- recording: your voice, yourself, your screen ---------------- */
  var rec = null;
  function pickType(kind) {
    var list = kind === "audio" ? ["audio/mp4", "audio/webm;codecs=opus", "audio/webm", "audio/ogg"] : ["video/mp4;codecs=avc1,mp4a.40.2", "video/mp4", "video/webm;codecs=vp9,opus", "video/webm;codecs=vp8,opus", "video/webm"];
    if (!window.MediaRecorder) return null;
    for (var i = 0; i < list.length; i++) if (MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(list[i])) return list[i];
    return "";
  }
  function record(kind) {
    if (rec) { rec.stop(); return; }
    if (!window.MediaRecorder || !navigator.mediaDevices) { Office.toast("This browser can't record.", true); return; }
    var get = kind === "screen" ? navigator.mediaDevices.getDisplayMedia({ video: true, audio: true }) : navigator.mediaDevices.getUserMedia(kind === "voice" ? { audio: { echoCancellation: true, noiseSuppression: true } } : { video: { width: { ideal: 1280 }, height: { ideal: 720 }, facingMode: "user" }, audio: true });
    get.then(function (stream) {
      var type = pickType(kind === "voice" ? "audio" : "video"), r = new MediaRecorder(stream, type ? { mimeType: type } : undefined), chunks = [], began = Date.now();
      rec = { r: r, kind: kind, stream: stream, stop: function () { if (r.state !== "inactive") r.stop(); } };
      r.ondataavailable = function (e) { if (e.data && e.data.size) chunks.push(e.data); };
      r.onstop = function () {
        stream.getTracks().forEach(function (tr) { tr.stop(); });
        var secs = (Date.now() - began) / 1000, blob = new Blob(chunks, { type: (r.mimeType || type || (kind === "voice" ? "audio/webm" : "video/webm")).split(";")[0] });
        var k = kind; rec = null; if (kind === "voice") pause();
        if (!blob.size) { showPanel(); return; }
        storeMedia(blob, kind === "voice" ? "audio" : "video", kind === "voice" ? "Voiceover" : kind === "screen" ? "Screen recording" : "Camera", secs).then(function (m) {
          if (!m.dur) m.dur = secs;
          if (k === "voice") {
            snap(); SC().voice = "idb:" + m.id; SC().vvol = 1;
            if (m.dur > SC().dur) { SC().dur = Math.min(120, Math.ceil(m.dur * 10) / 10); Office.toast("The scene is longer now, to fit your voice."); }
            timeline(); changed();
          } else {
            // a recording gets its own new scene
            snap(); var s = blankScene(); s.bg = { c: "#000000" }; s.trans = "fade"; proj.scenes.splice(cur + 1, 0, s); cur++;
            addMediaEl(m); SC().dur = Math.min(120, Math.max(1, Math.round(m.dur * 10) / 10)); timeline();
          }
          showPanel();
        });
      };
      r.start(250);
      if (kind === "voice") {
        // the scene plays once while you talk, then waits on its last moment until you press Stop
        t = startOf(cur); mode = "play"; play();
        var endAt = startOf(cur) + SC().dur, me = rec;
        (function hold() { if (rec !== me) return; if (playing && t >= endAt - 0.03) { pause(); t = endAt - 0.03; drawNow(); } requestAnimationFrame(hold); })();
      }
      if (kind === "screen") stream.getVideoTracks()[0].onended = function () { if (rec) rec.stop(); };
      showPanel();
    }).catch(function (e) { Office.toast(e && e.name === "NotAllowedError" ? "Recording wasn't allowed - check the browser's permission for the camera or microphone." : "Couldn't start recording: " + (e.message || e), true); });
  }

  /* ---------------- the panel on the left ---------------- */
  function mini(fn, w, h) { var c = document.createElement("canvas"); c.width = w * 2; c.height = h * 2; var g = c.getContext("2d"); g.scale(2, 2); fn(g); return c; }
  function showPanel() {
    var pn = $("panel");
    document.querySelectorAll(".rail [data-p]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.p === panel)); });
    pn.hidden = !panel; if (!panel || !proj) { fitSoon(); return; }
    var x = '<button type="button" class="x" data-close aria-label="Close the panel">✕</button>', h = "", sc = SC();
    if (panel === "scene") {
      h = x + "<h3>Scene " + (cur + 1) + " of " + proj.scenes.length + "</h3>" +
        '<div class="rng">Length<input type="range" id="durIn" min="0.5" max="30" step="0.5" value="' + Math.min(30, sc.dur) + '"><output>' + sc.dur + "s</output></div>" +
        "<h4>Caption</h4><textarea id=\"capIn\" rows=\"3\" placeholder=\"Words shown at the bottom, like subtitles\" style=\"width:100%;border:1px solid var(--rule);border-radius:10px;padding:8px;background:var(--tool);resize:vertical\">" + esc(sc.cap) + "</textarea>" +
        "<h4>How it comes in</h4><div class=\"gridp\" id=\"transG\">" + [["none", "Cut"], ["fade", "Fade"], ["slide", "Slide"], ["up", "Slide up"], ["zoom", "Zoom"], ["wipe", "Wipe"], ["circle", "Circle"]].map(function (tr) {
          return '<button type="button" class="pbtn ' + (sc.trans === tr[0] ? "" : "ghost") + '" data-tr="' + tr[0] + '" style="margin:0;padding:8px 4px;font-size:.82rem">' + tr[1] + "</button>";
        }).join("") + '</div><button type="button" class="pbtn ghost" data-trall style="margin-top:8px">Use this on every scene</button>' +
        "<h4>Background</h4><div class=\"gridp six\">" + COLOURS.map(function (c) { return '<button type="button" class="sw" data-bgc="' + c + '" style="background:' + c + '" aria-label="' + c + '"></button>'; }).join("") +
        GRADS.map(function (g, i) { return '<button type="button" class="sw" data-bgg="' + i + '" style="background:' + S.cssFill("lin:" + g.join(",")) + '" aria-label="Gradient ' + (i + 1) + '"></button>'; }).join("") + "</div>" +
        "<h4>Scenes</h4>" + '<button type="button" class="pbtn" data-sa="add">＋ New scene</button><button type="button" class="pbtn ghost" data-sa="dup">⧉ Copy this scene</button>' +
        '<div class="gridp two"><button type="button" class="pbtn ghost" data-sa="left" style="margin:0">◀ Earlier</button><button type="button" class="pbtn ghost" data-sa="right" style="margin:0">Later ▶</button></div>' +
        '<button type="button" class="pbtn ghost" data-sa="del" style="margin-top:8px">🗑 Delete this scene</button>' +
        '<button type="button" class="pbtn ghost" data-sa="tpl">✨ Add scenes from a template</button>';
    } else if (panel === "text") {
      h = x + "<h3>Text</h3>" +
        '<button type="button" class="addtext" data-tx="t" style="font:800 1.5rem var(--ui)">Add a title</button>' +
        '<button type="button" class="addtext" data-tx="s" style="font:700 1.1rem var(--ui)">Add a subtitle</button>' +
        '<button type="button" class="addtext" data-tx="b" style="font:400 .9rem var(--ui)">Add some words</button>' +
        '<button type="button" class="addtext" data-tx="l" style="font:700 .9rem var(--ui)"><span style="background:#e11d48;color:#fff;padding:2px 8px;border-radius:6px">Lower third</span> a name label</button>' +
        '<button type="button" class="addtext" data-tx="k" style="font:700 .95rem \'Courier Prime\',monospace">Typewriter words…</button>' +
        '<p class="s">Tip: pick some text, then choose how it comes in (Animation) at the top.</p>';
    } else if (panel === "el") {
      h = x + "<h3>Elements</h3><h4>Shapes</h4><div class=\"gridp four\" id=\"shapeG\"></div><h4>Stickers</h4><div class=\"gridp six\">" + EMOJI.map(function (e) { return '<button type="button" class="emo" data-emo="' + e + '">' + e + "</button>"; }).join("") + "</div>";
    } else if (panel === "media") {
      var ms = Object.keys(MEDIA).map(function (k) { return MEDIA[k]; }).filter(function (m) { return m.pid === proj.id && (m.kind === "image" || m.kind === "video"); });
      h = x + '<h3>Media</h3><label class="pbtn">⬆️ Add photos or video clips<input type="file" accept="image/*,video/*" multiple hidden id="mIn"></label>' +
        '<p class="s">Or drag them onto the screen. They stay on this device.</p>' +
        (ms.length ? '<h4>In this video</h4><div class="gridp" id="mediaG">' + ms.map(function (m) {
          return '<button type="button" class="tile" data-md="' + m.id + '" aria-label="Add ' + esc(m.name) + '">' + (m.kind === "image" ? '<img alt="" src="' + m.url + '">' : '<video muted playsinline preload="metadata" src="' + m.url + '#t=0.5"></video>') +
            (m.kind === "video" ? '<span class="badge">▶ ' + fmt(m.dur) + "</span>" : "") + "</button>";
        }).join("") + '</div><p class="s" style="margin-top:8px">Tap one to add it. Pick a photo on the screen and use “Background” to fill the scene with it.</p>' : "");
    } else if (panel === "rec") {
      var on = rec && rec.kind, can = !!(window.MediaRecorder && navigator.mediaDevices);
      h = x + "<h3>Record</h3>" + (can ? "" : '<p class="s">This browser can\'t record.</p>') +
        '<div class="recbox"><b>🎙 Voiceover for scene ' + (cur + 1) + '</b><p class="s" style="margin:6px 0">The scene plays while you talk. Press Stop when you\'ve finished - if you talk for longer, the scene gets longer.</p>' +
        '<button type="button" class="pbtn' + (on === "voice" ? "" : "") + '" data-rec="voice"' + (can && (!on || on === "voice") ? "" : " disabled") + ">" + (on === "voice" ? '<span class="dot"></span>Stop' : "● Record my voice") + "</button>" +
        (sc.voice ? '<div class="mrow"><span>🎙 Voiceover recorded</span><button type="button" data-vplay>▶</button><button type="button" data-vdel>Delete</button></div><div class="rng">Volume<input type="range" id="vvolIn" min="0" max="1" step="0.05" value="' + (sc.vvol == null ? 1 : sc.vvol) + '"><output>' + Math.round((sc.vvol == null ? 1 : sc.vvol) * 100) + "%</output></div>" : "") + "</div>" +
        '<div class="recbox"><b>📹 Record yourself</b><p class="s" style="margin:6px 0">Uses your camera and microphone. It goes in as a new scene.</p>' + (on === "camera" ? '<video id="camPrev" autoplay muted playsinline></video>' : "") +
        '<button type="button" class="pbtn" data-rec="camera"' + (can && (!on || on === "camera") ? "" : " disabled") + ">" + (on === "camera" ? '<span class="dot"></span>Stop' : "● Start the camera") + "</button></div>" +
        (navigator.mediaDevices && navigator.mediaDevices.getDisplayMedia ? '<div class="recbox"><b>🖥 Record your screen</b><p class="s" style="margin:6px 0">Show how to do something on the computer.</p><button type="button" class="pbtn" data-rec="screen"' + (can && (!on || on === "screen") ? "" : " disabled") + ">" + (on === "screen" ? '<span class="dot"></span>Stop' : "● Record my screen") + "</button></div>" : "");
    } else if (panel === "music") {
      var mu = proj.music;
      h = x + "<h3>Music</h3>" + (mu ? '<div class="mrow"><span>🎵 ' + esc(mu.name) + '</span><button type="button" data-mdel>Remove</button></div><div class="rng">Volume<input type="range" id="mvolIn" min="0" max="1" step="0.05" value="' + mu.vol + '"><output>' + Math.round(mu.vol * 100) + "%</output></div>" : '<p class="s">Music plays all the way through your video and gets quieter at the end.</p>') +
        "<h4>Made for you - free to use</h4><div class=\"mlist\">" + TUNES.map(function (tu, i) { return '<div class="mrow"><span>🎶 ' + tu.n + '</span><button type="button" data-tprev="' + i + '">▶ Listen</button><button type="button" data-tuse="' + i + '">Use</button></div>'; }).join("") + "</div>" +
        '<h4>Your own</h4><button type="button" class="pbtn ghost" data-mup>⬆️ Add a music file</button><p class="s">Only use music you made or that you\'re allowed to use.</p>';
    }
    pn.innerHTML = h;
    if (panel === "el") {
      var g = $("shapeG");
      S.SHAPES.forEach(function (k) {
        var b = document.createElement("button"); b.type = "button"; b.className = "tile"; b.dataset.shape = k; b.setAttribute("aria-label", k);
        b.appendChild(mini(function (gg) { S.drawEl(gg, { t: "shape", shape: k, x: 12, y: k === "line" ? 28 : 12, w: 40, h: k === "line" ? 8 : 40, fill: k === "line" ? "none" : "#e11d48", stroke: "#9f1239", sw: k === "line" ? 4 : 0, o: 1 }); }, 64, 64));
        g.appendChild(b);
      });
    }
    var cp = $("camPrev"); if (cp && rec) cp.srcObject = rec.stream;
    var mi = $("mIn"); if (mi) mi.onchange = function () { importFiles(this.files); this.value = ""; };
    var di = $("durIn"); if (di) { var dsnap = false; di.oninput = function () { if (!dsnap) { snap(); dsnap = true; } SC().dur = +di.value; di.nextElementSibling.textContent = di.value + "s"; tlSoon(); progress(); changed(); }; }
    var ci = $("capIn"); if (ci) { var csnap = false; ci.oninput = function () { if (!csnap) { snap(); csnap = true; } SC().cap = ci.value.slice(0, 400); draw(); changed(); }; }
    var vv = $("vvolIn"); if (vv) vv.oninput = function () { SC().vvol = +vv.value; vv.nextElementSibling.textContent = Math.round(vv.value * 100) + "%"; changed(); };
    var mv = $("mvolIn"); if (mv) mv.oninput = function () { proj.music.vol = +mv.value; mv.nextElementSibling.textContent = Math.round(mv.value * 100) + "%"; changed(); };
    fitSoon();
  }
  var tunePrev = null;
  $("panel").addEventListener("click", function (ev) {
    var b = ev.target.closest("button,label"); if (!b || b.disabled) return;
    var d = b.dataset, sc = SC();
    if (d.close != null) { panel = null; showPanel(); return; }
    if (d.tr) { snap(); sc.trans = d.tr; timeline(); showPanel(); changed(); return; }
    if (d.trall != null) { snap(); proj.scenes.forEach(function (s, i) { if (i) s.trans = sc.trans; }); timeline(); changed(); Office.toast("Every scene comes in the same way now."); return; }
    if (d.bgc) { snap(); sc.bg = { c: d.bgc }; draw(); tlSoon(); changed(); return; }
    if (d.bgg != null) { snap(); sc.bg = { g: GRADS[+d.bgg].slice() }; draw(); tlSoon(); changed(); return; }
    if (d.sa === "add") { addScene(false); return; }
    if (d.sa === "dup") { addScene(true); return; }
    if (d.sa === "left" || d.sa === "right") { var j = cur + (d.sa === "left" ? -1 : 1); if (j < 0 || j >= proj.scenes.length) return; snap(); var tmp = proj.scenes[cur]; proj.scenes[cur] = proj.scenes[j]; proj.scenes[j] = tmp; cur = j; t = startOf(cur); timeline(); showPanel(); changed(); return; }
    if (d.sa === "del") {
      snap(); stopEdit();
      if (proj.scenes.length < 2) { proj.scenes[0] = blankScene(); } else { proj.scenes.splice(cur, 1); cur = Math.max(0, cur - 1); }
      sel = null; t = startOf(cur); draw(); timeline(); ctxBar(); showPanel(); changed(); Office.toast("Scene deleted - Undo brings it back."); return;
    }
    if (d.sa === "tpl") { startDialog(true); return; }
    if (d.tx) {
      var m = { t: [.11, "Your title", { font: "Poppins", bold: true, anim: "pop" }], s: [.06, "A subtitle", { font: "Poppins", anim: "rise" }], b: [.045, "Some words about this scene", { font: "Inter", anim: "fade", w: proj.W * 0.7 }],
        l: [.045, "Your Name · Reporter", { font: "Poppins", bold: true, align: "left", bgc: "#e11d48", anim: "left", w: proj.W * 0.45 }], k: [.05, "Once upon a time…", { font: "Courier Prime", anim: "type", w: proj.W * 0.75 }] }[d.tx];
      var e = textEl(m[1], m[0], m[2]); if (d.tx === "l") { e.x = proj.W * 0.05; e.y = proj.H * 0.72; }
      add(e, true); return;
    }
    if (d.shape) { var s = Math.min(proj.W, proj.H) * 0.3; add({ id: uid(), t: "shape", shape: d.shape, fill: d.shape === "line" ? "none" : "#e11d48", stroke: d.shape === "line" ? "#ffffff" : undefined, sw: d.shape === "line" ? 6 : 0, x: (proj.W - s) / 2, y: (proj.H - (d.shape === "line" ? 10 : s)) / 2, w: s, h: d.shape === "line" ? 10 : s, r: 0, o: 1, anim: "pop" }); return; }
    if (d.emo) { var z = Math.min(proj.W, proj.H) * 0.25; add({ id: uid(), t: "emoji", ch: d.emo, x: (proj.W - z) / 2, y: (proj.H - z) / 2, w: z, h: z, r: 0, o: 1, anim: "pop" }); return; }
    if (d.md) { var mm = MEDIA[d.md]; if (mm) addMediaEl(mm); return; }
    if (d.rec) { record(d.rec); return; }
    if (d.vplay != null) { var vm = mediaOf(sc.voice); if (vm) { audioOn(); route(vm.el, 1); vm.el.currentTime = 0; vm.el.play(); } return; }
    if (d.vdel != null) { snap(); delete sc.voice; timeline(); showPanel(); changed(); return; }
    if (d.mdel != null) { snap(); proj.music = null; stopAll(); timeline(); showPanel(); changed(); return; }
    if (d.mup != null) { $("audioIn").click(); return; }
    if (d.tprev != null) {
      var tu = TUNES[+d.tprev]; if (tunePrev) { tunePrev.pause(); tunePrev = null; }
      b.textContent = "…"; makeTune(tu).then(function (bl) { b.textContent = "▶ Listen"; tunePrev = new Audio(URL.createObjectURL(bl)); tunePrev.play(); setTimeout(function () { if (tunePrev) tunePrev.pause(); }, 9000); }, function (e) { Office.toast(e.message, true); });
      return;
    }
    if (d.tuse != null) { var tu2 = TUNES[+d.tuse]; b.textContent = "…"; makeTune(tu2).then(function (bl) { return storeMedia(bl, "audio", tu2.n + " (made in Vids)"); }).then(setMusic, function (e) { Office.toast(e.message || e, true); }); return; }
  });
  document.querySelectorAll(".rail [data-p]").forEach(function (b) { b.onclick = function () { panel = panel === b.dataset.p && window.innerWidth < 860 ? null : b.dataset.p; showPanel(); }; });
  var fitT = null; function fitSoon() { clearTimeout(fitT); fitT = setTimeout(function () { fitZoom(); draw(); }, 40); }
  window.addEventListener("resize", fitSoon);

  /* ---------------- the bar for what's picked ---------------- */
  function ctxBar() {
    var c = $("ctx"), e = mode === "edit" ? selEl() : null, h = "";
    if (!proj) { c.innerHTML = ""; return; }
    if (!e) {
      c.innerHTML = '<span class="hint">Scene ' + (cur + 1) + " · " + SC().dur + "s · Double-tap words to change them</span><span class=\"sep\"></span>" +
        '<button type="button" class="cb" data-a="text">🅣 Add text</button><button type="button" class="cb" data-a="media">🖼 Add media</button><button type="button" class="cb" data-a="voice">🎙 Voiceover</button>' +
        '<span class="sep"></span><button type="button" class="cb" data-a="undo" title="Undo (Ctrl+Z)">↶</button><button type="button" class="cb" data-a="redo" title="Redo">↷</button>';
      return;
    }
    function colorBtn(key, val, label) { var solid = /^lin:/.test(val || "") ? val.slice(4).split(",")[0] : val; return '<label class="cb" title="' + label + '"><span class="chip" style="background:' + esc(S.cssFill(val || "transparent")) + '"></span><input type="color" data-col="' + key + '" value="' + (/^#[0-9a-f]{6}$/i.test(solid || "") ? solid : "#000000") + '" aria-label="' + label + '"></label>'; }
    if (e.t === "text") {
      h += '<select data-k="font" aria-label="Font">' + S.FONTS.map(function (f) { return '<option value="' + f[0] + '"' + (e.font === f[0] ? " selected" : "") + ">" + f[0] + "</option>"; }).join("") + "</select>" +
        '<input type="number" data-k="size" min="6" max="600" value="' + Math.round(e.size) + '" aria-label="Size">' + colorBtn("color", e.color, "Text colour") +
        '<button type="button" class="cb" data-tg="bold" aria-pressed="' + !!e.bold + '"><b>B</b></button><button type="button" class="cb" data-tg="italic" aria-pressed="' + !!e.italic + '"><i>I</i></button>' +
        '<button type="button" class="cb" data-a="align">' + ({ left: "⬅︎≡", center: "≡", right: "≡➡︎" }[e.align] || "≡") + "</button>" +
        '<button type="button" class="cb" data-tg="sh" aria-pressed="' + !!e.sh + '" title="Shadow">🌑</button>' +
        '<button type="button" class="cb" data-a="box" aria-pressed="' + !!e.bgc + '" title="Box behind the words">▭ Box</button>' + (e.bgc ? colorBtn("bgc", e.bgc, "Box colour") : "") +
        '<button type="button" class="cb" data-a="outline" aria-pressed="' + !!(e.stroke && e.sw) + '" title="Outline">◌ Outline</button>';
    } else if (e.t === "shape") {
      h += (e.shape === "line" ? colorBtn("stroke", e.stroke, "Line colour") : colorBtn("fill", e.fill, "Colour")) + '<button type="button" class="cb" data-tg="sh" aria-pressed="' + !!e.sh + '">🌑 Shadow</button>';
    } else if (e.t === "img") {
      h += '<button type="button" class="cb" data-a="asbg">🖼 Fill the scene</button><select data-k="mask" aria-label="Shape"><option value="">Box</option><option value="circle"' + (e.mask === "circle" ? " selected" : "") + '>Circle</option><option value="round"' + (e.mask === "round" ? " selected" : "") + ">Rounded</option></select>";
    } else if (e.t === "video") {
      var m = mediaOf(e.src);
      h += '<button type="button" class="cb" data-a="mute" aria-pressed="' + !!e.mute + '">' + (e.mute ? "🔇 Sound off" : "🔊 Sound on") + "</button>" +
        '<label class="hint">Start the clip at <input type="number" data-k="from" min="0" step="0.5" max="' + Math.floor(m && m.dur || 600) + '" value="' + (e.from || 0) + '" aria-label="Start the clip at (seconds)"> s</label>' +
        '<button type="button" class="cb" data-a="fit">⛶ Fill the screen</button>';
    } else if (e.t === "emoji") h += '<span class="hint">' + esc(e.ch) + "</span>";
    h += '<span class="sep"></span><label class="hint">✨ <select data-k="anim" aria-label="How it comes in">' + ANIMS.map(function (a) { return '<option value="' + a[0] + '"' + ((e.anim || "none") === a[0] ? " selected" : "") + ">" + a[1] + "</option>"; }).join("") + "</select></label>" +
      '<label class="hint">at <input type="number" data-k="at" min="0" step="0.25" max="120" value="' + (e.at || 0) + '" aria-label="Comes in after (seconds)"> s</label>' +
      '<button type="button" class="cb" data-a="preview" title="See it come in">▶ Preview</button>' +
      '<label class="hint" title="See-through">◐ <input type="range" data-k="o" min="0.05" max="1" step="0.05" value="' + (e.o == null ? 1 : e.o) + '" style="width:70px;accent-color:var(--a1)"></label>' +
      '<label class="hint" title="Turn">⟳ <input type="range" data-k="r" min="-180" max="180" step="1" value="' + (e.r || 0) + '" style="width:70px;accent-color:var(--a1)"></label>' +
      '<button type="button" class="cb" data-a="fwd" title="Bring forward">⬆︎</button><button type="button" class="cb" data-a="back" title="Send backward">⬇︎</button>' +
      '<button type="button" class="cb" data-a="dup" title="Copy">⧉</button><button type="button" class="cb" data-a="del" title="Delete">🗑</button><button type="button" class="cb" data-a="undo" title="Undo">↶</button>';
    c.innerHTML = h;
  }
  var csnap = false;
  function setProp(k, v, nosnap) { var e = selEl(); if (!e) return; if (!nosnap) snap(); e[k] = v; if (e.t === "text") e.h = S.textHeight(meas, e); draw(); placeEditor(); changed(); }
  $("ctx").addEventListener("input", function (ev) {
    var tg = ev.target, e = selEl(); if (!e) return;
    if (tg.dataset.col) { if (!csnap) { snap(); csnap = true; } e[tg.dataset.col] = tg.value; tg.previousElementSibling.style.background = tg.value; draw(); placeEditor(); changed(); }
    if (tg.dataset.k === "o" || tg.dataset.k === "r") { if (!csnap) { snap(); csnap = true; } e[tg.dataset.k] = +tg.value; draw(); changed(); }
  });
  $("ctx").addEventListener("change", function (ev) {
    var tg = ev.target, k = tg.dataset.k; csnap = false; if (!k || k === "o" || k === "r") return;
    if (k === "font") setProp("font", tg.value);
    else if (k === "size") setProp("size", S.num(tg.value, 40, 6, 600));
    else if (k === "anim") { setProp("anim", tg.value); previewEl(); }
    else if (k === "at") setProp("at", S.num(tg.value, 0, 0, 120));
    else if (k === "from") setProp("from", S.num(tg.value, 0, 0, 36000));
    else if (k === "mask") setProp("mask", tg.value || undefined);
  });
  $("ctx").addEventListener("click", function (ev) {
    var b = ev.target.closest("button"); if (!b) return;
    var a = b.dataset.a, e = selEl();
    if (b.dataset.tg) { setProp(b.dataset.tg, !e[b.dataset.tg]); ctxBar(); return; }
    if (a === "text") { panel = "text"; showPanel(); return; } if (a === "media") { panel = "media"; showPanel(); return; } if (a === "voice") { panel = "rec"; showPanel(); return; }
    if (a === "undo") { doUndo(); return; } if (a === "redo") { doRedo(); return; }
    if (!e) return;
    if (a === "align") { setProp("align", { left: "center", center: "right", right: "left" }[e.align] || "center"); ctxBar(); return; }
    if (a === "box") { setProp("bgc", e.bgc ? undefined : "#e11d48"); ctxBar(); return; }
    if (a === "outline") { if (e.stroke && e.sw) { snap(); delete e.stroke; e.sw = 0; } else { snap(); e.stroke = "#000000"; e.sw = Math.max(1, e.size / 20); } e.h = S.textHeight(meas, e); draw(); changed(); ctxBar(); return; }
    if (a === "asbg") { snap(); SC().bg = { img: e.src, dim: 0 }; SC().els = SC().els.filter(function (x) { return x.id !== e.id; }); sel = null; draw(); tlSoon(); ctxBar(); changed(); return; }
    if (a === "mute") { setProp("mute", !e.mute); var v = clipFor(e); if (v) v.muted = !!e.mute; ctxBar(); return; }
    if (a === "fit") { snap(); e.x = 0; e.y = 0; e.w = proj.W; e.h = proj.H; e.r = 0; draw(); changed(); return; }
    if (a === "preview") { previewEl(); return; }
    if (a === "fwd" || a === "back") { var els = SC().els, i = els.indexOf(e), j = i + (a === "fwd" ? 1 : -1); if (j < 0 || j >= els.length) return; snap(); els[i] = els[j]; els[j] = e; draw(); changed(); return; }
    if (a === "dup") { snap(); var c = JSON.parse(JSON.stringify(e)); c.id = uid(); c.x += 20; c.y += 20; SC().els.push(c); select(c.id); draw(); changed(); return; }
    if (a === "del") { snap(); stopEdit(); SC().els = SC().els.filter(function (x) { return x !== e; }); sel = null; draw(); ctxBar(); tlSoon(); changed(); return; }
  });
  /* play this scene from just before the picked thing comes in */
  function previewEl() {
    var e = selEl(); if (!e) return;
    t = startOf(cur) + Math.max(0, (e.at || 0) - 0.2); play();
    var until = startOf(cur) + Math.min(SC().dur, (e.at || 0) + 1.6), id = loopId;
    (function chk() { if (id !== loopId) return; if (t >= until) { pause(); mode = "edit"; t = startOf(cur); draw(); return; } requestAnimationFrame(chk); })();
  }

  /* ---------------- keys ---------------- */
  document.addEventListener("keydown", function (ev) {
    var a = document.activeElement, typing = a && (/INPUT|TEXTAREA|SELECT/.test(a.tagName) || a.isContentEditable);
    var mod = ev.ctrlKey || ev.metaKey, k = ev.key.toLowerCase();
    if (mod && k === "z" && !typing) { ev.preventDefault(); if (ev.shiftKey) doRedo(); else doUndo(); return; }
    if (mod && k === "y" && !typing) { ev.preventDefault(); doRedo(); return; }
    if (typing || !proj) return;
    if (k === " " && !(a && a.tagName === "BUTTON")) { ev.preventDefault(); if (playing) pause(); else play(); return; }
    var e = selEl(); if (!e) return;
    if (k === "delete" || k === "backspace") { ev.preventDefault(); snap(); SC().els = SC().els.filter(function (x) { return x !== e; }); sel = null; draw(); ctxBar(); changed(); return; }
    if (k === "enter" && e.t === "text") { ev.preventDefault(); startEdit(e); return; }
    var mv = { arrowleft: [-1, 0], arrowright: [1, 0], arrowup: [0, -1], arrowdown: [0, 1] }[k];
    if (mv) { ev.preventDefault(); snap(); var st = ev.shiftKey ? 10 : 2; e.x += mv[0] * st; e.y += mv[1] * st; draw(); changed(); }
  });

  /* ---------------- making the video file ---------------- */
  $("exportBtn").onclick = function () {
    if (!proj) return; pause(); stopEdit(); saveNow();
    var type = pickType("video");
    var d = document.createElement("dialog"); d.className = "of-dialog";
    if (type === null || !HTMLCanvasElement.prototype.captureStream) {
      d.innerHTML = '<div class="of-dhead"><h2>Make the video</h2><button type="button" class="of-x" data-x>✕</button></div><p>This browser can\'t make video files. Try Chrome, Edge or Safari on a computer or iPad.</p>';
    } else {
      var hd = proj.ratio === "1:1" ? [1080, 1080] : null;
      d.innerHTML = '<div class="of-dhead"><h2>🎬 Make the video</h2><button type="button" class="of-x" data-x aria-label="Close">✕</button></div>' +
        '<p class="of-where">It plays through once while it\'s being made (' + fmt(total()) + "), so keep this tab open and on screen.</p>" +
        '<p class="of-where">It\'s made in 720p (' + Math.round(proj.W * vq()) + " × " + Math.round(proj.H * vq()) + "), up to 30 minutes long" + (total() > 1800 ? ' - <b style="color:#dc2626">this one is ' + fmt(total()) + ", so make it shorter first</b>" : "") + ".</p>" +
        '<div class="of-new"><button type="button" class="of-btn" data-go>⬇ Make it</button></div><div class="prog" hidden><div></div></div><p class="of-where" id="exMsg"></p>';
    }
    document.body.appendChild(d); d.showModal();
    var q = vq(), busy = null;
    d.addEventListener("click", function (ev) {
      var b = ev.target.closest("button"); if (ev.target === d && !busy) { d.close(); d.remove(); return; } if (!b) return;
      if (b.dataset.x != null) { if (busy) busy.cancel(); d.close(); d.remove(); return; }
      if (b.dataset.q) { q = +b.dataset.q; d.querySelectorAll("[data-q]").forEach(function (x) { x.setAttribute("aria-pressed", String(x === b)); }); return; }
      if (b.dataset.go != null && !busy) { if (total() > 1800) { Office.toast("Videos can be up to 30 minutes - make it shorter first.", true); return; } b.disabled = true; busy = makeVideo(q, type, d); }
    });
  };
  /* how much to scale the screen by for a 720p video */
  function vq() { return Math.min(1, 1280 / Math.max(proj.W, proj.H), 720 / Math.min(proj.W, proj.H)); }
  function makeVideo(q, type, d) {
    var W = Math.round(proj.W * q / 2) * 2, H = Math.round(proj.H * q / 2) * 2, c = document.createElement("canvas"); c.width = W; c.height = H;
    c.style.cssText = "position:fixed;left:-99999px;top:0"; document.body.appendChild(c);
    var g = c.getContext("2d"), bar = d.querySelector(".prog"), fill = bar.firstChild, msg = d.querySelector("#exMsg"), cancelled = false;
    bar.hidden = false; msg.textContent = "Getting ready…";
    var allEls = []; proj.scenes.forEach(function (s) { allEls = allEls.concat(s.els); });
    var job = { cancel: function () { cancelled = true; } };
    S.fontsReady(allEls).then(function () {
      audioOn(); speaker.gain.value = 0; // quiet while it's being made
      var vs = c.captureStream(30), tracks = vs.getVideoTracks().slice();
      if (dest) dest.stream.getAudioTracks().forEach(function (tr) { tracks.push(tr); });
      var r = new MediaRecorder(new MediaStream(tracks), Object.assign(type ? { mimeType: type } : {}, { videoBitsPerSecond: 1.5e6, audioBitsPerSecond: 96000 }));
      var chunks = []; r.ondataavailable = function (e) { if (e.data && e.data.size) chunks.push(e.data); };
      var T = total(), started = 0;
      r.onstop = function () {
        speaker.gain.value = 1; stopAll(); c.remove();
        if (cancelled) { return; }
        var mime = (r.mimeType || type || "video/webm").split(";")[0], ext = /mp4/.test(mime) ? ".mp4" : ".webm", blob = new Blob(chunks, { type: mime });
        fill.style.width = "100%"; msg.innerHTML = "Done! " + Office.size(blob.size) + " · " + (ext === ".webm" ? "WebM video - plays in any browser, YouTube and most apps." : "MP4 video.");
        var v = document.createElement("video"); v.controls = true; v.playsInline = true; v.src = URL.createObjectURL(blob); v.style.cssText = "width:100%;border-radius:12px;margin-top:10px;background:#000"; d.appendChild(v);
        Office.download(blob, Office.safeName($("title").value.trim() || "My video", ext));
        // post it on the Launchpad, or keep it in My Files (in the cloud)
        var pb = document.createElement("div"); pb.className = "of-new"; pb.style.marginTop = "12px";
        pb.innerHTML = '<button type="button" class="of-btn" data-post="1">🎬 Post it on the Launchpad</button><button type="button" class="of-btn ghost" data-post="0">☁️ Save to My Files (just me)</button>';
        d.appendChild(pb);
        var pm = document.createElement("p"); pm.className = "of-where"; d.appendChild(pm);
        pb.onclick = function (ev) {
          var bt = ev.target.closest("[data-post]"); if (!bt) return;
          var show = bt.dataset.post === "1";
          pb.querySelectorAll("button").forEach(function (x) { x.disabled = true; }); pm.textContent = "Uploading " + Office.size(blob.size) + "…";
          Office.init("doc").then(function (st) {
            if (!st.user) throw new Error("Sign in on the Launchpad first.");
            return Office.postVideo(blob, $("title").value.trim() || "My video", "", total(), show, {
              w: W, h: H, noShrink: true, onstep: function (t) { pm.textContent = t; },
              onprog: function (f) { pm.textContent = pm.textContent.replace(/ \d+%$/, "") + " " + Math.round(f * 100) + "%"; } });
          }).then(function (r) {
            pm.innerHTML = (show ? (proj.ratio === "9:16" && total() <= 100 ? "Posted as a Short! " : "Posted! ") : "Saved! ") + '<a href="https://hosting.civilinitiatives.org/watch/?v=' + encodeURIComponent(r.id) + '" target="_blank" rel="noopener">Watch it →</a>' +
              (show ? " Only people signed in to the Launchpad can see it." : " Only you can see it - find it in My Files.");
          }).catch(function (e2) { pm.textContent = e2.message || String(e2); pb.querySelectorAll("button").forEach(function (x) { x.disabled = false; }); });
        };
        t = 0; mode = "edit"; draw(); progress();
      };
      sync(0, true); t = 0;
      r.start(500); started = performance.now();
      msg.textContent = "Making your video… keep this tab open.";
      (function tick() {
        if (cancelled) { r.stop(); return; }
        var tt = (performance.now() - started) / 1000;
        if (tt >= T + 0.15) { sync(T, true); setTimeout(function () { r.stop(); }, 120); return; }
        t = Math.min(tt, T); sync(t, false);
        g.setTransform(q, 0, 0, q, 0, 0); frame(g, t, q);
        fill.style.width = Math.round(t / T * 100) + "%";
        requestAnimationFrame(tick);
      })();
    }).catch(function (e) { msg.textContent = "Couldn't make it: " + (e.message || e); });
    return job;
  }

  /* ---------------- starting, opening, the list of videos ---------------- */
  function startDialog(addOnly) {
    var d = document.createElement("dialog"), ratio = proj ? proj.ratio : "16:9"; d.className = "of-dialog ofwide";
    d.innerHTML = '<div class="of-dhead"><h2>' + (addOnly ? "✨ Add scenes from a template" : "🎬 Make a video") + '</h2><button type="button" class="of-x" data-x aria-label="Close">✕</button></div>' +
      (addOnly ? "" : '<p class="of-where">What shape?</p><div class="ratios">' + [["16:9", "Wide", "YouTube, class screens"], ["9:16", "Tall", "Phones, Shorts, Reels"], ["1:1", "Square", "Posts"]].map(function (r) {
        var k = 34 / Math.max(RATIOS[r[0]][0], RATIOS[r[0]][1]);
        return '<button type="button" data-r="' + r[0] + '" aria-pressed="' + (r[0] === ratio) + '"><span class="shp" style="width:' + RATIOS[r[0]][0] * k + "px;height:" + RATIOS[r[0]][1] * k + 'px"></span><span><b>' + r[1] + "</b><br><small>" + r[2] + "</small></span></button>";
      }).join("") + "</div>") +
      '<p class="of-where">Start with</p><div class="tpls">' + (addOnly ? "" : '<button type="button" data-t="-1"><canvas></canvas><b>Blank</b><small>One empty scene</small></button>') +
      TPLS.map(function (tp, i) { return '<button type="button" data-t="' + i + '"><canvas></canvas><b>' + tp.n + "</b><small>" + tp.d + " · " + tp.sc.length + " scenes</small></button>"; }).join("") + "</div>";
    document.body.appendChild(d); d.showModal();
    function thumbs() {
      d.querySelectorAll(".tpls button").forEach(function (b) {
        var i = +b.dataset.t, c = b.querySelector("canvas"), R = RATIOS[ratio], k = 320 / R[0], keep = proj;
        c.width = 320; c.height = 180; var g = c.getContext("2d"); g.fillStyle = "#000"; g.fillRect(0, 0, 320, 180);
        proj = clean({ ratio: ratio }); var sc = i < 0 ? blankScene() : fromTpl(TPLS[i])[0];
        var kk = Math.min(320 / R[0], 180 / R[1]); g.translate((320 - R[0] * kk) / 2, (180 - R[1] * kk) / 2); g.scale(kk, kk);
        drawScene(g, sc, sc.dur, false); proj = keep;
      });
    }
    thumbs();
    d.addEventListener("click", function (ev) {
      var b = ev.target.closest("button");
      if ((ev.target === d || (b && b.dataset.x != null))) { d.close(); d.remove(); if (!proj) newProject("16:9", -1); return; }
      if (!b) return;
      if (b.dataset.r) { ratio = b.dataset.r; d.querySelectorAll("[data-r]").forEach(function (x) { x.setAttribute("aria-pressed", String(x === b)); }); thumbs(); return; }
      if (b.dataset.t != null) {
        d.close(); d.remove(); var i = +b.dataset.t;
        if (addOnly) { snap(); var sc = fromTpl(TPLS[i]); proj.scenes.splice.apply(proj.scenes, [cur + 1, 0].concat(sc)); cur++; t = startOf(cur); draw(); timeline(); showPanel(); changed(); }
        else newProject(ratio, i);
      }
    });
    d.addEventListener("cancel", function () { setTimeout(function () { d.remove(); if (!proj) newProject("16:9", -1); }, 0); });
  }
  function newProject(ratio, tpl) {
    saveNow(); pause();
    meta = { id: uid() + uid(), title: tpl >= 0 ? TPLS[tpl].n : "Untitled video" }; $("title").value = meta.title;
    proj = clean({ ratio: ratio }); if (tpl >= 0) proj.scenes = fromTpl(TPLS[tpl]);
    cur = 0; sel = null; t = 0; mode = "edit"; undo = []; redo = []; panel = window.innerWidth < 860 ? null : "scene";
    try { localStorage.setItem("hl.vids.last", meta.id); } catch (e) {}
    history.replaceState(null, "", "?id=" + meta.id);
    fitZoom(); draw(); timeline(); ctxBar(); showPanel(); saveNow();
  }
  function openProject(id) {
    return getOne("projects", id).then(function (r) {
      if (!r) { startDialog(false); return; }
      pause(); meta = { id: r.id, title: r.title }; $("title").value = r.title;
      proj = clean(r.body); cur = 0; sel = null; t = 0; mode = "edit"; undo = []; redo = [];
      try { localStorage.setItem("hl.vids.last", r.id); } catch (e) {}
      history.replaceState(null, "", "?id=" + r.id);
      $("status").textContent = "Saved on this device";
      return loadProjectMedia(r.id).then(function () { fitZoom(); draw(); timeline(); ctxBar(); showPanel(); });
    });
  }
  $("projBtn").onclick = function () {
    saveNow().then(function () { return all("projects"); }).then(function (rows) {
      rows = (rows || []).sort(function (a, b) { return b.updated - a.updated; });
      var d = document.createElement("dialog"); d.className = "of-dialog";
      d.innerHTML = '<div class="of-dhead"><h2>🎬 My videos</h2><button type="button" class="of-x" data-x aria-label="Close">✕</button></div>' +
        '<p class="of-where">Saved on this device, in this browser. They don\'t go to your account (videos are too big) - make the video file to keep or share it.</p>' +
        '<div class="of-new"><button type="button" class="of-btn" data-new>＋ New video</button></div><div class="plist">' +
        (rows.map(function (r) {
          return '<div class="row"><button type="button" class="op" data-op="' + esc(r.id) + '"><img alt="" src="' + esc(r.thumb || "") + '"><span><strong>' + esc(r.title || "Untitled video") + "</strong><small>" +
            (r.id === meta.id ? "Open now · " : "") + new Date(r.updated).toLocaleString("en-GB", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) + '</small></span></button><button type="button" class="of-del" data-del="' + esc(r.id) + '" aria-label="Delete">🗑</button></div>';
        }).join("") || '<p class="of-where">No videos yet.</p>') + "</div>";
      document.body.appendChild(d); d.showModal();
      d.addEventListener("click", function (ev) {
        var b = ev.target.closest("button"); if (ev.target === d || (b && b.dataset.x != null)) { d.close(); d.remove(); return; } if (!b) return;
        if (b.dataset.new != null) { d.close(); d.remove(); startDialog(false); return; }
        if (b.dataset.op) { d.close(); d.remove(); openProject(b.dataset.op); return; }
        if (b.dataset.del) {
          var id = b.dataset.del;
          if (b.dataset.sure !== "1") { b.dataset.sure = "1"; b.textContent = "Delete?"; return; }
          delOne("projects", id).then(function () { return all("media"); }).then(function (ms) {
            return Promise.all((ms || []).filter(function (m) { return m.pid === id; }).map(function (m) { if (MEDIA[m.id]) { URL.revokeObjectURL(MEDIA[m.id].url); delete MEDIA[m.id]; } return delOne("media", m.id); }));
          }).then(function () { b.closest(".row").remove(); Office.toast("Deleted."); if (id === meta.id) { d.close(); d.remove(); proj = null; startDialog(false); } });
        }
      });
    });
  };
  Office.themeToggle($("themeBtn"));
  $("title").addEventListener("input", changed);
  window.addEventListener("beforeunload", function () { saveNow(); });
  document.addEventListener("visibilitychange", function () { if (document.hidden) saveNow(); });
  Office.setImporter(function (f) { if (proj) importFiles([f]); });
  if (document.fonts) document.fonts.addEventListener("loadingdone", function () {
    if (!proj) return; proj.scenes.forEach(function (s) { s.els.forEach(function (e) { if (e.t === "text") e.h = S.textHeight(meas, e); }); }); draw(); tlSoon();
  });
  // start: the video from last time, or a new one
  var q = new URLSearchParams(location.search).get("id"), last = null;
  try { last = localStorage.getItem("hl.vids.last"); } catch (e) {}
  db().then(function (d) {
    if (!d) Office.toast("This browser won't let Vids keep your work - it will be lost when you close the tab. Make the video before you go.", true);
    if (q || last) return openProject(q || last); startDialog(false);
  }).catch(function () { startDialog(false); });
})();
