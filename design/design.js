/* Launchpad Design - make posters, social posts, thumbnails, invitations
   and certificates, like Canva. Pick a size, start from a template or a
   blank page, add text, shapes, stickers and pictures, then download.
   Saved to your account (and you can work on it together, live). */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var S = window.Scene, esc = Office.esc;
  var cv = $("cv"), ctx = cv.getContext("2d"), ov = $("ov"), stage = $("stage"), pagebox = $("pagebox");
  var doc = null, cur = 0, sel = null, zoom = 1, fit = true, panel = "tpl", editing = null;
  var undo = [], redo = [], file = { id: null, title: "Untitled design" };
  var meas = document.createElement("canvas").getContext("2d");
  function uid() { return Math.random().toString(36).slice(2, 10); }
  function P() { return doc.pages[cur]; }
  function find(id) { var e = null; (P().els || []).forEach(function (x) { if (x.id === id) e = x; }); return e; }
  function selEl() { return sel ? find(sel) : null; }
  function M() { return Math.min(doc.w, doc.h); }

  /* ---------------- sizes and templates ---------------- */
  var SIZES = [
    ["Instagram post", 1080, 1080, "Square social post"], ["Story", 1080, 1920, "Instagram, TikTok, Snapchat"],
    ["Poster", 1240, 1754, "A4 portrait - prints well"], ["Flyer", 874, 1240, "A5 portrait"],
    ["Presentation", 1920, 1080, "16:9 slide"], ["YouTube thumbnail", 1280, 720, "16:9"],
    ["Certificate", 1754, 1240, "A4 landscape"], ["Invitation", 1050, 1470, "5 × 7 card"],
    ["Banner", 1500, 500, "Website or channel banner"], ["Logo", 500, 500, "Square"],
    ["Bookmark", 600, 1800, "Tall and thin"], ["Phone wallpaper", 1080, 2340, "Tall"]
  ];
  /* Templates are drawn in fractions of the page, so each one fits any size.
     s = text size as a fraction of the page's shorter side. */
  var TPL = [
    { n: "Bold poster", bg: { g: ["#7c3aed", "#db2777", 135] }, els: [
      { t: "shape", shape: "circle", x: .62, y: -.08, w: .5, h: .5, fill: "#f9a8d4", o: .35 },
      { t: "shape", shape: "circle", x: -.15, y: .7, w: .45, h: .45, fill: "#c4b5fd", o: .35 },
      { t: "text", text: "SUMMER\nFEST", x: .08, y: .2, w: .84, s: .2, font: "Anton", color: "#ffffff", align: "center", lh: .95, sh: 1 },
      { t: "text", text: "Music · Food · Games", x: .1, y: .62, w: .8, s: .055, font: "Poppins", color: "#fdf2f8", align: "center", bold: 1 },
      { t: "text", text: "SAT 14 JUNE · 12 - 5PM", x: .2, y: .76, w: .6, s: .04, font: "Inter", color: "#7c3aed", align: "center", bold: 1, bgc: "#ffffff" },
      { t: "emoji", ch: "🎉", x: .74, y: .06, w: .16, h: .16 } ] },
    { n: "School event", bg: { c: "#ffffff" }, els: [
      { t: "shape", shape: "rect", x: 0, y: 0, w: 1, h: .3, fill: "lin:#0ea5e9,#2563eb,90" },
      { t: "text", text: "Open Evening", x: .08, y: .08, w: .84, s: .11, font: "Playfair Display", color: "#ffffff", bold: 1 },
      { t: "text", text: "Come and see what we've been learning!", x: .08, y: .36, w: .84, s: .05, font: "Inter", color: "#0f172a", bold: 1 },
      { t: "text", text: "Thursday 7pm in the main hall.\nProjects, performances and a bake sale.\nEveryone welcome.", x: .08, y: .5, w: .84, s: .04, font: "Inter", color: "#334155", lh: 1.5 },
      { t: "emoji", ch: "🏫", x: .72, y: .7, w: .2, h: .2 } ] },
    { n: "Book review", bg: { c: "#fef3c7" }, els: [
      { t: "text", text: "BOOK REVIEW", x: .08, y: .07, w: .84, s: .085, font: "Archivo Black", color: "#92400e", align: "center", ls: 4 },
      { t: "shape", shape: "round", x: .3, y: .2, w: .4, h: .3, fill: "#fde68a", stroke: "#92400e", sw: 4 },
      { t: "text", text: "Add the cover here", x: .3, y: .33, w: .4, s: .035, font: "Inter", color: "#92400e", align: "center" },
      { t: "text", text: "⭐⭐⭐⭐⭐", x: .2, y: .54, w: .6, s: .06, font: "Inter", color: "#111", align: "center" },
      { t: "text", text: "Title by Author", x: .08, y: .63, w: .84, s: .055, font: "Merriweather", color: "#451a03", align: "center", bold: 1 },
      { t: "text", text: "What it's about, your favourite part, and who you'd recommend it to.", x: .1, y: .73, w: .8, s: .036, font: "Merriweather", color: "#78350f", align: "center", lh: 1.4 } ] },
    { n: "Quote", bg: { c: "#111827" }, els: [
      { t: "text", text: "“", x: .06, y: .02, w: .3, s: .4, font: "Playfair Display", color: "#f59e0b" },
      { t: "text", text: "The more that you read, the more things you will know.", x: .1, y: .32, w: .8, s: .075, font: "Playfair Display", italic: 1, color: "#f9fafb", lh: 1.3 },
      { t: "shape", shape: "line", x: .1, y: .72, w: .2, h: .02, stroke: "#f59e0b", sw: 6 },
      { t: "text", text: "Dr. Seuss", x: .1, y: .76, w: .8, s: .045, font: "Inter", color: "#d1d5db", bold: 1 } ] },
    { n: "Science fair", bg: { g: ["#0f172a", "#1e3a8a", 160] }, els: [
      { t: "shape", shape: "circle", x: .55, y: .08, w: .38, h: .38, fill: "#22d3ee", o: .18 },
      { t: "emoji", ch: "🔬", x: .62, y: .14, w: .24, h: .24 },
      { t: "text", text: "SCIENCE\nFAIR", x: .08, y: .12, w: .56, s: .1, font: "Archivo Black", color: "#67e8f9", lh: 1 },
      { t: "text", text: "Experiments, inventions and big questions - by Year 7.", x: .08, y: .5, w: .84, s: .048, font: "Inter", color: "#e0f2fe", lh: 1.4 },
      { t: "text", text: "Friday · Science block", x: .08, y: .72, w: .84, s: .04, font: "Inter", color: "#0f172a", bold: 1, bgc: "#67e8f9" },
      { t: "emoji", ch: "🧪", x: .78, y: .74, w: .14, h: .14 } ] },
    { n: "Birthday", bg: { c: "#fde047" }, els: [
      { t: "shape", shape: "circle", x: .08, y: .08, w: .06, h: .06, fill: "#ec4899" }, { t: "shape", shape: "circle", x: .82, y: .14, w: .05, h: .05, fill: "#3b82f6" },
      { t: "shape", shape: "circle", x: .15, y: .82, w: .05, h: .05, fill: "#22c55e" }, { t: "shape", shape: "star", x: .78, y: .78, w: .1, h: .1, fill: "#f97316" },
      { t: "emoji", ch: "🎂", x: .36, y: .12, w: .28, h: .28 },
      { t: "text", text: "Happy Birthday!", x: .06, y: .46, w: .88, s: .12, font: "Pacifico", color: "#be185d", align: "center" },
      { t: "text", text: "Party at mine · Saturday 3pm", x: .1, y: .66, w: .8, s: .045, font: "Fredoka", color: "#831843", align: "center", bold: 1 } ] },
    { n: "Certificate", bg: { c: "#fffdf7" }, els: [
      { t: "shape", shape: "rect", x: .03, y: .04, w: .94, h: .92, fill: "none", stroke: "#b45309", sw: 10 },
      { t: "shape", shape: "rect", x: .055, y: .075, w: .89, h: .85, fill: "none", stroke: "#f59e0b", sw: 3 },
      { t: "text", text: "Certificate of Achievement", x: .1, y: .14, w: .8, s: .075, font: "Playfair Display", color: "#78350f", align: "center", bold: 1 },
      { t: "text", text: "PROUDLY PRESENTED TO", x: .1, y: .32, w: .8, s: .035, font: "Inter", color: "#92400e", align: "center", ls: 4 },
      { t: "text", text: "Your Name", x: .1, y: .41, w: .8, s: .11, font: "Pacifico", color: "#1c1917", align: "center" },
      { t: "text", text: "for outstanding effort and brilliant work this term.", x: .15, y: .6, w: .7, s: .038, font: "Merriweather", italic: 1, color: "#44403c", align: "center" },
      { t: "shape", shape: "badge", x: .43, y: .7, w: .14, h: .2, fill: "lin:#fbbf24,#d97706,160" },
      { t: "emoji", ch: "🏅", x: .455, y: .73, w: .09, h: .13 } ] },
    { n: "Gaming thumbnail", bg: { g: ["#020617", "#4c1d95", 120] }, els: [
      { t: "shape", shape: "blob", x: .52, y: .05, w: .5, h: .9, fill: "#a21caf", o: .45 },
      { t: "emoji", ch: "🎮", x: .58, y: .2, w: .36, h: .6 },
      { t: "text", text: "I WON\nEVERY\nROUND!", x: .05, y: .1, w: .55, s: .16, font: "Bangers", color: "#fde047", stroke: "#000000", sw: 6, lh: .95, sh: 1 },
      { t: "shape", shape: "arrow", x: .42, y: .62, w: .2, h: .12, fill: "#ef4444", r: -20 } ] },
    { n: "Big news", bg: { c: "#14b8a6" }, els: [
      { t: "emoji", ch: "📣", x: .1, y: .1, w: .22, h: .22 },
      { t: "text", text: "BIG NEWS", x: .08, y: .36, w: .84, s: .15, font: "Archivo Black", color: "#ffffff" },
      { t: "text", text: "Our class won the reading challenge!", x: .08, y: .58, w: .84, s: .06, font: "Inter", color: "#134e4a", bold: 1, bgc: "#ccfbf1" } ] },
    { n: "Minimal", bg: { c: "#fafaf9" }, els: [
      { t: "text", text: "Simple title", x: .1, y: .38, w: .8, s: .09, font: "Inter", color: "#1c1917", bold: 1 },
      { t: "shape", shape: "line", x: .1, y: .52, w: .12, h: .02, stroke: "#1c1917", sw: 4 },
      { t: "text", text: "A short line of words to go with it.", x: .1, y: .56, w: .8, s: .04, font: "Inter", color: "#57534e" } ] },
    { n: "Bake sale", bg: { c: "#fbcfe8" }, els: [
      { t: "shape", shape: "blob", x: .12, y: .08, w: .76, h: .5, fill: "#f9a8d4" },
      { t: "emoji", ch: "🧁", x: .32, y: .12, w: .36, h: .36 },
      { t: "text", text: "BAKE SALE", x: .06, y: .58, w: .88, s: .13, font: "Fredoka", color: "#9d174d", align: "center", bold: 1, sh: 1 },
      { t: "shape", shape: "badge", x: .72, y: .74, w: .2, h: .2, fill: "#facc15", r: 12 },
      { t: "text", text: "50p", x: .72, y: .8, w: .2, s: .06, font: "Archivo Black", color: "#9d174d", align: "center", r: 12 },
      { t: "text", text: "Break time · Friday", x: .1, y: .78, w: .6, s: .045, font: "Inter", color: "#831843", bold: 1 } ] },
    { n: "Newsletter", bg: { c: "#ffffff" }, els: [
      { t: "shape", shape: "rect", x: 0, y: 0, w: 1, h: .16, fill: "#1e293b" },
      { t: "text", text: "Class 7B News", x: .06, y: .04, w: .6, s: .07, font: "Merriweather", color: "#ffffff", bold: 1 },
      { t: "text", text: "WEEK 6", x: .7, y: .06, w: .24, s: .035, font: "Inter", color: "#fbbf24", bold: 1, align: "right", ls: 3 },
      { t: "text", text: "This week", x: .06, y: .22, w: .42, s: .05, font: "Merriweather", color: "#0f172a", bold: 1 },
      { t: "text", text: "Write what happened this week here. Trips, matches, good news and shout-outs.", x: .06, y: .3, w: .42, s: .032, font: "Inter", color: "#334155", lh: 1.5 },
      { t: "shape", shape: "round", x: .54, y: .22, w: .4, h: .3, fill: "#e2e8f0" },
      { t: "text", text: "Coming up", x: .06, y: .6, w: .88, s: .05, font: "Merriweather", color: "#0f172a", bold: 1 },
      { t: "text", text: "• Monday: maths quiz\n• Wednesday: library visit\n• Friday: non-uniform day", x: .06, y: .68, w: .88, s: .034, font: "Inter", color: "#334155", lh: 1.6 } ] }
  ];
  function fromTpl(t) {
    var W = doc.w, H = doc.h, m = M();
    var els = t.els.map(function (e) {
      var o = { id: uid(), t: e.t, x: e.x * W, y: e.y * H, w: e.w * W, h: (e.h || .1) * H, r: e.r || 0, o: e.o == null ? 1 : e.o };
      ["shape", "fill", "stroke", "font", "color", "align", "text", "ch", "bgc"].forEach(function (k) { if (e[k] != null) o[k] = e[k]; });
      if (e.sw) o.sw = e.sw * m / 1080;
      ["bold", "italic", "sh"].forEach(function (k) { if (e[k]) o[k] = true; });
      if (e.lh) o.lh = e.lh; if (e.ls) o.ls = e.ls * m / 1080;
      if (e.t === "text") { o.size = Math.round(e.s * m); o.h = S.textHeight(meas, o); }
      if (e.t === "emoji") { var sz = Math.min(o.w, o.h); o.x += (o.w - sz) / 2; o.w = o.h = sz; }
      if (e.shape === "circle" && e.w === e.h) { var c = Math.min(o.w, o.h); o.w = o.h = c; }
      return o;
    });
    return { bg: JSON.parse(JSON.stringify(t.bg)), els: els };
  }

  /* ---------------- the file ---------------- */
  function blank(w, h) { return { w: w, h: h, pages: [{ id: uid(), bg: { c: "#ffffff" }, els: [] }] }; }
  function clean(b) {
    if (!b || !Array.isArray(b.pages) || !b.pages.length) return blank(1080, 1080);
    var d = { w: S.num(b.w, 1080, 16, 8000), h: S.num(b.h, 1080, 16, 8000), pages: [] };
    var seen = {};
    b.pages.slice(0, 100).forEach(function (p, i) {
      if (!p || typeof p !== "object") return;
      var id = String(p.id || "").replace(/[^a-z0-9]/gi, "").slice(0, 16); if (!id || seen[id]) id = "pg" + i; seen[id] = 1;
      var es = {}, els = (Array.isArray(p.els) ? p.els : []).slice(0, 400).map(S.cleanEl).filter(function (e) { if (!e || es[e.id]) return false; es[e.id] = 1; return true; });
      d.pages.push({ id: id, bg: S.cleanBg(p.bg), els: els });
    });
    if (!d.pages.length) d.pages.push({ id: uid(), bg: { c: "#ffffff" }, els: [] });
    return d;
  }
  var saver = Office.autosaver($("status"), function () {
    file.title = $("title").value.trim() || "Untitled design";
    return { id: file.id, title: file.title, body: doc };
  }, {
    /* working together: each page and each thing on it is its own part */
    surface: ov,
    where: function () { return ""; },
    apply: function (b, title) {
      if (document.activeElement !== $("title")) $("title").value = title || "Untitled design";
      var pid = P() && P().id, sizeWas = doc.w + "x" + doc.h;
      doc = clean(b);
      var i = doc.pages.map(function (p) { return p.id; }).indexOf(pid); cur = i >= 0 ? i : Math.min(cur, doc.pages.length - 1);
      if (sel && !find(sel)) { sel = null; stopEdit(); }
      if (editing) {
        var e = find(editing), ta = $("ted");
        if (e && ta && ta.value !== e.text) {
          var o = ta.value, nt = e.text, at = ta.selectionStart, p = 0, q = 0, mn = Math.min(o.length, nt.length);
          while (p < mn && o.charCodeAt(p) === nt.charCodeAt(p)) p++;
          while (q < mn - p && o.charCodeAt(o.length - 1 - q) === nt.charCodeAt(nt.length - 1 - q)) q++;
          ta.value = nt; at = at <= p ? at : at >= o.length - q ? at + nt.length - o.length : nt.length - q; ta.selectionStart = ta.selectionEnd = at;
        }
      }
      if (sizeWas !== doc.w + "x" + doc.h) fitZoom();
      draw(); pages(); if (panel === "layers") showPanel(); ctxBar();
    }
  });
  function changed(keepRedo) {
    saver.poke(); thumbSoon(); if (panel === "layers") layersSoon();
  }
  function snap() { undo.push(JSON.stringify(doc)); if (undo.length > 80) undo.shift(); redo = []; }
  function doUndo() { if (!undo.length) return; redo.push(JSON.stringify(doc)); doc = clean(JSON.parse(undo.pop())); afterHistory(); }
  function doRedo() { if (!redo.length) return; undo.push(JSON.stringify(doc)); doc = clean(JSON.parse(redo.pop())); afterHistory(); }
  function afterHistory() { cur = Math.min(cur, doc.pages.length - 1); if (sel && !find(sel)) sel = null; stopEdit(); fitZoom(); draw(); pages(); ctxBar(); changed(); }

  /* ---------------- drawing the page ---------------- */
  var dpr = Math.min(2, window.devicePixelRatio || 1), rq = 0;
  function draw() {
    if (rq) return; rq = requestAnimationFrame(function () { rq = 0; drawNow(); });
  }
  function drawNow() {
    if (!doc) return;
    var W = doc.w, H = doc.h, sw = Math.round(W * zoom), sh = Math.round(H * zoom);
    if (cv.width !== Math.round(sw * dpr) || cv.height !== Math.round(sh * dpr)) { cv.width = Math.round(sw * dpr); cv.height = Math.round(sh * dpr); }
    cv.style.width = sw + "px"; cv.style.height = sh + "px"; pagebox.style.width = sw + "px"; pagebox.style.height = sh + "px";
    ctx.setTransform(zoom * dpr, 0, 0, zoom * dpr, 0, 0);
    S.drawPage(ctx, P(), W, H, { onload: draw, skip: editing });
    selBox();
  }
  function fitZoom() {
    if (!fit || !doc) return;
    var r = stage.getBoundingClientRect(), pad = window.innerWidth < 560 ? 40 : 80;
    zoom = Math.max(0.05, Math.min((r.width - pad) / doc.w, (r.height - pad) / doc.h, 2));
    $("zoomN") && ($("zoomN").textContent = Math.round(zoom * 100) + "%");
  }
  function setZoom(z) { fit = false; zoom = Math.max(0.05, Math.min(4, z)); $("zoomN").textContent = Math.round(zoom * 100) + "%"; draw(); }

  /* the box round what's picked, with handles to resize and turn it */
  var box = document.createElement("div"); box.id = "selbox"; box.hidden = true;
  [["nw", 0, 0], ["n", 50, 0, "side h"], ["ne", 100, 0], ["e", 100, 50, "side"], ["se", 100, 100], ["s", 50, 100, "side h"], ["sw", 0, 100], ["w", 0, 50, "side"]].forEach(function (h) {
    var d = document.createElement("div"); d.className = "hd " + (h[3] || ""); d.dataset.h = h[0]; d.style.left = h[1] + "%"; d.style.top = h[2] + "%"; box.appendChild(d);
  });
  var rotH = document.createElement("div"); rotH.className = "rot"; rotH.dataset.h = "rot"; rotH.textContent = "⟳"; rotH.title = "Turn"; box.appendChild(rotH);
  pagebox.appendChild(box);
  function selBox() {
    var e = selEl();
    if (!e) { box.hidden = true; return; }
    box.hidden = false; box.classList.toggle("locked", !!e.lock);
    box.style.left = e.x * zoom + "px"; box.style.top = e.y * zoom + "px";
    box.style.width = e.w * zoom + "px"; box.style.height = e.h * zoom + "px";
    box.style.transform = "rotate(" + (e.r || 0) + "deg)";
    // text and stickers keep their shape from the corners; sides only where it makes sense
    var sides = e.t === "text" ? ["e", "w"] : e.t === "emoji" ? [] : ["n", "e", "s", "w"];
    box.querySelectorAll(".hd").forEach(function (d) { var k = d.dataset.h; d.hidden = k.length === 1 && sides.indexOf(k) < 0; });
  }

  /* ---------------- picking, moving, resizing, turning ---------------- */
  function toPage(ev) { var r = cv.getBoundingClientRect(); return { x: (ev.clientX - r.left) / zoom, y: (ev.clientY - r.top) / zoom }; }
  function topAt(p) {
    var els = P().els;
    for (var i = els.length - 1; i >= 0; i--) if (S.hit(els[i], p.x, p.y, 4 / zoom)) return els[i];
    return null;
  }
  var act = null, guides = [];
  function showGuides(gs) {
    guides.forEach(function (g) { g.remove(); }); guides = [];
    gs.forEach(function (g) {
      var d = document.createElement("div"); d.className = "guide";
      if (g.v != null) { d.style.left = g.v * zoom + "px"; d.style.top = 0; d.style.width = "1px"; d.style.height = "100%"; }
      else { d.style.top = g.h * zoom + "px"; d.style.left = 0; d.style.height = "1px"; d.style.width = "100%"; }
      pagebox.appendChild(d); guides.push(d);
    });
  }
  function select(id) { if (editing && editing !== id) stopEdit(); sel = id; selBox(); ctxBar(); if (panel === "layers") showPanel(); }
  ov.addEventListener("pointerdown", function (ev) {
    if (ev.button > 0) return;
    ov.focus({ preventScroll: true });
    var p = toPage(ev), e = topAt(p);
    if (editing && e && e.id === editing) return;
    if (!e) { select(null); stopEdit(); return; }
    if (sel !== e.id) select(e.id);
    if (e.lock) return;
    act = { kind: "move", id: e.id, sx: p.x, sy: p.y, ox: e.x, oy: e.y, moved: false, pid: ev.pointerId };
    ov.setPointerCapture(ev.pointerId);
  });
  box.addEventListener("pointerdown", function (ev) {
    var h = ev.target.dataset && ev.target.dataset.h, e = selEl(); if (!h || !e || e.lock) return;
    ev.preventDefault(); ev.stopPropagation();
    var p = toPage(ev);
    act = { kind: h === "rot" ? "rot" : "size", h: h, id: e.id, sx: p.x, sy: p.y, o: JSON.parse(JSON.stringify(e)), moved: false };
    ev.target.setPointerCapture(ev.pointerId);
    ev.target.onpointermove = move; ev.target.onpointerup = ev.target.onpointercancel = function (e2) { ev.target.onpointermove = null; end(e2); };
  });
  ov.addEventListener("pointermove", function (ev) { if (act && act.kind === "move") move(ev); else hover(ev); });
  ov.addEventListener("pointerup", end); ov.addEventListener("pointercancel", end);
  function hover(ev) { var e = topAt(toPage(ev)); ov.style.cursor = e ? (e.lock ? "default" : "move") : "default"; }
  function move(ev) {
    if (!act) return;
    var e = find(act.id); if (!e) { act = null; return; }
    var p = toPage(ev), dx = p.x - act.sx, dy = p.y - act.sy;
    if (!act.moved) { if (Math.abs(dx) * zoom < 2 && Math.abs(dy) * zoom < 2) return; act.moved = true; snap(); }
    if (act.kind === "move") {
      var nx = act.ox + dx, ny = act.oy + dy, gs = [], T = 6 / zoom;
      // snap to the middle and edges of the page
      [[0, 0], [doc.w / 2, e.w / 2], [doc.w, e.w]].forEach(function (s) { if (Math.abs(nx + s[1] - s[0]) < T) { nx = s[0] - s[1]; gs.push({ v: s[0] }); } });
      [[0, 0], [doc.h / 2, e.h / 2], [doc.h, e.h]].forEach(function (s) { if (Math.abs(ny + s[1] - s[0]) < T) { ny = s[0] - s[1]; gs.push({ h: s[0] }); } });
      e.x = nx; e.y = ny; showGuides(gs);
    } else if (act.kind === "rot") {
      var cx = act.o.x + act.o.w / 2, cy = act.o.y + act.o.h / 2;
      var ang = Math.atan2(p.y - cy, p.x - cx) * 180 / Math.PI + 90;
      if (ev.shiftKey) ang = Math.round(ang / 15) * 15;
      else [0, 90, 180, 270, -90, 360].forEach(function (s) { if (Math.abs(ang - s) < 4) ang = s; });
      e.r = Math.round(((ang + 540) % 360) - 180);
    } else resize(e, act, p, ev.shiftKey);
    draw(); if (editing === e.id) placeEditor();
  }
  function resize(e, a, p, free) {
    var o = a.o, h = a.h, rad = (o.r || 0) * Math.PI / 180, cos = Math.cos(rad), sin = Math.sin(rad);
    // the pointer in the element's own (turned) frame
    var dx = p.x - a.sx, dy = p.y - a.sy, lx = dx * cos + dy * sin, ly = -dx * sin + dy * cos;
    var w = o.w, ht = o.h, ox = 0, oy = 0;
    var corner = h.length === 2, keep = (e.t === "text" || e.t === "emoji" || e.t === "img") && corner && !free;
    if (h.indexOf("e") >= 0) w = o.w + lx; if (h.indexOf("w") >= 0) { w = o.w - lx; ox = lx; }
    if (h.indexOf("s") >= 0) ht = o.h + ly; if (h.indexOf("n") >= 0) { ht = o.h - ly; oy = ly; }
    if (keep || e.t === "emoji") {
      var k = Math.max(w / o.w, ht / o.h); if (!isFinite(k) || k <= 0) k = 0.05;
      k = Math.max(k, 10 / Math.min(o.w, o.h));
      w = o.w * k; ht = o.h * k;
      ox = h.indexOf("w") >= 0 ? o.w - w : 0; oy = h.indexOf("n") >= 0 ? o.h - ht : 0;
    }
    w = Math.max(10, w); ht = Math.max(10, ht);
    if (e.t === "text") {
      if (corner) { e.size = Math.max(4, Math.round(o.size * w / o.w)); if (o.sw) e.sw = o.sw * w / o.w; if (o.ls) e.ls = o.ls * w / o.w; }
      e.w = w; ht = S.textHeight(meas, e); if (h.indexOf("n") >= 0) oy = o.h - ht;
      if (!corner) oy = 0;
    }
    // keep the opposite side still, allowing for the turn
    var cxo = ox + w / 2 - o.w / 2, cyo = oy + ht / 2 - o.h / 2;
    var ncx = o.x + o.w / 2 + cxo * cos - cyo * sin, ncy = o.y + o.h / 2 + cxo * sin + cyo * cos;
    e.w = w; e.h = ht; e.x = ncx - w / 2; e.y = ncy - ht / 2;
  }
  function end(ev) {
    if (!act) return;
    var a = act; act = null; showGuides([]);
    if (a.moved) { changed(); ctxBar(); }
  }
  ov.addEventListener("dblclick", function (ev) {
    var e = topAt(toPage(ev)); if (!e || e.lock) return;
    if (e.t === "text") { select(e.id); startEdit(e); }
    else if (e.t === "shape" && e.shape !== "line") { // type on a shape: a text box on top of it
      snap(); var t = textEl("Your text", Math.round(Math.min(e.w, e.h) / 5), { x: e.x + e.w * 0.1, w: e.w * 0.8, align: "center", color: lightOf(e.fill) ? "#111111" : "#ffffff" });
      t.y = e.y + (e.h - t.h) / 2; t.r = e.r; P().els.push(t); select(t.id); startEdit(t, true); changed();
    }
  });
  function lightOf(c) {
    var m = String(c || "").match(/#([0-9a-f]{6})/i); if (!m) return true;
    var n = parseInt(m[1], 16); return ((n >> 16) * 0.299 + ((n >> 8) & 255) * 0.587 + (n & 255) * 0.114) > 150;
  }

  /* ---------------- typing in a text box ---------------- */
  function startEdit(e, all) {
    stopEdit(); editing = e.id;
    var ta = document.createElement("textarea"); ta.id = "ted"; ta.spellcheck = true; ta.value = e.text;
    pagebox.appendChild(ta); placeEditor(); draw();
    ta.focus(); if (all) ta.select(); else ta.setSelectionRange(ta.value.length, ta.value.length);
    var before = JSON.stringify(doc), snapped = false;
    ta.addEventListener("input", function () {
      var x = find(editing); if (!x) return;
      if (!snapped) { undo.push(before); redo = []; snapped = true; }
      x.text = ta.value.slice(0, 5000); x.h = S.textHeight(meas, x); placeEditor(); selBox(); changed();
    });
    ta.addEventListener("keydown", function (k) { if (k.key === "Escape") { k.preventDefault(); stopEdit(); ov.focus(); } k.stopPropagation(); });
    ta.addEventListener("blur", function () { setTimeout(function () { if (document.activeElement !== ta) stopEdit(); }, 120); });
  }
  function placeEditor() {
    var ta = $("ted"), e = find(editing); if (!ta || !e) return;
    var z = zoom;
    ta.style.left = e.x * z + "px"; ta.style.top = e.y * z + "px"; ta.style.width = e.w * z + 2 + "px"; ta.style.height = Math.max(e.h, (e.size || 32) * (e.lh || 1.2)) * z + "px";
    ta.style.font = S.fontOf(e, z); ta.style.lineHeight = (e.lh || 1.2); ta.style.letterSpacing = (e.ls || 0) * z + "px";
    ta.style.textAlign = e.align || "left"; ta.style.color = /^lin:/.test(e.color) ? e.color.slice(4).split(",")[0] : e.color;
    ta.style.textTransform = e.upper ? "uppercase" : "none"; ta.style.textDecoration = e.under ? "underline" : "none";
    ta.style.transform = "rotate(" + (e.r || 0) + "deg)"; ta.style.opacity = e.o == null ? 1 : e.o;
    ta.style.paddingTop = (e.bgc ? (e.size || 32) * 0.25 * z : 0) + ((e.size || 32) * ((e.lh || 1.2) - 1) / 2) * 0 + "px";
    ta.style.background = e.bgc || "transparent";
    ta.style.webkitTextStroke = e.stroke && e.sw ? (e.sw * z) + "px " + e.stroke : "0";
  }
  function stopEdit() {
    var ta = $("ted"); if (!ta) { editing = null; return; }
    var id = editing; editing = null; ta.remove();
    var e = id && find(id);
    if (e && !e.text.trim()) { P().els = P().els.filter(function (x) { return x.id !== id; }); if (sel === id) sel = null; changed(); }
    draw(); ctxBar();
  }

  /* when a font arrives, text boxes are measured again */
  function reflow() {
    if (!doc) return; var any = false;
    doc.pages.forEach(function (p) { p.els.forEach(function (e) { if (e.t === "text") { var h = S.textHeight(meas, e); if (Math.abs(h - e.h) > 0.5) { e.h = h; any = true; } } }); });
    draw(); selBox(); placeEditor(); thumbSoon(); if (panel === "tpl" || panel === "text") showPanelSoon();
    if (any && file.id) saver.poke();
  }
  var spT = null; function showPanelSoon() { clearTimeout(spT); spT = setTimeout(showPanel, 300); }
  if (document.fonts) document.fonts.addEventListener("loadingdone", reflow);

  /* ---------------- adding things ---------------- */
  function textEl(text, size, more) {
    var e = { id: uid(), t: "text", text: text, font: "Inter", size: size, color: "#111111", align: "left", lh: 1.2, ls: 0, x: 0, y: 0, w: doc.w * 0.6, h: 10, r: 0, o: 1 };
    Object.keys(more || {}).forEach(function (k) { e[k] = more[k]; });
    e.h = S.textHeight(meas, e);
    if (!more || more.x == null) e.x = (doc.w - e.w) / 2;
    if (!more || more.y == null) e.y = (doc.h - e.h) / 2;
    return e;
  }
  function add(e, edit) {
    snap(); P().els.push(e); select(e.id); draw(); changed();
    if (edit && e.t === "text") startEdit(e, true);
  }
  function addText(kind) {
    var m = M(), k = { h: [.09, "Add a heading", { bold: true }], s: [.055, "Add a subheading", { bold: true }], b: [.035, "Add a little bit of body text", {}] }[kind];
    var more = Object.assign({ w: doc.w * (kind === "b" ? 0.6 : 0.75) }, k[2]); more.align = kind === "b" ? "left" : "center";
    add(textEl(k[1], Math.round(k[0] * m), more), true);
  }
  var COMBOS = [
    { n: "BIG SALE", s: .12, font: "Anton", color: "#dc2626", align: "center", sh: true },
    { n: "Hello there", s: .1, font: "Pacifico", color: "#7c3aed", align: "center" },
    { n: "Simply stated", s: .07, font: "Playfair Display", italic: true, color: "#1c1917", align: "center" },
    { n: "OUTLINED", s: .1, font: "Archivo Black", color: "#fde047", stroke: "#111111", sw: 4, align: "center" },
    { n: "Highlighted", s: .06, font: "Inter", bold: true, color: "#ffffff", bgc: "#7c3aed", align: "center" },
    { n: "S P A C E D", s: .05, font: "Inter", bold: true, color: "#0f172a", ls: 14, align: "center", upper: true },
    { n: "Rainbow", s: .1, font: "Fredoka", bold: true, color: "lin:#f97316,#db2777,90", align: "center" },
    { n: "Typewriter", s: .055, font: "Courier Prime", color: "#111827", align: "left" }
  ];
  function addCombo(c) {
    var m = M(), more = {}; Object.keys(c).forEach(function (k) { if (k !== "n" && k !== "s") more[k] = c[k]; });
    if (c.sw) more.sw = c.sw * m / 1080; if (c.ls) more.ls = c.ls * m / 1080;
    more.w = doc.w * 0.8;
    add(textEl(c.n, Math.round(c.s * m), more), true);
  }
  function addShape(k) {
    var s = M() * 0.32, e = { id: uid(), t: "shape", shape: k, fill: "#7c3aed", x: (doc.w - s) / 2, y: (doc.h - s) / 2, w: s, h: s, r: 0, o: 1, rad: 0, sw: 0 };
    if (k === "line") { e.h = Math.max(8, s * 0.06); e.y = (doc.h - e.h) / 2; e.w = s * 1.4; e.x = (doc.w - e.w) / 2; e.stroke = "#111111"; e.sw = Math.max(3, s * 0.03); e.fill = "none"; }
    if (k === "arrow") { e.h = s * 0.55; e.y = (doc.h - e.h) / 2; e.fill = "#ef4444"; }
    if (k === "speech") { e.fill = "#ffffff"; e.stroke = "#111111"; e.sw = Math.max(2, s * 0.02); e.w = s * 1.3; e.x = (doc.w - e.w) / 2; }
    add(e);
  }
  function addEmoji(ch) { var s = M() * 0.22; add({ id: uid(), t: "emoji", ch: ch, x: (doc.w - s) / 2, y: (doc.h - s) / 2, w: s, h: s, r: 0, o: 1 }); }
  /* pictures: made smaller so files stay quick to save; see-through bits kept */
  function picture(f, max) {
    max = max || 1400;
    return new Promise(function (ok, no) {
      if (!/^image\//.test(f.type)) { no(new Error("That isn't a picture.")); return; }
      var r = new FileReader();
      r.onload = function () {
        var im = new Image();
        im.onload = function () {
          var k = Math.min(1, max / Math.max(im.width, im.height)), c = document.createElement("canvas");
          c.width = Math.max(1, Math.round(im.width * k)); c.height = Math.max(1, Math.round(im.height * k));
          var g = c.getContext("2d"), see = /png|gif|webp|svg/.test(f.type);
          if (!see) { g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); }
          g.drawImage(im, 0, 0, c.width, c.height);
          var u = c.toDataURL(see ? "image/png" : "image/jpeg", 0.84);
          if (see && u.length > 900000) u = c.toDataURL("image/webp", 0.85);
          ok({ src: u, w: c.width, h: c.height });
        };
        im.onerror = function () { no(new Error("Couldn't read that picture.")); };
        im.src = r.result;
      };
      r.onerror = function () { no(new Error("Couldn't read that file.")); };
      r.readAsDataURL(f);
    });
  }
  function addPictures(files) {
    var list = Array.prototype.slice.call(files || []).slice(0, 12);
    if (!list.length) return;
    Office.toast("Adding " + (list.length > 1 ? list.length + " pictures" : "the picture") + "…");
    list.reduce(function (pr, f, i) {
      return pr.then(function () { return picture(f); }).then(function (r) {
        var k = Math.min(doc.w * 0.6 / r.w, doc.h * 0.6 / r.h), w = r.w * k, h = r.h * k;
        var e = { id: uid(), t: "img", src: r.src, x: (doc.w - w) / 2 + i * 20, y: (doc.h - h) / 2 + i * 20, w: w, h: h, r: 0, o: 1, rad: 0 };
        snap(); P().els.push(e); select(e.id); draw(); changed();
        if (panel === "up") showPanel();
      }).catch(function (e) { Office.toast(e.message, true); });
    }, Promise.resolve());
  }
  $("imgIn").onchange = function () { addPictures(this.files); this.value = ""; };
  $("bgIn").onchange = function () {
    var f = this.files[0]; this.value = ""; if (!f) return;
    picture(f, 2000).then(function (r) { snap(); P().bg = { img: r.src, dim: 0 }; draw(); changed(); showPanel(); }).catch(function (e) { Office.toast(e.message, true); });
  };
  // drop pictures straight onto the page
  stage.addEventListener("dragover", function (e) { e.preventDefault(); });
  stage.addEventListener("drop", function (e) { e.preventDefault(); if (e.dataTransfer && e.dataTransfer.files.length) addPictures(e.dataTransfer.files); });
  document.addEventListener("paste", function (e) {
    if (editing || /INPUT|TEXTAREA/.test((document.activeElement || {}).tagName)) return;
    var fs = []; Array.prototype.forEach.call((e.clipboardData || {}).items || [], function (it) { if (it.kind === "file") fs.push(it.getAsFile()); });
    if (fs.length) { e.preventDefault(); addPictures(fs); return; }
    if (clip) { e.preventDefault(); pasteClip(); }
  });

  /* ---------------- the panel on the left ---------------- */
  var EMOJI = ("⭐ 🌟 ✨ 💥 🔥 ❤️ 💜 💙 💚 💛 🧡 🖤 🎉 🎊 🎈 🎁 🎂 🧁 🍕 🍔 🍦 🍩 🍎 🍓 🌈 ☀️ 🌙 ⚡ ❄️ 🌊 🌸 🌻 🌲 🍀 🐶 🐱 🦊 🐼 🦁 🐸 🐝 🦋 🐢 🦄 🐉 🚀 ✈️ 🚗 ⚽ 🏀 🎾 🏆 🥇 🏅 🎮 🎧 🎵 🎨 📚 ✏️ 📐 🔬 🧪 🧠 💡 🏫 🎓 📣 📌 ✅ ❌ ❓ ❗ 👍 👏 🙌 😀 😂 😍 😎 🤩 🥳 😮 🤔 😴 👀 💯 ➡️ ⬆️ 🔔 ⏰ 📅 🌍 🏠 💬").split(" ");
  var COLOURS = ["#ffffff", "#f8fafc", "#e2e8f0", "#94a3b8", "#334155", "#0f172a", "#000000", "#fee2e2", "#fecaca", "#ef4444", "#b91c1c", "#ffedd5", "#fdba74", "#f97316", "#fef3c7", "#fde047", "#eab308", "#dcfce7", "#86efac", "#22c55e", "#15803d", "#ccfbf1", "#2dd4bf", "#0d9488", "#e0f2fe", "#7dd3fc", "#0ea5e9", "#2563eb", "#1e3a8a", "#ede9fe", "#c4b5fd", "#7c3aed", "#4c1d95", "#fce7f3", "#f9a8d4", "#db2777"];
  var GRADS = [["#7c3aed", "#db2777", 135], ["#06b6d4", "#3b82f6", 135], ["#f97316", "#facc15", 135], ["#22c55e", "#14b8a6", 135], ["#0f172a", "#334155", 180], ["#fbcfe8", "#c4b5fd", 135], ["#fde68a", "#fca5a5", 135], ["#1e3a8a", "#0ea5e9", 200], ["#020617", "#4c1d95", 120], ["#ecfccb", "#bef264", 160], ["#fff7ed", "#fed7aa", 180], ["#111827", "#ef4444", 135]];
  function mini(drawFn, w, h) {
    var c = document.createElement("canvas"), k = 2; c.width = w * k; c.height = h * k;
    var g = c.getContext("2d"); g.scale(k, k); drawFn(g); return c;
  }
  function tplThumb(t, w) {
    var hh = Math.round(w * doc.h / doc.w), saved = cur, keep = doc;
    var page = fromTpl(t);
    return mini(function (g) { g.scale(w / doc.w, hh / doc.h); S.drawPage(g, page, doc.w, doc.h, { onload: null, clean: true }); }, w, hh);
  }
  var layersT = null; function layersSoon() { clearTimeout(layersT); layersT = setTimeout(showPanel, 400); }
  function showPanel() {
    var pn = $("panel");
    document.querySelectorAll(".rail [data-p]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.p === panel)); });
    pn.hidden = !panel; if (!panel) { fitSoon(); return; }
    var h = "";
    var x = '<button type="button" class="x" data-close aria-label="Close the panel">✕</button>';
    if (panel === "tpl") {
      h = x + "<h3>Templates</h3><p style=\"color:var(--soft);font-size:.86rem;margin:0 0 10px\">Tap one to use it on this page. They fit any size.</p><div class=\"gridp two\" id=\"tplGrid\"></div>";
    } else if (panel === "el") {
      h = x + "<h3>Elements</h3><h4>Shapes</h4><div class=\"gridp four\" id=\"shapeGrid\"></div><h4>Stickers</h4><div class=\"gridp six\">" +
        EMOJI.map(function (e) { return '<button type="button" class="emo" data-emo="' + e + '" aria-label="' + e + '">' + e + "</button>"; }).join("") + "</div>";
    } else if (panel === "text") {
      h = x + '<h3>Text</h3><button type="button" class="pbtn" data-tx="h">🅣 Add a text box</button>' +
        '<button type="button" class="addtext" data-tx="h" style="font:800 1.5rem var(--ui)">Add a heading</button>' +
        '<button type="button" class="addtext" data-tx="s" style="font:700 1.1rem var(--ui)">Add a subheading</button>' +
        '<button type="button" class="addtext" data-tx="b" style="font:400 .9rem var(--ui)">Add a little bit of body text</button>' +
        '<h4>Font combinations</h4><div class="gridp two" id="comboGrid"></div>';
    } else if (panel === "up") {
      var pics = [], seen = {};
      doc.pages.forEach(function (p) { p.els.forEach(function (e) { if (e.t === "img" && !seen[e.src]) { seen[e.src] = 1; pics.push(e.src); } }); if (p.bg.img && !seen[p.bg.img]) { seen[p.bg.img] = 1; pics.push(p.bg.img); } });
      h = x + '<h3>Uploads</h3><label class="pbtn">⬆️ Upload pictures<input type="file" accept="image/*" multiple hidden id="upIn"></label>' +
        '<p style="color:var(--soft);font-size:.84rem;margin:0 0 6px">Or drag them onto the page, or paste one. Pictures are made smaller so the design saves quickly.</p>' +
        (pics.length ? '<h4>In this design</h4><div class="gridp">' + pics.map(function (s, i) { return '<button type="button" class="tile" data-pic="' + i + '"><img alt="" src="' + esc(s) + '"></button>'; }).join("") + "</div>" : "");
      showPanel._pics = pics;
    } else if (panel === "bg") {
      var bg = P().bg;
      h = x + "<h3>Background</h3><h4>Colour</h4><div class=\"gridp six\">" +
        COLOURS.map(function (c) { return '<button type="button" class="sw" data-bgc="' + c + '" style="background:' + c + '" aria-label="' + c + '"></button>'; }).join("") +
        '</div><label class="pbtn ghost" style="margin-top:10px">🎨 Any colour<input type="color" id="bgAny" value="' + (bg.c && /^#[0-9a-f]{6}$/i.test(bg.c) ? bg.c : "#ffffff") + '" style="width:0;height:0;opacity:0;border:0;padding:0"></label>' +
        "<h4>Gradients</h4><div class=\"gridp six\">" + GRADS.map(function (g, i) { return '<button type="button" class="sw" data-bgg="' + i + '" style="background:' + S.cssFill("lin:" + g.join(",")) + '" aria-label="Gradient ' + (i + 1) + '"></button>'; }).join("") + "</div>" +
        '<h4>Picture</h4><button type="button" class="pbtn ghost" data-bgimg>🖼 Use a picture as the background</button>' +
        (bg.img ? '<label style="display:grid;grid-template-columns:90px 1fr;gap:8px;align-items:center;font-size:.86rem">Darken<input type="range" min="0" max="0.8" step="0.05" value="' + (bg.dim || 0) + '" id="bgDim" style="accent-color:var(--a1)"></label>' : "") +
        '<button type="button" class="pbtn ghost" data-bgall>Use this background on every page</button>';
    } else if (panel === "layers") {
      var els = P().els.slice().reverse();
      h = x + "<h3>Layers</h3><p style=\"color:var(--soft);font-size:.86rem;margin:0 0 10px\">Top of the list is in front.</p><div class=\"layers\">" +
        (els.map(function (e) {
          var name = e.t === "text" ? (e.text || "Text").slice(0, 40) : e.t === "emoji" ? "Sticker " + e.ch : e.t === "img" ? "Picture" : "Shape: " + e.shape;
          return '<div style="display:flex;gap:4px;align-items:center"><button type="button" class="ly" data-ly="' + e.id + '" aria-current="' + (e.id === sel) + '"><canvas data-lc="' + e.id + '" width="88" height="88"></canvas><span>' + esc(name) + (e.lock ? " 🔒" : "") + '</span></button>' +
            '<button type="button" class="cb" data-up="' + e.id + '" aria-label="Forward">▲</button><button type="button" class="cb" data-dn="' + e.id + '" aria-label="Backward">▼</button></div>';
        }).join("") || "<p style=\"color:var(--soft)\">Nothing on this page yet.</p>") + "</div>";
    }
    pn.innerHTML = h;
    if (panel === "tpl") {
      var g = $("tplGrid");
      TPL.forEach(function (t, i) {
        var b = document.createElement("button"); b.type = "button"; b.className = "tile tpl"; b.dataset.tpl = i; b.setAttribute("aria-label", t.n);
        b.appendChild(tplThumb(t, 130)); var cap = document.createElement("span"); cap.className = "cap"; cap.textContent = t.n; b.appendChild(cap);
        g.appendChild(b);
      });
    }
    if (panel === "el") {
      var sg = $("shapeGrid");
      S.SHAPES.forEach(function (k) {
        var b = document.createElement("button"); b.type = "button"; b.className = "tile"; b.dataset.shape = k; b.setAttribute("aria-label", k);
        b.appendChild(mini(function (gg) {
          gg.translate(12, k === "line" ? 0 : 12);
          S.drawEl(gg, { t: "shape", shape: k, x: 0, y: k === "line" ? 26 : 0, w: 40, h: k === "line" ? 8 : 40, fill: k === "line" ? "none" : "#7c3aed", stroke: "#4c1d95", sw: k === "line" ? 4 : 0, o: 1 });
        }, 64, 64));
        sg.appendChild(b);
      });
    }
    if (panel === "text") {
      var cg = $("comboGrid");
      COMBOS.forEach(function (c, i) {
        var b = document.createElement("button"); b.type = "button"; b.className = "tile"; b.dataset.combo = i; b.setAttribute("aria-label", c.n);
        b.style.aspectRatio = "2"; b.style.background = c.color === "#ffffff" || c.color === "#fde047" ? "#334155" : "#fff";
        b.appendChild(mini(function (gg) {
          var e = { t: "text", text: c.n, font: c.font, size: 22, color: c.color, bold: c.bold, italic: c.italic, stroke: c.stroke, sw: c.sw ? 1.5 : 0, bgc: c.bgc, ls: c.ls ? 4 : 0, upper: c.upper, align: "center", w: 120, x: 4, y: 0, h: 40, lh: 1.1, o: 1, sh: c.sh };
          e.h = S.textHeight(gg, e); e.y = (64 - e.h) / 2; S.drawEl(gg, e);
        }, 128, 64));
        cg.appendChild(b);
      });
    }
    if (panel === "layers") pn.querySelectorAll("canvas[data-lc]").forEach(function (c) {
      var e = find(c.dataset.lc); if (!e) return;
      var g2 = c.getContext("2d"), k = Math.min(80 / e.w, 80 / e.h);
      g2.translate(44 - e.w * k / 2, 44 - e.h * k / 2); g2.scale(k, k); g2.translate(-e.x, -e.y);
      var cp = JSON.parse(JSON.stringify(e)); cp.r = 0; S.drawEl(g2, cp, { onload: layersSoon });
    });
    var ui = $("upIn"); if (ui) ui.onchange = function () { addPictures(this.files); this.value = ""; };
    var ba = $("bgAny"); if (ba) ba.oninput = function () { snap(); P().bg = { c: ba.value }; draw(); changed(); };
    var bd = $("bgDim"); if (bd) bd.oninput = function () { P().bg.dim = +bd.value; draw(); changed(); };
    fitSoon();
  }
  $("panel").addEventListener("click", function (ev) {
    var t = ev.target.closest("button,label"); if (!t) return;
    var d = t.dataset;
    if (d.close != null) { panel = null; showPanel(); return; }
    if (d.tpl != null) {
      var tp = TPL[+d.tpl], has = P().els.length;
      snap(); var made = fromTpl(tp);
      if (has) { // keep what's here: the template goes on a new page after this one
        doc.pages.splice(cur + 1, 0, { id: uid(), bg: made.bg, els: made.els }); cur++;
        Office.toast("Added “" + tp.n + "” as a new page.");
      } else { P().bg = made.bg; P().els = made.els; }
      sel = null; draw(); pages(); ctxBar(); changed(); return;
    }
    if (d.shape) { addShape(d.shape); return; }
    if (d.emo) { addEmoji(d.emo); return; }
    if (d.tx) { addText(d.tx); return; }
    if (d.combo != null) { addCombo(COMBOS[+d.combo]); return; }
    if (d.pic != null) {
      var s = showPanel._pics[+d.pic], m = new Image(); m.onload = function () {
        var k = Math.min(doc.w * 0.5 / m.naturalWidth, doc.h * 0.5 / m.naturalHeight), w = m.naturalWidth * k, h = m.naturalHeight * k;
        add({ id: uid(), t: "img", src: s, x: (doc.w - w) / 2, y: (doc.h - h) / 2, w: w, h: h, r: 0, o: 1 });
      }; m.src = s; return;
    }
    if (d.bgc) { snap(); P().bg = { c: d.bgc }; draw(); changed(); thumbSoon(); return; }
    if (d.bgg != null) { snap(); P().bg = { g: GRADS[+d.bgg].slice() }; draw(); changed(); return; }
    if (d.bgimg != null) { $("bgIn").click(); return; }
    if (d.bgall != null) { snap(); var b = JSON.stringify(P().bg); doc.pages.forEach(function (p) { p.bg = JSON.parse(b); }); draw(); pages(); changed(); Office.toast("Every page has this background now."); return; }
    if (d.ly) { select(d.ly); return; }
    if (d.up) { layer(d.up, 1); return; }
    if (d.dn) { layer(d.dn, -1); return; }
  });
  document.querySelectorAll(".rail [data-p]").forEach(function (b) {
    b.onclick = function () { panel = panel === b.dataset.p && window.innerWidth < 860 ? null : b.dataset.p; showPanel(); };
  });
  var fitT = null; function fitSoon() { clearTimeout(fitT); fitT = setTimeout(function () { if (fit) { fitZoom(); draw(); } }, 30); }

  /* ---------------- the bar for what's picked ---------------- */
  var openPop = null;
  function closePop() { if (openPop) { openPop.remove(); openPop = null; } }
  document.addEventListener("pointerdown", function (e) { if (openPop && !openPop.contains(e.target) && !e.target.closest("[data-pop]")) closePop(); });
  function ctxBar() {
    var c = $("ctx"), e = selEl(); closePop();
    var h = "";
    if (!e) {
      var bg = P().bg;
      h = '<span class="hint">Page ' + (cur + 1) + " of " + doc.pages.length + " · " + doc.w + " × " + doc.h + " px</span><span class=\"sep\"></span>" +
        '<button type="button" class="cb" data-act="bgpanel"><span class="chip" style="background:' + (bg.img ? "url(" + esc(bg.img) + ") center/cover" : bg.g ? S.cssFill("lin:" + bg.g.join(",")) : esc(bg.c)) + '"></span> Background</button>' +
        '<button type="button" class="cb" data-act="tplpanel">✨ Templates</button>' +
        '<button type="button" class="cb" data-act="texth">🅣 Add text</button>' +
        '<span class="sep"></span><button type="button" class="cb" data-act="undo" title="Undo (Ctrl+Z)">↶</button><button type="button" class="cb" data-act="redo" title="Redo (Ctrl+Y)">↷</button>';
      c.innerHTML = h; return;
    }
    function colorBtn(key, val, label) {
      var solid = /^lin:/.test(val || "") ? val.slice(4).split(",")[0] : val;
      return '<label class="cb" title="' + label + '"><span class="chip" style="background:' + esc(S.cssFill(val || "transparent")) + '"></span>' +
        '<input type="color" data-col="' + key + '" value="' + (/^#[0-9a-f]{6}$/i.test(solid || "") ? solid : "#000000") + '" aria-label="' + label + '"></label>';
    }
    if (e.t === "text") {
      h += '<select data-k="font" aria-label="Font">' + S.FONTS.map(function (f) { return '<option value="' + f[0] + '"' + (e.font === f[0] ? " selected" : "") + ' style="font-family:' + f[1].replace(/"/g, "'") + '">' + f[0] + "</option>"; }).join("") + "</select>" +
        '<button type="button" class="cb" data-act="sz-" aria-label="Smaller">−</button><input type="number" data-k="size" min="4" max="800" value="' + Math.round(e.size) + '" aria-label="Text size"><button type="button" class="cb" data-act="sz+" aria-label="Bigger">+</button>' +
        colorBtn("color", e.color, "Text colour") +
        '<button type="button" class="cb" data-tg="bold" aria-pressed="' + !!e.bold + '" title="Bold"><b>B</b></button>' +
        '<button type="button" class="cb" data-tg="italic" aria-pressed="' + !!e.italic + '" title="Italic"><i>I</i></button>' +
        '<button type="button" class="cb" data-tg="under" aria-pressed="' + !!e.under + '" title="Underline"><u>U</u></button>' +
        '<button type="button" class="cb" data-tg="upper" aria-pressed="' + !!e.upper + '" title="Capitals">aA</button>' +
        '<button type="button" class="cb" data-act="align" title="Line up the words">' + ({ left: "⬅︎≡", center: "≡", right: "≡➡︎" }[e.align] || "≡") + "</button>" +
        '<button type="button" class="cb" data-pop="spacing" title="Spacing">↕ Spacing</button>' +
        '<button type="button" class="cb" data-pop="effects" title="Effects">✨ Effects</button>';
    } else if (e.t === "shape") {
      if (e.shape !== "line") h += colorBtn("fill", e.fill, "Fill colour") + '<button type="button" class="cb" data-pop="grad" title="Gradient fill">🌈</button>';
      h += colorBtn("stroke", e.stroke || "#111111", "Border colour") +
        '<select data-k="sw" aria-label="Border">' + [0, 2, 4, 8, 12, 20, 32].map(function (n) { return '<option value="' + n + '"' + (Math.round(e.sw || 0) === n ? " selected" : "") + ">" + (n ? n + " px border" : "No border") + "</option>"; }).join("") + "</select>" +
        (e.shape === "rect" || e.shape === "round" ? '<button type="button" class="cb" data-pop="corners" title="Rounded corners">◜ Corners</button>' : "") +
        '<button type="button" class="cb" data-act="sh" aria-pressed="' + !!e.sh + '" title="Shadow">🌑 Shadow</button>';
    } else if (e.t === "img") {
      h += '<button type="button" class="cb" data-pop="crop" title="Zoom in and move the picture in its box">✂️ Crop</button>' +
        '<button type="button" class="cb" data-pop="filters" title="Brightness, colour and filters">🎛 Adjust</button>' +
        '<select data-k="mask" aria-label="Shape"><option value="">Box shape</option>' + ["rect", "circle", "round", "star", "heart", "hex", "blob", "tri"].map(function (k) { return '<option value="' + k + '"' + (e.mask === k ? " selected" : "") + ">" + { rect: "Square corners", circle: "Circle", round: "Rounded", star: "Star", heart: "Heart", hex: "Hexagon", blob: "Blob", tri: "Triangle" }[k] + "</option>"; }).join("") + "</select>" +
        colorBtn("stroke", e.stroke || "#ffffff", "Border colour") +
        '<select data-k="sw" aria-label="Border">' + [0, 4, 8, 16, 24].map(function (n) { return '<option value="' + n + '"' + (Math.round(e.sw || 0) === n ? " selected" : "") + ">" + (n ? n + " px border" : "No border") + "</option>"; }).join("") + "</select>" +
        '<button type="button" class="cb" data-act="sh" aria-pressed="' + !!e.sh + '" title="Shadow">🌑</button>' +
        '<button type="button" class="cb" data-act="asbg" title="Make it the page background">🖼 Background</button>';
    } else if (e.t === "emoji") {
      h += '<span class="hint">Sticker ' + esc(e.ch) + "</span>" + '<button type="button" class="cb" data-act="sh" aria-pressed="' + !!e.sh + '" title="Shadow">🌑 Shadow</button>';
    }
    h += '<span class="sep"></span><button type="button" class="cb" data-pop="look" title="See-through, position and more">◐ Transparency</button>' +
      '<button type="button" class="cb" data-pop="pos" title="Line it up on the page">⊞ Position</button>' +
      '<button type="button" class="cb" data-act="fwd" title="Bring forward">⬆︎</button><button type="button" class="cb" data-act="back" title="Send backward">⬇︎</button>' +
      '<button type="button" class="cb" data-act="flip" title="Flip">⇋</button>' +
      '<button type="button" class="cb" data-act="lock" aria-pressed="' + !!e.lock + '" title="' + (e.lock ? "Unlock" : "Lock so it can't be moved") + '">' + (e.lock ? "🔒" : "🔓") + "</button>" +
      '<button type="button" class="cb" data-act="dup" title="Copy it (Ctrl+D)">⧉</button><button type="button" class="cb" data-act="del" title="Delete">🗑</button>' +
      '<button type="button" class="cb" data-act="undo" title="Undo (Ctrl+Z)">↶</button>';
    c.innerHTML = h;
  }
  function setProp(k, v, nosnap) {
    var e = selEl(); if (!e) return;
    if (!nosnap) snap();
    e[k] = v;
    if (e.t === "text") e.h = S.textHeight(meas, e);
    draw(); placeEditor(); changed();
  }
  $("ctx").addEventListener("change", function (ev) {
    var t = ev.target, k = t.dataset.k; if (!k) return;
    if (k === "font") setProp("font", t.value);
    else if (k === "size") setProp("size", S.num(t.value, 32, 4, 800));
    else if (k === "sw") setProp("sw", +t.value);
    else if (k === "mask") setProp("mask", t.value || undefined);
    ov.focus({ preventScroll: true });
  });
  var colSnap = false;
  $("ctx").addEventListener("input", function (ev) {
    var t = ev.target; if (!t.dataset.col) return;
    var e = selEl(); if (!e) return;
    if (!colSnap) { snap(); colSnap = true; }
    e[t.dataset.col] = t.value; if (t.dataset.col === "stroke" && !e.sw) e.sw = e.t === "text" ? 2 : 4;
    t.previousElementSibling.style.background = t.value;
    draw(); placeEditor(); changed();
  });
  $("ctx").addEventListener("change", function (ev) { if (ev.target.dataset.col) { colSnap = false; ctxBar(); } });
  $("ctx").addEventListener("click", function (ev) {
    var b = ev.target.closest("button"); if (!b) return;
    var e = selEl(), a = b.dataset.act;
    if (b.dataset.tg) { setProp(b.dataset.tg, !e[b.dataset.tg]); ctxBar(); return; }
    if (b.dataset.pop) { pop(b, b.dataset.pop); return; }
    if (a === "bgpanel") { panel = "bg"; showPanel(); return; }
    if (a === "tplpanel") { panel = "tpl"; showPanel(); return; }
    if (a === "texth") { addText("h"); return; }
    if (a === "undo") { doUndo(); return; } if (a === "redo") { doRedo(); return; }
    if (!e) return;
    if (a === "sz-" || a === "sz+") { setProp("size", Math.max(4, Math.round(e.size * (a === "sz+" ? 1.12 : 0.89)))); ctxBar(); return; }
    if (a === "align") { setProp("align", { left: "center", center: "right", right: "left" }[e.align] || "center"); ctxBar(); return; }
    if (a === "sh") { setProp("sh", !e.sh); ctxBar(); return; }
    if (a === "flip") { setProp("fx", !e.fx); return; }
    if (a === "lock") { setProp("lock", !e.lock); selBox(); ctxBar(); return; }
    if (a === "fwd") { layer(e.id, 1); return; } if (a === "back") { layer(e.id, -1); return; }
    if (a === "dup") { dup(); return; } if (a === "del") { del(); return; }
    if (a === "asbg") { snap(); P().bg = { img: e.src, dim: 0 }; P().els = P().els.filter(function (x) { return x.id !== e.id; }); sel = null; draw(); ctxBar(); changed(); return; }
  });
  function pop(btn, kind) {
    closePop(); var e = selEl(); if (!e) return;
    var p = document.createElement("div"); p.className = "pop";
    function slider(label, key, min, max, step, val, fn) {
      return '<label>' + label + '<input type="range" data-s="' + key + '" min="' + min + '" max="' + max + '" step="' + step + '" value="' + val + '"><output>' + (Math.round(val * 100) / 100) + "</output></label>";
    }
    var fl = e.fl || {};
    if (kind === "spacing") p.innerHTML = slider("Letters", "ls", -5, 60, 1, e.ls || 0) + slider("Lines", "lh", 0.7, 3, 0.05, e.lh || 1.2);
    else if (kind === "effects") p.innerHTML = '<div class="row"><button type="button" data-fx="none">None</button><button type="button" data-fx="shadow">Shadow</button><button type="button" data-fx="outline">Outline</button><button type="button" data-fx="box">Highlight box</button><button type="button" data-fx="neon">Neon</button></div>' +
      slider("Outline", "sw", 0, 30, 0.5, e.sw || 0) + '<label>Outline colour<input type="color" data-c="stroke" value="' + (/^#[0-9a-f]{6}$/i.test(e.stroke || "") ? e.stroke : "#000000") + '"><span></span></label>' +
      '<label>Box colour<input type="color" data-c="bgc" value="' + (/^#[0-9a-f]{6}$/i.test(e.bgc || "") ? e.bgc : "#fde047") + '"><span></span></label>';
    else if (kind === "look") p.innerHTML = slider("See-through", "o", 0.05, 1, 0.05, e.o == null ? 1 : e.o) + slider("Turn", "r", -180, 180, 1, e.r || 0);
    else if (kind === "corners") p.innerHTML = slider("Rounded", "rad", 0, Math.round(Math.min(e.w, e.h) / 2), 1, e.rad || 0);
    else if (kind === "crop") p.innerHTML = slider("Zoom", "zoom", 1, 4, 0.05, e.zoom || 1) + slider("Left - right", "px", 0, 1, 0.01, e.px == null ? 0.5 : e.px) + slider("Up - down", "py", 0, 1, 0.01, e.py == null ? 0.5 : e.py) + slider("Rounded", "rad", 0, Math.round(Math.min(e.w, e.h) / 2), 1, e.rad || 0);
    else if (kind === "filters") p.innerHTML = '<div class="row"><button type="button" data-ff="none">Original</button><button type="button" data-ff="bw">Black & white</button><button type="button" data-ff="warm">Warm</button><button type="button" data-ff="old">Old photo</button><button type="button" data-ff="pop">Pop</button><button type="button" data-ff="soft">Dreamy</button></div>' +
      slider("Brightness", "fl.b", 20, 200, 1, fl.b == null ? 100 : fl.b) + slider("Contrast", "fl.c", 20, 200, 1, fl.c == null ? 100 : fl.c) + slider("Colour", "fl.s", 0, 250, 1, fl.s == null ? 100 : fl.s) +
      slider("Black & white", "fl.g", 0, 100, 1, fl.g || 0) + slider("Sepia", "fl.sp", 0, 100, 1, fl.sp || 0) + slider("Blur", "fl.bl", 0, 20, 0.5, fl.bl || 0);
    else if (kind === "grad") p.innerHTML = '<div class="gridp six" style="width:230px">' + GRADS.map(function (g, i) { return '<button type="button" class="sw" data-gf="' + i + '" style="background:' + S.cssFill("lin:" + g.join(",")) + '" aria-label="Gradient ' + (i + 1) + '"></button>'; }).join("") + "</div>";
    else if (kind === "pos") p.innerHTML = '<div class="row"><button type="button" data-al="l">⇤ Left</button><button type="button" data-al="c">↔ Centre</button><button type="button" data-al="r">Right ⇥</button></div><div class="row"><button type="button" data-al="t">⤒ Top</button><button type="button" data-al="m">↕ Middle</button><button type="button" data-al="b">Bottom ⤓</button></div><div class="row"><button type="button" data-al="front">Bring to front</button><button type="button" data-al="backmost">Send to back</button><button type="button" data-al="fill">Fill the page</button></div>';
    $("ctx").appendChild(p); openPop = p;
    var br = btn.getBoundingClientRect(), cr = $("ctx").getBoundingClientRect();
    p.style.left = Math.max(4, Math.min(br.left - cr.left, cr.width - p.offsetWidth - 8)) + "px";
    var sn = false;
    p.addEventListener("input", function (ev) {
      var t = ev.target, x = selEl(); if (!x) return;
      if (!sn) { snap(); sn = true; }
      if (t.dataset.s) {
        var k = t.dataset.s, v = +t.value;
        if (k.indexOf("fl.") === 0) { x.fl = x.fl || {}; x.fl[k.slice(3)] = v; } else x[k] = v;
        if (k === "sw" && v && !x.stroke) x.stroke = "#000000";
        t.nextElementSibling.textContent = Math.round(v * 100) / 100;
      } else if (t.dataset.c) { x[t.dataset.c] = t.value; if (t.dataset.c === "stroke" && !x.sw) x.sw = 2; }
      if (x.t === "text") x.h = S.textHeight(meas, x);
      draw(); placeEditor(); changed();
    });
    p.addEventListener("click", function (ev) {
      var b = ev.target.closest("button"), x = selEl(); if (!b || !x) return;
      var m = M();
      if (b.dataset.fx) {
        snap(); delete x.stroke; x.sw = 0; delete x.bgc; x.sh = false;
        if (b.dataset.fx === "shadow") x.sh = true;
        if (b.dataset.fx === "outline") { x.stroke = "#000000"; x.sw = Math.max(1, x.size / 18); }
        if (b.dataset.fx === "box") { x.bgc = "#fde047"; }
        if (b.dataset.fx === "neon") { x.color = "#f0abfc"; x.stroke = "#a21caf"; x.sw = Math.max(1, x.size / 24); x.sh = true; }
        x.h = S.textHeight(meas, x); draw(); changed(); pop(btn, kind); return;
      }
      if (b.dataset.ff) {
        snap(); x.fl = { none: {}, bw: { g: 100, c: 110 }, warm: { sp: 35, s: 130, b: 105 }, old: { sp: 70, c: 90, b: 95 }, pop: { s: 180, c: 125 }, soft: { b: 112, c: 85, bl: 1 } }[b.dataset.ff];
        draw(); changed(); pop(btn, kind); return;
      }
      if (b.dataset.gf != null) { snap(); x.fill = "lin:" + GRADS[+b.dataset.gf].join(","); draw(); changed(); ctxBar(); return; }
      if (b.dataset.al) {
        snap(); var al = b.dataset.al, els = P().els;
        if (al === "l") x.x = 0; if (al === "c") x.x = (doc.w - x.w) / 2; if (al === "r") x.x = doc.w - x.w;
        if (al === "t") x.y = 0; if (al === "m") x.y = (doc.h - x.h) / 2; if (al === "b") x.y = doc.h - x.h;
        if (al === "front") { P().els = els.filter(function (y) { return y !== x; }).concat([x]); }
        if (al === "backmost") { P().els = [x].concat(els.filter(function (y) { return y !== x; })); }
        if (al === "fill") { x.r = 0; x.x = 0; x.y = 0; x.w = doc.w; x.h = doc.h; if (x.t === "text") { x.h = S.textHeight(meas, x); x.y = (doc.h - x.h) / 2; } }
        draw(); changed();
      }
    });
  }
  function layer(id, d) {
    var els = P().els, i = els.map(function (e) { return e.id; }).indexOf(id), j = i + d;
    if (i < 0 || j < 0 || j >= els.length) return;
    snap(); var t = els[i]; els[i] = els[j]; els[j] = t; draw(); changed(); if (panel === "layers") showPanel();
  }
  var clip = null;
  function dup() {
    var e = selEl(); if (!e) return;
    snap(); var c = JSON.parse(JSON.stringify(e)); c.id = uid(); c.x += 24; c.y += 24; delete c.lock;
    P().els.push(c); select(c.id); draw(); changed();
  }
  function del() {
    var e = selEl(); if (!e) return;
    snap(); stopEdit(); P().els = P().els.filter(function (x) { return x.id !== e.id; }); sel = null; selBox(); draw(); ctxBar(); changed();
  }
  function pasteClip() { if (!clip) return; snap(); var c = JSON.parse(clip); c.id = uid(); c.x += 24; c.y += 24; P().els.push(c); select(c.id); draw(); changed(); clip = JSON.stringify(c); }

  /* ---------------- keys ---------------- */
  document.addEventListener("keydown", function (ev) {
    var t = document.activeElement, typing = t && (/INPUT|TEXTAREA|SELECT/.test(t.tagName) || t.isContentEditable);
    var mod = ev.ctrlKey || ev.metaKey, k = ev.key.toLowerCase();
    if (mod && k === "z" && !typing) { ev.preventDefault(); if (ev.shiftKey) doRedo(); else doUndo(); return; }
    if (mod && k === "y" && !typing) { ev.preventDefault(); doRedo(); return; }
    if (typing) return;
    var e = selEl();
    if (mod && k === "d") { ev.preventDefault(); dup(); return; }
    if (mod && k === "c" && e) { clip = JSON.stringify(e); return; }
    if (mod && k === "x" && e) { clip = JSON.stringify(e); del(); return; }
    if (!e) { if (k === "t") { ev.preventDefault(); addText("h"); } return; }
    if (k === "delete" || k === "backspace") { ev.preventDefault(); del(); return; }
    if (k === "escape") { select(null); return; }
    if (k === "enter" && e.t === "text") { ev.preventDefault(); startEdit(e); return; }
    var mv = { arrowleft: [-1, 0], arrowright: [1, 0], arrowup: [0, -1], arrowdown: [0, 1] }[k];
    if (mv && !e.lock) { ev.preventDefault(); snap(); var st = ev.shiftKey ? 10 : 1; e.x += mv[0] * st; e.y += mv[1] * st; draw(); changed(); }
  });

  /* ---------------- pages along the bottom ---------------- */
  var thumbT = null; function thumbSoon() { clearTimeout(thumbT); thumbT = setTimeout(thumbs, 500); }
  function pages() {
    var bar = $("pages"), h = "";
    doc.pages.forEach(function (p, i) { h += '<button type="button" class="pth" data-pg="' + i + '" aria-current="' + (i === cur) + '" aria-label="Page ' + (i + 1) + '"><canvas></canvas><span class="n">' + (i + 1) + "</span></button>"; });
    h += '<button type="button" class="cb" data-pa="add" title="Add a page">＋ Page</button><button type="button" class="cb" data-pa="dup" title="Copy this page">⧉</button>' +
      '<button type="button" class="cb" data-pa="left" title="Move this page earlier">◀</button><button type="button" class="cb" data-pa="right" title="Move this page later">▶</button>' +
      '<button type="button" class="cb" data-pa="del" title="Delete this page">🗑</button><span class="sp"></span>' +
      '<button type="button" class="cb" data-pa="zout" aria-label="Zoom out">−</button><button type="button" class="cb" data-pa="fit" id="zoomN" title="Fit to the screen">' + Math.round(zoom * 100) + '%</button><button type="button" class="cb" data-pa="zin" aria-label="Zoom in">+</button>';
    bar.innerHTML = h; thumbs();
  }
  function thumbs() {
    $("pages").querySelectorAll(".pth canvas").forEach(function (c, i) {
      var p = doc.pages[i]; if (!p) return;
      var hh = 60, ww = Math.max(20, Math.round(hh * doc.w / doc.h)); if (ww > 140) { ww = 140; hh = Math.round(ww * doc.h / doc.w); }
      c.width = ww * 2; c.height = hh * 2; c.style.width = ww + "px"; c.style.height = hh + "px";
      var g = c.getContext("2d"); g.scale(ww * 2 / doc.w, hh * 2 / doc.h); S.drawPage(g, p, doc.w, doc.h, { onload: thumbSoon, clean: true });
    });
  }
  $("pages").addEventListener("click", function (ev) {
    var b = ev.target.closest("button"); if (!b) return;
    if (b.dataset.pg != null) { stopEdit(); cur = +b.dataset.pg; sel = null; draw(); pages(); ctxBar(); if (panel === "layers" || panel === "bg") showPanel(); return; }
    var a = b.dataset.pa;
    if (a === "zin") setZoom(zoom * 1.25); else if (a === "zout") setZoom(zoom / 1.25); else if (a === "fit") { fit = true; fitZoom(); draw(); pages(); }
    else if (a === "add") { snap(); doc.pages.splice(cur + 1, 0, { id: uid(), bg: JSON.parse(JSON.stringify(P().bg.img ? { c: "#ffffff" } : P().bg)), els: [] }); cur++; sel = null; draw(); pages(); ctxBar(); changed(); }
    else if (a === "dup") { snap(); var c = JSON.parse(JSON.stringify(P())); c.id = uid(); c.els.forEach(function (e) { e.id = uid(); }); doc.pages.splice(cur + 1, 0, c); cur++; sel = null; draw(); pages(); changed(); }
    else if (a === "left" || a === "right") { var j = cur + (a === "left" ? -1 : 1); if (j < 0 || j >= doc.pages.length) return; snap(); var t = doc.pages[cur]; doc.pages[cur] = doc.pages[j]; doc.pages[j] = t; cur = j; pages(); changed(); }
    else if (a === "del") {
      if (doc.pages.length < 2) { snap(); P().els = []; P().bg = { c: "#ffffff" }; sel = null; draw(); pages(); changed(); return; }
      snap(); doc.pages.splice(cur, 1); cur = Math.max(0, cur - 1); sel = null; draw(); pages(); ctxBar(); changed(); Office.toast("Page deleted - Undo brings it back.");
    }
  });
  stage.addEventListener("wheel", function (ev) { if (!(ev.ctrlKey || ev.metaKey)) return; ev.preventDefault(); setZoom(zoom * (ev.deltaY < 0 ? 1.1 : 0.9)); }, { passive: false });
  window.addEventListener("resize", function () { fitSoon(); });

  /* ---------------- choosing a size, resizing ---------------- */
  function sizeDialog(resizing) {
    var d = document.createElement("dialog"); d.className = "of-dialog ofwide";
    d.innerHTML = '<div class="of-dhead"><h2>' + (resizing ? "📐 Resize your design" : "✨ What are you making?") + '</h2><button type="button" class="of-x" data-x aria-label="Close">✕</button></div>' +
      (resizing ? '<p class="of-where">Everything on every page is made bigger or smaller to fit the new size.</p>' : "") +
      '<div class="sizes">' + SIZES.map(function (s, i) {
        var k = 56 / Math.max(s[1], s[2]);
        return '<button type="button" data-sz="' + i + '"><span class="shp" style="width:' + Math.max(8, s[1] * k) + "px;height:" + Math.max(8, s[2] * k) + 'px"></span><b>' + s[0] + "</b><small>" + s[1] + " × " + s[2] + " · " + s[3] + "</small></button>";
      }).join("") + '</div><form class="custom"><b>Custom size</b><input type="number" name="w" min="16" max="8000" value="' + (doc ? doc.w : 1200) + '" aria-label="Width"> × <input type="number" name="h" min="16" max="8000" value="' + (doc ? doc.h : 800) + '" aria-label="Height"> px <button class="of-btn" type="submit">' + (resizing ? "Resize" : "Make it") + "</button></form>";
    document.body.appendChild(d); d.showModal();
    function pick(w, h) {
      d.close(); d.remove();
      if (resizing) resizeTo(w, h);
      else { fresh(w, h); }
    }
    d.addEventListener("click", function (ev) {
      if (ev.target === d || ev.target.closest("[data-x]")) { d.close(); d.remove(); if (!doc) fresh(1080, 1080); return; }
      var b = ev.target.closest("[data-sz]"); if (b) { var s = SIZES[+b.dataset.sz]; pick(s[1], s[2]); }
    });
    d.querySelector("form").onsubmit = function (ev) { ev.preventDefault(); pick(S.num(this.w.value, 1080, 16, 8000), S.num(this.h.value, 1080, 16, 8000)); };
    d.addEventListener("cancel", function () { setTimeout(function () { d.remove(); if (!doc) fresh(1080, 1080); }, 0); });
  }
  function resizeTo(w, h) {
    snap(); var k = Math.min(w / doc.w, h / doc.h), ox = (w - doc.w * k) / 2, oy = (h - doc.h * k) / 2;
    doc.pages.forEach(function (p) {
      p.els.forEach(function (e) {
        e.x = e.x * k + ox; e.y = e.y * k + oy; e.w *= k; e.h *= k;
        if (e.size) e.size = Math.max(4, Math.round(e.size * k)); if (e.sw) e.sw *= k; if (e.ls) e.ls *= k; if (e.rad) e.rad *= k;
        // full-width bands stay full width
        if (e.t === "shape" && e.w >= doc.w * k * 0.98) { e.x = 0; e.w = w; }
        if (e.t === "text") e.h = S.textHeight(meas, e);
      });
    });
    doc.w = w; doc.h = h; fit = true; fitZoom(); draw(); pages(); ctxBar(); changed(); if (panel === "tpl") showPanel();
    Office.toast("Resized to " + w + " × " + h + ".");
  }
  $("resizeBtn").onclick = function () { sizeDialog(true); };

  /* ---------------- downloading ---------------- */
  function render(p, scale) {
    var c = document.createElement("canvas"); c.width = Math.round(doc.w * scale); c.height = Math.round(doc.h * scale);
    var g = c.getContext("2d"); g.scale(scale, scale); S.drawPage(g, p, doc.w, doc.h, { clean: true }); return c;
  }
  function allSrcs() { var s = []; doc.pages.forEach(function (p) { if (p.bg.img) s.push(p.bg.img); p.els.forEach(function (e) { if (e.src) s.push(e.src); }); }); return s; }
  function scaleFor(kind) { var big = Math.max(doc.w, doc.h); return Math.min(kind === "pdf" ? 2 : 2, 4096 / big, Math.max(1, 2000 / big)); }
  function blobOf(c, type, q) { return new Promise(function (ok) { c.toBlob(ok, type, q); }); }
  $("exportBtn").onclick = function (e) { e.stopPropagation(); $("exportMenu").hidden = !$("exportMenu").hidden; };
  document.addEventListener("click", function (e) { if (!e.target.closest("#exportMenu,#exportBtn")) $("exportMenu").hidden = true; });
  $("exportMenu").addEventListener("click", function (ev) {
    var b = ev.target.closest("[data-ex]"); if (!b) return;
    $("exportMenu").hidden = true; stopEdit();
    var kind = b.dataset.ex, name = $("title").value.trim() || "Design";
    Office.toast("Getting it ready…");
    var allEls = []; doc.pages.forEach(function (p) { allEls = allEls.concat(p.els); });
    Promise.all([S.ready(allSrcs()), S.fontsReady(allEls)]).then(function () {
      var sc = scaleFor(kind);
      if (kind === "png" || kind === "jpg") {
        return blobOf(render(P(), sc), kind === "png" ? "image/png" : "image/jpeg", 0.92).then(function (bl) {
          Office.download(bl, Office.safeName(name + (doc.pages.length > 1 ? " page " + (cur + 1) : ""), "." + kind));
        });
      }
      if (kind === "zip") {
        var z = new JSZip();
        return doc.pages.reduce(function (pr, p, i) {
          return pr.then(function () { return blobOf(render(p, sc), "image/png"); }).then(function (bl) { z.file("page-" + (i + 1) + ".png", bl); });
        }, Promise.resolve()).then(function () { return z.generateAsync({ type: "blob" }); }).then(function (bl) { Office.download(bl, Office.safeName(name, ".zip")); });
      }
      if (kind === "pdf") {
        var pdf;
        return PDFLib.PDFDocument.create().then(function (d) {
          pdf = d;
          return doc.pages.reduce(function (pr, p) {
            return pr.then(function () { return blobOf(render(p, sc), "image/jpeg", 0.93); }).then(function (bl) { return bl.arrayBuffer(); }).then(function (buf) { return pdf.embedJpg(buf); }).then(function (im) {
              var pw = doc.w * 0.75, ph = doc.h * 0.75, pg = pdf.addPage([pw, ph]); pg.drawImage(im, { x: 0, y: 0, width: pw, height: ph });
            });
          }, Promise.resolve());
        }).then(function () { pdf.setTitle(name); return pdf.save({ useObjectStreams: false }); }).then(function (bytes) {
          Office.download(new Blob([bytes], { type: "application/pdf" }), Office.safeName(name, ".pdf"));
        });
      }
      if (kind === "print") {
        var w = window.open("", "_blank"); if (!w) { Office.toast("Allow pop-ups to print.", true); return; }
        var imgsH = doc.pages.map(function (p) { return '<img src="' + render(p, Math.min(2, sc)).toDataURL("image/jpeg", 0.92) + '">'; }).join("");
        w.document.write("<!doctype html><title>" + esc(name) + "</title><style>@page{margin:0;size:" + doc.w + "px " + doc.h + "px}body{margin:0}img{display:block;width:100%;page-break-after:always}</style>" + imgsH + "<script>onload=function(){print()}<\/script>");
        w.document.close();
      }
    }).catch(function (e) { Office.toast("Couldn't make that: " + (e.message || e), true); });
  });

  /* ---------------- files ---------------- */
  function load(f) {
    stopEdit(); file = { id: f.id, title: f.title || "Untitled design" }; $("title").value = file.title;
    doc = clean(f.body); cur = 0; sel = null; undo = []; redo = []; fit = true;
    history.replaceState(null, "", "?id=" + encodeURIComponent(f.id));
    try { localStorage.setItem("hl.design.last", f.id); } catch (e) {}
    fitZoom(); draw(); pages(); ctxBar(); showPanel();
    $("status").textContent = Office.isCloud() ? "Saved to your account" : "Saved on this device";
    saver.opened();
  }
  function fresh(w, h) {
    stopEdit(); file = { id: Office.uuid(), title: "Untitled design" }; $("title").value = file.title;
    doc = blank(w || 1080, h || 1080); cur = 0; sel = null; undo = []; redo = []; fit = true; panel = window.innerWidth < 860 ? null : "tpl";
    history.replaceState(null, "", "?id=" + encodeURIComponent(file.id));
    try { localStorage.setItem("hl.design.last", file.id); } catch (e) {}
    fitZoom(); draw(); pages(); ctxBar(); showPanel();
    saver.now(); saver.opened();
  }
  function openId(id) {
    $("status").textContent = "Opening…";
    return Office.get(id).then(function (f) { if (f) load(f); else sizeDialog(false); }).catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); sizeDialog(false); });
  }
  doc = blank(1080, 1080); fitZoom(); draw(); pages(); ctxBar();
  Office.init("design").then(function (st) {
    var q = new URLSearchParams(location.search), id = q.get("id"), last = null;
    try { last = localStorage.getItem("hl.design.last"); } catch (e) {}
    if (!st.user) Office.toast("Not signed in - your designs are saved in this browser.");
    if (q.get("new") != null) { doc = null; sizeDialog(false); return; }
    if (id || last) return openId(id || last);
    return Office.list().then(function (rows) { if (rows.length) return openId(rows[0].id); doc = null; sizeDialog(false); });
  });
  $("filesBtn").onclick = function () {
    saver.now();
    Office.filesDialog({
      heading: "Your designs", noun: "design", icon: "🖌", current: file.id, accept: "image/*,.png,.jpg,.jpeg,.webp,.gif,.svg",
      onOpen: openId, onNew: function () { sizeDialog(false); }, onDeletedCurrent: function () { sizeDialog(false); },
      onImport: function (f) { addPictures([f]); }
    });
  };
  Office.setImporter(function (f) { addPictures([f]); });
  Office.themeToggle($("themeBtn"));
  $("title").addEventListener("input", saver.poke);
  window.addEventListener("beforeunload", function () { saver.now(); });
  showPanel();
})();
