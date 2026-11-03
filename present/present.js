/* Launchpad Present - slides. A slide is 960 x 540 units (1 unit = 1 pt
   in PowerPoint), scaled to fit the screen. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var SW = 960, SH = 540;

  var THEMES = {
    classic:  { name: "Classic",  bg: ["#ffffff"], title: "#1f2937", text: "#374151", accent: "#2563eb", tf: "Calibri, Carlito, sans-serif", bf: "Calibri, Carlito, sans-serif" },
    midnight: { name: "Midnight", bg: ["#0f172a", "#1e3a5f"], title: "#ffffff", text: "#cbd5e1", accent: "#38bdf8", tf: "Calibri, Carlito, sans-serif", bf: "Calibri, Carlito, sans-serif" },
    aurora:   { name: "Aurora",   bg: ["#4338ca", "#db2777"], title: "#ffffff", text: "#fdf2f8", accent: "#fde047", tf: "'Trebuchet MS', sans-serif", bf: "'Trebuchet MS', sans-serif" },
    glass:    { name: "Glass",    bg: ["#dbeafe", "#fae8ff"], title: "#1e1b4b", text: "#334155", accent: "#7c3aed", tf: "'Segoe UI', system-ui, sans-serif", bf: "'Segoe UI', system-ui, sans-serif" },
    forest:   { name: "Forest",   bg: ["#ecfdf5"], title: "#065f46", text: "#1f2937", accent: "#16a34a", tf: "Georgia, serif", bf: "Calibri, Carlito, sans-serif" },
    sunset:   { name: "Sunset",   bg: ["#f97316", "#be123c"], title: "#ffffff", text: "#fff7ed", accent: "#fde68a", tf: "Arial, sans-serif", bf: "Arial, sans-serif" },
    paper:    { name: "Paper",    bg: ["#fdf6e3"], title: "#7c2d12", text: "#44403c", accent: "#b45309", tf: "Georgia, serif", bf: "Georgia, serif" },
    space:    { name: "Space",    bg: ["#000000", "#312e81"], title: "#e0e7ff", text: "#c7d2fe", accent: "#f472b6", tf: "Verdana, sans-serif", bf: "Verdana, sans-serif" }
  };
  function bgCss(t) { return t.bg.length > 1 ? "linear-gradient(135deg," + t.bg[0] + "," + t.bg[1] + ")" : t.bg[0]; }

  var file = { id: null, title: "Untitled presentation" };
  var deck = null, cur = 0, selId = null, editingId = null;
  var undo = [], redo = [];
  var saver = Office.autosaver($("status"), function () {
    file.title = $("title").value.trim() || "Untitled presentation";
    return { id: file.id, title: file.title, body: deck };
  });
  function uid() { return Math.random().toString(36).slice(2, 10); }
  function T() { return THEMES[deck.theme] || THEMES.classic; }
  function S() { return deck.slides[cur]; }

  /* ---------- layouts ---------- */
  function tx(x, y, w, h, size, role, html, extra) {
    return Object.assign({ id: uid(), type: "text", x: x, y: y, w: w, h: h, size: size, role: role, html: html || "", align: "left" }, extra || {});
  }
  function layout(kind) {
    var els;
    if (kind === "title") els = [tx(80, 160, 800, 120, 58, "title", "", { align: "center", ph: "Click to add a title" }), tx(120, 295, 720, 70, 26, "body", "", { align: "center", ph: "Click to add a subtitle" })];
    else if (kind === "content") els = [tx(60, 36, 840, 84, 40, "title", "", { ph: "Click to add a title" }), tx(60, 135, 840, 370, 26, "body", "", { ph: "• Click to add points" })];
    else if (kind === "two") els = [tx(60, 36, 840, 84, 40, "title", "", { ph: "Click to add a title" }), tx(60, 135, 410, 370, 24, "body", "", { ph: "• Left side" }), tx(490, 135, 410, 370, 24, "body", "", { ph: "• Right side" })];
    else if (kind === "picture") els = [tx(60, 36, 840, 84, 40, "title", "", { ph: "Click to add a title" }), tx(60, 135, 380, 370, 24, "body", "", { ph: "• Say something about the picture" })];
    else if (kind === "section") els = [tx(80, 210, 800, 100, 52, "title", "", { ph: "Section title" }), { id: uid(), type: "shape", shape: "rect", x: 80, y: 318, w: 140, h: 8, accent: true }];
    else els = [];
    return { id: uid(), layout: kind, trans: "fade", notes: "", els: els };
  }

  /* ---------- drawing a slide ---------- */
  var SHAPES = {
    rect: '<rect x="0" y="0" width="100" height="100"/>',
    round: '<rect x="0" y="0" width="100" height="100" rx="14" ry="14"/>',
    ellipse: '<ellipse cx="50" cy="50" rx="50" ry="50"/>',
    triangle: '<polygon points="50,0 100,100 0,100"/>',
    star: '<polygon points="50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35"/>',
    arrow: '<polygon points="0,30 62,30 62,0 100,50 62,100 62,70 0,70"/>',
    line: '<line x1="0" y1="50" x2="100" y2="50" stroke-width="4" vector-effect="non-scaling-stroke"/>'
  };
  function elColor(e, t) { return e.color || (e.type === "shape" ? (e.accent || !e.color ? t.accent : e.color) : e.role === "title" ? t.title : t.text); }
  function renderSlide(s, live) {
    var t = T(), d = document.createElement("div");
    d.className = "slide"; d.style.background = s.bg || bgCss(t);
    s.els.forEach(function (e) {
      var el = document.createElement("div");
      el.className = "el " + e.type; el.dataset.id = e.id;
      el.style.left = e.x + "px"; el.style.top = e.y + "px"; el.style.width = e.w + "px"; el.style.height = e.h + "px";
      if (e.type === "text") {
        el.innerHTML = e.html || "";
        if (live && e.ph) el.dataset.ph = e.ph;
        el.style.fontSize = e.size + "px"; el.style.textAlign = e.align || "left";
        el.style.color = elColor(e, t); el.style.fontFamily = e.role === "title" ? t.tf : t.bf;
        if (e.role === "title") el.style.fontWeight = "700";
        if (!live && !e.html) el.style.display = "none";
      } else if (e.type === "img") {
        var im = document.createElement("img"); im.src = e.src; im.alt = ""; el.appendChild(im);
      } else if (e.type === "shape") {
        var c = e.color || t.accent;
        el.innerHTML = '<svg viewBox="0 0 100 100" preserveAspectRatio="none" fill="' + Office.esc(c) + '" stroke="' + Office.esc(c) + '">' + (SHAPES[e.shape] || SHAPES.rect) + "</svg>";
      }
      d.appendChild(el);
    });
    return d;
  }

  var stage = $("stage"), slideEl = null, scale = 1;
  function fit() {
    var w = $("stagewrap").clientWidth - 40, h = $("stagewrap").clientHeight - 40;
    scale = Math.max(0.1, Math.min(w / SW, h / SH));
    stage.style.width = SW * scale + "px"; stage.style.height = SH * scale + "px";
    if (slideEl) { slideEl.style.transform = "scale(" + scale + ")"; slideEl.style.transformOrigin = "0 0"; }
    drawSel();
  }
  function draw() {
    stage.innerHTML = "";
    slideEl = renderSlide(S(), true); slideEl.classList.add("edit");
    stage.appendChild(slideEl);
    $("notes").value = S().notes || ""; $("trans").value = S().trans || "fade";
    fit(); thumbs();
  }
  function thumbs() {
    var box = $("thumbs"); box.innerHTML = "";
    deck.slides.forEach(function (s, i) {
      var b = document.createElement("button"); b.type = "button"; b.className = "thumb"; b.dataset.i = i;
      b.setAttribute("aria-label", "Slide " + (i + 1)); if (i === cur) b.setAttribute("aria-current", "true");
      b.innerHTML = '<span class="no">' + (i + 1) + '</span><span class="mini"></span>';
      b.querySelector(".mini").appendChild(renderSlide(s, false));
      box.appendChild(b);
    });
    var c = box.children[cur]; if (c && c.scrollIntoView) c.scrollIntoView({ block: "nearest" });
  }
  var thumbT = null;
  function thumbSoon() { clearTimeout(thumbT); thumbT = setTimeout(function () {
    var b = $("thumbs").children[cur]; if (!b) return; var m = b.querySelector(".mini"); m.innerHTML = ""; m.appendChild(renderSlide(S(), false)); }, 250); }
  $("thumbs").addEventListener("click", function (e) { var b = e.target.closest(".thumb"); if (b) goSlide(+b.dataset.i); });
  function goSlide(i) { finishEdit(); cur = Math.max(0, Math.min(deck.slides.length - 1, i)); selId = null; draw(); }

  /* ---------- undo ---------- */
  function mark() { undo.push(JSON.stringify({ deck: deck, cur: cur })); if (undo.length > 80) undo.shift(); redo = []; }
  function changed() { saver.poke(); thumbSoon(); }
  function restore(from, to) {
    finishEdit(); var s = from.pop(); if (!s) return;
    to.push(JSON.stringify({ deck: deck, cur: cur }));
    var o = JSON.parse(s); deck = o.deck; cur = Math.min(o.cur, deck.slides.length - 1); selId = null; draw(); saver.poke();
  }
  $("undoBtn").onclick = function () { restore(undo, redo); };
  $("redoBtn").onclick = function () { restore(redo, undo); };

  /* ---------- selecting, moving, resizing ---------- */
  function find(id) { return S().els.filter(function (e) { return e.id === id; })[0]; }
  function domOf(id) { return slideEl && slideEl.querySelector('.el[data-id="' + id + '"]'); }
  var selBox = document.createElement("div"); selBox.className = "selbox";
  ["nw", "n", "ne", "e", "se", "s", "sw", "w"].forEach(function (h) { var i = document.createElement("i"); i.dataset.h = h; selBox.appendChild(i); });
  var HPOS = { nw: [0, 0], n: [.5, 0], ne: [1, 0], e: [1, .5], se: [1, 1], s: [.5, 1], sw: [0, 1], w: [0, .5] };
  function drawSel() {
    var e = selId && find(selId);
    if (!e) { selBox.remove(); return; }
    stage.appendChild(selBox);
    selBox.style.left = e.x * scale + "px"; selBox.style.top = e.y * scale + "px"; selBox.style.width = e.w * scale + "px"; selBox.style.height = e.h * scale + "px";
    Array.prototype.forEach.call(selBox.children, function (i) { var p = HPOS[i.dataset.h]; i.style.left = p[0] * 100 + "%"; i.style.top = p[1] * 100 + "%"; });
    var tt = e.type === "text";
    $("textTools").style.opacity = tt || e.type === "shape" ? 1 : .45;
  }
  function select(id) { if (editingId && editingId !== id) finishEdit(); selId = id; drawSel(); }

  function pt(ev) { var r = stage.getBoundingClientRect(); return { x: (ev.clientX - r.left) / scale, y: (ev.clientY - r.top) / scale }; }
  var drag = null, guides = [];
  function clearGuides() { guides.forEach(function (g) { g.remove(); }); guides = []; }
  function guide(vertical, at) {
    var g = document.createElement("div"); g.className = "guide";
    if (vertical) { g.style.left = at * scale + "px"; g.style.top = 0; g.style.width = "1px"; g.style.height = "100%"; }
    else { g.style.top = at * scale + "px"; g.style.left = 0; g.style.height = "1px"; g.style.width = "100%"; }
    stage.appendChild(g); guides.push(g);
  }
  stage.addEventListener("pointerdown", function (ev) {
    var h = ev.target.closest(".selbox i");
    if (h) { ev.preventDefault(); var e0 = find(selId); mark(); drag = { mode: "resize", h: h.dataset.h, start: pt(ev), o: Object.assign({}, e0) }; stage.setPointerCapture(ev.pointerId); return; }
    var el = ev.target.closest(".el");
    if (!el) { finishEdit(); select(null); return; }
    var id = el.dataset.id;
    if (editingId === id) return;                       // clicking inside the text being typed
    ev.preventDefault();
    var e = find(id);
    if (selId === id && e.type === "text" && ev.pointerType !== "mouse") { startEdit(id); return; } // second tap types
    select(id);
    mark(); drag = { mode: "move", start: pt(ev), o: Object.assign({}, e), moved: false };
    stage.setPointerCapture(ev.pointerId);
  });
  stage.addEventListener("pointermove", function (ev) {
    if (!drag) return;
    var p = pt(ev), dx = p.x - drag.start.x, dy = p.y - drag.start.y, e = find(selId); if (!e) return;
    if (Math.abs(dx) + Math.abs(dy) > 2) drag.moved = true;
    clearGuides();
    if (drag.mode === "move") {
      var nx = drag.o.x + dx, ny = drag.o.y + dy, cx = nx + e.w / 2, cy = ny + e.h / 2;
      if (Math.abs(cx - SW / 2) < 8) { nx = SW / 2 - e.w / 2; guide(true, SW / 2); }
      if (Math.abs(cy - SH / 2) < 8) { ny = SH / 2 - e.h / 2; guide(false, SH / 2); }
      if (Math.abs(nx) < 8) { nx = 0; guide(true, 0); }
      if (Math.abs(ny) < 8) { ny = 0; guide(false, 0); }
      if (Math.abs(nx + e.w - SW) < 8) { nx = SW - e.w; guide(true, SW); }
      if (Math.abs(ny + e.h - SH) < 8) { ny = SH - e.h; guide(false, SH); }
      e.x = Math.round(nx); e.y = Math.round(ny);
    } else {
      var o = drag.o, H = drag.h, x = o.x, y = o.y, w = o.w, hh = o.h;
      if (/e/.test(H)) w = o.w + dx; if (/s/.test(H)) hh = o.h + dy;
      if (/w/.test(H)) { w = o.w - dx; x = o.x + dx; } if (/n/.test(H)) { hh = o.h - dy; y = o.y + dy; }
      if ((e.type === "img") && H.length === 2 && ev.shiftKey === false) { // pictures keep their shape from the corners
        var ratio = o.w / o.h; if (w / hh > ratio) w = hh * ratio; else hh = w / ratio;
        if (/w/.test(H)) x = o.x + o.w - w; if (/n/.test(H)) y = o.y + o.h - hh;
      }
      var min = e.shape === "line" ? 4 : 20;
      if (w < 20) { if (/w/.test(H)) x -= 20 - w; w = 20; } if (hh < min) { if (/n/.test(H)) y -= min - hh; hh = min; }
      e.x = Math.round(x); e.y = Math.round(y); e.w = Math.round(w); e.h = Math.round(hh);
    }
    var d = domOf(selId); d.style.left = e.x + "px"; d.style.top = e.y + "px"; d.style.width = e.w + "px"; d.style.height = e.h + "px";
    drawSel();
  });
  stage.addEventListener("pointerup", function () {
    clearGuides();
    if (drag) { if (!drag.moved) undo.pop(); else changed(); }
    drag = null;
  });
  // pointer capture sends the double-click to the stage, so find the box under the pointer
  stage.addEventListener("dblclick", function (ev) {
    var hit = document.elementsFromPoint(ev.clientX, ev.clientY).filter(function (n) { return n.classList && n.classList.contains("el") && n.classList.contains("text"); })[0];
    if (hit) startEdit(hit.dataset.id);
  });

  function startEdit(id) {
    var e = find(id), d = domOf(id); if (!e || e.type !== "text") return;
    finishEdit(); select(id); editingId = id; mark();
    d.contentEditable = "true"; d.focus();
    var r = document.createRange(); r.selectNodeContents(d); r.collapse(false); var s = getSelection(); s.removeAllRanges(); s.addRange(r);
    if (!e.html && e.role === "body" && /^•/.test(e.ph || "")) { document.execCommand("insertUnorderedList"); }
    d.oninput = function () { e.html = Office.cleanHTML(d.innerHTML); changed(); };
    d.onpaste = function (ev) { ev.preventDefault(); document.execCommand("insertText", false, (ev.clipboardData || window.clipboardData).getData("text/plain")); };
  }
  function finishEdit() {
    if (!editingId) return;
    var e = find(editingId), d = domOf(editingId);
    if (d) { d.contentEditable = "false"; d.oninput = null; if (e) { var h = Office.cleanHTML(d.innerHTML); if (/^(<ul><li><br><\/li><\/ul>|<br>|<p><br><\/p>|<div><br><\/div>)$/.test(h)) h = ""; e.html = h; if (!h) d.innerHTML = ""; } }
    editingId = null; changed();
  }

  /* ---------- the ribbon ---------- */
  function menu(btnId, menuId) {
    var b = $(btnId), m = $(menuId);
    b.onclick = function (ev) { ev.stopPropagation(); document.querySelectorAll(".menu").forEach(function (x) { if (x !== m) x.hidden = true; }); m.hidden = !m.hidden; };
    document.addEventListener("click", function (ev) { if (!m.hidden && !m.contains(ev.target) && ev.target !== b) m.hidden = true; });
    return m;
  }
  menu("newBtn", "newMenu").onclick = function (ev) {
    var b = ev.target.closest("[data-lay]"); if (!b) return; this.hidden = true;
    finishEdit(); mark(); deck.slides.splice(cur + 1, 0, layout(b.dataset.lay)); cur++; selId = null; draw(); changed();
  };
  menu("shapeBtn", "shapeMenu").onclick = function (ev) {
    var b = ev.target.closest("[data-shape]"); if (!b) return; this.hidden = true;
    var sh = b.dataset.shape, e = { id: uid(), type: "shape", shape: sh, x: 380, y: 190, w: 200, h: sh === "line" ? 6 : sh === "arrow" ? 110 : 160 };
    if (sh === "line") e.y = 268;
    add(e);
  };
  function add(e) { finishEdit(); mark(); S().els.push(e); selId = e.id; draw(); changed(); }
  $("addText").onclick = function () { var e = tx(330, 220, 300, 80, 28, "body", "", { ph: "Type here" }); add(e); startEdit(e.id); };
  $("imgIn").onchange = function () {
    var f = this.files[0]; this.value = ""; if (!f) return;
    Office.shrinkImage(f, 1600).then(function (src) {
      var im = new Image(); im.onload = function () {
        var w = im.width, h = im.height, k = Math.min(520 / w, 380 / h, 1) || 1; w *= k; h *= k;
        var s = S(), x = (SW - w) / 2, y = (SH - h) / 2 + 30;
        if (s.layout === "picture" && !s.els.some(function (e) { return e.type === "img"; })) { k = Math.min(420 / im.width, 370 / im.height); w = im.width * k; h = im.height * k; x = 480 + (420 - w) / 2; y = 135 + (370 - h) / 2; }
        add({ id: uid(), type: "img", src: src, x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h) });
      }; im.src = src;
    }).catch(function (e) { Office.toast(e.message, true); });
  };
  document.querySelectorAll("[data-cmd]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.onclick = function () {
      var e = selId && find(selId); if (!e || e.type !== "text") { Office.toast("Pick a text box first."); return; }
      if (!editingId) { // no words picked: change the whole box
        startEdit(e.id); var d = domOf(e.id), r = document.createRange(); r.selectNodeContents(d); var s = getSelection(); s.removeAllRanges(); s.addRange(r);
      }
      document.execCommand(b.dataset.cmd); var d2 = domOf(editingId); if (d2) { find(editingId).html = Office.cleanHTML(d2.innerHTML); changed(); }
    };
  });
  function resize(k) {
    var e = selId && find(selId); if (!e || e.type !== "text") return;
    mark(); e.size = Math.max(10, Math.min(140, Math.round(e.size * k))); var d = domOf(e.id); d.style.fontSize = e.size + "px"; changed();
  }
  $("szUp").onclick = function () { resize(1.15); }; $("szDown").onclick = function () { resize(1 / 1.15); };
  document.querySelectorAll("[data-alg]").forEach(function (b) {
    b.onclick = function () { var e = selId && find(selId); if (!e || e.type !== "text") return; mark(); e.align = b.dataset.alg; domOf(e.id).style.textAlign = e.align; changed(); };
  });
  $("colIn").addEventListener("input", function () { $("colBar").style.background = this.value; });
  $("colIn").addEventListener("change", function () {
    var e = selId && find(selId), v = this.value; if (!e) return;
    if (editingId && !getSelection().isCollapsed) { document.execCommand("foreColor", false, v); find(editingId).html = Office.cleanHTML(domOf(editingId).innerHTML); changed(); return; }
    mark(); e.color = v; delete e.accent; draw(); select(e.id); changed();
  });
  function layer(front) { var s = S(), e = selId && find(selId); if (!e) return; mark(); s.els.splice(s.els.indexOf(e), 1); if (front) s.els.push(e); else s.els.unshift(e); draw(); changed(); }
  $("fwdBtn").onclick = function () { layer(true); }; $("backBtn").onclick = function () { layer(false); };
  function dup() { var e = selId && find(selId); if (!e) return; var c = JSON.parse(JSON.stringify(e)); c.id = uid(); c.x += 20; c.y += 20; add(c); }
  function del() { var s = S(), e = selId && find(selId); if (!e) return; finishEdit(); mark(); s.els.splice(s.els.indexOf(e), 1); selId = null; draw(); changed(); }
  $("dupBtn").onclick = dup; $("delBtn").onclick = del;
  $("trans").onchange = function () { mark(); S().trans = this.value; changed(); };
  $("notes").addEventListener("input", function () { S().notes = this.value.slice(0, 5000); saver.poke(); });

  $("sDup").onclick = function () { finishEdit(); mark(); var c = JSON.parse(JSON.stringify(S())); c.id = uid(); c.els.forEach(function (e) { e.id = uid(); }); deck.slides.splice(cur + 1, 0, c); goSlide(cur + 1); changed(); };
  $("sDel").onclick = function () {
    var b = this; if (deck.slides.length < 2) { Office.toast("A presentation needs at least one slide."); return; }
    if (!b.dataset.armed) { b.dataset.armed = 1; b.textContent = "Sure?"; setTimeout(function () { delete b.dataset.armed; b.textContent = "🗑"; }, 2500); return; }
    delete b.dataset.armed; b.textContent = "🗑"; finishEdit(); mark(); deck.slides.splice(cur, 1); goSlide(Math.min(cur, deck.slides.length - 1)); changed();
  };
  function moveSlide(d) { var j = cur + d; if (j < 0 || j >= deck.slides.length) return; finishEdit(); mark(); var s = deck.slides.splice(cur, 1)[0]; deck.slides.splice(j, 0, s); goSlide(j); changed(); }
  $("sUp").onclick = function () { moveSlide(-1); }; $("sDown").onclick = function () { moveSlide(1); };

  $("designBtn").onclick = function () {
    var d = document.createElement("dialog"); d.className = "of-dialog"; d.style.width = "min(720px, calc(100vw - 32px))";
    d.innerHTML = '<div class="of-dhead"><h2>Design</h2><button type="button" class="of-x" aria-label="Close">✕</button></div><p class="of-where">Changes every slide. Colours you picked yourself stay.</p><div class="thpick"></div>';
    var box = d.querySelector(".thpick");
    Object.keys(THEMES).forEach(function (k) {
      var t = THEMES[k], b = document.createElement("button"); b.type = "button"; b.dataset.t = k; b.setAttribute("aria-pressed", String(deck.theme === k));
      b.style.background = bgCss(t);
      b.innerHTML = '<span style="color:' + t.title + ";font-family:" + Office.esc(t.tf) + '">' + Office.esc(t.name) + '<i style="position:absolute;left:12%;right:62%;bottom:16%;height:6px;border-radius:3px;background:' + t.accent + '"></i></span>';
      box.appendChild(b);
    });
    document.body.appendChild(d); d.showModal();
    d.onclick = function (ev) {
      if (ev.target.closest(".of-x") || ev.target === d) { d.close(); d.remove(); return; }
      var b = ev.target.closest("[data-t]"); if (!b) return;
      mark(); deck.theme = b.dataset.t; draw(); changed();
      box.querySelectorAll("button").forEach(function (x) { x.setAttribute("aria-pressed", String(x === b)); });
    };
  };

  /* ---------- keys ---------- */
  document.addEventListener("keydown", function (ev) {
    if ($("show")) return;
    var tag = (document.activeElement && document.activeElement.tagName) || "", typing = editingId || /INPUT|TEXTAREA|SELECT/.test(tag);
    var mod = ev.ctrlKey || ev.metaKey;
    if (ev.key === "F5" || (mod && ev.key === "Enter")) { ev.preventDefault(); play(ev.shiftKey ? cur : 0); return; }
    if (ev.key === "Escape" && editingId) { finishEdit(); drawSel(); return; }
    if (mod && ev.key.toLowerCase() === "s") { ev.preventDefault(); saver.now(); return; }
    if (typing) return;
    if (mod && ev.key.toLowerCase() === "z") { ev.preventDefault(); restore(ev.shiftKey ? redo : undo, ev.shiftKey ? undo : redo); return; }
    if (mod && ev.key.toLowerCase() === "y") { ev.preventDefault(); restore(redo, undo); return; }
    if (mod && ev.key.toLowerCase() === "d") { ev.preventDefault(); dup(); return; }
    if (mod && ev.key.toLowerCase() === "m") { ev.preventDefault(); deck.slides.splice(cur + 1, 0, layout("content")); goSlide(cur + 1); changed(); return; }
    if ((ev.key === "Delete" || ev.key === "Backspace") && selId) { ev.preventDefault(); del(); return; }
    if (ev.key === "Enter" && selId) { ev.preventDefault(); startEdit(selId); return; }
    var mv = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[ev.key];
    if (mv && selId) { ev.preventDefault(); var e = find(selId), k = ev.shiftKey ? 10 : 2; e.x += mv[0] * k; e.y += mv[1] * k; var d = domOf(selId); d.style.left = e.x + "px"; d.style.top = e.y + "px"; drawSel(); changed(); return; }
    if (ev.key === "PageDown" || (mv && !selId && (ev.key === "ArrowDown" || ev.key === "ArrowRight"))) { ev.preventDefault(); goSlide(cur + 1); return; }
    if (ev.key === "PageUp" || (mv && !selId && (ev.key === "ArrowUp" || ev.key === "ArrowLeft"))) { ev.preventDefault(); goSlide(cur - 1); return; }
  });

  /* ---------- presenting ---------- */
  $("playBtn").onclick = function () { play(0); };
  function play(from) {
    finishEdit(); select(null);
    var box = document.createElement("div"); box.id = "show"; box.setAttribute("role", "dialog"); box.setAttribute("aria-label", "Presenting");
    box.innerHTML = '<div class="bar"><span id="showN"></span><button type="button" data-a="prev" aria-label="Previous">◀</button><button type="button" data-a="next" aria-label="Next">▶</button><button type="button" data-a="end">End show</button></div>';
    document.body.appendChild(box);
    var i = from, holder = null, mouseT = null;
    function showAt(n, back) {
      if (n < 0) return; if (n >= deck.slides.length) { end(); return; }
      i = n; var s = deck.slides[i], sl = renderSlide(s, false);
      var k = Math.min(innerWidth / SW, innerHeight / SH);
      var w = document.createElement("div"); w.style.width = SW * k + "px"; w.style.height = SH * k + "px"; w.style.overflow = "hidden";
      sl.style.transform = "scale(" + k + ")"; sl.style.transformOrigin = "0 0";
      w.appendChild(sl);
      if (s.trans && s.trans !== "none") w.className = "t-" + s.trans + (back ? " t-back" : "");
      if (holder) holder.remove(); holder = w; box.insertBefore(w, box.firstChild);
      box.querySelector("#showN").textContent = (i + 1) + " / " + deck.slides.length;
    }
    function end() {
      document.removeEventListener("keydown", key, true); window.removeEventListener("resize", re);
      if (document.fullscreenElement) document.exitFullscreen().catch(function () {});
      box.remove(); goSlide(i);
    }
    function key(ev) {
      ev.stopPropagation();
      if (/^(ArrowRight|ArrowDown|PageDown| |Enter|n)$/.test(ev.key)) { ev.preventDefault(); showAt(i + 1); }
      else if (/^(ArrowLeft|ArrowUp|PageUp|Backspace|p)$/.test(ev.key)) { ev.preventDefault(); showAt(i - 1, true); }
      else if (ev.key === "Home") showAt(0); else if (ev.key === "End") showAt(deck.slides.length - 1);
      else if (ev.key === "Escape") { ev.preventDefault(); end(); }
    }
    function re() { showAt(i); }
    var sx = null;
    box.addEventListener("pointerdown", function (ev) { sx = ev.clientX; });
    box.addEventListener("pointerup", function (ev) {
      var a = ev.target.closest("[data-a]");
      if (a) { if (a.dataset.a === "end") end(); else if (a.dataset.a === "prev") showAt(i - 1, true); else showAt(i + 1); return; }
      if (sx !== null && Math.abs(ev.clientX - sx) > 50) { if (ev.clientX < sx) showAt(i + 1); else showAt(i - 1, true); }
      else if (ev.clientX < innerWidth * 0.25) showAt(i - 1, true); else showAt(i + 1);
      sx = null;
    });
    box.addEventListener("pointermove", function () { box.classList.add("mouse"); clearTimeout(mouseT); mouseT = setTimeout(function () { box.classList.remove("mouse"); }, 2200); });
    document.addEventListener("keydown", key, true); window.addEventListener("resize", re);
    if (box.requestFullscreen) box.requestFullscreen().catch(function () {}).then(function () { setTimeout(re, 120); });
    showAt(i);
  }

  /* ---------- export ---------- */
  var exM = menu("exportBtn", "exportMenu");
  exM.onclick = function (ev) {
    var b = ev.target.closest("[data-ex]"); if (!b) return; this.hidden = true; finishEdit();
    var t = $("title").value.trim() || "Untitled presentation";
    if (b.dataset.ex === "pdf") {
      var pa = $("printArea"); pa.innerHTML = "";
      deck.slides.forEach(function (s) { var pg = document.createElement("div"); pg.className = "pg"; pg.appendChild(renderSlide(s, false)); pa.appendChild(pg); });
      setTimeout(function () { window.print(); }, 200); return;
    }
    Office.toast("Making your PowerPoint…");
    Pptx.build(deck, THEMES, t).then(function (blob) { Office.download(blob, Office.safeName(t, ".pptx")); })
      .catch(function (e) { Office.toast("Couldn't make the PowerPoint: " + (e.message || e), true); });
  };

  /* ---------- files ---------- */
  function clean(d) {
    d = d && d.slides && d.slides.length ? d : { theme: "aurora", slides: [layout("title")] };
    d.theme = THEMES[d.theme] ? d.theme : "classic";
    d.slides = d.slides.slice(0, 200).map(function (s) {
      s.id = s.id || uid(); s.els = (s.els || []).slice(0, 80).filter(function (e) { return e && /^(text|img|shape)$/.test(e.type); });
      s.els.forEach(function (e) {
        e.id = e.id || uid(); ["x", "y", "w", "h", "size"].forEach(function (k) { e[k] = +e[k] || 0; });
        if (e.type === "text") e.html = Office.cleanHTML(e.html || "");
        if (e.type === "img" && !/^data:image\/(png|jpe?g|gif|webp);base64,/i.test(e.src || "")) e.src = "";
        if (e.color && !/^#[0-9a-f]{3,8}$/i.test(e.color)) delete e.color;
      });
      s.els = s.els.filter(function (e) { return e.type !== "img" || e.src; });
      return s;
    });
    return d;
  }
  function load(f) {
    file = { id: f.id, title: f.title || "Untitled presentation" }; $("title").value = file.title;
    deck = clean(f.body); cur = 0; selId = null; undo = []; redo = []; draw();
    history.replaceState(null, "", "?id=" + encodeURIComponent(f.id));
    try { localStorage.setItem("hl.present.last", f.id); } catch (e) {}
    $("status").textContent = Office.isCloud() ? "Saved to your account" : "Saved on this device";
  }
  function fresh() {
    file = { id: Office.uuid(), title: "Untitled presentation" }; $("title").value = file.title;
    deck = clean(null); cur = 0; selId = null; undo = []; redo = []; draw();
    history.replaceState(null, "", "?id=" + encodeURIComponent(file.id));
    try { localStorage.setItem("hl.present.last", file.id); } catch (e) {}
    saver.now();
  }
  function openId(id) {
    $("status").textContent = "Opening…";
    return Office.get(id).then(function (f) { if (f) load(f); else fresh(); })
      .catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); fresh(); });
  }
  deck = clean(null); draw();
  Office.init("deck").then(function (st) {
    var q = new URLSearchParams(location.search).get("id"), last = null;
    try { last = localStorage.getItem("hl.present.last"); } catch (e) {}
    if (!st.user) Office.toast("Not signed in - your work is saved in this browser.");
    if (q || last) return openId(q || last);
    return Office.list().then(function (rows) { if (rows.length) return openId(rows[0].id); fresh(); });
  });
  $("filesBtn").onclick = function () {
    finishEdit(); saver.now();
    Office.filesDialog({
      heading: "Your presentations", noun: "presentation", icon: "📙", current: file.id, accept: ".pptx",
      onOpen: openId, onNew: fresh, onDeletedCurrent: fresh,
      onImport: function (f) {
        Office.toast("Opening " + f.name + "…");
        Pptx.read(f).then(function (d) { fresh(); deck = clean(d); file.title = f.name.replace(/\.[^.]+$/, ""); $("title").value = file.title; draw(); saver.now(); Office.toast("Opened " + f.name); })
          .catch(function (e) { Office.toast("Couldn't open that PowerPoint: " + (e.message || e), true); });
      }
    });
  };
  Office.themeToggle($("themeBtn"));
  $("title").addEventListener("input", saver.poke);
  window.addEventListener("resize", fit);
  window.addEventListener("beforeunload", function () { finishEdit(); saver.now(); });
})();
