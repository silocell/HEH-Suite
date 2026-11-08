/* Launchpad Draw - a whiteboard: pens, shapes, arrows, text, sticky notes
   and pictures on a board you can move around and zoom. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var NS = "http://www.w3.org/2000/svg";
  var board = $("board"), svg = $("svg"), world = $("world"), ui = $("ui");
  var items = [], sel = null, tool = "select", colour = "#111827", fill = false;
  var view = { x: 0, y: 0, z: 1 }, bg = "dots";
  var undo = [], redo = [];
  var file = { id: null, title: "Untitled board" };
  var NOTE = { "#111827": "#fef08a", "#2563eb": "#bfdbfe", "#dc2626": "#fecaca", "#16a34a": "#bbf7d0", "#ea580c": "#fed7aa", "#7c3aed": "#ddd6fe" };
  var saver = Office.autosaver($("status"), function () {
    file.title = $("title").value.trim() || "Untitled board";
    return { id: file.id, title: file.title, body: { items: items, bg: bg, view: view } };
  }, {
    /* working together: someone else's change came in (each shape is its own part) */
    local: ["view"], surface: board, ring: "shadow",
    where: function () { return sel ? '#world [data-i="' + sel + '"]' : ""; },
    apply: function (b, title) {
      if (document.activeElement !== $("title")) $("title").value = title || "Untitled board";
      var keepView = view; clean(b); view = keepView;
      if (sel && !find(sel)) sel = null;
      draw(); $("hint").hidden = items.length > 0;
      if (edState) {
        var n = find(edState.it.id), ta = edState.ta;
        if (!n) { edState = null; ta.remove(); return; }
        edState.it = n;
        if (ta.value !== n.text) {
          // keep the cursor in the same place in the words
          var o = ta.value, nt = n.text, at = ta.selectionStart, p = 0, q = 0, mn = Math.min(o.length, nt.length);
          while (p < mn && o.charCodeAt(p) === nt.charCodeAt(p)) p++;
          while (q < mn - p && o.charCodeAt(o.length - 1 - q) === nt.charCodeAt(nt.length - 1 - q)) q++;
          ta.value = nt; at = at <= p ? at : at >= o.length - q ? at + nt.length - o.length : nt.length - q;
          ta.selectionStart = ta.selectionEnd = at;
        }
      }
    }
  });
  var edState = null;
  function uid() { return Math.random().toString(36).slice(2, 10); }
  function esc(s) { return Office.esc(s); }

  /* ---------- view ---------- */
  function applyView() {
    var t = "translate(" + view.x + "," + view.y + ") scale(" + view.z + ")";
    world.setAttribute("transform", t); ui.setAttribute("transform", t);
    var bgR = $("bgRect"); bgR.setAttribute("transform", t);
    bgR.setAttribute("fill", bg === "dots" ? "url(#pdots)" : bg === "grid" ? "url(#pgrid)" : "transparent");
    $("zN").textContent = Math.round(view.z * 100) + "%";
    drawSel();
  }
  function toWorld(cx, cy) { var r = board.getBoundingClientRect(); return [(cx - r.left - view.x) / view.z, (cy - r.top - view.y) / view.z]; }
  function zoomAt(f, cx, cy) {
    var r = board.getBoundingClientRect(), sx = cx - r.left, sy = cy - r.top;
    var nz = Math.max(0.1, Math.min(6, view.z * f));
    view.x = sx - (sx - view.x) * nz / view.z; view.y = sy - (sy - view.y) * nz / view.z; view.z = nz; applyView();
  }

  /* ---------- drawing items ---------- */
  var measure = document.createElement("canvas").getContext("2d");
  function wrap(text, width, fs) {
    measure.font = fs + 'px "Segoe UI", system-ui, sans-serif';
    var out = [];
    String(text || "").split("\n").forEach(function (para) {
      var words = para.split(/\s+/), line = "";
      words.forEach(function (w) {
        var t = line ? line + " " + w : w;
        if (measure.measureText(t).width > width && line) { out.push(line); line = w; } else line = t;
      });
      out.push(line);
    });
    return out;
  }
  function textSvg(it, x, y, w, h, colour, centre) {
    if (!it.text) return "";
    var fs = it.fs || 18, lines = wrap(it.text, Math.max(20, w - 16), fs), lh = fs * 1.25, total = lines.length * lh;
    var top = centre ? y + (h - total) / 2 + fs * 0.9 : y + 8 + fs * 0.9;
    return '<text font-family="Segoe UI, system-ui, sans-serif" font-size="' + fs + '" fill="' + colour + '"' + (centre ? ' text-anchor="middle"' : "") + ">" +
      lines.map(function (l, i) { return '<tspan x="' + (centre ? x + w / 2 : x + 8) + '" y="' + (top + i * lh).toFixed(1) + '">' + esc(l) + "</tspan>"; }).join("") + "</text>";
  }
  function box(it) { // normalised box
    if (it.t === "pen") {
      var x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      for (var i = 0; i < it.p.length; i += 2) { x0 = Math.min(x0, it.p[i]); y0 = Math.min(y0, it.p[i + 1]); x1 = Math.max(x1, it.p[i]); y1 = Math.max(y1, it.p[i + 1]); }
      return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
    }
    return { x: Math.min(it.x, it.x + it.w), y: Math.min(it.y, it.y + it.h), w: Math.abs(it.w), h: Math.abs(it.h) };
  }
  function itemSvg(it) {
    var c = it.c || "#111827", sw = it.sw || 4, b = box(it), f = it.f ? c : "transparent", fo = it.f ? ' fill-opacity="0.16"' : "";
    var common = ' stroke="' + c + '" stroke-width="' + sw + '" stroke-linecap="round" stroke-linejoin="round"';
    var s = "";
    if (it.t === "pen") {
      var d = ""; for (var i = 0; i < it.p.length; i += 2) d += (i ? "L" : "M") + it.p[i].toFixed(1) + " " + it.p[i + 1].toFixed(1);
      if (it.p.length === 2) d += "l0.1 0.1";
      s = '<path d="' + d + '" fill="none"' + common + (it.hl ? ' stroke-opacity="0.38"' : "") + '/><path d="' + d + '" fill="none" stroke="transparent" stroke-width="' + Math.max(14, sw + 10) + '"/>';
    } else if (it.t === "rect") s = '<rect x="' + b.x + '" y="' + b.y + '" width="' + b.w + '" height="' + b.h + '" rx="10" fill="' + f + '"' + fo + common + "/>" + textSvg(it, b.x, b.y, b.w, b.h, c, true);
    else if (it.t === "ellipse") s = '<ellipse cx="' + (b.x + b.w / 2) + '" cy="' + (b.y + b.h / 2) + '" rx="' + b.w / 2 + '" ry="' + b.h / 2 + '" fill="' + f + '"' + fo + common + "/>" + textSvg(it, b.x + b.w * 0.12, b.y, b.w * 0.76, b.h, c, true);
    else if (it.t === "diamond") s = '<polygon points="' + (b.x + b.w / 2) + "," + b.y + " " + (b.x + b.w) + "," + (b.y + b.h / 2) + " " + (b.x + b.w / 2) + "," + (b.y + b.h) + " " + b.x + "," + (b.y + b.h / 2) + '" fill="' + f + '"' + fo + common + "/>" + textSvg(it, b.x + b.w * 0.2, b.y, b.w * 0.6, b.h, c, true);
    else if (it.t === "line" || it.t === "arrow") {
      var x2 = it.x + it.w, y2 = it.y + it.h;
      s = '<line x1="' + it.x + '" y1="' + it.y + '" x2="' + x2 + '" y2="' + y2 + '"' + common + '/><line x1="' + it.x + '" y1="' + it.y + '" x2="' + x2 + '" y2="' + y2 + '" stroke="transparent" stroke-width="16"/>';
      if (it.t === "arrow") {
        var ang = Math.atan2(it.h, it.w), L = 12 + sw * 2.2;
        s += '<polygon points="' + x2 + "," + y2 + " " + (x2 - L * Math.cos(ang - 0.45)) + "," + (y2 - L * Math.sin(ang - 0.45)) + " " + (x2 - L * Math.cos(ang + 0.45)) + "," + (y2 - L * Math.sin(ang + 0.45)) + '" fill="' + c + '"/>';
      }
    } else if (it.t === "text") s = '<rect x="' + b.x + '" y="' + b.y + '" width="' + b.w + '" height="' + b.h + '" fill="transparent"/>' + textSvg(it, b.x, b.y, b.w, b.h, c, false);
    else if (it.t === "note") s = '<rect x="' + (b.x + 3) + '" y="' + (b.y + 5) + '" width="' + b.w + '" height="' + b.h + '" rx="6" fill="#000" opacity="0.12"/><rect x="' + b.x + '" y="' + b.y + '" width="' + b.w + '" height="' + b.h + '" rx="6" fill="' + (it.bgc || "#fef08a") + '"/>' + textSvg(it, b.x, b.y, b.w, b.h, "#1f2937", true);
    else if (it.t === "img") s = '<image href="' + it.src + '" x="' + b.x + '" y="' + b.y + '" width="' + b.w + '" height="' + b.h + '" preserveAspectRatio="none"/>';
    return '<g data-i="' + it.id + '">' + s + "</g>";
  }
  function draw() { world.innerHTML = items.map(itemSvg).join(""); drawSel(); }
  function drawOne(it) {
    var g = world.querySelector('[data-i="' + it.id + '"]');
    var t = document.createElementNS(NS, "g"); t.innerHTML = itemSvg(it); var n = t.firstChild;
    if (g) g.replaceWith(n); else world.appendChild(n);
  }
  function find(id) { return items.filter(function (i) { return i.id === id; })[0]; }
  function drawSel() {
    ui.innerHTML = "";
    var it = sel && find(sel); if (!it) return;
    var b = box(it), p = 6 / view.z, r = 7 / view.z;
    ui.innerHTML = '<rect x="' + (b.x - p) + '" y="' + (b.y - p) + '" width="' + (b.w + 2 * p) + '" height="' + (b.h + 2 * p) + '" fill="none" stroke="#0d9488" stroke-width="' + (1.5 / view.z) + '" stroke-dasharray="' + (5 / view.z) + '"/>' +
      (it.t === "pen" ? "" : '<circle data-h="se" cx="' + (it.t === "line" || it.t === "arrow" ? it.x + it.w : b.x + b.w + p) + '" cy="' + (it.t === "line" || it.t === "arrow" ? it.y + it.h : b.y + b.h + p) + '" r="' + r + '" fill="#fff" stroke="#0d9488" stroke-width="' + (2 / view.z) + '" style="cursor:nwse-resize"/>');
  }

  /* ---------- undo ---------- */
  function mark() { undo.push(JSON.stringify(items)); if (undo.length > 100) undo.shift(); redo = []; }
  function changed() { saver.poke(); $("hint").hidden = items.length > 0; }
  function restore(from, to) { var s = from.pop(); if (!s) return; to.push(JSON.stringify(items)); items = JSON.parse(s); sel = null; draw(); changed(); }
  $("undoBtn").onclick = function () { restore(undo, redo); };
  $("redoBtn").onclick = function () { restore(redo, undo); };

  /* ---------- tools ---------- */
  function setTool(t) {
    tool = t; if (t !== "select") { sel = null; drawSel(); }
    document.querySelectorAll("[data-tool]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.tool === t)); });
    board.classList.toggle("pan", t === "hand");
    board.style.cursor = t === "hand" ? "" : t === "select" ? "default" : t === "erase" ? "cell" : t === "text" ? "text" : "crosshair";
  }
  document.querySelectorAll("[data-tool]").forEach(function (b) { b.onclick = function () { setTool(b.dataset.tool); }; });
  function pickColour(c) {
    colour = c; $("cBar").style.background = c;
    document.querySelectorAll(".sw").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.c === c)); });
    var it = sel && find(sel);
    if (it) { mark(); if (it.t === "note") it.bgc = NOTE[c] || lighten(c); else it.c = c; drawOne(it); changed(); }
  }
  function lighten(h) { var n = parseInt(h.slice(1), 16), r = n >> 16 & 255, g = n >> 8 & 255, b = n & 255; return "rgb(" + [r, g, b].map(function (v) { return Math.round(v + (255 - v) * 0.7); }).join(",") + ")"; }
  document.querySelectorAll(".sw").forEach(function (b) { b.onclick = function () { pickColour(b.dataset.c); }; });
  $("cIn").addEventListener("input", function () { $("cBar").style.background = this.value; });
  $("cIn").addEventListener("change", function () { pickColour(this.value); });
  $("sw").onchange = function () { var it = sel && find(sel); if (it && it.sw !== undefined) { mark(); it.sw = +this.value; drawOne(it); drawSel(); changed(); } };
  $("fillBtn").onclick = function () {
    fill = !fill; this.setAttribute("aria-pressed", String(fill));
    var it = sel && find(sel); if (it && /rect|ellipse|diamond/.test(it.t)) { mark(); it.f = fill; drawOne(it); changed(); }
  };
  $("bg").onchange = function () { bg = this.value; applyView(); saver.poke(); };
  function del() { if (!sel) return; mark(); items = items.filter(function (i) { return i.id !== sel; }); sel = null; draw(); changed(); }
  $("delBtn").onclick = del;
  function dup() { var it = sel && find(sel); if (!it) return; mark(); var c = JSON.parse(JSON.stringify(it)); c.id = uid(); shift(c, 24, 24); items.push(c); sel = c.id; draw(); changed(); }
  $("dupBtn").onclick = dup;
  $("fwdBtn").onclick = function () { var it = sel && find(sel); if (!it) return; mark(); items.splice(items.indexOf(it), 1); items.push(it); draw(); changed(); };
  function shift(it, dx, dy) { if (it.t === "pen") for (var i = 0; i < it.p.length; i += 2) { it.p[i] += dx; it.p[i + 1] += dy; } else { it.x += dx; it.y += dy; } }

  /* ---------- pointer ---------- */
  var ptrs = {}, act = null, spaceDown = false;
  board.addEventListener("pointerdown", function (e) {
    if (e.target === $("ed")) return;
    board.focus({ preventScroll: true });
    ptrs[e.pointerId] = [e.clientX, e.clientY];
    var ids = Object.keys(ptrs);
    if (ids.length === 2) { // two fingers: move and zoom the board
      if (act && act.it && act.made) { items.pop(); draw(); }
      var a = ptrs[ids[0]], b = ptrs[ids[1]];
      act = { mode: "pinch", d: Math.hypot(a[0] - b[0], a[1] - b[1]), mx: (a[0] + b[0]) / 2, my: (a[1] + b[1]) / 2 }; return;
    }
    if (e.target.closest(".zoombar")) return;
    board.setPointerCapture(e.pointerId);
    var w = toWorld(e.clientX, e.clientY), x = w[0], y = w[1];
    if (tool === "hand" || spaceDown || e.button === 1) { act = { mode: "pan", sx: e.clientX, sy: e.clientY, vx: view.x, vy: view.y }; board.classList.add("panning"); return; }
    var hit = e.target.closest("[data-i]"), h = e.target.closest("[data-h]");
    if (tool === "select") {
      if (h && sel) { mark(); var it0 = find(sel); act = { mode: "resize", it: it0, o: JSON.parse(JSON.stringify(it0)), x: x, y: y }; return; }
      if (hit) { sel = hit.dataset.i; drawSel(); mark(); act = { mode: "move", it: find(sel), x: x, y: y, moved: false }; return; }
      sel = null; drawSel(); act = { mode: "pan", sx: e.clientX, sy: e.clientY, vx: view.x, vy: view.y }; return;
    }
    if (tool === "erase") { act = { mode: "erase" }; mark(); if (hit) { items = items.filter(function (i) { return i.id !== hit.dataset.i; }); draw(); changed(); } return; }
    if (tool === "text" || tool === "note") e.preventDefault();
    mark();
    var sw = +$("sw").value, it;
    if (tool === "pen" || tool === "hl") it = { id: uid(), t: "pen", hl: tool === "hl", c: tool === "hl" ? "#facc15" : colour, sw: tool === "hl" ? sw * 4 : sw * (e.pointerType === "pen" && e.pressure ? 0.6 + e.pressure : 1), p: [x, y] };
    else if (tool === "text") { it = { id: uid(), t: "text", x: x, y: y - 14, w: 240, h: 40, c: colour, fs: 22, text: "" }; items.push(it); draw(); act = null; editText(it, true); setTool("select"); sel = it.id; return; }
    else if (tool === "note") { it = { id: uid(), t: "note", x: x - 90, y: y - 70, w: 180, h: 140, bgc: NOTE[colour] || lighten(colour), fs: 18, text: "" }; items.push(it); draw(); act = null; editText(it, true); setTool("select"); sel = it.id; return; }
    else it = { id: uid(), t: tool, x: x, y: y, w: 0, h: 0, c: colour, sw: sw, f: fill && /rect|ellipse|diamond/.test(tool), fs: 18, text: "" };
    items.push(it); act = { mode: "make", it: it, x: x, y: y, made: true }; drawOne(it);
  });
  board.addEventListener("pointermove", function (e) {
    if (!ptrs[e.pointerId]) return;
    ptrs[e.pointerId] = [e.clientX, e.clientY];
    if (!act) return;
    if (act.mode === "pinch") {
      var ids = Object.keys(ptrs); if (ids.length < 2) return;
      var a = ptrs[ids[0]], b = ptrs[ids[1]], d = Math.hypot(a[0] - b[0], a[1] - b[1]), mx = (a[0] + b[0]) / 2, my = (a[1] + b[1]) / 2;
      view.x += mx - act.mx; view.y += my - act.my; act.mx = mx; act.my = my;
      zoomAt(d / act.d, mx, my); act.d = d; return;
    }
    if (act.mode === "pan") { view.x = act.vx + e.clientX - act.sx; view.y = act.vy + e.clientY - act.sy; applyView(); return; }
    var w = toWorld(e.clientX, e.clientY), x = w[0], y = w[1];
    if (act.mode === "erase") { var t = document.elementFromPoint(e.clientX, e.clientY), g = t && t.closest && t.closest("[data-i]"); if (g) { items = items.filter(function (i) { return i.id !== g.dataset.i; }); draw(); changed(); } return; }
    if (act.mode === "move") { shift(act.it, x - act.x, y - act.y); act.x = x; act.y = y; act.moved = true; drawOne(act.it); drawSel(); return; }
    if (act.mode === "resize") {
      var it = act.it, o = act.o;
      if (it.t === "line" || it.t === "arrow") { it.w = o.w + x - act.x; it.h = o.h + y - act.y; }
      else {
        var b = box(o); it.x = b.x; it.y = b.y; it.w = Math.max(20, b.w + x - act.x); it.h = Math.max(16, b.h + y - act.y);
        if (it.t === "img" && !e.shiftKey) it.h = it.w * b.h / b.w;
      }
      drawOne(it); drawSel(); return;
    }
    if (act.mode === "make") {
      var m = act.it;
      if (m.t === "pen") { var L = m.p.length; if (Math.hypot(m.p[L - 2] - x, m.p[L - 1] - y) < 1.5 / view.z || L > 8000) return; m.p.push(x, y); }
      else { m.w = x - act.x; m.h = y - act.y; if (e.shiftKey && !/line|arrow/.test(m.t)) { var s = Math.max(Math.abs(m.w), Math.abs(m.h)); m.w = s * Math.sign(m.w || 1); m.h = s * Math.sign(m.h || 1); } }
      drawOne(m);
    }
  });
  function up(e) {
    delete ptrs[e.pointerId];
    board.classList.remove("panning");
    if (!act) return;
    if (act.mode === "move" && !act.moved) undo.pop();
    if (act.mode === "make") {
      var m = act.it;
      if (m.t !== "pen" && Math.abs(m.w) < 6 && Math.abs(m.h) < 6) { // a tap: a shape of a normal size
        if (/line|arrow/.test(m.t)) { m.w = 160; m.h = 0; } else { m.x -= 70; m.y -= 45; m.w = 140; m.h = 90; }
      }
      if (m.t !== "pen") { var b = box(m); if (!/line|arrow/.test(m.t)) { m.x = b.x; m.y = b.y; m.w = b.w; m.h = b.h; } sel = m.id; setTool("select"); }
      drawOne(m); drawSel();
    }
    if (act.mode !== "pan" && act.mode !== "pinch") changed();
    act = null;
  }
  board.addEventListener("pointerup", up); board.addEventListener("pointercancel", up);
  board.addEventListener("dblclick", function (e) {
    var g = document.elementsFromPoint(e.clientX, e.clientY).map(function (n) { return n.closest && n.closest("[data-i]"); }).filter(Boolean)[0];
    var it = g && find(g.dataset.i);
    if (it && /rect|ellipse|diamond|note|text/.test(it.t)) { sel = it.id; drawSel(); editText(it); }
    else if (!it) { var w = toWorld(e.clientX, e.clientY); mark(); var t = { id: uid(), t: "text", x: w[0], y: w[1] - 14, w: 240, h: 40, c: colour, fs: 22, text: "" }; items.push(t); draw(); editText(t, true); sel = t.id; }
  });
  board.addEventListener("wheel", function (e) {
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) zoomAt(Math.exp(-e.deltaY * 0.01), e.clientX, e.clientY);
    else { view.x -= e.deltaX; view.y -= e.deltaY; applyView(); }
  }, { passive: false });
  // Safari's pinch on a trackpad
  board.addEventListener("gesturechange", function (e) { e.preventDefault(); zoomAt(e.scale / (board._gs || 1), e.clientX, e.clientY); board._gs = e.scale; });
  board.addEventListener("gesturestart", function (e) { e.preventDefault(); board._gs = 1; });

  /* typing in shapes, notes and text boxes */
  function editText(it, isNew) {
    var old = $("ed"); if (old) old.blur();
    var b = box(it), r = board.getBoundingClientRect();
    var ta = document.createElement("textarea"); ta.id = "ed"; ta.value = it.text || "";
    ta.style.left = (b.x * view.z + view.x) + "px"; ta.style.top = (b.y * view.z + view.y) + "px";
    ta.style.width = Math.max(160, b.w * view.z) + "px"; ta.style.height = Math.max(44, b.h * view.z) + "px";
    ta.style.fontSize = (it.fs || 18) * view.z + "px";
    if (it.t === "note") ta.style.background = it.bgc;
    board.appendChild(ta);
    // after the tap that made it has finished, or the board takes the focus back
    ta.focus(); setTimeout(function () { ta.focus(); }, 0);
    if (!isNew) mark();
    edState = { it: it, ta: ta };
    ta.addEventListener("input", function () {
      it = edState && edState.ta === ta ? edState.it : it;
      it.text = ta.value.slice(0, 4000);
      if (it.t === "text") { var lines = wrap(it.text, it.w - 16, it.fs).length; it.h = Math.max(40, lines * it.fs * 1.25 + 16); ta.style.height = it.h * view.z + "px"; }
      drawOne(it); saver.poke();
    });
    ta.addEventListener("keydown", function (e) { if (e.key === "Escape" || (e.key === "Enter" && (e.ctrlKey || e.metaKey))) ta.blur(); e.stopPropagation(); });
    ta.addEventListener("blur", function () {
      if (edState && edState.ta === ta) { it = edState.it; edState = null; }
      ta.remove();
      if (it.t === "text" && !it.text.trim()) { items = items.filter(function (i) { return i !== it; }); sel = null; }
      draw(); changed();
    });
  }
  $("imgIn").onchange = function () {
    var f = this.files[0]; this.value = ""; if (!f) return;
    Office.shrinkImage(f, 1400).then(function (src) {
      var im = new Image(); im.onload = function () {
        var r = board.getBoundingClientRect(), c = toWorld(r.left + r.width / 2, r.top + r.height / 2), k = Math.min(420 / im.width, 320 / im.height, 1) / view.z;
        mark(); var it = { id: uid(), t: "img", src: src, x: c[0] - im.width * k / 2, y: c[1] - im.height * k / 2, w: im.width * k, h: im.height * k };
        items.push(it); sel = it.id; setTool("select"); draw(); changed();
      }; im.src = src;
    }).catch(function (e) { Office.toast(e.message, true); });
  };
  board.addEventListener("paste", function (e) {
    var f = Array.prototype.filter.call((e.clipboardData && e.clipboardData.files) || [], function (x) { return /^image\//.test(x.type); })[0];
    if (f) { e.preventDefault(); var dt = new DataTransfer(); dt.items.add(f); $("imgIn").files = dt.files; $("imgIn").onchange(); }
  });

  /* ---------- keys ---------- */
  document.addEventListener("keydown", function (e) {
    if (/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName) || $("ed")) return;
    var mod = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    if (e.key === " ") { spaceDown = true; board.classList.add("pan"); e.preventDefault(); return; }
    if (mod && k === "z") { e.preventDefault(); restore(e.shiftKey ? redo : undo, e.shiftKey ? undo : redo); return; }
    if (mod && k === "y") { e.preventDefault(); restore(redo, undo); return; }
    if (mod && k === "d") { e.preventDefault(); dup(); return; }
    if (mod && k === "s") { e.preventDefault(); saver.now(); return; }
    if (mod) return;
    if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); del(); return; }
    if (e.key === "Enter" && sel) { var it = find(sel); if (it && /rect|ellipse|diamond|note|text/.test(it.t)) { e.preventDefault(); editText(it); } return; }
    var keys = { v: "select", h: "hand", p: "pen", r: "rect", o: "ellipse", a: "arrow", l: "line", t: "text", n: "note", e: "erase", d: "diamond" };
    if (keys[k]) setTool(keys[k]);
    var mv = { arrowleft: [-1, 0], arrowright: [1, 0], arrowup: [0, -1], arrowdown: [0, 1] }[k];
    if (mv && sel) { e.preventDefault(); var s = find(sel), st = e.shiftKey ? 20 : 2; shift(s, mv[0] * st, mv[1] * st); drawOne(s); drawSel(); changed(); }
  });
  document.addEventListener("keyup", function (e) { if (e.key === " ") { spaceDown = false; board.classList.remove("pan"); } });

  /* ---------- zoom buttons ---------- */
  function centre() { var r = board.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; }
  $("zIn").onclick = function () { var c = centre(); zoomAt(1.25, c[0], c[1]); };
  $("zOut").onclick = function () { var c = centre(); zoomAt(0.8, c[0], c[1]); };
  $("zN").onclick = function () { var c = centre(); zoomAt(1 / view.z, c[0], c[1]); };
  function bounds() {
    if (!items.length) return null;
    var x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    items.forEach(function (it) { var b = box(it), p = (it.sw || 0) + 10; x0 = Math.min(x0, b.x - p); y0 = Math.min(y0, b.y - p); x1 = Math.max(x1, b.x + b.w + p); y1 = Math.max(y1, b.y + b.h + p); });
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }
  function fit() {
    var b = bounds(), r = board.getBoundingClientRect(); if (!b) { view = { x: r.width / 2, y: r.height / 2, z: 1 }; applyView(); return; }
    var z = Math.max(0.1, Math.min(2, Math.min((r.width - 80) / b.w, (r.height - 80) / b.h)));
    view = { z: z, x: (r.width - b.w * z) / 2 - b.x * z, y: (r.height - b.h * z) / 2 - b.y * z }; applyView();
  }
  $("zFit").onclick = fit;

  /* ---------- starting points ---------- */
  var tplM = $("tplMenu"); $("tplBtn").onclick = function (e) { e.stopPropagation(); tplM.hidden = !tplM.hidden; };
  document.addEventListener("click", function (e) { if (!tplM.hidden && !tplM.contains(e.target)) tplM.hidden = true; if (!exM.hidden && !exM.contains(e.target)) exM.hidden = true; });
  tplM.onclick = function (e) {
    var b = e.target.closest("[data-tpl]"); if (!b) return; tplM.hidden = true;
    var c = toWorld.apply(null, centre()), cx = c[0], cy = c[1], add = [], t = b.dataset.tpl;
    function S(tp, x, y, w, h, text, col, extra) { return Object.assign({ id: uid(), t: tp, x: x, y: y, w: w, h: h, c: col || "#111827", sw: 3, fs: 18, text: text || "", f: true }, extra || {}); }
    function A(x1, y1, x2, y2, col) { return { id: uid(), t: "arrow", x: x1, y: y1, w: x2 - x1, h: y2 - y1, c: col || "#64748b", sw: 3 }; }
    if (t === "mind") {
      add.push(S("ellipse", cx - 110, cy - 50, 220, 100, "Big idea", "#7c3aed", { fs: 22 }));
      [["Idea 1", -320, -170, "#2563eb"], ["Idea 2", 200, -170, "#16a34a"], ["Idea 3", -320, 110, "#ea580c"], ["Idea 4", 200, 110, "#dc2626"]].forEach(function (n) {
        add.push(S("rect", cx + n[1], cy + n[2], 140, 70, n[0], n[3]));
        add.push({ id: uid(), t: "line", x: cx + (n[1] < 0 ? -95 : 95), y: cy + (n[2] < 0 ? -30 : 30), w: (n[1] < 0 ? n[1] + 140 + 95 : n[1] - 95), h: (n[2] < 0 ? n[2] + 70 + 30 : n[2] - 30), c: n[3], sw: 3 });
      });
    } else if (t === "flow") {
      add.push(S("ellipse", cx - 80, cy - 260, 160, 70, "Start", "#16a34a"), A(cx, cy - 190, cx, cy - 140));
      add.push(S("rect", cx - 90, cy - 140, 180, 70, "Do something", "#2563eb"), A(cx, cy - 70, cx, cy - 20));
      add.push(S("diamond", cx - 110, cy - 20, 220, 120, "A question?", "#ea580c"), A(cx, cy + 100, cx, cy + 150), A(cx + 110, cy + 40, cx + 220, cy + 40));
      add.push(S("text", cx + 125, cy + 6, 60, 30, "No", "#64748b", { f: false }), S("text", cx + 6, cy + 108, 60, 30, "Yes", "#64748b", { f: false }));
      add.push(S("rect", cx + 220, cy + 5, 170, 70, "Try again", "#7c3aed"), S("ellipse", cx - 80, cy + 150, 160, 70, "End", "#dc2626"));
    } else if (t === "kanban") {
      [["To do", "#dc2626", "#fecaca"], ["Doing", "#ea580c", "#fed7aa"], ["Done", "#16a34a", "#bbf7d0"]].forEach(function (col, i) {
        var x = cx - 330 + i * 230;
        add.push(S("rect", x, cy - 220, 210, 440, "", col[1], { f: true }), S("text", x + 10, cy - 210, 190, 40, col[0], col[1], { fs: 24, f: false }));
        add.push({ id: uid(), t: "note", x: x + 20, y: cy - 150, w: 170, h: 110, bgc: col[2], fs: 16, text: "A task" });
      });
    } else if (t === "venn") {
      add.push(S("ellipse", cx - 230, cy - 150, 300, 300, "", "#2563eb"), S("ellipse", cx - 70, cy - 150, 300, 300, "", "#dc2626"));
      add.push(S("text", cx - 200, cy - 190, 200, 40, "Group A", "#2563eb", { fs: 22, f: false }), S("text", cx + 60, cy - 190, 200, 40, "Group B", "#dc2626", { fs: 22, f: false }), S("text", cx - 50, cy - 12, 120, 40, "Both", "#111827", { fs: 20, f: false }));
    }
    mark(); items = items.concat(add); sel = null; draw(); changed(); fit();
  };

  /* ---------- export ---------- */
  var exM = $("exportMenu"); $("exportBtn").onclick = function (e) { e.stopPropagation(); exM.hidden = !exM.hidden; };
  function standalone() {
    var b = bounds() || { x: 0, y: 0, w: 400, h: 300 };
    return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="' + [b.x, b.y, b.w, b.h].map(function (n) { return n.toFixed(1); }).join(" ") + '" width="' + Math.round(b.w) + '" height="' + Math.round(b.h) + '">' +
      '<rect x="' + b.x + '" y="' + b.y + '" width="' + b.w + '" height="' + b.h + '" fill="#ffffff"/>' + items.map(itemSvg).join("") + "</svg>";
  }
  exM.onclick = function (e) {
    var b = e.target.closest("[data-ex]"); if (!b) return; exM.hidden = true;
    var t = $("title").value.trim() || "Board";
    if (!items.length) { Office.toast("The board is empty."); return; }
    if (b.dataset.ex === "pdf") { sel = null; drawSel(); fit(); setTimeout(function () { window.print(); }, 150); return; }
    var s = standalone();
    if (b.dataset.ex === "svg") { Office.download(new Blob([s], { type: "image/svg+xml" }), Office.safeName(t, ".svg")); return; }
    var bb = bounds(), k = Math.min(3, 4000 / Math.max(bb.w, bb.h)), img = new Image();
    img.onload = function () {
      var c = document.createElement("canvas"); c.width = Math.round(bb.w * k); c.height = Math.round(bb.h * k);
      var g = c.getContext("2d"); g.drawImage(img, 0, 0, c.width, c.height);
      c.toBlob(function (bl) { Office.download(bl, Office.safeName(t, ".png")); });
    };
    img.onerror = function () { Office.toast("Couldn't make the picture.", true); };
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(s);
  };

  /* ---------- files ---------- */
  function clean(b) {
    b = b || {};
    var ok = /^(pen|rect|ellipse|diamond|line|arrow|text|note|img)$/;
    items = (Array.isArray(b.items) ? b.items : []).filter(function (i) { return i && ok.test(i.t); }).slice(0, 5000).map(function (i) {
      ["x", "y", "w", "h", "sw", "fs"].forEach(function (k) { if (k in i) i[k] = +i[k] || 0; });
      if (i.c && !/^#[0-9a-f]{3,8}$/i.test(i.c)) i.c = "#111827";
      if (i.bgc && !/^(#[0-9a-f]{3,8}|rgb\(\d+,\d+,\d+\))$/i.test(i.bgc)) i.bgc = "#fef08a";
      if (i.t === "img" && !/^data:image\/(png|jpe?g|gif|webp);base64,/i.test(i.src || "")) i.src = "";
      if (i.t === "pen") i.p = (i.p || []).map(Number);
      i.id = String(i.id || uid()).replace(/[^a-z0-9]/gi, "").slice(0, 16) || uid();
      i.text = String(i.text || "").slice(0, 4000);
      return i;
    }).filter(function (i) { return i.t !== "img" || i.src; });
    bg = /^(dots|grid|plain)$/.test(b.bg) ? b.bg : "dots"; $("bg").value = bg;
    if (b.view && isFinite(b.view.z)) view = { x: +b.view.x || 0, y: +b.view.y || 0, z: Math.max(0.1, Math.min(6, +b.view.z || 1)) };
    else { var r = board.getBoundingClientRect(); view = { x: r.width / 2, y: r.height / 2, z: 1 }; }
  }
  function load(f) {
    file = { id: f.id, title: f.title || "Untitled board" }; $("title").value = file.title;
    clean(f.body); undo = []; redo = []; sel = null; draw(); applyView(); $("hint").hidden = items.length > 0;
    history.replaceState(null, "", "?id=" + encodeURIComponent(f.id));
    try { localStorage.setItem("hl.draw.last", f.id); } catch (e) {}
    $("status").textContent = Office.isCloud() ? "Saved to your account" : "Saved on this device";
    saver.opened();
  }
  function fresh() {
    file = { id: Office.uuid(), title: "Untitled board" }; $("title").value = file.title;
    clean(null); undo = []; redo = []; sel = null; draw(); applyView(); $("hint").hidden = false;
    history.replaceState(null, "", "?id=" + encodeURIComponent(file.id));
    try { localStorage.setItem("hl.draw.last", file.id); } catch (e) {}
    saver.now(); saver.opened();
  }
  function openId(id) {
    $("status").textContent = "Opening…";
    return Office.get(id).then(function (f) { if (f) load(f); else fresh(); }).catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); fresh(); });
  }
  clean(null); applyView();
  Office.init("draw").then(function (st) {
    var q = new URLSearchParams(location.search).get("id"), last = null;
    try { last = localStorage.getItem("hl.draw.last"); } catch (e) {}
    if (!st.user) Office.toast("Not signed in - your boards are saved in this browser.");
    if (q || last) return openId(q || last);
    return Office.list().then(function (rows) { if (rows.length) return openId(rows[0].id); fresh(); });
  });
  $("filesBtn").onclick = function () {
    saver.now();
    Office.filesDialog({
      heading: "Your boards", noun: "board", icon: "🎨", current: file.id, accept: ".svg,.png,.jpg,.jpeg",
      onOpen: openId, onNew: fresh, onDeletedCurrent: fresh,
      onImport: function (f) { // a picture goes onto the board
        var dt = new DataTransfer(); dt.items.add(f); $("imgIn").files = dt.files; $("imgIn").onchange();
      }
    });
  };
  Office.themeToggle($("themeBtn"));
  $("title").addEventListener("input", saver.poke);
  window.addEventListener("resize", drawSel);
  window.addEventListener("beforeunload", function () { saver.now(); });
})();
