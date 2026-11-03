/* Launchpad Notebook - sections, pages, notes, checklists and drawing. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var note = $("note"), ptitle = $("ptitle"), titleIn = $("title");
  var COLS = ["#7c3aed", "#2563eb", "#16a34a", "#ea580c", "#db2777", "#0891b2", "#ca8a04", "#dc2626"];
  var file = { id: null, title: "My Notebook" }, nb = null, sec = null, page = null;
  var saver = Office.autosaver($("status"), function () {
    flush();
    file.title = titleIn.value.trim() || "My Notebook";
    return { id: file.id, title: file.title, body: nb };
  });
  function uid() { return Math.random().toString(36).slice(2, 10); }
  function now() { return new Date().toISOString(); }
  function newPage(title, html) { return { id: uid(), title: title || "", html: html || "", ink: [], created: now(), updated: now() }; }
  function starter() {
    var welcome = newPage("Welcome to your Notebook",
      "<p>This is your notebook - sections down the left, pages next to them. Everything saves by itself.</p>" +
      '<ul class="todo"><li class="done">Open the Notebook</li><li>Make a section for each subject</li><li>Tap the box to tick things off</li></ul>' +
      "<p>⭐ Tags help you find things later. Try <b>✏️ Draw</b> to write or sketch with a finger or pencil.</p>");
    return { sections: [{ id: uid(), name: "Homework", color: COLS[0], pages: [welcome] }, { id: uid(), name: "Revision", color: COLS[1], pages: [newPage("Revision plan")] }] };
  }
  function clean(b) {
    if (!b || !b.sections || !b.sections.length) return starter();
    b.sections = b.sections.slice(0, 60).map(function (s, i) {
      return { id: s.id || uid(), name: String(s.name || "Section").slice(0, 60), color: /^#[0-9a-f]{6}$/i.test(s.color) ? s.color : COLS[i % COLS.length],
        pages: (s.pages || []).slice(0, 300).map(function (p) {
          return { id: p.id || uid(), title: String(p.title || "").slice(0, 120), html: Office.cleanHTML(p.html || ""), ink: Array.isArray(p.ink) ? p.ink.slice(0, 3000) : [],
            created: p.created || now(), updated: p.updated || now() };
        }) };
    });
    return b;
  }

  /* ---------- the lists ---------- */
  function drawSecs() {
    $("secList").innerHTML = nb.sections.map(function (s) {
      return '<button type="button" class="sec" data-s="' + s.id + '"' + (sec === s ? ' aria-current="true"' : "") + '><span class="dot" style="background:' + s.color + '"></span>' +
        Office.esc(s.name) + '<span class="n">' + s.pages.length + "</span></button>";
    }).join("");
  }
  function strip(h) { var d = document.createElement("div"); d.innerHTML = String(h).replace(/<\/(p|li|div|h\d)>|<br>/g, " $&"); return (d.textContent || "").replace(/\s+/g, " ").trim(); }
  function drawPages() {
    var q = $("search").value.trim().toLowerCase(), rows = [];
    if (q) {
      nb.sections.forEach(function (s) { s.pages.forEach(function (p) {
        var t = strip(p.html), i = t.toLowerCase().indexOf(q);
        if (i >= 0 || p.title.toLowerCase().indexOf(q) >= 0) rows.push({ s: s, p: p, snip: i >= 0 ? "…" + t.slice(Math.max(0, i - 20), i + 50) : s.name });
      }); });
      $("pagesHead").textContent = rows.length + (rows.length === 1 ? " match" : " matches");
    } else {
      sec.pages.forEach(function (p) { rows.push({ s: sec, p: p, snip: strip(p.html).slice(0, 60) || "Empty page" }); });
      $("pagesHead").textContent = sec.name;
    }
    $("pageList").innerHTML = rows.map(function (r) {
      return '<button type="button" class="pg" data-s="' + r.s.id + '" data-p="' + r.p.id + '"' + (page === r.p ? ' aria-current="true"' : "") + ">" +
        "<strong>" + Office.esc(r.p.title || "Untitled page") + "</strong><small>" + Office.esc(r.snip) + "</small></button>";
    }).join("") || '<p style="color:var(--soft);padding:8px">' + (q ? "Nothing found." : "No pages yet - tap ＋") + "</p>";
  }
  $("secList").addEventListener("click", function (e) {
    var b = e.target.closest(".sec"); if (!b) return;
    var s = nb.sections.filter(function (x) { return x.id === b.dataset.s; })[0];
    if (s === sec && !b.dataset.renaming) { rename(b, s); return; }
    flush(); sec = s; $("search").value = ""; openPage(sec.pages[0] || null); drawSecs(); drawPages();
  });
  function rename(b, s) {
    b.dataset.renaming = 1;
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>Section</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
      '<input id="snm" maxlength="60" style="width:100%;height:40px;border:1px solid var(--rule);border-radius:10px;padding:0 12px;background:var(--tool);margin:12px 0">' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px">' + COLS.map(function (c) { return '<button type="button" data-c="' + c + '" aria-label="Colour" style="width:32px;height:32px;border-radius:9px;border:' + (c === s.color ? "3px solid var(--ink)" : "0") + ";background:" + c + '"></button>'; }).join("") + "</div>" +
      '<div class="of-new"><button type="button" class="of-btn" id="sOk">Save</button><button type="button" class="of-btn ghost" id="sDel">Delete section</button></div>';
    document.body.appendChild(d); d.showModal();
    var inp = d.querySelector("#snm"); inp.value = s.name; inp.focus(); inp.select(); var col = s.color;
    function close() { d.close(); d.remove(); delete b.dataset.renaming; }
    d.onclick = function (e) {
      if (e.target.closest(".of-x") || e.target === d) { close(); return; }
      var c = e.target.closest("[data-c]");
      if (c) { col = c.dataset.c; d.querySelectorAll("[data-c]").forEach(function (x) { x.style.border = x === c ? "3px solid var(--ink)" : "0"; }); return; }
      if (e.target.id === "sOk") { s.name = inp.value.trim().slice(0, 60) || s.name; s.color = col; close(); drawSecs(); drawPages(); saver.poke(); }
      if (e.target.id === "sDel") {
        if (nb.sections.length < 2) { Office.toast("A notebook needs at least one section."); return; }
        var t = e.target; if (!t.dataset.armed) { t.dataset.armed = 1; t.textContent = "Delete it and its " + s.pages.length + " pages?"; t.style.background = "var(--bad)"; t.style.color = "#fff"; return; }
        nb.sections.splice(nb.sections.indexOf(s), 1); sec = nb.sections[0]; close(); openPage(sec.pages[0] || null); drawSecs(); drawPages(); saver.poke();
      }
    };
    inp.onkeydown = function (e) { if (e.key === "Enter") d.querySelector("#sOk").click(); };
  }
  $("addSec").onclick = function () {
    if (nb.sections.length >= 60) return;
    flush(); var s = { id: uid(), name: "New section", color: COLS[nb.sections.length % COLS.length], pages: [newPage("")] };
    nb.sections.push(s); sec = s; openPage(s.pages[0]); drawSecs(); drawPages(); saver.poke();
    rename($("secList").querySelector('[data-s="' + s.id + '"]'), s);
  };
  $("pageList").addEventListener("click", function (e) {
    var b = e.target.closest(".pg"); if (!b) return;
    flush();
    var s = nb.sections.filter(function (x) { return x.id === b.dataset.s; })[0], p = s.pages.filter(function (x) { return x.id === b.dataset.p; })[0];
    sec = s; openPage(p); drawSecs(); drawPages();
    var q = $("search").value.trim(); if (q) markHits(q);
  });
  $("addPage").onclick = function () {
    if (sec.pages.length >= 300) return;
    flush(); var p = newPage(""); sec.pages.unshift(p); $("search").value = ""; openPage(p); drawSecs(); drawPages(); saver.poke(); ptitle.focus();
  };
  $("search").addEventListener("input", drawPages);

  function markHits(q) {
    var low = q.toLowerCase(), w = document.createTreeWalker(note, NodeFilter.SHOW_TEXT), list = [], first = null;
    while (w.nextNode()) list.push(w.currentNode);
    list.forEach(function (n) {
      var i = n.nodeValue.toLowerCase().indexOf(low); if (i < 0) return;
      var r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + q.length);
      var m = document.createElement("mark"); m.className = "hit"; r.surroundContents(m); if (!first) first = m;
    });
    if (first) first.scrollIntoView({ block: "center" });
    setTimeout(function () { note.querySelectorAll("mark.hit").forEach(function (m) { m.replaceWith(document.createTextNode(m.textContent)); }); note.normalize(); }, 2500);
  }

  /* ---------- a page ---------- */
  function openPage(p) {
    page = p;
    var has = !!p;
    ptitle.disabled = !has; note.contentEditable = has ? "true" : "false";
    ptitle.value = p ? p.title : ""; note.innerHTML = p ? p.html : "";
    $("pdate").textContent = p ? "Made " + new Date(p.created).toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric" }) : "";
    $("cw").scrollTop = 0;
    sizeInk(); redrawInk();
    try { localStorage.setItem("hl.nb.cur", JSON.stringify({ s: sec && sec.id, p: p && p.id })); } catch (e) {}
  }
  function flush() {
    if (!page) return;
    var c = note.cloneNode(true); c.querySelectorAll("mark.hit").forEach(function (m) { m.replaceWith(document.createTextNode(m.textContent)); });
    var h = c.innerHTML, t = ptitle.value.slice(0, 120);
    if (h !== page.html || t !== page.title) { page.html = h; page.title = t; page.updated = now(); }
  }
  var listT = null;
  function edited() { saver.poke(); clearTimeout(listT); listT = setTimeout(function () { flush(); drawPages(); sizeInk(); }, 600); }
  note.addEventListener("input", edited);
  ptitle.addEventListener("input", edited);
  ptitle.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); note.focus(); } });
  note.addEventListener("paste", function (e) {
    var cd = e.clipboardData; if (!cd) return;
    var imgs = Array.prototype.filter.call(cd.files || [], function (f) { return /^image\//.test(f.type); });
    e.preventDefault();
    if (imgs.length) { imgs.forEach(addImage); return; }
    var html = cd.getData("text/html");
    if (html) document.execCommand("insertHTML", false, Office.cleanHTML(html.replace(/<!--[\s\S]*?-->/g, "")));
    else document.execCommand("insertText", false, cd.getData("text/plain"));
  });
  // tick a checklist item by tapping its box
  note.addEventListener("click", function (e) {
    var li = e.target.closest("ul.todo > li"); if (!li) return;
    var r = li.getBoundingClientRect();
    if (e.clientX - r.left < 26) { e.preventDefault(); li.classList.toggle("done"); edited(); }
    var a = e.target.closest("a"); if (a && (e.ctrlKey || e.metaKey)) window.open(a.href, "_blank", "noopener");
  });

  /* ---------- the ribbon ---------- */
  document.querySelectorAll("[data-cmd]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.onclick = function () { note.focus(); document.execCommand(b.dataset.cmd); edited(); };
  });
  document.querySelectorAll("[data-block]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.onclick = function () { note.focus(); document.execCommand("formatBlock", false, "<" + b.dataset.block + ">"); edited(); };
  });
  document.querySelectorAll("[data-tag]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.onclick = function () { note.focus(); document.execCommand("insertText", false, b.dataset.tag + " "); edited(); };
  });
  $("dateBtn").addEventListener("mousedown", function (e) { e.preventDefault(); });
  $("dateBtn").onclick = function () { note.focus(); document.execCommand("insertText", false, new Date().toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" }) + " "); edited(); };
  $("hl").addEventListener("input", function () { $("hlBar").style.background = this.value; });
  $("hl").addEventListener("change", function () { note.focus(); document.execCommand("hiliteColor", false, this.value); edited(); });
  $("todoBtn").addEventListener("mousedown", function (e) { e.preventDefault(); });
  $("todoBtn").onclick = todo;
  function todo() {
    note.focus();
    var s = getSelection(), n = s.anchorNode && (s.anchorNode.nodeType === 3 ? s.anchorNode.parentNode : s.anchorNode);
    var ul = n && n.closest && n.closest("ul");
    if (ul && note.contains(ul)) { ul.classList.toggle("todo"); edited(); return; }
    if (n && n.closest && n.closest("ol")) document.execCommand("insertOrderedList");
    document.execCommand("insertUnorderedList");
    n = s.anchorNode && (s.anchorNode.nodeType === 3 ? s.anchorNode.parentNode : s.anchorNode);
    ul = n && n.closest && n.closest("ul"); if (ul) ul.classList.add("todo");
    edited();
  }
  document.addEventListener("keydown", function (e) {
    var mod = e.ctrlKey || e.metaKey;
    if (mod && e.key === "1") { e.preventDefault(); todo(); }
    if (mod && e.key.toLowerCase() === "s") { e.preventDefault(); saver.now(); }
  });
  $("imgIn").onchange = function () { if (this.files[0]) addImage(this.files[0]); this.value = ""; };
  function addImage(f) {
    Office.shrinkImage(f, 1200).then(function (src) { note.focus(); document.execCommand("insertHTML", false, '<img src="' + src + '" alt="">'); edited(); setTimeout(sizeInk, 300); })
      .catch(function (e) { Office.toast(e.message, true); });
  }

  /* ---------- drawing ---------- */
  var cv = $("ink"), g = cv.getContext("2d"), sheetEl = $("sheetEl"), drawing = false, pen = "#111827", erasing = false, stroke = null;
  function sizeInk() {
    var w = sheetEl.offsetWidth, h = Math.max(sheetEl.scrollHeight, $("cw").clientHeight), r = window.devicePixelRatio || 1;
    var maxY = 0; if (page) page.ink.forEach(function (s) { for (var i = 1; i < s.p.length; i += 2) if (s.p[i] > maxY) maxY = s.p[i]; });
    h = Math.max(h, maxY + 200);
    if (cv.width !== Math.round(w * r) || cv.height !== Math.round(h * r)) {
      cv.width = Math.round(w * r); cv.height = Math.round(h * r); cv.style.width = w + "px"; cv.style.height = h + "px";
      sheetEl.style.minHeight = h + "px"; redrawInk();
    }
  }
  function paint(s) {
    if (s.p.length < 2) return;
    g.save(); g.lineCap = "round"; g.lineJoin = "round";
    if (s.hl) { g.globalAlpha = .35; g.strokeStyle = "#facc15"; g.lineWidth = s.w * 4; } else { g.strokeStyle = s.c; g.lineWidth = s.w; }
    g.beginPath(); g.moveTo(s.p[0], s.p[1]);
    if (s.p.length === 2) g.lineTo(s.p[0] + .1, s.p[1] + .1);
    for (var i = 2; i < s.p.length; i += 2) g.lineTo(s.p[i], s.p[i + 1]);
    g.stroke(); g.restore();
  }
  function redrawInk() {
    var r = window.devicePixelRatio || 1; g.setTransform(r, 0, 0, r, 0, 0); g.clearRect(0, 0, cv.width, cv.height);
    if (page) page.ink.forEach(paint);
  }
  $("drawBtn").onclick = function () {
    drawing = !drawing; this.setAttribute("aria-pressed", String(drawing));
    $("inkbar").hidden = !drawing; sheetEl.classList.toggle("drawing", drawing); sizeInk();
  };
  document.querySelectorAll("[data-pen]").forEach(function (b) {
    b.onclick = function () {
      erasing = b.dataset.pen === "eraser"; if (!erasing) pen = b.dataset.pen;
      document.querySelectorAll("[data-pen]").forEach(function (x) { x.setAttribute("aria-pressed", String(x === b)); });
    };
  });
  function at(e) { var r = cv.getBoundingClientRect(); return [Math.round(e.clientX - r.left), Math.round(e.clientY - r.top)]; }
  function eraseAt(p) {
    if (!page) return; var before = page.ink.length;
    page.ink = page.ink.filter(function (s) { for (var i = 0; i < s.p.length; i += 2) if (Math.abs(s.p[i] - p[0]) < 12 && Math.abs(s.p[i + 1] - p[1]) < 12) return false; return true; });
    if (page.ink.length !== before) { redrawInk(); saver.poke(); }
  }
  cv.addEventListener("pointerdown", function (e) {
    if (!drawing || !page) return; e.preventDefault(); cv.setPointerCapture(e.pointerId);
    var p = at(e);
    if (erasing) { stroke = { erase: true }; eraseAt(p); return; }
    var w = +$("penSize").value; if (e.pointerType === "pen" && e.pressure) w = Math.max(1, w * (0.5 + e.pressure));
    stroke = { c: pen === "hl" ? "#facc15" : pen, w: Math.round(w * 10) / 10, hl: pen === "hl", p: p.slice() };
  });
  cv.addEventListener("pointermove", function (e) {
    if (!stroke) return; var p = at(e);
    if (stroke.erase) { eraseAt(p); return; }
    var L = stroke.p.length; if (Math.abs(stroke.p[L - 2] - p[0]) + Math.abs(stroke.p[L - 1] - p[1]) < 2) return;
    if (L > 4000) return;
    stroke.p.push(p[0], p[1]);
    g.save(); g.lineCap = "round"; g.lineJoin = "round";
    if (stroke.hl) { g.globalAlpha = .35; g.strokeStyle = "#facc15"; g.lineWidth = stroke.w * 4; } else { g.strokeStyle = stroke.c; g.lineWidth = stroke.w; }
    g.beginPath(); g.moveTo(stroke.p[L - 2], stroke.p[L - 1]); g.lineTo(p[0], p[1]); g.stroke(); g.restore();
  });
  function endStroke() {
    if (stroke && !stroke.erase && page) { if (page.ink.length < 3000) page.ink.push(stroke); page.updated = now(); if (stroke.hl) redrawInk(); saver.poke(); }
    stroke = null;
  }
  cv.addEventListener("pointerup", endStroke); cv.addEventListener("pointercancel", endStroke);
  $("inkClear").onclick = function () {
    if (!page || !page.ink.length) return; var b = this;
    if (!b.dataset.armed) { b.dataset.armed = 1; b.textContent = "Clear it all?"; setTimeout(function () { delete b.dataset.armed; b.textContent = "Clear drawing"; }, 2500); return; }
    page.ink = []; redrawInk(); saver.poke(); delete b.dataset.armed; b.textContent = "Clear drawing";
  };
  window.addEventListener("resize", function () { sizeInk(); });

  /* ---------- export ---------- */
  var exBtn = $("exportBtn"), exMenu = $("exportMenu");
  exBtn.onclick = function (e) { e.stopPropagation(); exMenu.hidden = !exMenu.hidden; };
  document.addEventListener("click", function (e) { if (!exMenu.hidden && !exMenu.contains(e.target)) exMenu.hidden = true; });
  function inkImage(p) {
    if (!p || !p.ink.length) return "";
    var maxX = 0, maxY = 0; p.ink.forEach(function (s) { for (var i = 0; i < s.p.length; i += 2) { maxX = Math.max(maxX, s.p[i]); maxY = Math.max(maxY, s.p[i + 1]); } });
    var c = document.createElement("canvas"); c.width = maxX + 20; c.height = maxY + 20; var keep = g; g = c.getContext("2d");
    g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); p.ink.forEach(paint); g = keep;
    return '<p><img src="' + c.toDataURL("image/png") + '" alt="Drawing"></p>';
  }
  exMenu.onclick = function (e) {
    var b = e.target.closest("[data-ex]"); if (!b) return; exMenu.hidden = true; flush();
    if (b.dataset.ex === "pdf") { window.print(); return; }
    if (b.dataset.ex === "docx") {
      if (!page) return;
      var box = document.createElement("div"); box.innerHTML = "<h1>" + Office.esc(page.title || "Untitled page") + "</h1>" + page.html + inkImage(page);
      box.querySelectorAll("ul.todo > li").forEach(function (li) { li.insertBefore(document.createTextNode(li.classList.contains("done") ? "☑ " : "☐ "), li.firstChild); });
      document.body.appendChild(box); box.style.position = "fixed"; box.style.left = "-9999px";
      Docx.build(box, page.title || "Notebook page").then(function (blob) { box.remove(); Office.download(blob, Office.safeName(page.title || "Notebook page", ".docx")); })
        .catch(function (er) { box.remove(); Office.toast("Couldn't make the Word file: " + (er.message || er), true); });
      return;
    }
    var html = "<!DOCTYPE html><html lang=\"en-GB\"><head><meta charset=\"utf-8\"><title>" + Office.esc(titleIn.value) + "</title><style>body{font:15px/1.6 system-ui,sans-serif;max-width:780px;margin:40px auto;padding:0 20px}h2.sec{border-bottom:3px solid;padding-bottom:4px;margin-top:2em}img{max-width:100%}ul.todo{list-style:none}ul.todo li.done{text-decoration:line-through;color:#777}ul.todo li::before{content:'☐ '}ul.todo li.done::before{content:'☑ '}</style></head><body><h1>" + Office.esc(titleIn.value) + "</h1>";
    nb.sections.forEach(function (s) {
      html += '<h2 class="sec" style="border-color:' + s.color + '">' + Office.esc(s.name) + "</h2>";
      s.pages.forEach(function (p) { html += "<h3>" + Office.esc(p.title || "Untitled page") + "</h3>" + p.html + inkImage(p); });
    });
    Office.download(new Blob([html + "</body></html>"], { type: "text/html;charset=utf-8" }), Office.safeName(titleIn.value || "Notebook", ".html"));
  };

  /* ---------- files ---------- */
  function use(body, keepCur) {
    nb = clean(body); sec = nb.sections[0]; var p = sec.pages[0] || null;
    if (keepCur) { try { var c = JSON.parse(localStorage.getItem("hl.nb.cur") || "{}");
      nb.sections.forEach(function (s) { if (s.id === c.s) { sec = s; p = s.pages[0] || null; s.pages.forEach(function (x) { if (x.id === c.p) p = x; }); } }); } catch (e) {} }
    openPage(p); drawSecs(); drawPages();
  }
  function load(f) {
    file = { id: f.id, title: f.title || "My Notebook" }; titleIn.value = file.title; use(f.body, true);
    history.replaceState(null, "", "?id=" + encodeURIComponent(f.id));
    try { localStorage.setItem("hl.nb.last", f.id); } catch (e) {}
    $("status").textContent = Office.isCloud() ? "Saved to your account" : "Saved on this device";
  }
  function fresh() {
    file = { id: Office.uuid(), title: "My Notebook" }; titleIn.value = file.title; use(null);
    history.replaceState(null, "", "?id=" + encodeURIComponent(file.id));
    try { localStorage.setItem("hl.nb.last", file.id); } catch (e) {}
    saver.now();
  }
  function openId(id) {
    $("status").textContent = "Opening…";
    return Office.get(id).then(function (f) { if (f) load(f); else fresh(); })
      .catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); fresh(); });
  }
  use(null);
  Office.init("note").then(function (st) {
    var q = new URLSearchParams(location.search).get("id"), last = null;
    try { last = localStorage.getItem("hl.nb.last"); } catch (e) {}
    if (!st.user) Office.toast("Not signed in - your notes are saved in this browser.");
    if (q || last) return openId(q || last);
    return Office.list().then(function (rows) { if (rows.length) return openId(rows[0].id); fresh(); });
  });
  $("filesBtn").onclick = function () {
    flush(); saver.now();
    Office.filesDialog({
      heading: "Your notebooks", noun: "notebook", icon: "📓", current: file.id, accept: ".docx,.txt,.md,.html",
      onOpen: openId, onNew: fresh, onDeletedCurrent: fresh,
      onImport: function (f) { // a file becomes a new page in this notebook
        var name = f.name.replace(/\.[^.]+$/, "");
        function addPg(html) { flush(); var p = newPage(name, Office.cleanHTML(html)); sec.pages.unshift(p); openPage(p); drawSecs(); drawPages(); saver.poke(); Office.toast("Added " + f.name + " as a page"); }
        if (/\.docx$/i.test(f.name)) { Docx.read(f).then(addPg).catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); }); return; }
        var r = new FileReader(); r.onload = function () { var t = String(r.result); addPg(/\.html?$/i.test(f.name) ? t : t.split(/\r?\n/).map(function (l) { return "<p>" + (Office.esc(l) || "<br>") + "</p>"; }).join("")); };
        r.readAsText(f);
      }
    });
  };
  Office.themeToggle($("themeBtn"));
  titleIn.addEventListener("input", saver.poke);
  window.addEventListener("beforeunload", function () { flush(); saver.now(); });
})();
