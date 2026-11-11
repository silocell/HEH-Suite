/* Launchpad Write - the editor, laid out like Word: a File menu, ribbon
   tabs (Home, Insert, Layout, Review, View) and Table/Picture tabs that
   appear when you're in a table or pick a picture. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var page = $("page"), titleIn = $("title"), statusEl = $("status");
  var MARGINS = { normal: 96, narrow: 48, moderate: 72, wide: 144 };       // px at 96 dpi
  var SIZES = { A4: [794, 1123], Letter: [816, 1056], A5: [559, 794] };
  var DEF_LAYOUT = { margins: "normal", orient: "portrait", size: "A4", cols: 1, ls: 1.5, pc: "#ffffff" };
  var file = { id: null, title: "Untitled document", body: { html: "" } };
  var layout = Object.assign({}, DEF_LAYOUT), comments = {};
  function cleanDoc() {
    var c = page.cloneNode(true);
    c.querySelectorAll("mark.find").forEach(function (m) { m.replaceWith(document.createTextNode(m.textContent)); });
    c.querySelectorAll("img.picked").forEach(function (i) { i.classList.remove("picked"); if (!i.className) i.removeAttribute("class"); });
    c.querySelectorAll("span.cmt.now").forEach(function (s) { s.classList.remove("now"); });
    return c;
  }
  var saver = Office.autosaver(statusEl, function () {
    file.title = titleIn.value.trim() || "Untitled document";
    // comments whose words were deleted go too
    var live = {}; page.querySelectorAll("span.cmt[data-c]").forEach(function (s) { live[s.dataset.c] = 1; });
    Object.keys(comments).forEach(function (k) { if (!live[k]) delete comments[k]; });
    file.body = { html: cleanDoc().innerHTML, zoom: +$("zoom").value, layout: layout, comments: comments };
    return file;
  }, {
    /* working together: someone else's change came in */
    html: true, surface: page, local: ["zoom"],
    apply: function (b, title) {
      if (document.activeElement !== titleIn) titleIn.value = title || "Untitled document";
      Office.patchBlocks(page, Office.cleanHTML(b.html || "") || "<p><br></p>");
      if (JSON.stringify(b.layout || {}) !== JSON.stringify(layout)) useLayout(b.layout);
      if (JSON.stringify(b.comments || {}) !== JSON.stringify(comments)) useComments(b.comments);
      fixToc(); count(); navSoon(); pages();
    }
  });

  try { document.execCommand("styleWithCSS", false, true); } catch (e) {}
  try { document.execCommand("defaultParagraphSeparator", false, "p"); } catch (e) {}

  /* ================= tabs ================= */
  var curTab = "home";
  function showTab(t) {
    curTab = t;
    document.querySelectorAll(".tabsbar [data-tab]").forEach(function (b) { b.setAttribute("aria-selected", String(b.dataset.tab === t)); });
    document.querySelectorAll(".ribbon[data-panel]").forEach(function (p) { p.hidden = p.dataset.panel !== t; });
    closeDds();
  }
  document.querySelectorAll(".tabsbar [data-tab]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.onclick = function () { showTab(b.dataset.tab); };
  });

  /* drop-downs on the ribbon */
  function closeDds(except) { document.querySelectorAll(".dd").forEach(function (d) { if (d !== except) d.hidden = true; }); }
  function dd(btnId, ddId) {
    var b = $(btnId), d = $(ddId);
    b.addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
    b.onclick = function (e) { e.stopPropagation(); closeDds(d); d.hidden = !d.hidden; };
    d.addEventListener("mousedown", function (e) { if (e.target.tagName !== "INPUT") e.preventDefault(); });
    return d;
  }
  document.addEventListener("click", function (e) { if (!e.target.closest(".dd")) closeDds(); });

  /* ================= opening and starting files ================= */
  function applyLayout() {
    var sz = SIZES[layout.size] || SIZES.A4, w = sz[0], h = sz[1];
    if (layout.orient === "landscape") { var t = w; w = h; h = t; }
    var r = document.documentElement.style;
    r.setProperty("--pw", w + "px"); r.setProperty("--ph", h + "px");
    r.setProperty("--pm", (MARGINS[layout.margins] || 96) + "px"); r.setProperty("--ls", layout.ls);
    r.setProperty("--paper", layout.pc);
    page.classList.toggle("cols2", +layout.cols === 2); page.classList.toggle("cols3", +layout.cols === 3);
    var ps = document.getElementById("printSize") || document.head.appendChild(Object.assign(document.createElement("style"), { id: "printSize" }));
    ps.textContent = "@media print{@page{size:" + layout.size + " " + layout.orient + ";margin:0}}";
    document.querySelectorAll("[data-mg]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.mg === layout.margins)); });
    document.querySelectorAll("[data-or]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.or === layout.orient)); });
    document.querySelectorAll("[data-sz]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.sz === layout.size)); });
    document.querySelectorAll("[data-cols]").forEach(function (b) { b.setAttribute("aria-pressed", String(+b.dataset.cols === +layout.cols)); });
    document.querySelectorAll("[data-ls]").forEach(function (b) { b.setAttribute("aria-pressed", String(+b.dataset.ls === +layout.ls)); });
    document.querySelectorAll("[data-pc]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.pc === layout.pc)); });
    pages();
  }
  function useLayout(l) {
    l = l || {};
    layout = {
      margins: MARGINS[l.margins] ? l.margins : "normal", orient: l.orient === "landscape" ? "landscape" : "portrait",
      size: SIZES[l.size] ? l.size : "A4", cols: [1, 2, 3].indexOf(+l.cols) >= 0 ? +l.cols : 1,
      ls: [1, 1.15, 1.5, 2].indexOf(+l.ls) >= 0 ? +l.ls : 1.5, pc: /^#[0-9a-f]{6}$/i.test(l.pc || "") ? l.pc : "#ffffff"
    };
    applyLayout();
  }
  function useComments(c) {
    comments = {};
    if (c && typeof c === "object") Object.keys(c).slice(0, 500).forEach(function (k) {
      if (/^[a-z0-9]{1,16}$/.test(k) && c[k]) comments[k] = { text: String(c[k].text || "").slice(0, 2000), at: String(c[k].at || "") };
    });
    drawComments();
  }
  function load(f) {
    file = { id: f.id, title: f.title || "Untitled document", body: f.body || {} };
    titleIn.value = file.title; titleIn.dataset.auto = "";
    page.innerHTML = Office.cleanHTML(file.body.html || "");
    useLayout(file.body.layout); useComments(file.body.comments);
    setZoom(file.body.zoom || 100);
    try { localStorage.setItem("hl.write.last", f.id); } catch (e) {}
    history.replaceState(null, "", "?id=" + encodeURIComponent(f.id));
    fixToc(); count(); navList(); statusEl.textContent = Office.isCloud() ? "Saved to your account" : "Saved on this device";
    saver.opened();
  }
  function fresh(title, html, lay) {
    file = { id: Office.uuid(), title: title || "Untitled document", body: { html: html || "" } };
    titleIn.value = file.title; titleIn.dataset.auto = ""; page.innerHTML = html || "<p><br></p>";
    useLayout(lay); useComments(null); count(); navList();
    history.replaceState(null, "", "?id=" + encodeURIComponent(file.id));
    try { localStorage.setItem("hl.write.last", file.id); } catch (e) {}
    saver.now(); saver.opened(); page.focus();
  }
  function openId(id) {
    statusEl.textContent = "Opening…";
    return Office.get(id).then(function (f) { if (f) load(f); else { Office.toast("That document has been deleted, or it isn't shared with you. Here's a new one.", true); fresh(); } })
      .catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); fresh(); });
  }
  applyLayout();
  Office.setImporter(function (f) { importFile(f); });
  Office.init("doc").then(function (st) {
    var q = new URLSearchParams(location.search).get("id"), last = null;
    try { last = localStorage.getItem("hl.write.last"); } catch (e) {}
    if (!st.user && !st.demo) Office.toast("Not signed in - your work is saved in this browser.");
    if (q || last) return openId(q || last);
    return Office.list().then(function (rows) { if (rows.length) return openId(rows[0].id); fresh(); });
  });

  var ACCEPT = ".docx,.odt,.rtf,.txt,.html,.htm,.md,.markdown";
  function filesWin() {
    saver.now();
    Office.filesDialog({
      heading: "Your documents", noun: "document", icon: "📘", current: file.id, accept: ACCEPT,
      onOpen: openId, onNew: function () { fresh(); }, onDeletedCurrent: function () { fresh(); }, onImport: importFile
    });
  }
  $("filesBtn").onclick = filesWin;
  Office.themeToggle($("themeBtn"));
  titleIn.addEventListener("input", saver.poke);
  titleIn.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); page.focus(); } });

  /* ================= typing ================= */
  /* Like Word: a document nobody has named takes its name from the first line. */
  var DEFAULT_T = "Untitled document";
  function autoName() {
    if (titleIn.value.trim() && titleIn.value !== DEFAULT_T && titleIn.dataset.auto !== "1") return;
    var first = ""; for (var n = page.firstElementChild; n && !first; n = n.nextElementSibling) first = (n.textContent || "").replace(/\s+/g, " ").trim();
    var name = first.slice(0, 60).replace(/[.,;:!?\s]+$/, "");
    if (name.length < 2) { if (titleIn.dataset.auto === "1") { titleIn.value = DEFAULT_T; } return; }
    if (titleIn.value !== name) { titleIn.value = name; titleIn.dataset.auto = "1"; }
  }
  titleIn.addEventListener("input", function () { titleIn.dataset.auto = ""; });
  /* AutoCorrect: the usual slips put right as you type, and a lone "i" made "I".
     Ctrl+Z straight after puts your spelling back. */
  var FIX = { teh:"the", hte:"the", adn:"and", nad:"and", taht:"that", thier:"their", recieve:"receive", recieved:"received", beleive:"believe",
    becuase:"because", becasue:"because", wiht:"with", whcih:"which", alot:"a lot", definately:"definitely", seperate:"separate", untill:"until",
    wich:"which", freind:"friend", freinds:"friends", goverment:"government", enviroment:"environment", tommorow:"tomorrow", tomorow:"tomorrow",
    occured:"occurred", begining:"beginning", wierd:"weird", truely:"truly", basicly:"basically", finaly:"finally", realy:"really",
    dont:"don't", doesnt:"doesn't", didnt:"didn't", cant:"can't", wont:"won't", isnt:"isn't", wasnt:"wasn't", im:"I'm", ive:"I've", youre:"you're", i:"I" };
  function autoCorrect(e) {
    if (!e || e.inputType !== "insertText" || !/^[ .,;:!?]$/.test(e.data || "")) return;
    var sel = getSelection(); if (!sel.rangeCount || !sel.isCollapsed) return;
    var node = sel.anchorNode, off = sel.anchorOffset; if (!node || node.nodeType !== 3) return;
    var before = node.data.slice(0, off - 1), m = before.match(/(^|[^A-Za-z'])([A-Za-z']+)$/); if (!m) return;
    var word = m[2], fix = FIX[word.toLowerCase()] || word;
    // the start of a sentence gets a capital letter, as in Word
    var lead = before.slice(0, before.length - word.length), blk = node.parentElement && node.parentElement.closest("p,li,h1,h2,h3,blockquote,td,th,div");
    var startOfBlock = !lead.trim() && (!blk || blk.textContent.replace(/\u00a0/g, " ").trimStart().indexOf(word) === 0);
    var sentence = /[.!?]\s+$/.test(lead) || startOfBlock;
    if (sentence && /^[a-z][a-z']*$/.test(fix)) fix = fix[0].toUpperCase() + fix.slice(1);
    else if (/^[A-Z]/.test(word) && fix !== "I") fix = fix[0].toUpperCase() + fix.slice(1);
    if (fix === word) return;
    var start = off - 1 - word.length, r = document.createRange();
    r.setStart(node, start); r.setEnd(node, off - 1); sel.removeAllRanges(); sel.addRange(r);
    document.execCommand("insertText", false, fix);
    var after = document.createRange(); after.setStart(sel.anchorNode, Math.min(sel.anchorOffset + 1, sel.anchorNode.length || 0)); after.collapse(true);
    sel.removeAllRanges(); sel.addRange(after);
  }
  page.addEventListener("input", function (e) { autoCorrect(e); tidy(); count(); autoName(); saver.poke(); navSoon(); });
  document.addEventListener("selectionchange", function () { syncTools(); context(); pageOfCaret(); });

  page.addEventListener("paste", function (e) {
    var cd = e.clipboardData; if (!cd) return;
    var imgs = Array.prototype.filter.call(cd.files || [], function (f) { return /^image\//.test(f.type); });
    if (imgs.length) { e.preventDefault(); imgs.forEach(insertImageFile); return; }
    var html = cd.getData("text/html");
    e.preventDefault();
    if (html) document.execCommand("insertHTML", false, Office.cleanHTML(html.replace(/<!--[\s\S]*?-->/g, "")));
    else document.execCommand("insertText", false, cd.getData("text/plain"));
  });
  page.addEventListener("drop", function (e) {
    var fs = e.dataTransfer && e.dataTransfer.files;
    if (fs && fs.length && /^image\//.test(fs[0].type)) { e.preventDefault(); insertImageFile(fs[0]); }
  });
  page.addEventListener("click", function (e) {
    var a = e.target.closest("a");
    if (a && (e.ctrlKey || e.metaKey) && a.href) window.open(a.href, "_blank", "noopener");
    var img = e.target.closest("img"); pickImg(img || null);
    var li = e.target.closest("ul.todo > li");
    if (li && e.clientX - li.getBoundingClientRect().left < 24) { li.classList.toggle("done"); saver.poke(); }
    var toc = e.target.closest("ul.toc > li[data-h]");
    if (toc) { var hs = headings(), h = hs[+toc.getAttribute("data-h")]; if (h) h.scrollIntoView({ block: "start", behavior: "smooth" }); }
    var c = e.target.closest("span.cmt[data-c]"); if (c) { openComments(); focusComment(c.dataset.c); }
  });

  document.addEventListener("keydown", function (e) {
    var mod = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    if (e.key === "Escape" && document.body.classList.contains("focus")) { focusMode(false); return; }
    if (e.key === "Escape" && $("backstage")) { closeBackstage(); return; }
    if (mod && k === "s") { e.preventDefault(); saver.now(); }
    else if (mod && (k === "f" || k === "h")) { e.preventDefault(); openFind(); }
    else if (mod && k === "p") { e.preventDefault(); window.print(); }
    else if (mod && k === "k") { e.preventDefault(); keep(); linkBox(); }
    else if (mod && k === "o") { e.preventDefault(); filesWin(); }
    else if (mod && e.key === "Enter" && document.activeElement === page) { e.preventDefault(); insertBlock('<hr class="pb">'); saver.poke(); }
    else if (mod && e.altKey && /^[123]$/.test(e.key)) { e.preventDefault(); applyStyle("h" + e.key); }
    else if (mod && e.key === "]") { e.preventDefault(); bump(1); }
    else if (mod && e.key === "[") { e.preventDefault(); bump(-1); }
    else if (e.key === "Escape" && !$("findbar").hidden) closeFind();
    else if ((e.key === "Delete" || e.key === "Backspace") && picked && document.activeElement === page) { e.preventDefault(); picked.remove(); pickImg(null); saver.poke(); }
    else if (e.key === "Tab" && document.activeElement === page) {
      e.preventDefault();
      var b = selBlock(); if (b && b.closest("li")) document.execCommand(e.shiftKey ? "outdent" : "indent");
      else if (inCell()) moveCell(e.shiftKey ? -1 : 1);
      else document.execCommand("insertText", false, "    ");
    }
  });

  /* ================= selection helpers ================= */
  var saved = null;   // the selection, kept while a menu or colour picker has focus
  function keep() { var s = getSelection(); if (s.rangeCount && page.contains(s.anchorNode)) saved = s.getRangeAt(0).cloneRange(); }
  function back() { page.focus(); if (saved) { var s = getSelection(); s.removeAllRanges(); try { s.addRange(saved); } catch (e) {} } }
  page.addEventListener("keyup", keep); page.addEventListener("mouseup", keep); page.addEventListener("blur", keep); page.addEventListener("touchend", function () { setTimeout(keep, 0); });
  function selBlock() {
    var s = getSelection(); if (!s.rangeCount) return null;
    var n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentNode;
    while (n && n !== page && !/^(P|H1|H2|H3|BLOCKQUOTE|LI|DIV|TD|TH|PRE)$/.test(n.tagName)) n = n.parentNode;
    return n && n !== page && page.contains(n) ? n : null;
  }
  function topBlock() {
    var s = getSelection(); if (!s.rangeCount) return null;
    var n = s.anchorNode; if (n === page) return page.childNodes[s.anchorOffset - 1] || null;
    while (n && n.parentNode !== page) n = n.parentNode;
    return n;
  }
  function insertBlock(html) {
    back();
    var t = document.createElement("div"); t.innerHTML = html;
    var at = topBlock(), after = document.createElement("p"); after.innerHTML = "<br>";
    var frag = document.createDocumentFragment(); while (t.firstChild) frag.appendChild(t.firstChild); frag.appendChild(after);
    var first = frag.firstChild;
    if (at && at.nodeType === 1 && at.textContent.trim() === "" && !at.querySelector("img,hr,table")) at.replaceWith(frag);
    else if (at && page.contains(at)) at.after(frag); else page.appendChild(frag);
    var cell = first && first.querySelector ? first.querySelector("td,th") : null;
    var r = document.createRange(); r.selectNodeContents(cell || after); r.collapse(true);
    var sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
    count(); pages(); navSoon();
  }
  function tidy() {
    var bad = page.querySelectorAll("p > ul, p > ol, p > table, p > p, p > h1, p > h2, p > h3, p > blockquote, p > div, p > pre");
    if (!bad.length) return;
    var s = getSelection(), keepR = s.rangeCount ? s.getRangeAt(0) : null;
    Array.prototype.slice.call(page.querySelectorAll("p")).forEach(function (p) {
      if (!p.isConnected || !p.querySelector(":scope > ul, :scope > ol, :scope > table, :scope > p, :scope > h1, :scope > h2, :scope > h3, :scope > blockquote, :scope > div, :scope > pre")) return;
      var frag = document.createDocumentFragment(), cur = null;
      Array.prototype.slice.call(p.childNodes).forEach(function (n) {
        if (n.nodeType === 1 && /^(UL|OL|TABLE|P|H1|H2|H3|BLOCKQUOTE|DIV|PRE)$/.test(n.tagName)) { cur = null; frag.appendChild(n); }
        else if (n.nodeType === 3 && !n.nodeValue.trim() && !cur) { }
        else if (n.nodeType === 1 && n.tagName === "BR" && !cur) { }
        else { if (!cur) { cur = document.createElement("p"); frag.appendChild(cur); } cur.appendChild(n); }
      });
      p.replaceWith(frag);
    });
    if (keepR) { try { s.removeAllRanges(); s.addRange(keepR); } catch (e) {} }
  }
  function inCell() { var b = selBlock(); return b && b.closest && b.closest("td,th"); }
  function moveCell(d) {
    var c = inCell(); if (!c) return;
    var cells = Array.prototype.slice.call(c.closest("table").querySelectorAll("td,th")), i = cells.indexOf(c) + d;
    if (i >= cells.length) { tableOp("rowBelow", cells[cells.length - 1]); cells = Array.prototype.slice.call(c.closest("table").querySelectorAll("td,th")); }
    var t = cells[Math.max(0, i)]; if (!t) return;
    var r = document.createRange(); r.selectNodeContents(t); r.collapse(false);
    var s = getSelection(); s.removeAllRanges(); s.addRange(r); saver.poke();
  }

  /* ================= Home ================= */
  document.querySelectorAll("[data-cmd]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.addEventListener("click", function () {
      back();
      var c = b.dataset.cmd;
      if (c === "insertHorizontalRule") insertBlock("<hr>");
      else { try { if (!document.execCommand(c, false, null) && (c === "cut" || c === "copy")) Office.toast("Use " + (/Mac|iP/.test(navigator.platform) ? "⌘" : "Ctrl") + (c === "cut" ? "+X" : "+C") + " to " + c + "."); } catch (er) {} }
      tidy(); page.focus(); syncTools(); saver.poke(); count();
    });
  });
  $("pasteBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("pasteBtn").onclick = function () {
    if (!navigator.clipboard || !navigator.clipboard.readText) { Office.toast("Use Ctrl+V (or ⌘V) to paste."); return; }
    navigator.clipboard.readText().then(function (t) { back(); document.execCommand("insertText", false, t); saver.poke(); })
      .catch(function () { Office.toast("Use Ctrl+V (or ⌘V) to paste."); });
  };

  function applyStyle(v) {
    back();
    var tag = { p: "p", title: "p", subtitle: "p", h1: "h1", h2: "h2", h3: "h3", blockquote: "blockquote", pre: "pre" }[v] || "p";
    document.execCommand("formatBlock", false, "<" + tag + ">");
    var b = selBlock();
    if (b) { b.classList.remove("doc-title", "doc-subtitle"); if (v === "title") b.classList.add("doc-title"); if (v === "subtitle") b.classList.add("doc-subtitle"); if (!b.className) b.removeAttribute("class"); }
    tidy(); syncTools(); saver.poke(); navSoon(); page.focus();
  }
  document.querySelectorAll("[data-style]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
    b.onclick = function () { applyStyle(b.dataset.style); };
  });
  $("font").addEventListener("mousedown", keep);
  $("font").onchange = function () { back(); document.execCommand("fontName", false, this.value); saver.poke(); page.focus(); };
  $("size").addEventListener("mousedown", keep);
  function applySize(pt) {
    back();
    document.execCommand("styleWithCSS", false, false);
    document.execCommand("fontSize", false, "7");
    page.querySelectorAll('font[size="7"]').forEach(function (f) {
      var s = document.createElement("span"); s.style.fontSize = pt + "pt";
      while (f.firstChild) s.appendChild(f.firstChild); f.replaceWith(s);
    });
    page.querySelectorAll("span").forEach(function (s) { if (s.style.fontSize === "xxx-large") s.style.fontSize = pt + "pt"; });
    document.execCommand("styleWithCSS", false, true);
    $("size").value = String(pt); saver.poke(); page.focus(); pages();
  }
  $("size").onchange = function () { applySize(+this.value); };
  var STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
  function curPt() {
    var s = getSelection(), n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentNode;
    if (!n || !page.contains(n)) return 12;
    return Math.round(parseFloat(getComputedStyle(n).fontSize) * 0.75);
  }
  function bump(d) {
    keep(); var now = curPt(), i = 0; while (i < STEPS.length - 1 && STEPS[i] < now) i++;
    if (STEPS[i] > now && d > 0) i--; var n = STEPS[Math.max(0, Math.min(STEPS.length - 1, i + d))]; applySize(n);
  }
  ["growBtn", "shrinkBtn"].forEach(function (id, j) {
    $(id).addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
    $(id).onclick = function () { bump(j ? -1 : 1); };
  });
  function colour(inp, bar, cmd) {
    inp.addEventListener("mousedown", keep); inp.addEventListener("touchstart", keep);
    inp.addEventListener("input", function () { bar.style.background = inp.value; });
    inp.addEventListener("change", function () { back(); document.execCommand(cmd, false, inp.value); saver.poke(); });
  }
  colour($("fg"), $("fgBar"), "foreColor");
  colour($("hl"), $("hlBar"), "hiliteColor");
  dd("caseBtn", "caseDd").onclick = function (e) {
    var b = e.target.closest("[data-case]"); if (!b) return; this.hidden = true; back();
    var s = getSelection(), t = s.toString(); if (!t) { Office.toast("Select some words first."); return; }
    var m = b.dataset.case, o = t.toLowerCase();
    if (m === "upper") o = t.toUpperCase();
    else if (m === "title") o = o.replace(/(^|[\s\-("'“‘])([a-zà-þ])/g, function (x, a, c) { return a + c.toUpperCase(); });
    else if (m === "sentence") o = o.replace(/(^\s*|[.!?]\s+)([a-zà-þ])/g, function (x, a, c) { return a + c.toUpperCase(); }).replace(/\bi\b/g, "I");
    document.execCommand("insertText", false, o); saver.poke();
  };
  function lsPick(e) { var b = e.target.closest("[data-ls]"); if (!b) return; closeDds(); layout.ls = +b.dataset.ls; applyLayout(); saver.poke(); }
  dd("lsBtn", "lsDd").onclick = lsPick;
  $("todoBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("todoBtn").onclick = function () {
    back();
    var s = getSelection(), n = s.anchorNode && (s.anchorNode.nodeType === 3 ? s.anchorNode.parentNode : s.anchorNode);
    var ul = n && n.closest && n.closest("ul");
    if (ul && page.contains(ul)) { ul.classList.toggle("todo"); saver.poke(); return; }
    document.execCommand("insertUnorderedList"); tidy();
    n = s.anchorNode && (s.anchorNode.nodeType === 3 ? s.anchorNode.parentNode : s.anchorNode);
    ul = n && n.closest && n.closest("ul"); if (ul) ul.classList.add("todo");
    saver.poke();
  };

  function syncTools() {
    if (!page.contains(getSelection().anchorNode)) return;
    ["bold", "italic", "underline", "strikeThrough", "superscript", "subscript", "insertUnorderedList", "insertOrderedList",
     "justifyLeft", "justifyCenter", "justifyRight", "justifyFull"].forEach(function (c) {
      document.querySelectorAll('[data-cmd="' + c + '"]').forEach(function (b) {
        try { b.setAttribute("aria-pressed", document.queryCommandState(c) ? "true" : "false"); } catch (e) {}
      });
    });
    var b = selBlock(), v = "p";
    if (b) { var t = b.tagName.toLowerCase(); if (b.classList.contains("doc-title")) v = "title"; else if (b.classList.contains("doc-subtitle")) v = "subtitle"; else if (/^(h[123]|blockquote|pre)$/.test(t)) v = t; }
    document.querySelectorAll("[data-style]").forEach(function (x) { x.setAttribute("aria-pressed", String(x.dataset.style === v)); });
    var pt = curPt(); if (STEPS.indexOf(pt) >= 0 && document.activeElement !== $("size")) $("size").value = String(pt);
  }

  /* ================= Insert ================= */
  $("pbBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("pbBtn").onclick = function () { insertBlock('<hr class="pb">'); saver.poke(); };
  function headings() { return Array.prototype.slice.call(page.querySelectorAll("h1, h2, h3")).filter(function (h) { return !h.closest("ul.toc"); }); }
  $("tocBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("tocBtn").onclick = function () {
    var hs = headings();
    if (!hs.length) { Office.toast("Add some headings first (Home → Styles → Heading 1)."); return; }
    var html = '<ul class="toc"><li class="toch">Contents</li>' + hs.map(function (h, i) {
      var lvl = +h.tagName[1]; return '<li data-h="' + i + '" style="margin-left:' + ((lvl - 1) * 22) + 'px">' + Office.esc(h.textContent.trim() || "(empty heading)") + "</li>";
    }).join("") + "</ul>";
    var old = page.querySelector("ul.toc");
    if (old) { var t = document.createElement("div"); t.innerHTML = html; old.replaceWith(t.firstChild); Office.toast("Contents updated."); }
    else insertBlock(html);
    fixToc(); saver.poke();
  };
  // cleanHTML drops data-h, so put it back from the order of the entries
  function fixToc() { page.querySelectorAll("ul.toc").forEach(function (u) { var i = 0; u.querySelectorAll("li:not(.toch)").forEach(function (li) { li.setAttribute("data-h", i++); }); }); }

  /* table size picker, like Word's grid */
  var tg = $("tgrid"), tdd = dd("tableBtn", "tableDd");
  for (var i = 0; i < 80; i++) { var sp = document.createElement("span"); sp.dataset.r = Math.floor(i / 10) + 1; sp.dataset.c = i % 10 + 1; tg.appendChild(sp); }
  function hoverT(r, c) { tg.querySelectorAll("span").forEach(function (s) { s.classList.toggle("on", +s.dataset.r <= r && +s.dataset.c <= c); }); $("tlbl").textContent = r ? c + " × " + r + " table" : "Pick a size"; }
  tg.addEventListener("pointerover", function (e) { var s = e.target.closest("span"); if (s) hoverT(+s.dataset.r, +s.dataset.c); });
  tg.addEventListener("click", function (e) {
    var s = e.target.closest("span"); if (!s) return; tdd.hidden = true;
    var r = +s.dataset.r, c = +s.dataset.c, html = "<table><tbody>";
    for (var i = 0; i < r; i++) { html += "<tr>"; for (var j = 0; j < c; j++) html += (i === 0 && r > 1 ? "<th><br></th>" : "<td><br></td>"); html += "</tr>"; }
    insertBlock(html + "</tbody></table>"); saver.poke(); hoverT(0, 0);
  });

  $("imgIn").addEventListener("click", keep);
  $("imgIn").onchange = function () { if (this.files[0]) insertImageFile(this.files[0]); this.value = ""; };
  function insertImageFile(f) {
    Office.shrinkImage(f, 1400).then(function (src) {
      back(); document.execCommand("insertHTML", false, '<img src="' + src + '" alt="">'); saver.poke(); setTimeout(pages, 200);
    }).catch(function (e) { Office.toast(e.message, true); });
  }
  $("linkBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("linkBtn").onclick = linkBox;
  function linkBox() {
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>Add a link</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
      '<p class="of-where">Paste a web address. Select some words first to turn them into the link.</p>' +
      '<input id="lnkIn" style="width:100%;height:40px;border:1px solid var(--rule);border-radius:10px;padding:0 12px;background:var(--tool)" placeholder="https://…">' +
      '<div class="of-new" style="margin-top:12px"><button type="button" class="of-btn" id="lnkOk">Add link</button></div>';
    document.body.appendChild(d); d.showModal();
    var inp = d.querySelector("input"); inp.focus();
    function done() { d.close(); d.remove(); }
    d.querySelector(".of-x").onclick = done;
    function ok() {
      var u = inp.value.trim(); if (!u) { done(); return; }
      if (!/^(https?:|mailto:)/i.test(u)) u = "https://" + u;
      done(); back();
      var s = getSelection();
      if (s.isCollapsed) document.execCommand("insertHTML", false, '<a href="' + Office.esc(u) + '">' + Office.esc(u) + "</a>");
      else document.execCommand("createLink", false, u);
      page.querySelectorAll("a").forEach(function (a) { a.target = "_blank"; a.rel = "noopener noreferrer"; });
      saver.poke();
    }
    d.querySelector("#lnkOk").onclick = ok;
    inp.onkeydown = function (e) { if (e.key === "Enter") ok(); };
  }
  $("dateBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("dateBtn").onclick = function () { back(); document.execCommand("insertText", false, new Date().toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric" })); saver.poke(); };
  var SYM = "± × ÷ ≠ ≈ ≤ ≥ √ ∛ π ∞ ° ² ³ ½ ¼ ¾ ‰ ∑ ∫ Δ θ α β γ λ μ σ Ω φ → ← ↑ ↓ ↔ ⇒ ∴ ∵ ∈ ∉ ∩ ∪ ⊂ ∅ ∠ ⊥ ∥ £ € $ ¥ ¢ © ® ™ § ¶ • … – — « » “ ” ‘ ’ ✓ ✗ ★ ☆ ♥ ♪ ☺ ☀ ☁ ♀ ♂ é è ê ë à á â ä ç ñ ö ü ß".split(" ");
  var sg = $("symgrid"); SYM.forEach(function (c) { var b = document.createElement("button"); b.type = "button"; b.textContent = c; b.title = c; sg.appendChild(b); });
  dd("symBtn", "symDd").onclick = function (e) { var b = e.target.closest("button"); if (!b) return; back(); document.execCommand("insertText", false, b.textContent); saver.poke(); keep(); };

  /* ================= Layout ================= */
  dd("mgBtn", "mgDd").onclick = function (e) { var b = e.target.closest("[data-mg]"); if (!b) return; closeDds(); layout.margins = b.dataset.mg; applyLayout(); saver.poke(); };
  dd("orBtn", "orDd").onclick = function (e) { var b = e.target.closest("[data-or]"); if (!b) return; closeDds(); layout.orient = b.dataset.or; applyLayout(); saver.poke(); };
  dd("szBtn", "szDd").onclick = function (e) { var b = e.target.closest("[data-sz]"); if (!b) return; closeDds(); layout.size = b.dataset.sz; applyLayout(); saver.poke(); };
  dd("colBtn", "colDd").onclick = function (e) { var b = e.target.closest("[data-cols]"); if (!b) return; closeDds(); layout.cols = +b.dataset.cols; applyLayout(); saver.poke(); };
  dd("pcBtn", "pcDd").onclick = function (e) { var b = e.target.closest("[data-pc]"); if (!b) return; closeDds(); layout.pc = b.dataset.pc; applyLayout(); saver.poke(); };
  dd("ls2Btn", "ls2Dd").onclick = lsPick;

  /* page lines: where each page would end when printed */
  var pagesT = null;
  function pages() {
    clearTimeout(pagesT);
    pagesT = setTimeout(function () {
      var wrap = $("pagewrap"); wrap.querySelectorAll(".pmark").forEach(function (m) { m.remove(); });
      var ph = parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--ph")) || 1123;
      var h = page.scrollHeight, n = Math.max(1, Math.ceil((h - 4) / ph));
      // manual page breaks start a new page too
      n += page.querySelectorAll("hr.pb").length;
      for (var k = 1; k * ph < h - 4; k++) {
        var m = document.createElement("div"); m.className = "pmark"; m.style.top = k * ph + "px";
        m.innerHTML = "<span>Page " + (k + 1) + "</span>"; wrap.appendChild(m);
      }
      totalPages = n; pageOfCaret();
    }, 150);
  }
  var totalPages = 1;
  function pageOfCaret() {
    var s = getSelection(), n = 1;
    if (s.rangeCount && page.contains(s.anchorNode)) {
      var r = s.getRangeAt(0).getBoundingClientRect(), top = page.getBoundingClientRect().top;
      var ph = parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--ph")) || 1123, z = (+$("zoom").value || 100) / 100;
      if (r.top || r.bottom) n = Math.floor((r.top - top) / (ph * z)) + 1;
      var before = 0; page.querySelectorAll("hr.pb").forEach(function (hr) { if (hr.compareDocumentPosition(s.anchorNode) & Node.DOCUMENT_POSITION_FOLLOWING) before++; });
      n = Math.max(1, n + before);
    }
    $("pgN").textContent = "Page " + Math.min(n, totalPages) + " of " + totalPages;
  }
  new ResizeObserver(pages).observe(page);

  /* ================= Review ================= */
  $("spellBtn").onclick = function () { var on = page.spellcheck = !page.spellcheck; this.setAttribute("aria-pressed", String(on)); page.blur(); page.focus(); };
  function stats() {
    var t = page.innerText || "", words = (t.match(/[A-Za-z0-9À-ɏ'’-]+/g) || []).length;
    return { words: words, chars: t.replace(/\n/g, "").length, charsNoSpace: t.replace(/\s/g, "").length,
      paras: Array.prototype.filter.call(page.querySelectorAll("p,h1,h2,h3,li,blockquote,pre"), function (p) { return p.textContent.trim(); }).length,
      lines: Math.round(page.scrollHeight / (16 * layout.ls)), pages: totalPages, mins: Math.max(1, Math.round(words / 200)) };
  }
  function wordCount() {
    var s = stats(), d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>Word count</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
      '<dl class="info" style="display:grid;grid-template-columns:1fr auto;gap:4px 24px;margin:14px 0">' +
      [["Pages", s.pages], ["Words", s.words], ["Characters (no spaces)", s.charsNoSpace], ["Characters (with spaces)", s.chars], ["Paragraphs", s.paras], ["Reading time", "about " + s.mins + " min"], ["Speaking time", "about " + Math.max(1, Math.round(s.words / 130)) + " min"]]
        .map(function (r) { return "<dt>" + r[0] + "</dt><dd>" + r[1] + "</dd>"; }).join("") + "</dl>";
    document.body.appendChild(d); d.showModal();
    d.onclick = function (e) { if (e.target === d || e.target.closest(".of-x")) { d.close(); d.remove(); } };
  }
  $("wcBtn").onclick = wordCount; $("wc").onclick = wordCount;
  $("speakBtn").onclick = function () {
    var b = this;
    if (!("speechSynthesis" in window)) { Office.toast("This browser can't read aloud.", true); return; }
    if (speechSynthesis.speaking) { speechSynthesis.cancel(); b.setAttribute("aria-pressed", "false"); return; }
    var t = getSelection().toString() || page.innerText; if (!t.trim()) return;
    var u = new SpeechSynthesisUtterance(t.slice(0, 30000)); u.lang = "en-GB"; u.rate = 1;
    var v = speechSynthesis.getVoices().filter(function (x) { return /en-GB/i.test(x.lang); })[0]; if (v) u.voice = v;
    u.onend = u.onerror = function () { b.setAttribute("aria-pressed", "false"); };
    b.setAttribute("aria-pressed", "true"); speechSynthesis.speak(u);
  };
  /* Dictate: the browser's own speech recognition types what you say */
  var rec = null;
  $("dictBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("dictBtn").onclick = function () {
    var b = this, SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (rec) { rec.stop(); return; }
    if (!SR) { Office.toast("This browser can't take dictation - try Chrome, Edge or Safari.", true); return; }
    rec = new SR(); rec.lang = "en-GB"; rec.continuous = true; rec.interimResults = false;
    rec.onresult = function (ev) {
      for (var i = ev.resultIndex; i < ev.results.length; i++) if (ev.results[i].isFinal) {
        var t = ev.results[i][0].transcript.trim(); if (!t) continue;
        back(); t = t.replace(/\b(full stop|period)\b/gi, ".").replace(/\bcomma\b/gi, ",").replace(/\bquestion mark\b/gi, "?").replace(/\bnew line\b/gi, "\n").replace(/\s+([.,?])/g, "$1");
        var before = getSelection().anchorNode && getSelection().anchorNode.textContent || "";
        document.execCommand("insertText", false, (/\S$/.test(before) ? " " : "") + t); keep(); saver.poke(); count();
      }
    };
    rec.onerror = function (ev) { if (ev.error === "not-allowed") Office.toast("Allow the microphone to dictate.", true); };
    rec.onend = function () { rec = null; b.setAttribute("aria-pressed", "false"); Office.toast("Dictation stopped."); };
    try { rec.start(); b.setAttribute("aria-pressed", "true"); Office.toast("Listening… say “full stop” or “comma” for punctuation. Tap again to stop."); }
    catch (er) { rec = null; Office.toast("Couldn't start dictation.", true); }
  };

  /* comments */
  function cid() { return Math.random().toString(36).slice(2, 10); }
  function addComment() {
    back(); var s = getSelection();
    if (!s.rangeCount || s.isCollapsed || !page.contains(s.anchorNode)) { Office.toast("Select the words you want to comment on first."); return; }
    var id = cid(), r = s.getRangeAt(0), span = document.createElement("span"); span.className = "cmt"; span.setAttribute("data-c", id);
    try { r.surroundContents(span); }
    catch (e) { span.appendChild(r.extractContents()); r.insertNode(span); }
    comments[id] = { text: "", at: new Date().toISOString() };
    openComments(); drawComments(); focusComment(id, true); saver.poke();
  }
  ["cmtBtn", "cmtBtn2"].forEach(function (x) { $(x).addEventListener("mousedown", function (e) { e.preventDefault(); keep(); }); $(x).onclick = addComment; });
  function openComments() { $("cPane").hidden = false; $("cpBtn").setAttribute("aria-pressed", "true"); }
  $("cpBtn").onclick = function () { var h = !$("cPane").hidden; $("cPane").hidden = h; this.setAttribute("aria-pressed", String(!h)); drawComments(); };
  $("cClose").onclick = function () { $("cPane").hidden = true; $("cpBtn").setAttribute("aria-pressed", "false"); };
  function drawComments() {
    var box = $("cList"), spans = page.querySelectorAll("span.cmt[data-c]"), seen = {};
    box.innerHTML = "";
    spans.forEach(function (s) {
      var id = s.dataset.c; if (seen[id] || !comments[id]) return; seen[id] = 1;
      var c = comments[id], d = document.createElement("div"); d.className = "cbox"; d.dataset.c = id;
      d.innerHTML = "<q>" + Office.esc(s.textContent.slice(0, 80)) + '</q><textarea aria-label="Comment" placeholder="Type your comment"></textarea>' +
        '<div class="row"><span>' + Office.esc(c.at ? new Date(c.at).toLocaleDateString("en-GB", { day: "numeric", month: "short" }) : "") + '</span><button type="button" data-res="1">✓ Resolve</button></div>';
      d.querySelector("textarea").value = c.text;
      d.querySelector("textarea").oninput = function () { c.text = this.value.slice(0, 2000); saver.poke(); };
      d.querySelector("textarea").onfocus = function () { focusComment(id); };
      d.querySelector("[data-res]").onclick = function () {
        page.querySelectorAll('span.cmt[data-c="' + id + '"]').forEach(function (x) { while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x); x.remove(); });
        delete comments[id]; page.normalize(); drawComments(); saver.poke();
      };
      box.appendChild(d);
    });
    if (!box.children.length) box.innerHTML = '<p style="color:var(--soft);font-size:.88rem">No comments. Select some words and press 💬 Comment.</p>';
  }
  function focusComment(id, type) {
    page.querySelectorAll("span.cmt.now").forEach(function (x) { x.classList.remove("now"); });
    page.querySelectorAll('span.cmt[data-c="' + id + '"]').forEach(function (x) { x.classList.add("now"); });
    document.querySelectorAll(".cbox").forEach(function (b) { b.classList.toggle("now", b.dataset.c === id); });
    var b = document.querySelector('.cbox[data-c="' + id + '"]'); if (b) { b.scrollIntoView({ block: "nearest" }); if (type) b.querySelector("textarea").focus(); }
  }

  /* ================= View ================= */
  $("navBtn").onclick = function () { var h = !$("navPane").hidden; $("navPane").hidden = h; this.setAttribute("aria-pressed", String(!h)); navList(); };
  $("navClose").onclick = function () { $("navPane").hidden = true; $("navBtn").setAttribute("aria-pressed", "false"); };
  var navT = null; function navSoon() { clearTimeout(navT); navT = setTimeout(navList, 400); }
  function navList() {
    if ($("navPane").hidden) return;
    var hs = headings(), ul = $("navList");
    ul.innerHTML = hs.map(function (h, i) { return '<li><button type="button" class="l' + h.tagName[1] + '" data-i="' + i + '">' + Office.esc(h.textContent.trim() || "(empty heading)") + "</button></li>"; }).join("") ||
      '<li style="color:var(--soft);font-size:.88rem;padding:6px">Headings you add will show here, so you can jump between them.</li>';
    ul.onclick = function (e) { var b = e.target.closest("[data-i]"); if (b && hs[+b.dataset.i]) hs[+b.dataset.i].scrollIntoView({ block: "start", behavior: "smooth" }); };
  }
  $("linesBtn").onclick = function () { var on = document.body.classList.toggle("nolines"); this.setAttribute("aria-pressed", String(!on)); };
  function focusMode(on) {
    document.body.classList.toggle("focus", on);
    var x = $("focusExit");
    if (on && !x) { x = document.createElement("button"); x.id = "focusExit"; x.className = "of-btn"; x.textContent = "Exit focus (Esc)"; x.onclick = function () { focusMode(false); }; document.body.appendChild(x); }
    if (!on && x) x.remove();
    if (on && document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(function () {});
    if (!on && document.fullscreenElement) document.exitFullscreen().catch(function () {});
    page.focus();
  }
  $("focusBtn").onclick = function () { focusMode(true); };
  function setZoom(z) {
    z = Math.max(50, Math.min(200, Math.round(z))); $("zoom").value = z; $("zoomN").textContent = z + "%";
    $("zoomer").style.zoom = z / 100;
  }
  $("zoom").oninput = function () { setZoom(+this.value); saver.poke(); };
  $("zIn").onclick = function () { setZoom(+$("zoom").value + 10); saver.poke(); };
  $("zOut").onclick = function () { setZoom(+$("zoom").value - 10); saver.poke(); };
  $("z100").onclick = function () { setZoom(100); saver.poke(); };
  $("zFit").onclick = function () { var w = $("desk").clientWidth - 40, pw = parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--pw")) || 794; setZoom(Math.floor(w / pw * 100)); saver.poke(); };

  /* ================= Table and Picture tabs ================= */
  var picked = null;
  function pickImg(img) {
    if (picked && picked !== img) picked.classList.remove("picked");
    picked = img; if (img) img.classList.add("picked");
    context();
  }
  function context() {
    var cell = inCell(), wasTab = curTab;
    $("tableTab").hidden = !cell; $("picTab").hidden = !picked;
    if (!cell && wasTab === "table") showTab("home");
    if (!picked && wasTab === "picture") showTab("home");
    if (picked && wasTab !== "picture" && context.lastPic !== picked) showTab("picture");
    context.lastPic = picked;
  }
  function tableOp(op, cellArg) {
    var c = cellArg || inCell() || (saved && (function () { var n = saved.startContainer; if (n.nodeType === 3) n = n.parentNode; return n.closest && n.closest("td,th"); })());
    if (!c) { Office.toast("Put the cursor in a table first."); return; }
    var tr = c.closest("tr"), tbl = c.closest("table"), idx = Array.prototype.indexOf.call(tr.children, c);
    function blankCell(tag) { var x = document.createElement(tag || "td"); x.innerHTML = "<br>"; return x; }
    if (op === "rowAbove" || op === "rowBelow") {
      var nr = document.createElement("tr"); Array.prototype.forEach.call(tr.children, function () { nr.appendChild(blankCell()); });
      if (op === "rowAbove") tr.before(nr); else tr.after(nr);
    } else if (op === "colLeft" || op === "colRight") {
      tbl.querySelectorAll("tr").forEach(function (r) { var ref = r.children[idx]; var n = blankCell(ref && ref.tagName === "TH" ? "th" : "td"); if (!ref) r.appendChild(n); else if (op === "colLeft") ref.before(n); else ref.after(n); });
    } else if (op === "delRow") { if (tbl.querySelectorAll("tr").length > 1) tr.remove(); else tbl.remove(); }
    else if (op === "delCol") { var any = false; tbl.querySelectorAll("tr").forEach(function (r) { if (r.children[idx]) r.children[idx].remove(); if (r.children.length) any = true; }); if (!any) tbl.remove(); }
    else if (op === "delTable") tbl.remove();
    else if (op === "header") {
      var first = tbl.querySelector("tr"), isH = first.children[0] && first.children[0].tagName === "TH";
      Array.prototype.slice.call(first.children).forEach(function (x) { var n = document.createElement(isH ? "td" : "th"); while (x.firstChild) n.appendChild(x.firstChild); if (x.getAttribute("style")) n.setAttribute("style", x.getAttribute("style")); x.replaceWith(n); });
    }
    context(); saver.poke(); pages();
  }
  document.querySelectorAll("[data-tb]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
    b.onclick = function () { tableOp(b.dataset.tb); };
  });
  $("cellBg").addEventListener("mousedown", keep);
  $("cellBg").addEventListener("change", function () { back(); var c = inCell(); if (c) { c.style.backgroundColor = this.value; saver.poke(); } });
  document.querySelectorAll("[data-img]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.onclick = function () {
      if (!picked) return;
      if (b.dataset.img === "del") { picked.remove(); pickImg(null); saver.poke(); pages(); return; }
      var full = page.clientWidth - 2 * (parseFloat(getComputedStyle(page).paddingLeft) || 0);
      picked.setAttribute("width", Math.round(full * +b.dataset.img / 100)); saver.poke(); pages();
    };
  });
  document.querySelectorAll("[data-ial]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.onclick = function () {
      if (!picked) return;
      var p = picked.parentNode; if (p === page) { var w = document.createElement("p"); picked.before(w); w.appendChild(picked); p = w; }
      p.style.textAlign = b.dataset.ial; saver.poke();
    };
  });

  /* ================= File (backstage) ================= */
  $("fileTab").onclick = openBackstage;
  function closeBackstage() { var b = $("backstage"); if (b) b.remove(); page.focus(); }
  function openBackstage() {
    saver.now();
    var bs = document.createElement("div"); bs.id = "backstage"; bs.setAttribute("role", "dialog"); bs.setAttribute("aria-label", "File");
    bs.innerHTML = '<nav><button type="button" data-bk="back">← Back</button><button type="button" data-bk="home" aria-selected="true">Home</button><button type="button" data-bk="save">Save a copy</button><button type="button" data-bk="info">Info</button><button type="button" data-bk="print">Print</button></nav><section id="bkBody"></section>';
    document.body.appendChild(bs);
    function show(w) {
      bs.querySelectorAll("[data-bk]").forEach(function (b) { b.setAttribute("aria-selected", String(b.dataset.bk === w)); });
      var s = bs.querySelector("#bkBody");
      if (w === "home") s.innerHTML = "<h2>" + Office.esc(titleIn.value) + '</h2><div class="bgrid">' +
        '<button type="button" data-a="new"><i>📄</i><span><b>New blank document</b><small>Start fresh</small></span></button>' +
        '<button type="button" data-a="tpl-essay"><i>📝</i><span><b>Essay</b><small>Title, intro, paragraphs, conclusion</small></span></button>' +
        '<button type="button" data-a="tpl-report"><i>🔬</i><span><b>Science report</b><small>Aim, method, results, conclusion</small></span></button>' +
        '<button type="button" data-a="tpl-letter"><i>✉️</i><span><b>Letter</b><small>Addresses, date, greeting</small></span></button>' +
        '<button type="button" data-a="open"><i>📂</i><span><b>Open</b><small>Your documents, or a file from this device</small></span></button></div>';
      else if (w === "save") s.innerHTML = '<h2>Save a copy</h2><div class="bgrid">' +
        '<button type="button" data-ex="docx"><i>📘</i><span><b>Word document</b><small>.docx - Word, Pages, Google Docs</small></span></button>' +
        '<button type="button" data-ex="pdf"><i>📄</i><span><b>PDF</b><small>Print, then “Save as PDF”</small></span></button>' +
        '<button type="button" data-ex="odt"><i>📗</i><span><b>OpenDocument</b><small>.odt - LibreOffice and OpenOffice</small></span></button>' +
        '<button type="button" data-ex="rtf"><i>📃</i><span><b>Rich text</b><small>.rtf - opens almost anywhere</small></span></button>' +
        '<button type="button" data-ex="md"><i>Ⓜ️</i><span><b>Markdown</b><small>.md - for websites and notes apps</small></span></button>' +
        '<button type="button" data-ex="txt"><i>📝</i><span><b>Plain text</b><small>.txt - just the words</small></span></button>' +
        '<button type="button" data-ex="html"><i>🌐</i><span><b>Web page</b><small>.html</small></span></button></div>';
      else if (w === "info") { var st = stats(); s.innerHTML = '<h2>Info</h2><dl class="info"><dt>Name</dt><dd>' + Office.esc(titleIn.value) + "</dd><dt>Saved</dt><dd>" + (Office.isCloud() ? "In your Launchpad account" : "In this browser") + "</dd><dt>Size</dt><dd>" + st.pages + " pages · " + st.words + " words</dd><dt>Page</dt><dd>" + layout.size + " · " + layout.orient + " · " + layout.margins + " margins</dd><dt>Comments</dt><dd>" + Object.keys(comments).length + "</dd></dl>"; }
      else if (w === "print") { closeBackstage(); setTimeout(function () { window.print(); }, 100); }
    }
    bs.onclick = function (e) {
      var nb = e.target.closest("[data-bk]");
      if (nb) { if (nb.dataset.bk === "back") closeBackstage(); else show(nb.dataset.bk); return; }
      var a = e.target.closest("[data-a]"), x = e.target.closest("[data-ex]");
      if (a) { closeBackstage(); if (a.dataset.a === "new") fresh(); else if (a.dataset.a === "open") filesWin(); else template(a.dataset.a); }
      if (x) { closeBackstage(); exportAs(x.dataset.ex); }
    };
    show("home");
  }
  function template(t) {
    var T = {
      "tpl-essay": ["Essay", '<p class="doc-title">Essay title</p><p class="doc-subtitle">Your name · ' + new Date().toLocaleDateString("en-GB") + "</p><h1>Introduction</h1><p>Say what the essay is about and what you will argue.</p><h1>Main points</h1><h2>Point one</h2><p>Make the point, give evidence, explain it.</p><h2>Point two</h2><p>Make the point, give evidence, explain it.</p><h1>Conclusion</h1><p>Sum up your argument and answer the question.</p>"],
      "tpl-report": ["Science report", '<p class="doc-title">Experiment title</p><p class="doc-subtitle">Your name · ' + new Date().toLocaleDateString("en-GB") + "</p><h1>Aim</h1><p>What are you trying to find out?</p><h1>Hypothesis</h1><p>What do you think will happen, and why?</p><h1>Equipment</h1><ul><li>Item</li><li>Item</li></ul><h1>Method</h1><ol><li>Step one</li><li>Step two</li></ol><h1>Results</h1><table><tbody><tr><th>Test</th><th>Result</th></tr><tr><td>1</td><td><br></td></tr><tr><td>2</td><td><br></td></tr></tbody></table><h1>Conclusion</h1><p>What did you find out? Was your hypothesis right?</p>"],
      "tpl-letter": ["Letter", '<p style="text-align:right">Your address<br>Town<br>Postcode</p><p style="text-align:right">' + new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" }) + "</p><p>Their name<br>Their address</p><p>Dear …,</p><p>Write your letter here.</p><p>Yours sincerely,</p><p>Your name</p>"]
    }[t];
    if (T) fresh(T[0], Office.cleanHTML(T[1]));
  }

  /* ================= export ================= */
  function toMarkdown(root) {
    function inl(n) {
      var o = "";
      n.childNodes.forEach(function (c) {
        if (c.nodeType === 3) { o += c.nodeValue.replace(/ /g, " ").replace(/([*_`\[\]])/g, "\\$1"); return; }
        if (c.nodeType !== 1) return;
        var t = c.tagName, s = c.style || {};
        if (t === "BR") o += "  \n";
        else if (t === "B" || t === "STRONG" || /bold|[6-9]00/.test(s.fontWeight || "")) o += "**" + inl(c) + "**";
        else if (t === "I" || t === "EM" || s.fontStyle === "italic") o += "*" + inl(c) + "*";
        else if (t === "S" || t === "STRIKE") o += "~~" + inl(c) + "~~";
        else if (t === "CODE") o += "`" + c.textContent + "`";
        else if (t === "A") o += "[" + inl(c) + "](" + c.getAttribute("href") + ")";
        else if (t === "IMG") o += "![" + (c.alt || "picture") + "](picture)";
        else o += inl(c);
      });
      return o;
    }
    var out = [];
    function list(el, depth, ordered) {
      var i = 1;
      Array.prototype.forEach.call(el.children, function (li) {
        if (li.tagName !== "LI") return;
        var mark = el.classList.contains("todo") ? "- [" + (li.classList.contains("done") ? "x" : " ") + "] " : ordered ? (i++) + ". " : "- ";
        var clone = li.cloneNode(true); clone.querySelectorAll("ul,ol").forEach(function (x) { x.remove(); });
        out.push("  ".repeat(depth) + mark + inl(clone).trim());
        Array.prototype.forEach.call(li.children, function (c) { if (c.tagName === "UL" || c.tagName === "OL") list(c, depth + 1, c.tagName === "OL"); });
      });
    }
    Array.prototype.forEach.call(root.children, function (b) {
      var t = b.tagName;
      if (b.classList.contains("doc-title")) out.push("# " + inl(b).trim(), "");
      else if (/^H[1-3]$/.test(t)) out.push("#".repeat(+t[1] + (root.querySelector(".doc-title") ? 1 : 0)) + " " + inl(b).trim(), "");
      else if (t === "UL" && b.classList.contains("toc")) { }
      else if (t === "UL" || t === "OL") { list(b, 0, t === "OL"); out.push(""); }
      else if (t === "BLOCKQUOTE") out.push("> " + inl(b).trim(), "");
      else if (t === "PRE") out.push("```", b.textContent, "```", "");
      else if (t === "HR") out.push(b.classList.contains("pb") ? "<div style=\"page-break-after:always\"></div>" : "---", "");
      else if (t === "TABLE") {
        var rows = Array.prototype.map.call(b.querySelectorAll("tr"), function (tr) { return Array.prototype.map.call(tr.children, function (c) { return inl(c).trim().replace(/\|/g, "\\|") || " "; }); });
        if (rows.length) { out.push("| " + rows[0].join(" | ") + " |", "|" + rows[0].map(function () { return " --- "; }).join("|") + "|"); rows.slice(1).forEach(function (r) { out.push("| " + r.join(" | ") + " |"); }); out.push(""); }
      } else { var x = inl(b).trim(); if (x) out.push(x, ""); }
    });
    return out.join("\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
  }
  function exportAs(kind) {
    var t = titleIn.value.trim() || "Untitled document", doc = cleanDoc();
    doc.querySelectorAll("span.cmt").forEach(function (s) { while (s.firstChild) s.parentNode.insertBefore(s.firstChild, s); s.remove(); });
    if (kind === "pdf") { window.print(); return; }
    if (kind === "txt") { Office.download(new Blob([page.innerText], { type: "text/plain;charset=utf-8" }), Office.safeName(t, ".txt")); return; }
    if (kind === "md") { Office.download(new Blob([toMarkdown(doc)], { type: "text/markdown;charset=utf-8" }), Office.safeName(t, ".md")); return; }
    if (kind === "html") {
      var h = "<!DOCTYPE html><html lang=\"en-GB\"><head><meta charset=\"utf-8\"><title>" + Office.esc(t) + "</title>" +
        "<style>body{font:12pt/" + layout.ls + " Calibri,Carlito,sans-serif;max-width:720px;margin:40px auto;padding:0 20px;background:" + layout.pc + "}img{max-width:100%}table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px 8px}.doc-title{font-size:28pt;font-weight:300}.doc-subtitle{font-size:15pt;color:#5a5a5a}hr.pb{break-after:page;border:0}</style></head><body>" +
        doc.innerHTML + "</body></html>";
      Office.download(new Blob([h], { type: "text/html;charset=utf-8" }), Office.safeName(t, ".html")); return;
    }
    if (kind === "rtf") { Office.download(new Blob([Odf.writeRtf(doc)], { type: "application/rtf" }), Office.safeName(t, ".rtf")); return; }
    if (kind === "odt") {
      document.body.appendChild(doc); doc.style.cssText = "position:fixed;left:-99999px;top:0;width:600px";
      Odf.writeOdt(doc, t, layout).then(function (b) { doc.remove(); Office.download(b, Office.safeName(t, ".odt")); })
        .catch(function (e) { doc.remove(); Office.toast("Couldn't make the file: " + (e.message || e), true); });
      return;
    }
    Office.toast("Making your Word file…");
    document.body.appendChild(doc); doc.style.cssText = "position:fixed;left:-99999px;top:0;width:600px";
    Docx.build(doc, t, layout).then(function (blob) { doc.remove(); Office.download(blob, Office.safeName(t, ".docx")); })
      .catch(function (e) { doc.remove(); Office.toast("Couldn't make the Word file: " + (e.message || e), true); });
  }

  /* ================= opening files ================= */
  function mdToHtml(md) {
    var e = Office.esc, lines = md.replace(/\r\n?/g, "\n").split("\n"), out = [], i = 0;
    function inl(s) {
      s = e(s);
      return s.replace(/!\[([^\]]*)\]\([^)]*\)/g, "[$1]").replace(/\[([^\]]+)\]\((https?:[^)\s]+)\)/g, '<a href="$2">$1</a>')
        .replace(/\*\*([^*]+)\*\*|__([^_]+)__/g, function (m, a, b) { return "<b>" + (a || b) + "</b>"; })
        .replace(/(^|[^*])\*([^*]+)\*/g, "$1<i>$2</i>").replace(/~~([^~]+)~~/g, "<s>$1</s>").replace(/`([^`]+)`/g, "<code>$1</code>").replace(/\\([*_`\[\]])/g, "$1");
    }
    while (i < lines.length) {
      var l = lines[i], m;
      if (/^```/.test(l)) { var code = []; i++; while (i < lines.length && !/^```/.test(lines[i])) code.push(lines[i++]); i++; out.push("<pre>" + e(code.join("\n")) + "</pre>"); continue; }
      if ((m = /^(#{1,6})\s+(.*)$/.exec(l))) { var n = Math.min(3, m[1].length); out.push(m[1].length === 1 && !out.length ? '<p class="doc-title">' + inl(m[2]) + "</p>" : "<h" + n + ">" + inl(m[2]) + "</h" + n + ">"); i++; continue; }
      if (/^\s*([-*+]|\d+\.)\s+/.test(l)) {
        var ordered = /^\s*\d+\./.test(l), todo = /^\s*[-*+]\s+\[[ xX]\]/.test(l), items = [];
        while (i < lines.length && /^\s*([-*+]|\d+\.)\s+/.test(lines[i])) {
          var t = lines[i].replace(/^\s*([-*+]|\d+\.)\s+/, ""), done = /^\[[xX]\]/.test(t); t = t.replace(/^\[[ xX]\]\s*/, "");
          items.push("<li" + (todo && done ? ' class="done"' : "") + ">" + inl(t) + "</li>"); i++;
        }
        out.push("<" + (ordered ? "ol" : "ul") + (todo ? ' class="todo"' : "") + ">" + items.join("") + "</" + (ordered ? "ol" : "ul") + ">"); continue;
      }
      if (/^\s*>/.test(l)) { var q = []; while (i < lines.length && /^\s*>/.test(lines[i])) q.push(lines[i++].replace(/^\s*>\s?/, "")); out.push("<blockquote>" + inl(q.join(" ")) + "</blockquote>"); continue; }
      if (/^\s*\|.*\|\s*$/.test(l)) {
        var rows = []; while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) rows.push(lines[i++]);
        rows = rows.filter(function (r) { return !/^\s*\|[\s:|-]+\|\s*$/.test(r); });
        out.push("<table><tbody>" + rows.map(function (r, j) { return "<tr>" + r.trim().slice(1, -1).split("|").map(function (c) { return (j ? "<td>" : "<th>") + inl(c.trim()) + (j ? "</td>" : "</th>"); }).join("") + "</tr>"; }).join("") + "</tbody></table>"); continue;
      }
      if (/^\s*(---|\*\*\*|___)\s*$/.test(l)) { out.push("<hr>"); i++; continue; }
      if (!l.trim()) { i++; continue; }
      var para = []; while (i < lines.length && lines[i].trim() && !/^(#|```|\s*([-*+]|\d+\.)\s|\s*>|\s*\|)/.test(lines[i])) para.push(lines[i++]);
      if (!para.length) para.push(lines[i++]);
      out.push("<p>" + para.map(inl).join("<br>") + "</p>");
    }
    return out.join("");
  }
  function importFile(f) {
    var name = f.name.replace(/\.[^.]+$/, "") || "Imported document";
    if (f.size > 20 * 1024 * 1024) { Office.toast("That file is too big (over 20 MB).", true); return; }
    if (/\.docx$/i.test(f.name)) {
      Office.toast("Opening " + f.name + "…");
      Docx.read(f).then(function (html) { fresh(name, Office.cleanHTML(html), Docx.lastLayout); fixToc(); Office.toast("Opened " + f.name); })
        .catch(function (e) { Office.toast("Couldn't open that Word file: " + (e.message || e), true); });
      return;
    }
    if (/\.odt$/i.test(f.name)) {
      Office.toast("Opening " + f.name + "…");
      Odf.readOdt(f).then(function (html) { fresh(name, Office.cleanHTML(html)); Office.toast("Opened " + f.name); })
        .catch(function (e) { Office.toast("Couldn't open that file: " + (e.message || e), true); });
      return;
    }
    var r = new FileReader();
    r.onload = function () {
      var t = String(r.result);
      if (/\.rtf$/i.test(f.name)) { try { fresh(name, Office.cleanHTML(Odf.readRtf(t))); } catch (e) { Office.toast("Couldn't open that file: " + e.message, true); } return; }
      if (/\.html?$/i.test(f.name)) fresh(name, Office.cleanHTML(t.replace(/^[\s\S]*<body[^>]*>|<\/body>[\s\S]*$/gi, "")));
      else if (/\.(md|markdown)$/i.test(f.name)) fresh(name, Office.cleanHTML(mdToHtml(t)));
      else fresh(name, t.split(/\r?\n/).map(function (l) { return "<p>" + (Office.esc(l) || "<br>") + "</p>"; }).join(""));
    };
    r.readAsText(f);
  }

  /* ================= counts ================= */
  var countT = null;
  function count() {
    clearTimeout(countT);
    countT = setTimeout(function () {
      var t = page.innerText || "", w = (t.match(/[A-Za-z0-9À-ɏ'’-]+/g) || []).length;
      $("wc").textContent = w + (w === 1 ? " word" : " words");
      $("cc").textContent = t.replace(/\n/g, "").length + " characters";
      $("rt").textContent = w ? "≈ " + Math.max(1, Math.round(w / 200)) + " min read" : "";
    }, 120);
  }

  /* ================= find and replace ================= */
  var hits = [], at = -1;
  function openFind() {
    $("findbar").hidden = false;
    var s = getSelection().toString(); if (s && s.length < 60) $("findIn").value = s;
    $("findIn").focus(); $("findIn").select(); find();
  }
  function closeFind() { clearMarks(); $("findbar").hidden = true; page.focus(); }
  function clearMarks() {
    page.querySelectorAll("mark.find").forEach(function (m) { m.replaceWith(document.createTextNode(m.textContent)); });
    page.normalize(); hits = []; at = -1;
  }
  function find() {
    clearMarks();
    var q = $("findIn").value; if (!q) { $("findN").textContent = ""; return; }
    var low = q.toLowerCase(), walker = document.createTreeWalker(page, NodeFilter.SHOW_TEXT), nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(function (n) {
      var t = n.nodeValue, i = t.toLowerCase().indexOf(low); if (i < 0) return;
      var frag = document.createDocumentFragment(), last = 0;
      while (i >= 0) {
        frag.appendChild(document.createTextNode(t.slice(last, i)));
        var m = document.createElement("mark"); m.className = "find"; m.textContent = t.slice(i, i + q.length);
        frag.appendChild(m); hits.push(m); last = i + q.length; i = t.toLowerCase().indexOf(low, last);
      }
      frag.appendChild(document.createTextNode(t.slice(last)));
      n.replaceWith(frag);
    });
    go(0);
  }
  function go(i) {
    if (!hits.length) { $("findN").textContent = $("findIn").value ? "No matches" : ""; return; }
    if (hits[at]) hits[at].classList.remove("now");
    at = (i + hits.length) % hits.length; hits[at].classList.add("now");
    hits[at].scrollIntoView({ block: "center", behavior: "smooth" });
    $("findN").textContent = (at + 1) + " of " + hits.length;
  }
  $("findBtn").onclick = openFind;
  $("findClose").onclick = closeFind;
  $("findIn").oninput = find;
  $("findIn").onkeydown = function (e) { if (e.key === "Enter") { e.preventDefault(); go(at + (e.shiftKey ? -1 : 1)); } };
  $("findNext").onclick = function () { go(at + 1); };
  $("findPrev").onclick = function () { go(at - 1); };
  $("replOne").onclick = function () {
    if (!hits[at]) return;
    hits[at].replaceWith(document.createTextNode($("replIn").value));
    var keepAt = at; find(); if (hits.length) go(Math.min(keepAt, hits.length - 1)); saver.poke(); count();
  };
  $("replAll").onclick = function () {
    var n = hits.length; if (!n) return;
    hits.forEach(function (m) { m.replaceWith(document.createTextNode($("replIn").value)); });
    page.normalize(); hits = []; at = -1; find();
    Office.toast("Replaced " + n + (n === 1 ? " match" : " matches")); saver.poke(); count();
  };
  window.addEventListener("beforeunload", function () { clearMarks(); saver.now(); });
})();
