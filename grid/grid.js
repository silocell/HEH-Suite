/* Launchpad Grid - the spreadsheet. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var F = Formula, CN = F.colName, RN = F.refName;
  var wrap = $("wrap"), table = $("sheet"), ed = $("ed"), fbar = $("fbar"), nameBox = $("nameBox");
  var titleIn = $("title"), statusEl = $("status");
  var MAXR = 5000, MAXC = 52, DEFW = 96;

  function blankSheet(n) { return { name: "Sheet" + n, cells: {}, w: {}, rows: 100, cols: 26 }; }
  var file = { id: null, title: "Untitled spreadsheet" };
  var book = { sheets: [blankSheet(1)], active: 0 };
  var sh = book.sheets[0];
  var sel = { r: 0, c: 0, r2: 0, c2: 0 };   // r,c = the active cell; r2,c2 = the other corner
  var undo = [], redo = [];
  var saver = Office.autosaver(statusEl, function () {
    file.title = titleIn.value.trim() || "Untitled spreadsheet";
    return { id: file.id, title: file.title, body: { sheets: book.sheets, active: book.active } };
  });

  /* ================= values ================= */
  var vals = {}, busy = {};
  function literal(raw) {
    if (raw == null || raw === "") return null;
    var s = String(raw).trim();
    if (/^(true|false)$/i.test(s)) return s.toUpperCase() === "TRUE";
    var t = s.replace(/^(-?)£/, "$1").replace(/,(?=\d{3}(\D|$))/g, "");
    if (/^-?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return +t;
    if (/^-?(\d+\.?\d*|\.\d+)%$/.test(t)) return parseFloat(t) / 100;
    return String(raw);
  }
  function value(k) {
    if (k in vals) return vals[k];
    var c = sh.cells[k], raw = c && c.v;
    if (raw == null || raw === "") return (vals[k] = null);
    if (typeof raw === "string" && raw[0] === "=" && raw.length > 1) {
      if (busy[k]) return new F.Err("#CIRC!");
      busy[k] = 1;
      var v = F.run(raw.slice(1), ctx);
      delete busy[k];
      return (vals[k] = v);
    }
    return (vals[k] = literal(raw));
  }
  var ctx = { cell: function (r, c) { if (r >= MAXR || c >= MAXC) return new F.Err("#REF!"); return value(RN(r, c)); } };
  function valueIn(si, k) { var keep = sh, kv = vals; sh = book.sheets[si]; vals = {}; var v = value(k); sh = keep; vals = kv; return v; }

  function show(v, nf) {
    if (v === null) return "";
    if (F.isErr(v)) return v.err;
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") {
      var o2 = { minimumFractionDigits: 2, maximumFractionDigits: 2 };
      if (nf === "gbp") return (v < 0 ? "-£" : "£") + Math.abs(v).toLocaleString("en-GB", o2);
      if (nf === "n2") return v.toLocaleString("en-GB", o2);
      if (nf === "int") return Math.round(v).toLocaleString("en-GB");
      if (nf === "pct") { var p = v * 100; return (Number.isInteger(+p.toFixed(6)) ? Math.round(p) : +p.toFixed(2)) + "%"; }
      var a = Math.abs(v);
      if (a !== 0 && (a >= 1e11 || a < 1e-6)) return v.toExponential(4).replace("e+", "E+").replace("e-", "E-");
      return F.fmtNum(+v.toPrecision(11));
    }
    return String(v);
  }

  /* ================= drawing the grid ================= */
  var tds = [], colHeads = [], rowHeads = [], drawn = {};
  function width(c) { return sh.w[CN(c)] || DEFW; }
  function build() {
    var h = '<colgroup><col style="width:46px">';
    for (var c = 0; c < sh.cols; c++) h += '<col style="width:' + width(c) + 'px">';
    h += '</colgroup><thead><tr><th class="corner" title="Select everything"></th>';
    for (c = 0; c < sh.cols; c++) h += '<th data-c="' + c + '">' + CN(c) + '<span class="rs" data-rs="' + c + '"></span></th>';
    h += "</tr></thead><tbody>";
    for (var r = 0; r < sh.rows; r++) {
      h += '<tr><th class="rh" data-r="' + r + '">' + (r + 1) + "</th>";
      for (c = 0; c < sh.cols; c++) h += "<td></td>";
      h += "</tr>";
    }
    table.innerHTML = h + "</tbody>";
    tds = []; rowHeads = []; drawn = {};
    var trs = table.tBodies[0].rows;
    for (r = 0; r < trs.length; r++) { rowHeads.push(trs[r].cells[0]); tds.push(Array.prototype.slice.call(trs[r].cells, 1)); }
    colHeads = Array.prototype.slice.call(table.tHead.rows[0].cells, 1);
    paintAll(); place();
  }
  function paintCell(k) {
    var p = F.parseRef(k); if (!p || !tds[p.r] || !tds[p.r][p.c]) return;
    var td = tds[p.r][p.c], c = sh.cells[k] || {}, v = value(k);
    td.textContent = show(v, c.nf);
    var cls = F.isErr(v) ? "e" : typeof v === "number" ? "n" : typeof v === "boolean" ? "t" : "";
    td.className = cls;
    var st = td.style;
    st.fontWeight = c.b ? "700" : ""; st.fontStyle = c.i ? "italic" : ""; st.textDecoration = c.u ? "underline" : "";
    st.textAlign = c.al ? { l: "left", c: "center", r: "right" }[c.al] : "";
    st.background = c.bg || ""; st.color = c.fg || "";
    td.title = F.isErr(v) ? errHelp(v.err) : "";
    if (c.v || c.b || c.i || c.u || c.bg || c.fg || c.al) drawn[k] = 1; else delete drawn[k];
  }
  function paintAll() {
    vals = {};
    var keys = {}; Object.keys(drawn).forEach(function (k) { keys[k] = 1; }); Object.keys(sh.cells).forEach(function (k) { keys[k] = 1; });
    Object.keys(keys).forEach(paintCell);
    stats();
    $("cellsN").textContent = Object.keys(sh.cells).length ? Object.keys(sh.cells).length + " cells used" : "";
  }
  function errHelp(e) {
    return { "#DIV/0!": "Dividing by zero (or an empty cell)", "#VALUE!": "A word where a number was needed, or a typo in the formula",
      "#REF!": "Points at a cell that isn't there any more", "#NAME?": "That function name isn't known - check the spelling",
      "#N/A": "The lookup didn't find a match", "#NUM!": "The number is too big or not possible", "#CIRC!": "The formula uses its own cell, going round in a circle" }[e] || "";
  }

  /* ================= selection ================= */
  function norm() { return { r1: Math.min(sel.r, sel.r2), r2: Math.max(sel.r, sel.r2), c1: Math.min(sel.c, sel.c2), c2: Math.max(sel.c, sel.c2) }; }
  function rectOf(r1, c1, r2, c2) {
    var a = tds[r1] && tds[r1][c1], b = tds[r2] && tds[r2][c2]; if (!a || !b) return null;
    return { x: a.offsetLeft, y: a.offsetTop, w: b.offsetLeft + b.offsetWidth - a.offsetLeft, h: b.offsetTop + b.offsetHeight - a.offsetTop };
  }
  function setBox(el, r) { if (!r) { el.style.display = "none"; return; } el.style.display = ""; el.style.left = r.x + "px"; el.style.top = r.y + "px"; el.style.width = r.w + "px"; el.style.height = r.h + "px"; }
  var lastHeads = [];
  function place() {
    var n = norm();
    setBox($("selBox"), n.r1 === n.r2 && n.c1 === n.c2 ? null : rectOf(n.r1, n.c1, n.r2, n.c2));
    setBox($("actBox"), rectOf(sel.r, sel.c, sel.r, sel.c));
    var rr = rectOf(n.r1, n.c1, n.r2, n.c2), fh = $("fillH");
    if (rr) { fh.style.display = ""; fh.style.left = (rr.x + rr.w) + "px"; fh.style.top = (rr.y + rr.h) + "px"; } else fh.style.display = "none";
    lastHeads.forEach(function (h) { h.classList.remove("on"); }); lastHeads = [];
    for (var c = n.c1; c <= n.c2; c++) if (colHeads[c]) { colHeads[c].classList.add("on"); lastHeads.push(colHeads[c]); }
    for (var r = n.r1; r <= n.r2; r++) if (rowHeads[r]) { rowHeads[r].classList.add("on"); lastHeads.push(rowHeads[r]); }
    var k = RN(sel.r, sel.c), cell = sh.cells[k] || {};
    nameBox.value = n.r1 === n.r2 && n.c1 === n.c2 ? k : RN(n.r1, n.c1) + ":" + RN(n.r2, n.c2);
    if (document.activeElement !== fbar) fbar.value = cell.v == null ? "" : cell.v;
    document.querySelectorAll("[data-tog]").forEach(function (b) { b.setAttribute("aria-pressed", cell[b.dataset.tog] ? "true" : "false"); });
    document.querySelectorAll("[data-al]").forEach(function (b) { b.setAttribute("aria-pressed", cell.al === b.dataset.al ? "true" : "false"); });
    $("nf").value = cell.nf || "";
    stats();
  }
  function stats() {
    var n = norm(), cnt = 0, sum = 0, nn = 0;
    if (n.r1 === n.r2 && n.c1 === n.c2) { $("stat").innerHTML = ""; return; }
    for (var r = n.r1; r <= Math.min(n.r2, sh.rows - 1); r++) for (var c = n.c1; c <= Math.min(n.c2, sh.cols - 1); c++) {
      var v = value(RN(r, c)); if (v !== null && v !== "") cnt++; if (typeof v === "number") { sum += v; nn++; }
    }
    $("stat").innerHTML = nn ? "Sum <b>" + Office.esc(show(sum)) + "</b> &nbsp; Average <b>" + Office.esc(show(sum / nn)) + "</b> &nbsp; Count <b>" + cnt + "</b>"
      : cnt ? "Count <b>" + cnt + "</b>" : "";
  }
  function go(r, c, extend) {
    r = Math.max(0, Math.min(sh.rows - 1, r)); c = Math.max(0, Math.min(sh.cols - 1, c));
    if (extend) { sel.r2 = r; sel.c2 = c; } else { sel.r = sel.r2 = r; sel.c = sel.c2 = c; }
    if (r > sh.rows - 15 && sh.rows < MAXR) grow(Math.min(MAXR, sh.rows + 100), sh.cols);
    place(); scrollTo(extend ? sel.r2 : sel.r, extend ? sel.c2 : sel.c);
  }
  function scrollTo(r, c) {
    var td = tds[r] && tds[r][c]; if (!td) return;
    var top = td.offsetTop - 26, left = td.offsetLeft - 47;
    if (top < wrap.scrollTop) wrap.scrollTop = top;
    else if (td.offsetTop + td.offsetHeight > wrap.scrollTop + wrap.clientHeight) wrap.scrollTop = td.offsetTop + td.offsetHeight - wrap.clientHeight + 2;
    if (left < wrap.scrollLeft) wrap.scrollLeft = left;
    else if (td.offsetLeft + td.offsetWidth > wrap.scrollLeft + wrap.clientWidth) wrap.scrollLeft = td.offsetLeft + td.offsetWidth - wrap.clientWidth + 2;
  }
  function grow(rows, cols) { if (rows <= sh.rows && cols <= sh.cols) return; sh.rows = Math.max(sh.rows, rows); sh.cols = Math.max(sh.cols, cols); build(); }
  wrap.addEventListener("scroll", function () {
    if (wrap.scrollTop + wrap.clientHeight > wrap.scrollHeight - 200 && sh.rows < MAXR) grow(Math.min(MAXR, sh.rows + 100), sh.cols);
  });

  /* ================= changing cells (with undo) ================= */
  function clone(c) { return c ? JSON.parse(JSON.stringify(c)) : undefined; }
  function empty(c) { return !c || ((c.v == null || c.v === "") && !c.b && !c.i && !c.u && !c.bg && !c.fg && !c.nf && !c.al); }
  /* changes: {key: newCell or undefined} on the current sheet */
  function apply(changes, noUndo) {
    var rec = { s: book.active, ch: [] };
    Object.keys(changes).forEach(function (k) {
      var before = clone(sh.cells[k]), after = changes[k];
      if (empty(after)) after = undefined;
      if (JSON.stringify(before) === JSON.stringify(after)) return;
      rec.ch.push({ k: k, b: before, a: clone(after) });
      if (after) sh.cells[k] = after; else delete sh.cells[k];
    });
    if (!rec.ch.length) return;
    if (!noUndo) { undo.push(rec); if (undo.length > 200) undo.shift(); redo = []; }
    paintAll(); place(); saver.poke();
  }
  function replay(from, to, useAfter) {
    var rec = from.pop(); if (!rec) return;
    if (rec.whole) { // a structural change: swap the whole sheet
      var cur = JSON.parse(JSON.stringify(book.sheets[rec.s]));
      book.sheets[rec.s] = JSON.parse(JSON.stringify(rec.sheet)); to.push({ s: rec.s, whole: true, sheet: cur });
      switchSheet(rec.s); saver.poke(); return;
    }
    if (rec.s !== book.active) switchSheet(rec.s);
    rec.ch.forEach(function (c) { var v = useAfter ? c.a : c.b; if (v) sh.cells[c.k] = clone(v); else delete sh.cells[c.k]; });
    to.push(rec); paintAll(); place(); saver.poke();
  }
  function snapshot() { undo.push({ s: book.active, whole: true, sheet: JSON.parse(JSON.stringify(sh)) }); redo = []; }
  $("undoBtn").onclick = function () { replay(undo, redo, false); };
  $("redoBtn").onclick = function () { replay(redo, undo, true); };

  function setRaw(r, c, raw) {
    var k = RN(r, c), cell = clone(sh.cells[k]) || {};
    raw = String(raw);
    if (raw === "") delete cell.v; else cell.v = raw;
    if (!cell.nf && /^-?(\d+\.?\d*|\.\d+)%$/.test(raw.trim())) cell.nf = "pct";
    if (!cell.nf && /^-?£\s?[\d,]+(\.\d+)?$/.test(raw.trim())) cell.nf = "gbp";
    var ch = {}; ch[k] = cell; apply(ch);
  }
  function eachSel(fn) {
    var n = norm(), ch = {};
    for (var r = n.r1; r <= n.r2; r++) for (var c = n.c1; c <= n.c2; c++) { var k = RN(r, c); ch[k] = fn(clone(sh.cells[k]) || {}, k, r, c); }
    apply(ch);
  }

  /* ================= editing ================= */
  var editing = false, editFrom = "cell", startText = "";
  function startEdit(text, fromBar) {
    var td = tds[sel.r][sel.c]; editing = true; editFrom = fromBar ? "bar" : "cell"; ptRef = null;
    sel.r2 = sel.r; sel.c2 = sel.c; place();
    var cell = sh.cells[RN(sel.r, sel.c)];
    startText = cell && cell.v != null ? String(cell.v) : "";
    if (fromBar) { fbar.focus(); return; }
    ed.hidden = false;
    ed.style.left = td.offsetLeft + "px"; ed.style.top = td.offsetTop + "px";
    ed.style.minWidth = td.offsetWidth + "px"; ed.style.width = Math.max(td.offsetWidth, 160) + "px"; ed.style.height = td.offsetHeight + "px";
    ed.style.fontWeight = td.style.fontWeight; ed.style.fontStyle = td.style.fontStyle;
    ed.value = text != null ? text : startText; fbar.value = ed.value;
    ed.focus(); var L = ed.value.length; ed.setSelectionRange(L, L); hint();
  }
  function stopEdit(commit, move) {
    if (!editing) return;
    var text = editFrom === "bar" ? fbar.value : ed.value;
    editing = false; ptRef = null; ed.hidden = true; $("hint").hidden = true;
    if (commit && text !== startText) setRaw(sel.r, sel.c, autoClose(text));
    else place();
    wrap.focus({ preventScroll: true });
    if (move) go(sel.r + move[0], sel.c + move[1]);
  }
  function autoClose(t) { // =SUM(A1:A3 -> =SUM(A1:A3)
    if (t[0] !== "=") return t;
    var open = (t.match(/\(/g) || []).length - (t.match(/\)/g) || []).length, q = (t.match(/"/g) || []).length;
    if (q % 2) t += '"';
    while (open-- > 0) t += ")";
    return t;
  }
  ed.addEventListener("input", function () { fbar.value = ed.value; ed.style.width = Math.max(parseInt(ed.style.minWidth, 10), ed.value.length * 8 + 20) + "px"; hint(); });
  fbar.addEventListener("focus", function () { if (!editing) startEdit(null, true); });
  fbar.addEventListener("input", function () { if (editing && editFrom === "cell") ed.value = fbar.value; hint(); });
  function keyIn(e, box) {
    if (e.key === "Enter") { e.preventDefault(); stopEdit(true, [e.shiftKey ? -1 : 1, 0]); }
    else if (e.key === "Tab") { e.preventDefault(); stopEdit(true, [0, e.shiftKey ? -1 : 1]); }
    else if (e.key === "Escape") { e.preventDefault(); stopEdit(false); }
    else if (box === ed && /^Arrow/.test(e.key) && !pointing() && ed.dataset.typed === "1") {
      e.preventDefault(); stopEdit(true, { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]);
    }
  }
  ed.addEventListener("keydown", function (e) { keyIn(e, ed); });
  fbar.addEventListener("keydown", function (e) { keyIn(e, fbar); });
  ed.addEventListener("blur", function () { setTimeout(function () { if (editing && editFrom === "cell" && document.activeElement !== ed && document.activeElement !== fbar && !pointerDown) stopEdit(true); }, 0); });
  fbar.addEventListener("blur", function () { setTimeout(function () { if (editing && editFrom === "bar" && document.activeElement !== fbar && !pointerDown) stopEdit(true); }, 0); });

  /* While typing a formula, clicking a cell puts its name in. */
  function curText() { return editFrom === "bar" ? fbar.value : ed.value; }
  function pointing() { var t = curText(); return editing && t[0] === "=" && /[=(,+\-*\/^&<>:]$/.test(t); }
  function insertRef(txt) {
    var box = editFrom === "bar" ? fbar : ed;
    box.value = curText().replace(/(\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?)?$/i, "") + txt;
    if (box === ed) fbar.value = ed.value; else if (editing) ed.value = fbar.value;
    hint();
  }
  var DESC = {
    SUM: "SUM(cells) - adds them up", AVERAGE: "AVERAGE(cells) - the mean", MIN: "MIN(cells) - the smallest", MAX: "MAX(cells) - the biggest",
    COUNT: "COUNT(cells) - how many numbers", COUNTA: "COUNTA(cells) - how many aren't empty", COUNTBLANK: "COUNTBLANK(cells) - how many are empty",
    IF: "IF(test, if true, if false)", IFERROR: "IFERROR(formula, show this instead)", ROUND: "ROUND(number, decimal places)",
    ROUNDUP: "ROUNDUP(number, places)", ROUNDDOWN: "ROUNDDOWN(number, places)", COUNTIF: 'COUNTIF(cells, ">5") - count the ones that match',
    SUMIF: 'SUMIF(cells, "apple", cells to add)', AVERAGEIF: 'AVERAGEIF(cells, ">0", cells to average)',
    VLOOKUP: "VLOOKUP(find this, table, column number, FALSE)", HLOOKUP: "HLOOKUP(find this, table, row number, FALSE)",
    INDEX: "INDEX(cells, row, column)", MATCH: "MATCH(find this, cells, 0) - its position", MEDIAN: "MEDIAN(cells) - the middle value",
    MODE: "MODE(cells) - the most common", STDEV: "STDEV(cells) - standard deviation", PRODUCT: "PRODUCT(cells) - multiplies them",
    SQRT: "SQRT(number) - square root", POWER: "POWER(number, power)", ABS: "ABS(number) - without the minus", INT: "INT(number) - rounds down to a whole number",
    MOD: "MOD(number, divisor) - the remainder", PI: "PI() - 3.14159…", RAND: "RAND() - a random number 0-1", RANDBETWEEN: "RANDBETWEEN(low, high)",
    AND: "AND(test, test, …) - all true?", OR: "OR(test, test, …) - any true?", NOT: "NOT(test)", CONCAT: "CONCAT(text, text, …) - joins text",
    CONCATENATE: "CONCATENATE(text, text, …)", LEN: "LEN(text) - how many characters", UPPER: "UPPER(text)", LOWER: "LOWER(text)", PROPER: "PROPER(text) - Capitals Like This",
    TRIM: "TRIM(text) - removes extra spaces", LEFT: "LEFT(text, how many)", RIGHT: "RIGHT(text, how many)", MID: "MID(text, start, how many)",
    TODAY: "TODAY() - today's date", NOW: "NOW() - the date and time", RANK: "RANK(number, cells) - its place, 1 = biggest"
  };
  function hint() {
    var t = curText(), h = $("hint");
    var m = /([A-Z]+)\((?:[^()]*)$/i.exec(t.slice(0)), w = /=?([A-Z]{2,})$/i.exec(t);
    if (t[0] !== "=") { h.hidden = true; return; }
    var name = m ? m[1].toUpperCase() : null, html = "";
    if (name && DESC[name]) html = "<b>" + Office.esc(DESC[name].split(" - ")[0]) + "</b>" + (DESC[name].indexOf(" - ") > 0 ? " - " + Office.esc(DESC[name].split(" - ")[1]) : "");
    else if (w && !m) {
      var hits = F.names.filter(function (n) { return n.indexOf(w[1].toUpperCase()) === 0; }).slice(0, 5);
      if (hits.length) html = hits.map(function (n) { return "<b>" + n + "</b>"; }).join(" · ");
    }
    if (pointing()) html += (html ? "<br>" : "") + "Tap a cell to put it in the formula";
    if (!html) { h.hidden = true; return; }
    var td = tds[sel.r][sel.c]; h.innerHTML = html; h.hidden = false;
    h.style.left = td.offsetLeft + "px"; h.style.top = (td.offsetTop + td.offsetHeight + 4) + "px";
  }

  /* ================= mouse and touch ================= */
  var pointerDown = false, dragMode = null, fillFrom = null, ptRef = null;
  function cellAt(x, y) {
    var el = document.elementFromPoint(x, y); if (!el) return null;
    var td = el.closest && el.closest("td"); if (!td || !table.contains(td)) return null;
    return { r: td.parentNode.rowIndex - 1, c: td.cellIndex - 1 };
  }
  wrap.addEventListener("pointerdown", function (e) {
    if (e.target === ed) return;
    var rs = e.target.closest(".rs"); if (rs) { resizeStart(e, +rs.dataset.rs); return; }
    if (e.target.id === "fillH") { e.preventDefault(); pointerDown = true; dragMode = "fill"; fillFrom = norm(); wrap.setPointerCapture(e.pointerId); return; }
    var th = e.target.closest("th");
    if (th) {
      if (editing) stopEdit(true);
      e.preventDefault(); pointerDown = true; wrap.focus({ preventScroll: true });
      if (th.classList.contains("corner")) { sel = { r: 0, c: 0, r2: sh.rows - 1, c2: sh.cols - 1 }; place(); pointerDown = false; return; }
      if (th.dataset.c != null) { var c = +th.dataset.c; if (e.shiftKey) sel.c2 = c; else { sel.c = sel.c2 = c; sel.r = 0; } sel.r2 = sh.rows - 1; dragMode = "cols"; }
      else if (th.dataset.r != null) { var r = +th.dataset.r; if (e.shiftKey) sel.r2 = r; else { sel.r = sel.r2 = r; sel.c = 0; } sel.c2 = sh.cols - 1; dragMode = "rows"; }
      place(); wrap.setPointerCapture(e.pointerId); return;
    }
    var at = cellAt(e.clientX, e.clientY); if (!at) return;
    if (editing && pointing() || (editing && ptRef)) {
      e.preventDefault(); pointerDown = true; dragMode = "point"; ptRef = { r: at.r, c: at.c };
      insertRef(RN(at.r, at.c)); wrap.setPointerCapture(e.pointerId); return;
    }
    if (editing) stopEdit(true);
    e.preventDefault(); wrap.focus({ preventScroll: true });
    pointerDown = true; dragMode = "sel";
    if (e.shiftKey) { sel.r2 = at.r; sel.c2 = at.c; } else { sel.r = sel.r2 = at.r; sel.c = sel.c2 = at.c; }
    place();
    if (e.pointerType === "mouse") wrap.setPointerCapture(e.pointerId);
  });
  wrap.addEventListener("pointermove", function (e) {
    if (!pointerDown) return;
    if (dragMode === "resize") { resizeMove(e); return; }
    if (e.pointerType !== "mouse" && dragMode === "sel") return; // a finger drag scrolls
    var at = cellAt(e.clientX, Math.max(e.clientY, wrap.getBoundingClientRect().top + 30)); if (!at) return;
    if (dragMode === "sel") { if (sel.r2 !== at.r || sel.c2 !== at.c) { sel.r2 = at.r; sel.c2 = at.c; place(); } }
    else if (dragMode === "cols") { sel.c2 = at.c; place(); }
    else if (dragMode === "rows") { sel.r2 = at.r; place(); }
    else if (dragMode === "point") { insertRef(RN(Math.min(ptRef.r, at.r), Math.min(ptRef.c, at.c)) + (at.r !== ptRef.r || at.c !== ptRef.c ? ":" + RN(Math.max(ptRef.r, at.r), Math.max(ptRef.c, at.c)) : "")); }
    else if (dragMode === "fill") {
      var f = fillFrom, down = at.r - f.r2, right = at.c - f.c2;
      var n = { r1: f.r1, c1: f.c1, r2: f.r2, c2: f.c2 };
      if (down > 0 && down >= right) n.r2 = at.r; else if (right > 0) n.c2 = at.c;
      setBox($("selBox"), rectOf(n.r1, n.c1, n.r2, n.c2)); fillTo = n;
    }
  });
  var fillTo = null;
  wrap.addEventListener("pointerup", function () {
    if (dragMode === "resize") resizeEnd();
    if (dragMode === "fill" && fillTo) doFill(fillFrom, fillTo);
    if (dragMode === "point") { (editFrom === "bar" ? fbar : ed).focus(); }
    pointerDown = false; dragMode = null; fillTo = null;
  });
  wrap.addEventListener("dblclick", function (e) {
    var at = cellAt(e.clientX, e.clientY); if (!at || e.target.closest("th")) return;
    sel.r = sel.r2 = at.r; sel.c = sel.c2 = at.c; place(); startEdit(null);
  });
  // a tap on an already-selected cell starts editing on touch screens
  var lastTap = null;
  wrap.addEventListener("pointerup", function (e) {
    if (e.pointerType === "mouse" || editing) return;
    var at = cellAt(e.clientX, e.clientY); if (!at) return;
    if (lastTap && lastTap.r === at.r && lastTap.c === at.c && Date.now() - lastTap.t < 600) { startEdit(null); lastTap = null; return; }
    lastTap = { r: at.r, c: at.c, t: Date.now() };
  });

  /* column widths */
  var rz = null;
  function resizeStart(e, c) { e.preventDefault(); e.stopPropagation(); pointerDown = true; dragMode = "resize"; rz = { c: c, x: e.clientX, w: width(c) }; wrap.setPointerCapture(e.pointerId); }
  function resizeMove(e) { var w = Math.max(30, Math.min(600, rz.w + e.clientX - rz.x)); table.querySelectorAll("col")[rz.c + 1].style.width = w + "px"; rz.now = w; place(); }
  function resizeEnd() { if (rz && rz.now) { snapshot(); sh.w[CN(rz.c)] = rz.now; saver.poke(); } rz = null; }

  /* ================= fill handle ================= */
  function doFill(f, t) {
    var ch = {}, vertical = t.r2 > f.r2;
    if (vertical) {
      for (var c = f.c1; c <= f.c2; c++) {
        var src = []; for (var r = f.r1; r <= f.r2; r++) src.push(sh.cells[RN(r, c)]);
        var series = numericSeries(src);
        for (r = f.r2 + 1; r <= t.r2; r++) {
          var i = (r - f.r1) % src.length, from = src[i], cell = clone(from) || {};
          if (series) cell.v = String(+(series.start + series.step * (r - f.r1)).toPrecision(12));
          else if (from && typeof from.v === "string" && from.v[0] === "=") cell.v = "=" + F.shift(from.v.slice(1), r - (f.r1 + i), 0);
          ch[RN(r, c)] = cell;
        }
      }
    } else {
      for (var r2 = f.r1; r2 <= f.r2; r2++) {
        var src2 = []; for (var c2 = f.c1; c2 <= f.c2; c2++) src2.push(sh.cells[RN(r2, c2)]);
        var s2 = numericSeries(src2);
        for (c2 = f.c2 + 1; c2 <= t.c2; c2++) {
          var j = (c2 - f.c1) % src2.length, fr = src2[j], cl = clone(fr) || {};
          if (s2) cl.v = String(+(s2.start + s2.step * (c2 - f.c1)).toPrecision(12));
          else if (fr && typeof fr.v === "string" && fr.v[0] === "=") cl.v = "=" + F.shift(fr.v.slice(1), 0, c2 - (f.c1 + j));
          ch[RN(r2, c2)] = cl;
        }
      }
    }
    apply(ch); sel.r2 = t.r2; sel.c2 = t.c2; sel.r = t.r1; sel.c = t.c1; place();
  }
  function numericSeries(src) { // 1,2 -> 3,4,5 ; a single number just repeats
    if (src.length < 2) return null;
    var n = src.map(function (c) { var v = c && literal(c.v); return typeof v === "number" && !(typeof c.v === "string" && c.v[0] === "=") ? v : null; });
    if (n.some(function (v) { return v === null; })) return null;
    var step = n[1] - n[0];
    for (var i = 2; i < n.length; i++) if (Math.abs(n[i] - n[i - 1] - step) > 1e-9) return null;
    return { start: n[0], step: step };
  }

  /* ================= keyboard ================= */
  wrap.addEventListener("keydown", function (e) {
    if (editing || e.target !== wrap) return;
    var mod = e.ctrlKey || e.metaKey, k = e.key;
    var mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
    if (mv) {
      e.preventDefault();
      var r = e.shiftKey ? sel.r2 : sel.r, c = e.shiftKey ? sel.c2 : sel.c;
      if (mod) { // jump to the edge of the data
        var step = function () { r += mv[0]; c += mv[1]; };
        var filled = function () { return r >= 0 && c >= 0 && r < sh.rows && c < sh.cols && value(RN(r, c)) !== null; };
        step(); if (!filled()) { while (r > 0 && c > 0 && r < sh.rows - 1 && c < sh.cols - 1 && !filled()) step(); }
        else { while (filled()) step(); r -= mv[0]; c -= mv[1]; }
      } else { r += mv[0]; c += mv[1]; }
      go(r, c, e.shiftKey); return;
    }
    if (k === "Enter") { e.preventDefault(); go(sel.r + (e.shiftKey ? -1 : 1), sel.c); return; }
    if (k === "Tab") { e.preventDefault(); go(sel.r, sel.c + (e.shiftKey ? -1 : 1)); return; }
    if (k === "Home") { e.preventDefault(); go(mod ? 0 : sel.r, 0); return; }
    if (k === "PageDown") { e.preventDefault(); go(sel.r + 20, sel.c, e.shiftKey); return; }
    if (k === "PageUp") { e.preventDefault(); go(sel.r - 20, sel.c, e.shiftKey); return; }
    if (k === "F2") { e.preventDefault(); startEdit(null); return; }
    if (k === "Delete" || k === "Backspace") { e.preventDefault(); eachSel(function (c) { delete c.v; return c; }); return; }
    if (k === "Escape") { $("copyBox").hidden = true; clip = null; return; }
    if (mod) {
      var l = k.toLowerCase();
      if (l === "z") { e.preventDefault(); if (e.shiftKey) replay(redo, undo, true); else replay(undo, redo, false); return; }
      if (l === "y") { e.preventDefault(); replay(redo, undo, true); return; }
      if (l === "b" || l === "i" || l === "u") { e.preventDefault(); toggle(l); return; }
      if (l === "a") { e.preventDefault(); sel = { r: 0, c: 0, r2: sh.rows - 1, c2: sh.cols - 1 }; place(); return; }
      if (l === "d") { e.preventDefault(); var n = norm(); if (n.r2 > n.r1) doFill({ r1: n.r1, r2: n.r1, c1: n.c1, c2: n.c2 }, n); return; }
      if (l === "s") { e.preventDefault(); saver.now(); return; }
      if (l === "p") { e.preventDefault(); window.print(); return; }
      return;
    }
    if (k.length === 1 && !e.altKey) { e.preventDefault(); startEdit(k); ed.dataset.typed = "1"; return; }
  });
  ed.addEventListener("focus", function () { ed.dataset.typed = ed.value === startText ? "0" : "1"; });

  nameBox.addEventListener("keydown", function (e) {
    if (e.key !== "Enter") return; e.preventDefault();
    var p = nameBox.value.toUpperCase().replace(/\$/g, "").split(":"), a = F.parseRef(p[0]), b = p[1] ? F.parseRef(p[1]) : a;
    if (!a || !b || a.r >= MAXR || a.c >= MAXC || b.r >= MAXR || b.c >= MAXC) { Office.toast("That isn't a cell name, like B4 or A1:C10", true); return; }
    grow(Math.max(a.r, b.r) + 20, Math.max(a.c, b.c) + 1);
    sel = { r: a.r, c: a.c, r2: b.r, c2: b.c }; place(); scrollTo(a.r, a.c); wrap.focus();
  });
  nameBox.addEventListener("focus", function () { nameBox.select(); });

  /* ================= copy and paste ================= */
  var clip = null;
  function selText() {
    var n = norm(), rows = [];
    for (var r = n.r1; r <= n.r2; r++) {
      var row = []; for (var c = n.c1; c <= n.c2; c++) { var cell = sh.cells[RN(r, c)]; row.push(show(value(RN(r, c)), cell && cell.nf).replace(/[\t\n]/g, " ")); }
      rows.push(row.join("\t"));
    }
    return rows.join("\n");
  }
  function copy(cut) {
    var n = norm(), cells = [];
    for (var r = n.r1; r <= n.r2; r++) { var row = []; for (var c = n.c1; c <= n.c2; c++) row.push(clone(sh.cells[RN(r, c)])); cells.push(row); }
    clip = { n: n, cells: cells, text: selText(), cut: cut, s: book.active };
    var cb = $("copyBox"); cb.hidden = false; setBox(cb, rectOf(n.r1, n.c1, n.r2, n.c2));
    return clip.text;
  }
  document.addEventListener("copy", function (e) { if (editing || document.activeElement !== wrap) return; e.preventDefault(); e.clipboardData.setData("text/plain", copy(false)); });
  document.addEventListener("cut", function (e) { if (editing || document.activeElement !== wrap) return; e.preventDefault(); e.clipboardData.setData("text/plain", copy(true)); });
  document.addEventListener("paste", function (e) {
    if (editing || document.activeElement !== wrap) return;
    e.preventDefault();
    var text = e.clipboardData.getData("text/plain") || "", ch = {};
    if (clip && text.replace(/\r/g, "") === clip.text) {
      var dr = sel.r - clip.n.r1, dc = sel.c - clip.n.c1;
      if (clip.cut && clip.s === book.active) for (var r0 = clip.n.r1; r0 <= clip.n.r2; r0++) for (var c0 = clip.n.c1; c0 <= clip.n.c2; c0++) ch[RN(r0, c0)] = undefined;
      clip.cells.forEach(function (row, i) {
        row.forEach(function (cell, j) {
          var nc = clone(cell);
          if (nc && typeof nc.v === "string" && nc.v[0] === "=" && !clip.cut) nc.v = "=" + F.shift(nc.v.slice(1), dr, dc);
          ch[RN(sel.r + i, sel.c + j)] = nc;
        });
      });
      grow(sel.r + clip.cells.length + 5, Math.min(MAXC, Math.max(sh.cols, sel.c + clip.cells[0].length)));
      apply(ch); sel.r2 = sel.r + clip.cells.length - 1; sel.c2 = sel.c + clip.cells[0].length - 1; place();
      if (clip.cut) { clip = null; $("copyBox").hidden = true; }
      return;
    }
    var rows = parseTable(text);
    if (!rows.length) return;
    rows = rows.slice(0, MAXR - sel.r);
    rows.forEach(function (row, i) { row.slice(0, MAXC - sel.c).forEach(function (v, j) { var k = RN(sel.r + i, sel.c + j), cell = clone(sh.cells[k]) || {}; if (v === "") delete cell.v; else cell.v = v; ch[k] = cell; }); });
    var w = Math.max.apply(null, rows.map(function (r) { return r.length; }));
    grow(sel.r + rows.length + 5, Math.min(MAXC, Math.max(sh.cols, sel.c + w)));
    apply(ch); sel.r2 = sel.r + rows.length - 1; sel.c2 = Math.min(MAXC - 1, sel.c + w - 1); place();
  });
  function parseTable(text) { // tab- or comma-separated, with "quotes"
    text = text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
    if (!text) return [];
    var sep = text.indexOf("\t") >= 0 ? "\t" : (text.indexOf(",") >= 0 && text.indexOf("\n") >= 0 ? "," : "\t");
    var rows = [], row = [], cur = "", q = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
      if (ch === '"' && cur === "") { q = true; continue; }
      if (ch === sep) { row.push(cur); cur = ""; continue; }
      if (ch === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; continue; }
      cur += ch;
    }
    row.push(cur); rows.push(row);
    return rows;
  }

  /* ================= ribbon ================= */
  function toggle(p) { var on = !(sh.cells[RN(sel.r, sel.c)] || {})[p]; eachSel(function (c) { if (on) c[p] = 1; else delete c[p]; return c; }); }
  document.querySelectorAll("[data-tog]").forEach(function (b) { b.onclick = function () { toggle(b.dataset.tog); wrap.focus(); }; });
  document.querySelectorAll("[data-al]").forEach(function (b) {
    b.onclick = function () { var a = b.dataset.al, on = (sh.cells[RN(sel.r, sel.c)] || {}).al !== a; eachSel(function (c) { if (on) c.al = a; else delete c.al; return c; }); wrap.focus(); };
  });
  $("nf").onchange = function () { var v = this.value; eachSel(function (c) { if (v) c.nf = v; else delete c.nf; return c; }); wrap.focus(); };
  function colourPick(inp, bar, prop) {
    inp.addEventListener("input", function () { bar.style.background = inp.value; });
    inp.addEventListener("change", function () { var v = inp.value; eachSel(function (c) { c[prop] = v; return c; }); wrap.focus(); });
  }
  colourPick($("fg"), $("fgBar"), "fg"); colourPick($("bg"), $("bgBar"), "bg");
  $("clearFmt").onclick = function () { eachSel(function (c) { return c.v != null ? { v: c.v } : undefined; }); wrap.focus(); };

  $("sumBtn").onclick = function () {
    var n = norm(), r = n.r1, c = n.c1;
    if (n.r2 > n.r1 || n.c2 > n.c1) { // a range picked: put the total under it
      var ch = {};
      for (c = n.c1; c <= n.c2; c++) { var k = RN(n.r2 + 1, c), cell = clone(sh.cells[k]) || {}; cell.v = "=SUM(" + RN(n.r1, c) + ":" + RN(n.r2, c) + ")"; cell.b = 1; ch[k] = cell; }
      apply(ch); go(n.r2 + 1, n.c1); return;
    }
    var top = r - 1; while (top >= 0 && typeof value(RN(top, c)) === "number") top--;
    var guess = top + 1 < r ? RN(top + 1, c) + ":" + RN(r - 1, c) : "";
    if (!guess) { var left = c - 1; while (left >= 0 && typeof value(RN(r, left)) === "number") left--; if (left + 1 < c) guess = RN(r, left + 1) + ":" + RN(r, c - 1); }
    startEdit("=SUM(" + guess + (guess ? ")" : ""));
  };

  $("fxBtn").onclick = function () {
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>Insert a function</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
      '<input class="q" placeholder="Search - try average, count, lookup" aria-label="Search functions"><ul class="fnlist"></ul>';
    document.body.appendChild(d); d.showModal();
    var ul = d.querySelector("ul"), q = d.querySelector("input");
    function draw() {
      var s = q.value.trim().toLowerCase();
      ul.innerHTML = Object.keys(DESC).filter(function (n) { return !s || n.toLowerCase().indexOf(s) >= 0 || DESC[n].toLowerCase().indexOf(s) >= 0; })
        .map(function (n) { var p = DESC[n].split(" - "); return '<li><button type="button" data-f="' + n + '"><b>' + Office.esc(p[0]) + "</b><small>" + Office.esc(p[1] || "") + "</small></button></li>"; }).join("");
    }
    draw(); q.oninput = draw; q.focus();
    d.onclick = function (e) {
      if (e.target.closest(".of-x") || e.target === d) { d.close(); d.remove(); return; }
      var b = e.target.closest("[data-f]"); if (!b) return;
      d.close(); d.remove(); startEdit("=" + b.dataset.f + "(" + (b.dataset.f === "PI" || b.dataset.f === "TODAY" || b.dataset.f === "NOW" || b.dataset.f === "RAND" ? ")" : ""));
    };
  };

  /* sort the rows of the selection (or the block of data around the cell) */
  function region() {
    var n = norm(); if (n.r2 > n.r1) return n;
    var r1 = sel.r, r2 = sel.r, c1 = sel.c, c2 = sel.c, has = function (r, c) { return r >= 0 && c >= 0 && r < sh.rows && c < sh.cols && sh.cells[RN(r, c)] && sh.cells[RN(r, c)].v != null; };
    var grew = true;
    while (grew) {
      grew = false;
      for (var c = c1; c <= c2; c++) { if (has(r1 - 1, c)) { r1--; grew = true; break; } }
      for (c = c1; c <= c2; c++) { if (has(r2 + 1, c)) { r2++; grew = true; break; } }
      for (var r = r1; r <= r2; r++) { if (has(r, c1 - 1)) { c1--; grew = true; break; } }
      for (r = r1; r <= r2; r++) { if (has(r, c2 + 1)) { c2++; grew = true; break; } }
    }
    return { r1: r1, r2: r2, c1: c1, c2: c2 };
  }
  function sortBy(desc) {
    var g = region(), col = sel.c, head = false;
    if (g.r2 <= g.r1) return;
    var a = value(RN(g.r1, col)), b = value(RN(g.r1 + 1, col));
    if (typeof a === "string" && typeof b === "number") head = true;  // a heading row stays on top
    var start = g.r1 + (head ? 1 : 0), rows = [];
    for (var r = start; r <= g.r2; r++) { var row = []; for (var c = g.c1; c <= g.c2; c++) row.push(clone(sh.cells[RN(r, c)])); rows.push({ key: value(RN(r, col)), row: row }); }
    rows.sort(function (x, y) {
      var p = x.key, q = y.key;
      if (p === null && q === null) return 0; if (p === null) return 1; if (q === null) return -1;
      var t = typeof p === "number" && typeof q === "number" ? p - q : String(p).localeCompare(String(q), "en-GB", { numeric: true, sensitivity: "base" });
      return desc ? -t : t;
    });
    var ch = {};
    rows.forEach(function (o, i) { o.row.forEach(function (cell, j) { ch[RN(start + i, g.c1 + j)] = cell; }); });
    apply(ch); Office.toast("Sorted " + rows.length + " rows by column " + CN(col) + (head ? " (kept the heading row)" : ""));
  }
  $("sortUp").onclick = function () { sortBy(false); wrap.focus(); };
  $("sortDown").onclick = function () { sortBy(true); wrap.focus(); };

  /* rows and columns: insert and delete, fixing formulas that point past them */
  function restructure(axis, at, count) { // count > 0 inserts, < 0 deletes
    snapshot();
    var cells = {}, moveRef = function (r) {
      var pos = axis === "r" ? r.r : r.c;
      if (count < 0 && pos >= at && pos < at - count) return null;
      if (pos >= at) pos += count;
      return axis === "r" ? { r: pos, c: r.c, ar: r.ar, ac: r.ac } : { r: r.r, c: pos, ar: r.ar, ac: r.ac };
    };
    book.sheets.forEach(function (s) {
      if (s !== sh) return;
      Object.keys(s.cells).forEach(function (k) {
        var p = F.parseRef(k), cell = s.cells[k], np = moveRef(p);
        if (!np) return;
        if (typeof cell.v === "string" && cell.v[0] === "=") cell.v = "=" + F.mapRefs(cell.v.slice(1), moveRef);
        cells[RN(np.r, np.c)] = cell;
      });
      s.cells = cells;
      if (axis === "c") {
        var w = {}; Object.keys(s.w).forEach(function (col) { var n = F.colNum(col); if (count < 0 && n >= at && n < at - count) return; w[CN(n >= at ? n + count : n)] = s.w[col]; }); s.w = w;
        s.cols = Math.max(26, Math.min(MAXC, s.cols + Math.max(0, count)));
      } else s.rows = Math.max(100, Math.min(MAXR, s.rows + Math.max(0, count)));
    });
    build(); saver.poke();
  }
  $("addRow").onclick = function () { var n = norm(); restructure("r", n.r1, n.r2 - n.r1 + 1); };
  $("addCol").onclick = function () { var n = norm(); if (sh.cols >= MAXC) { Office.toast("A sheet can have up to " + MAXC + " columns.", true); return; } restructure("c", n.c1, Math.min(MAXC - sh.cols, n.c2 - n.c1 + 1)); };
  $("delRow").onclick = function () { var n = norm(); restructure("r", n.r1, -(n.r2 - n.r1 + 1)); go(n.r1, sel.c); };
  $("delCol").onclick = function () { var n = norm(); restructure("c", n.c1, -(n.c2 - n.c1 + 1)); go(sel.r, n.c1); };

  /* ================= charts ================= */
  $("chartBtn").onclick = function () {
    var n = norm();
    if (n.r1 === n.r2 && n.c1 === n.c2) { var g = region(); if (g.r2 > g.r1 || g.c2 > g.c1) n = g; }
    var data = [];
    for (var r = n.r1; r <= Math.min(n.r2, n.r1 + 60); r++) { var row = []; for (var c = n.c1; c <= Math.min(n.c2, n.c1 + 8); c++) row.push(value(RN(r, c))); data.push(row); }
    var headRow = data[0].some(function (v) { return typeof v === "string"; }) && data.length > 1 && data[1].some(function (v) { return typeof v === "number"; });
    var labelCol = data.slice(headRow ? 1 : 0).every(function (row) { return typeof row[0] !== "number"; }) && data[0].length > 1;
    var heads = headRow ? data[0] : null, body = headRow ? data.slice(1) : data;
    var labels = body.map(function (row, i) { return labelCol ? show(row[0]) : String(i + 1); });
    var series = [];
    for (var j = labelCol ? 1 : 0; j < (data[0] || []).length; j++) series.push({ name: heads ? show(heads[j]) : "Series " + (series.length + 1), v: body.map(function (row) { return typeof row[j] === "number" ? row[j] : 0; }) });
    if (!series.length || !body.length || series.every(function (s) { return s.v.every(function (x) { return x === 0; }); })) {
      Office.toast("Pick some cells with numbers first (labels in the first column).", true); return;
    }
    var d = document.createElement("dialog"); d.className = "of-dialog"; d.style.width = "min(820px, calc(100vw - 32px))";
    d.innerHTML = '<div class="of-dhead"><h2>Chart</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
      '<div class="chartbox"><div class="of-new" style="margin:0"><select id="ctype" style="height:36px;border:1px solid var(--rule);border-radius:9px;background:var(--tool);padding:0 8px">' +
      '<option value="bar">Bar chart</option><option value="line">Line graph</option><option value="pie">Pie chart</option><option value="scatter">Scatter graph</option></select>' +
      '<input id="ctitle" placeholder="Chart title" value="' + Office.esc(titleIn.value) + '" style="flex:1;min-width:160px;height:36px;border:1px solid var(--rule);border-radius:9px;background:var(--tool);padding:0 10px">' +
      '<button type="button" class="of-btn" id="cdl">⬇ Save picture</button></div><canvas width="1520" height="860" aria-label="Chart"></canvas></div>';
    document.body.appendChild(d); d.showModal();
    var cv = d.querySelector("canvas"), draw = function () { drawChart(cv, d.querySelector("#ctype").value, d.querySelector("#ctitle").value, labels, series); };
    d.querySelector("#ctype").onchange = draw; d.querySelector("#ctitle").oninput = draw; draw();
    d.querySelector("#cdl").onclick = function () { cv.toBlob(function (b) { Office.download(b, Office.safeName(d.querySelector("#ctitle").value || "Chart", ".png")); }); };
    d.querySelector(".of-x").onclick = function () { d.close(); d.remove(); };
  };
  var PAL = ["#2563eb", "#16a34a", "#f59e0b", "#dc2626", "#7c3aed", "#0891b2", "#db2777", "#65a30d"];
  function drawChart(cv, type, title, labels, series) {
    var g = cv.getContext("2d"), W = cv.width, H = cv.height;
    g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
    g.fillStyle = "#111"; g.font = "600 40px system-ui, sans-serif"; g.textAlign = "center"; g.fillText(title || "", W / 2, 60);
    g.font = "26px system-ui, sans-serif";
    var nice = function (v) { return Math.abs(v) >= 1000 ? (+v.toFixed(0)).toLocaleString("en-GB") : String(+v.toPrecision(4)); };
    if (type === "pie") {
      var s = series[0], tot = s.v.reduce(function (a, b) { return a + Math.max(0, b); }, 0) || 1, a0 = -Math.PI / 2, cx = W * 0.36, cy = H * 0.55, R = Math.min(W, H) * 0.34;
      s.v.forEach(function (v, i) {
        var a1 = a0 + Math.max(0, v) / tot * Math.PI * 2;
        g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, R, a0, a1); g.closePath(); g.fillStyle = PAL[i % PAL.length]; g.fill();
        g.strokeStyle = "#fff"; g.lineWidth = 4; g.stroke();
        var mid = (a0 + a1) / 2, pct = Math.round(v / tot * 100);
        if (pct >= 4) { g.fillStyle = "#fff"; g.font = "600 28px system-ui"; g.textAlign = "center"; g.fillText(pct + "%", cx + Math.cos(mid) * R * 0.65, cy + Math.sin(mid) * R * 0.65 + 10); }
        a0 = a1;
      });
      g.textAlign = "left"; g.font = "28px system-ui";
      labels.slice(0, 14).forEach(function (l, i) { var y = 150 + i * 46; g.fillStyle = PAL[i % PAL.length]; g.fillRect(W * 0.66, y - 22, 28, 28); g.fillStyle = "#222"; g.fillText((l + "").slice(0, 26) + "  " + nice(s.v[i]), W * 0.66 + 42, y); });
      return;
    }
    var L = 130, Rm = 50, T = 100, B = 120, pw = W - L - Rm, ph = H - T - B;
    var all = [], xs = labels.map(function (l) { return parseFloat(l); });
    series.forEach(function (s) { all.push.apply(all, s.v); });
    var lo = Math.min(0, Math.min.apply(null, all)), hi = Math.max(0, Math.max.apply(null, all)); if (hi === lo) hi = lo + 1;
    var stepRaw = (hi - lo) / 5, mag = Math.pow(10, Math.floor(Math.log10(stepRaw))), step = [1, 2, 2.5, 5, 10].map(function (k) { return k * mag; }).filter(function (k) { return k >= stepRaw; })[0];
    lo = Math.floor(lo / step) * step; hi = Math.ceil(hi / step) * step;
    var Y = function (v) { return T + ph - (v - lo) / (hi - lo) * ph; };
    g.strokeStyle = "#e5e7eb"; g.lineWidth = 2; g.fillStyle = "#555"; g.textAlign = "right"; g.font = "24px system-ui";
    for (var v = lo; v <= hi + step / 2; v += step) { g.beginPath(); g.moveTo(L, Y(v)); g.lineTo(L + pw, Y(v)); g.stroke(); g.fillText(nice(v), L - 14, Y(v) + 8); }
    g.strokeStyle = "#9ca3af"; g.beginPath(); g.moveTo(L, Y(0)); g.lineTo(L + pw, Y(0)); g.stroke();
    var n = labels.length, slot = pw / n;
    var scatter = type === "scatter" && xs.every(function (x) { return isFinite(x); });
    var xlo = scatter ? Math.min.apply(null, xs) : 0, xhi = scatter ? Math.max.apply(null, xs) : 1; if (xhi === xlo) xhi = xlo + 1;
    var X = function (i) { return scatter ? L + (xs[i] - xlo) / (xhi - xlo) * pw : L + slot * i + slot / 2; };
    g.fillStyle = "#444"; g.textAlign = "center"; g.font = "24px system-ui";
    var every = Math.ceil(n / 16);
    labels.forEach(function (l, i) { if (i % every === 0) g.fillText(String(l).slice(0, 12), X(i), T + ph + 40); });
    series.forEach(function (s, si) {
      var col = PAL[si % PAL.length]; g.fillStyle = col; g.strokeStyle = col;
      if (type === "bar") {
        var bw = slot * 0.75 / series.length;
        s.v.forEach(function (v, i) { var x = L + slot * i + slot * 0.125 + bw * si; g.fillRect(x, Math.min(Y(v), Y(0)), bw - 4, Math.abs(Y(v) - Y(0))); });
      } else {
        if (type === "line") { g.lineWidth = 6; g.beginPath(); s.v.forEach(function (v, i) { if (i) g.lineTo(X(i), Y(v)); else g.moveTo(X(i), Y(v)); }); g.stroke(); }
        s.v.forEach(function (v, i) { g.beginPath(); g.arc(X(i), Y(v), type === "scatter" ? 10 : 8, 0, Math.PI * 2); g.fill(); });
      }
    });
    if (series.length > 1 || series[0].name) {
      var lx = L; g.textAlign = "left"; g.font = "26px system-ui";
      series.forEach(function (s, si) { g.fillStyle = PAL[si % PAL.length]; g.fillRect(lx, H - 50, 26, 26); g.fillStyle = "#222"; var t = String(s.name).slice(0, 24); g.fillText(t, lx + 36, H - 28); lx += 70 + g.measureText(t).width; });
    }
  }

  /* ================= sheets (tabs) ================= */
  function switchSheet(i) {
    if (editing) stopEdit(true);
    book.active = i; sh = book.sheets[i]; sel = { r: 0, c: 0, r2: 0, c2: 0 }; $("copyBox").hidden = true;
    build(); tabs(); wrap.scrollTop = wrap.scrollLeft = 0;
  }
  function tabs() {
    var nav = $("tabs");
    nav.innerHTML = book.sheets.map(function (s, i) {
      return '<button type="button" class="tab" role="tab" data-i="' + i + '" aria-selected="' + (i === book.active) + '" title="' + (i === book.active ? "Double-click to rename" : "") + '">' + Office.esc(s.name) + "</button>";
    }).join("") + '<button type="button" class="tab" id="addSheet" title="Add a sheet" aria-label="Add a sheet">＋</button>' +
      (book.sheets.length > 1 ? '<button type="button" class="tab" id="delSheet" title="Delete this sheet" aria-label="Delete this sheet">🗑</button>' : "");
  }
  $("tabs").addEventListener("click", function (e) {
    var t = e.target.closest(".tab"); if (!t) return;
    if (t.id === "addSheet") {
      if (book.sheets.length >= 20) { Office.toast("Up to 20 sheets in one file.", true); return; }
      var n = book.sheets.length + 1; while (book.sheets.some(function (s) { return s.name === "Sheet" + n; })) n++;
      book.sheets.push(blankSheet(n)); switchSheet(book.sheets.length - 1); saver.poke(); return;
    }
    if (t.id === "delSheet") {
      if (!t.dataset.armed) { t.dataset.armed = 1; t.textContent = "Delete " + sh.name + "?"; setTimeout(function () { if (t.isConnected) tabs(); }, 3000); return; }
      book.sheets.splice(book.active, 1); undo = []; redo = []; switchSheet(Math.max(0, book.active - 1)); saver.poke(); return;
    }
    var i = +t.dataset.i; if (i !== book.active) switchSheet(i);
  });
  $("tabs").addEventListener("dblclick", function (e) {
    var t = e.target.closest(".tab[data-i]"); if (!t) return;
    var i = +t.dataset.i, inp = document.createElement("input"); inp.value = book.sheets[i].name; inp.maxLength = 31;
    t.textContent = ""; t.appendChild(inp); inp.focus(); inp.select();
    var done = function (ok) { if (ok && inp.value.trim()) { book.sheets[i].name = inp.value.trim().slice(0, 31); saver.poke(); } tabs(); };
    inp.onkeydown = function (ev) { if (ev.key === "Enter") done(true); if (ev.key === "Escape") done(false); };
    inp.onblur = function () { done(true); };
  });

  /* ================= files ================= */
  function useBook(b) {
    var sheets = (b && b.sheets && b.sheets.length ? b.sheets : [blankSheet(1)]).slice(0, 20).map(function (s, i) {
      return { name: String(s.name || "Sheet" + (i + 1)).slice(0, 31), cells: s.cells && typeof s.cells === "object" ? s.cells : {}, w: s.w || {},
        rows: Math.max(100, Math.min(MAXR, +s.rows || 100)), cols: Math.max(26, Math.min(MAXC, +s.cols || 26)) };
    });
    book = { sheets: sheets, active: Math.min(sheets.length - 1, Math.max(0, +b.active || 0)) };
    undo = []; redo = []; switchSheet(book.active);
  }
  function load(f) {
    file = { id: f.id, title: f.title || "Untitled spreadsheet" }; titleIn.value = file.title;
    useBook(f.body || {});
    history.replaceState(null, "", "?id=" + encodeURIComponent(f.id));
    try { localStorage.setItem("hl.grid.last", f.id); } catch (e) {}
    statusEl.textContent = Office.isCloud() ? "Saved to your account" : "Saved on this device";
  }
  function fresh(title, b) {
    file = { id: Office.uuid(), title: title || "Untitled spreadsheet" }; titleIn.value = file.title;
    useBook(b || {});
    history.replaceState(null, "", "?id=" + encodeURIComponent(file.id));
    try { localStorage.setItem("hl.grid.last", file.id); } catch (e) {}
    saver.now(); wrap.focus();
  }
  function openId(id) {
    statusEl.textContent = "Opening…";
    return Office.get(id).then(function (f) { if (f) load(f); else fresh(); })
      .catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); fresh(); });
  }
  build(); tabs();
  Office.init("sheet").then(function (st) {
    var q = new URLSearchParams(location.search).get("id"), last = null;
    try { last = localStorage.getItem("hl.grid.last"); } catch (e) {}
    if (!st.user) Office.toast("Not signed in - your work is saved in this browser.");
    if (q || last) return openId(q || last);
    return Office.list().then(function (rows) { if (rows.length) return openId(rows[0].id); fresh(); });
  });
  $("filesBtn").onclick = function () {
    saver.now();
    Office.filesDialog({
      heading: "Your spreadsheets", noun: "spreadsheet", icon: "📗", current: file.id, accept: ".xlsx,.csv,.tsv,.txt",
      onOpen: openId, onNew: function () { fresh(); }, onDeletedCurrent: function () { fresh(); }, onImport: importFile
    });
  };
  Office.themeToggle($("themeBtn"));
  titleIn.addEventListener("input", saver.poke);
  titleIn.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); wrap.focus(); } });

  function importFile(f) {
    var name = f.name.replace(/\.[^.]+$/, "") || "Imported spreadsheet";
    if (f.size > 20 * 1024 * 1024) { Office.toast("That file is too big (over 20 MB).", true); return; }
    if (/\.xlsx$/i.test(f.name)) {
      Office.toast("Opening " + f.name + "…");
      Xlsx.read(f).then(function (sheets) { fresh(name, { sheets: sheets, active: 0 }); Office.toast("Opened " + f.name); })
        .catch(function (e) { Office.toast("Couldn't open that Excel file: " + (e.message || e), true); });
      return;
    }
    var r = new FileReader();
    r.onload = function () {
      var rows = parseTable(String(r.result)).slice(0, MAXR), cells = {}, w = 0;
      rows.forEach(function (row, i) { row.slice(0, MAXC).forEach(function (v, j) { if (v !== "") cells[RN(i, j)] = { v: v }; }); w = Math.max(w, row.length); });
      fresh(name, { sheets: [{ name: "Sheet1", cells: cells, w: {}, rows: rows.length + 20, cols: Math.min(MAXC, Math.max(26, w)) }] });
    };
    r.readAsText(f);
  }

  var exBtn = $("exportBtn"), exMenu = $("exportMenu");
  exBtn.onclick = function (e) { e.stopPropagation(); exMenu.hidden = !exMenu.hidden; exBtn.setAttribute("aria-expanded", String(!exMenu.hidden)); };
  document.addEventListener("click", function (e) { if (!exMenu.hidden && !exMenu.contains(e.target)) { exMenu.hidden = true; exBtn.setAttribute("aria-expanded", "false"); } });
  exMenu.onclick = function (e) {
    var b = e.target.closest("[data-ex]"); if (!b) return; exMenu.hidden = true;
    var t = titleIn.value.trim() || "Untitled spreadsheet";
    if (b.dataset.ex === "pdf") { window.print(); return; }
    if (b.dataset.ex === "csv") {
      var maxR = -1, maxC = -1;
      Object.keys(sh.cells).forEach(function (k) { var p = F.parseRef(k); if (p && sh.cells[k].v != null) { maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); } });
      var lines = [];
      for (var r = 0; r <= maxR; r++) {
        var row = [];
        for (var c = 0; c <= maxC; c++) { var cell = sh.cells[RN(r, c)], s = show(value(RN(r, c)), cell && cell.nf); row.push(/[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s); }
        lines.push(row.join(","));
      }
      Office.download(new Blob(["﻿" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" }), Office.safeName(t + (book.sheets.length > 1 ? " - " + sh.name : ""), ".csv"));
      return;
    }
    Office.toast("Making your Excel file…");
    Xlsx.build(book.sheets, valueIn).then(function (blob) { Office.download(blob, Office.safeName(t, ".xlsx")); })
      .catch(function (er) { Office.toast("Couldn't make the Excel file: " + (er.message || er), true); });
  };
  window.addEventListener("beforeunload", function () { if (editing) stopEdit(true); saver.now(); });
  window.addEventListener("resize", place);
})();
