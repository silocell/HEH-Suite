/* Launchpad Write - the editor. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var page = $("page"), titleIn = $("title"), statusEl = $("status");
  var file = { id: null, title: "Untitled document", body: { html: "" } };
  var saver = Office.autosaver(statusEl, function () {
    file.title = titleIn.value.trim() || "Untitled document";
    var c = page.cloneNode(true);
    c.querySelectorAll("mark.find").forEach(function (m) { m.replaceWith(document.createTextNode(m.textContent)); });
    file.body = { html: c.innerHTML, zoom: +$("zoom").value };
    return file;
  });

  try { document.execCommand("styleWithCSS", false, true); } catch (e) {}
  try { document.execCommand("defaultParagraphSeparator", false, "p"); } catch (e) {}

  /* ---------- opening and starting files ---------- */
  function load(f) {
    file = { id: f.id, title: f.title || "Untitled document", body: f.body || {} };
    titleIn.value = file.title;
    page.innerHTML = Office.cleanHTML(file.body.html || "");
    setZoom(file.body.zoom || 100);
    try { localStorage.setItem("hl.write.last", f.id); } catch (e) {}
    history.replaceState(null, "", "?id=" + encodeURIComponent(f.id));
    count(); statusEl.textContent = Office.isCloud() ? "Saved to your account" : "Saved on this device";
  }
  function fresh(title, html) {
    file = { id: Office.uuid(), title: title || "Untitled document", body: { html: html || "" } };
    titleIn.value = file.title; page.innerHTML = html || "<p><br></p>"; count();
    history.replaceState(null, "", "?id=" + encodeURIComponent(file.id));
    try { localStorage.setItem("hl.write.last", file.id); } catch (e) {}
    saver.now(); page.focus();
  }
  function openId(id) {
    statusEl.textContent = "Opening…";
    return Office.get(id).then(function (f) { if (f) load(f); else fresh(); })
      .catch(function (e) { Office.toast("Couldn't open it: " + (e.message || e), true); fresh(); });
  }

  Office.init("doc").then(function (st) {
    var q = new URLSearchParams(location.search).get("id"), last = null;
    try { last = localStorage.getItem("hl.write.last"); } catch (e) {}
    if (!st.user) Office.toast("Not signed in - your work is saved in this browser.");
    if (q || last) return openId(q || last);
    return Office.list().then(function (rows) { if (rows.length) return openId(rows[0].id); fresh(); });
  });

  $("filesBtn").onclick = function () {
    saver.now();
    Office.filesDialog({
      heading: "Your documents", noun: "document", icon: "📘", current: file.id,
      accept: ".docx,.txt,.html,.htm,.md",
      onOpen: openId, onNew: function () { fresh(); },
      onDeletedCurrent: function () { fresh(); },
      onImport: importFile
    });
  };
  Office.themeToggle($("themeBtn"));
  titleIn.addEventListener("input", saver.poke);
  titleIn.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); page.focus(); } });

  /* ---------- typing ---------- */
  page.addEventListener("input", function () { tidy(); count(); saver.poke(); });
  document.addEventListener("selectionchange", syncTools);

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
  });

  document.addEventListener("keydown", function (e) {
    var mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === "s") { e.preventDefault(); saver.now(); }
    else if (mod && e.key.toLowerCase() === "f") { e.preventDefault(); openFind(); }
    else if (mod && e.key.toLowerCase() === "p") { e.preventDefault(); window.print(); }
    else if (mod && e.key.toLowerCase() === "k") { e.preventDefault(); addLink(); }
    else if (e.key === "Escape" && !$("findbar").hidden) closeFind();
    else if (e.key === "Tab" && document.activeElement === page) {
      e.preventDefault();
      var li = selBlock(); if (li && li.closest("li")) document.execCommand(e.shiftKey ? "outdent" : "indent");
      else if (inCell()) moveCell(e.shiftKey ? -1 : 1);
      else document.execCommand("insertText", false, "    ");
    }
  });

  /* ---------- the ribbon ---------- */
  var saved = null;   // the selection, kept while a menu or colour picker has focus
  function keep() { var s = getSelection(); if (s.rangeCount && page.contains(s.anchorNode)) saved = s.getRangeAt(0).cloneRange(); }
  function back() { page.focus(); if (saved) { var s = getSelection(); s.removeAllRanges(); s.addRange(saved); } }
  page.addEventListener("keyup", keep); page.addEventListener("mouseup", keep); page.addEventListener("blur", keep);

  document.querySelectorAll("[data-cmd]").forEach(function (b) {
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.addEventListener("click", function () {
      back();
      if (b.dataset.cmd === "insertHorizontalRule") insertBlock("<hr>");
      else document.execCommand(b.dataset.cmd, false, null);
      tidy(); page.focus(); syncTools(); saver.poke(); count();
    });
  });

  $("style").onchange = function () {
    back();
    var v = this.value;
    if (v === "title") { document.execCommand("formatBlock", false, "<p>"); var b = selBlock(); if (b) b.className = "doc-title"; }
    else { document.execCommand("formatBlock", false, "<" + v + ">"); var b2 = selBlock(); if (b2) b2.classList.remove("doc-title"); }
    saver.poke(); page.focus();
  };
  $("font").onchange = function () { back(); document.execCommand("fontName", false, this.value); saver.poke(); page.focus(); };
  $("size").onchange = function () {
    back(); var pt = this.value;
    document.execCommand("styleWithCSS", false, false);
    document.execCommand("fontSize", false, "7");
    page.querySelectorAll('font[size="7"]').forEach(function (f) {
      var s = document.createElement("span"); s.style.fontSize = pt + "pt";
      while (f.firstChild) s.appendChild(f.firstChild); f.replaceWith(s);
    });
    page.querySelectorAll("span").forEach(function (s) { if (s.style.fontSize === "xxx-large") s.style.fontSize = pt + "pt"; });
    document.execCommand("styleWithCSS", false, true);
    saver.poke(); page.focus();
  };
  function colour(inp, bar, cmd) {
    inp.addEventListener("mousedown", keep);
    inp.addEventListener("input", function () { bar.style.background = inp.value; });
    inp.addEventListener("change", function () {
      back(); document.execCommand(cmd, false, inp.value); saver.poke();
    });
  }
  colour($("fg"), $("fgBar"), "foreColor");
  colour($("hl"), $("hlBar"), "hiliteColor");

  function selBlock() {
    var s = getSelection(); if (!s.rangeCount) return null;
    var n = s.anchorNode; if (n && n.nodeType === 3) n = n.parentNode;
    while (n && n !== page && !/^(P|H1|H2|H3|BLOCKQUOTE|LI|DIV|TD|TH)$/.test(n.tagName)) n = n.parentNode;
    return n && n !== page ? n : null;
  }
  /* The block straight under the page that holds the cursor. */
  function topBlock() {
    var s = getSelection(); if (!s.rangeCount) return null;
    var n = s.anchorNode; if (n === page) return page.childNodes[s.anchorOffset - 1] || null;
    while (n && n.parentNode !== page) n = n.parentNode;
    return n;
  }
  /* Tables and other big blocks go after the current line (never inside it),
     with an empty line after them to keep typing. */
  function insertBlock(html) {
    var t = document.createElement("div"); t.innerHTML = html;
    var at = topBlock(), after = document.createElement("p"); after.innerHTML = "<br>";
    var frag = document.createDocumentFragment(); while (t.firstChild) frag.appendChild(t.firstChild); frag.appendChild(after);
    var first = frag.firstChild;
    if (at && at.nodeType === 1 && at.textContent.trim() === "" && !at.querySelector("img")) at.replaceWith(frag);
    else if (at) at.after(frag); else page.appendChild(frag);
    var cell = first && first.querySelector ? first.querySelector("td,th") : null;
    var r = document.createRange(); r.selectNodeContents(cell || after); r.collapse(true);
    var sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
  }
  /* Browsers sometimes put a list or table inside a paragraph, which
     isn't allowed. Lift them out so the page (and the Word file) stay tidy. */
  function tidy() {
    var bad = page.querySelectorAll("p > ul, p > ol, p > table, p > p, p > h1, p > h2, p > h3, p > blockquote, p > div");
    if (!bad.length) return;
    var s = getSelection(), keepR = s.rangeCount ? s.getRangeAt(0) : null;
    Array.prototype.slice.call(page.querySelectorAll("p")).forEach(function (p) {
      if (!p.isConnected || !p.querySelector(":scope > ul, :scope > ol, :scope > table, :scope > p, :scope > h1, :scope > h2, :scope > h3, :scope > blockquote, :scope > div")) return;
      var frag = document.createDocumentFragment(), cur = null;
      Array.prototype.slice.call(p.childNodes).forEach(function (n) {
        if (n.nodeType === 1 && /^(UL|OL|TABLE|P|H1|H2|H3|BLOCKQUOTE|DIV)$/.test(n.tagName)) { cur = null; frag.appendChild(n); }
        else if (n.nodeType === 3 && !n.nodeValue.trim() && !cur) { /* drop stray spaces */ }
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
    if (i >= cells.length) { // Tab in the last cell adds a row, like Word
      var tr = c.closest("tr"), nr = tr.cloneNode(true);
      nr.querySelectorAll("td,th").forEach(function (x) { var td = document.createElement("td"); td.innerHTML = "<br>"; x.replaceWith(td); });
      tr.after(nr); cells = Array.prototype.slice.call(c.closest("table").querySelectorAll("td,th"));
    }
    var t = cells[Math.max(0, i)]; if (!t) return;
    var r = document.createRange(); r.selectNodeContents(t); r.collapse(false);
    var s = getSelection(); s.removeAllRanges(); s.addRange(r); saver.poke();
  }
  function syncTools() {
    if (!page.contains(getSelection().anchorNode)) return;
    ["bold", "italic", "underline", "strikeThrough", "superscript", "subscript", "insertUnorderedList", "insertOrderedList",
     "justifyLeft", "justifyCenter", "justifyRight", "justifyFull"].forEach(function (c) {
      var b = document.querySelector('[data-cmd="' + c + '"]');
      try { if (b) b.setAttribute("aria-pressed", document.queryCommandState(c) ? "true" : "false"); } catch (e) {}
    });
    var b = selBlock(), v = "p";
    if (b) { var t = b.tagName.toLowerCase(); if (b.classList.contains("doc-title")) v = "title"; else if (/^h[123]$|^blockquote$/.test(t)) v = t; }
    $("style").value = v;
  }

  function addLink() { keep(); linkBox(); }
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

  $("imgIn").addEventListener("click", keep);
  $("imgIn").onchange = function () { if (this.files[0]) insertImageFile(this.files[0]); this.value = ""; };
  function insertImageFile(f) {
    Office.shrinkImage(f, 1400).then(function (src) {
      back(); document.execCommand("insertHTML", false, '<img src="' + src + '" alt="">'); saver.poke();
    }).catch(function (e) { Office.toast(e.message, true); });
  }

  $("tableBtn").addEventListener("mousedown", function (e) { e.preventDefault(); keep(); });
  $("tableBtn").onclick = function () {
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>Insert a table</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
      '<p class="of-where">Tab moves to the next cell. Tab in the last cell adds a row.</p>' +
      '<div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap"><label>Columns <input type="number" id="tc" value="3" min="1" max="12" style="width:70px;height:36px;border:1px solid var(--rule);border-radius:9px;padding:0 8px;background:var(--tool)"></label>' +
      '<label>Rows <input type="number" id="tr" value="3" min="1" max="60" style="width:70px;height:36px;border:1px solid var(--rule);border-radius:9px;padding:0 8px;background:var(--tool)"></label>' +
      '<label><input type="checkbox" id="th" checked> Header row</label></div>' +
      '<div class="of-new" style="margin-top:14px"><button type="button" class="of-btn" id="tOk">Insert</button></div>';
    document.body.appendChild(d); d.showModal();
    function done() { d.close(); d.remove(); }
    d.querySelector(".of-x").onclick = done;
    d.querySelector("#tOk").onclick = function () {
      var c = Math.max(1, Math.min(12, +d.querySelector("#tc").value || 3)), r = Math.max(1, Math.min(60, +d.querySelector("#tr").value || 3));
      var h = d.querySelector("#th").checked, html = "<table><tbody>";
      for (var i = 0; i < r; i++) { html += "<tr>"; for (var j = 0; j < c; j++) html += (h && i === 0 ? "<th><br></th>" : "<td><br></td>"); html += "</tr>"; }
      html += "</tbody></table>";
      done(); back(); insertBlock(html); saver.poke();
    };
  };

  /* ---------- export menu ---------- */
  var exBtn = $("exportBtn"), exMenu = $("exportMenu");
  exBtn.onclick = function (e) { e.stopPropagation(); exMenu.hidden = !exMenu.hidden; exBtn.setAttribute("aria-expanded", String(!exMenu.hidden)); };
  document.addEventListener("click", function (e) { if (!exMenu.hidden && !exMenu.contains(e.target)) { exMenu.hidden = true; exBtn.setAttribute("aria-expanded", "false"); } });
  exMenu.onclick = function (e) {
    var b = e.target.closest("[data-ex]"); if (!b) return;
    exMenu.hidden = true;
    var t = titleIn.value.trim() || "Untitled document";
    if (b.dataset.ex === "pdf") { window.print(); return; }
    if (b.dataset.ex === "txt") { Office.download(new Blob([page.innerText], { type: "text/plain;charset=utf-8" }), Office.safeName(t, ".txt")); return; }
    if (b.dataset.ex === "html") {
      var doc = "<!DOCTYPE html><html lang=\"en-GB\"><head><meta charset=\"utf-8\"><title>" + Office.esc(t) + "</title>" +
        "<style>body{font:12pt/1.5 Calibri,Carlito,sans-serif;max-width:720px;margin:40px auto;padding:0 20px}img{max-width:100%}table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px 8px}.doc-title{font-size:28pt;font-weight:300}</style></head><body>" +
        page.innerHTML + "</body></html>";
      Office.download(new Blob([doc], { type: "text/html;charset=utf-8" }), Office.safeName(t, ".html")); return;
    }
    if (b.dataset.ex === "docx") {
      Office.toast("Making your Word file…");
      Docx.build(page, t).then(function (blob) { Office.download(blob, Office.safeName(t, ".docx")); })
        .catch(function (e) { Office.toast("Couldn't make the Word file: " + (e.message || e), true); });
    }
  };

  function importFile(f) {
    var name = f.name.replace(/\.[^.]+$/, "") || "Imported document";
    if (f.size > 20 * 1024 * 1024) { Office.toast("That file is too big (over 20 MB).", true); return; }
    if (/\.docx$/i.test(f.name)) {
      Office.toast("Opening " + f.name + "…");
      Docx.read(f).then(function (html) { fresh(name, Office.cleanHTML(html)); Office.toast("Opened " + f.name); })
        .catch(function (e) { Office.toast("Couldn't open that Word file: " + (e.message || e), true); });
      return;
    }
    var r = new FileReader();
    r.onload = function () {
      var t = String(r.result);
      if (/\.html?$/i.test(f.name)) fresh(name, Office.cleanHTML(t.replace(/^[\s\S]*<body[^>]*>|<\/body>[\s\S]*$/gi, "")));
      else fresh(name, t.split(/\r?\n/).map(function (l) { return "<p>" + (Office.esc(l) || "<br>") + "</p>"; }).join(""));
    };
    r.readAsText(f);
  }

  /* ---------- counts and zoom ---------- */
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
  function setZoom(z) {
    z = Math.max(50, Math.min(200, z)); $("zoom").value = z; $("zoomN").textContent = z + "%";
    $("zoomer").style.zoom = z / 100;
  }
  $("zoom").oninput = function () { setZoom(+this.value); saver.poke(); };

  /* ---------- find and replace ---------- */
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
  // marks never get saved
  window.addEventListener("beforeunload", function () { clearMarks(); saver.now(); });

  /* ---------- read aloud ---------- */
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
})();
