/* =====================================================================
   Launchpad Write and Grid - the shared parts.

   Saving: signed in, files go to your Launchpad account (the
   office_files table, which only its owner can read or change), so
   they follow you to any device. Signed out, or before the table
   exists, they stay in this browser. Files made on this device while
   signed out are moved into your account the next time you sign in.

   Needs, in order: supabase-js, auth-store.js, then this file.
   ===================================================================== */
(function () {
  var URL_ = "https://hehs-live.duckdns.org";
  var KEY_ = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoiYW5vbiIsImlzcyI6InN1cGFiYXNlIiwiaWF0IjoxNzkyODMyMjkwLCJleHAiOjIxMDgxOTIyOTB9.7gFhQa5RBwVtiOPLO7Qf2-aVEbkd5xgpLiN1J06Gybo";
  var sb = null;
  try {
    sb = window.supabase.createClient(URL_, KEY_, {
      auth: { persistSession: true, autoRefreshToken: true,
              storage: (window.HEHAuth && window.HEHAuth.storage) || undefined }
    });
  } catch (e) { sb = null; }

  var kind = "doc", cloud = false, user = null;

  function uuid() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
      var r = Math.random() * 16 | 0; return (c === "x" ? r : (r & 3 | 8)).toString(16);
    });
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  /* ---- this browser ---- */
  function lkey() { return "hl.office." + kind; }
  function lall() {
    try { return JSON.parse(localStorage.getItem(lkey())) || {}; } catch (e) { return {}; }
  }
  function lput(all) {
    try { localStorage.setItem(lkey(), JSON.stringify(all)); return true; }
    catch (e) { return false; }
  }

  /* ---- start up: who is signed in, and is the table there? ---- */
  function init(k) {
    kind = k;
    if (!sb) return Promise.resolve({ user: null, cloud: false });
    return sb.auth.getSession().then(function (r) {
      user = r && r.data && r.data.session ? r.data.session.user : null;
      if (!user) return { user: null, cloud: false };
      return sb.from("office_files").select("id").limit(1).then(function (q) {
        cloud = !q.error;
        if (cloud) return moveLocalUp().then(function () { return { user: user, cloud: true }; });
        return { user: user, cloud: false };
      });
    }).catch(function () { return { user: null, cloud: false }; });
  }

  function moveLocalUp() {
    var all = lall(), ids = Object.keys(all);
    if (!ids.length) return Promise.resolve();
    var rows = ids.map(function (id) {
      var f = all[id];
      return { id: f.id, kind: kind, title: f.title || "Untitled", body: f.body || {} };
    });
    return sb.from("office_files").upsert(rows, { onConflict: "id" }).then(function (r) {
      if (!r.error) { try { localStorage.removeItem(lkey()); } catch (e) {} }
    });
  }

  function list() {
    if (cloud) {
      // files shared with you come too (user_id isn't yours)
      return sb.from("office_files").select("id,title,updated_at,user_id").eq("kind", kind)
        .order("updated_at", { ascending: false }).limit(300).then(function (r) {
          if (r.error) throw r.error;
          return (r.data || []).map(function (x) { x.shared = !!(user && x.user_id && x.user_id !== user.id); return x; });
        });
    }
    var all = lall();
    return Promise.resolve(Object.keys(all).map(function (id) {
      return { id: id, title: all[id].title, updated_at: all[id].updated_at };
    }).sort(function (a, b) { return a.updated_at < b.updated_at ? 1 : -1; }));
  }

  function get(id) {
    if (cloud) {
      return sb.from("office_files").select("id,title,body,updated_at").eq("id", id).maybeSingle()
        .then(function (r) { if (r.error) throw r.error; return r.data; });
    }
    return Promise.resolve(lall()[id] || null);
  }

  function save(f) {
    if (!f.id) f.id = uuid();
    var now = new Date().toISOString();
    if (cloud) {
      return sb.from("office_files").upsert({ id: f.id, kind: kind, title: f.title || "Untitled", body: f.body },
        { onConflict: "id" }).then(function (r) {
          // the server doesn't know this kind of file yet (it needs the
          // latest update) - keep it in this browser until it does
          if (r.error && /kind_check|check constraint/i.test(r.error.message || "")) { cloud = false; return save(f); }
          if (r.error) throw r.error; return f;
        });
    }
    var all = lall();
    all[f.id] = { id: f.id, title: f.title || "Untitled", body: f.body, updated_at: now };
    if (!lput(all)) return Promise.reject(new Error("This browser is out of room. Sign in to save to your account."));
    return Promise.resolve(f);
  }

  function remove(id) {
    if (cloud) {
      return sb.from("office_files").delete().eq("id", id).then(function (r) { if (r.error) throw r.error; });
    }
    var all = lall(); delete all[id]; lput(all); return Promise.resolve();
  }

  /* ---- little helpers both apps use ---- */
  function download(blob, name) {
    // in the Android app, the app saves it (a web view can't download blobs)
    if (window.SuiteAndroid && window.SuiteAndroid.save) {
      var fr = new FileReader();
      fr.onload = function () { var u = String(fr.result); window.SuiteAndroid.save(name, blob.type || "application/octet-stream", u.slice(u.indexOf(",") + 1)); };
      fr.readAsDataURL(blob); return;
    }
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  function safeName(t, ext) {
    return (String(t || "Untitled").replace(/[\\/:*?"<>|\u0000-\u001f]+/g, " ").trim().slice(0, 80) || "Untitled") + ext;
  }
  var toastT = null;
  function toast(msg, bad) {
    var el = document.getElementById("ofToast");
    if (!el) { el = document.createElement("div"); el.id = "ofToast"; el.setAttribute("role", "status"); document.body.appendChild(el); }
    el.textContent = msg; el.className = bad ? "show bad" : "show";
    clearTimeout(toastT); toastT = setTimeout(function () { el.className = ""; }, 3200);
  }
  function when(iso) {
    if (!iso) return "";
    var d = new Date(iso), now = new Date(), s = (now - d) / 1000;
    if (s < 60) return "just now";
    if (s < 3600) return Math.floor(s / 60) + " min ago";
    if (d.toDateString() === now.toDateString()) return "today " + d.toTimeString().slice(0, 5);
    return d.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: d.getFullYear() === now.getFullYear() ? undefined : "numeric" });
  }

  /* ---- shared storage: 10 GB for everyone, shown as a bar ---- */
  function size(n) { return n >= 1073741824 ? (n / 1073741824).toFixed(1) + " GB" : n >= 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB"; }
  function storage() {
    if (!cloud || !sb) return Promise.resolve(null);
    return sb.rpc("storage_info").then(function (r) { return r.error ? null : r.data; }, function () { return null; });
  }
  function storageBar(el) {
    storage().then(function (s) {
      if (!s || !el) return;
      var pct = Math.min(100, s.used / s.limit * 100), col = pct > 90 ? "#dc2626" : pct > 70 ? "#f59e0b" : "var(--a1)";
      el.innerHTML = '<div style="display:flex;justify-content:space-between;font-size:.82rem;color:var(--soft);margin-bottom:4px"><span>☁️ Suite storage: <b style="color:var(--ink)">' + size(s.used) + "</b> of " + size(s.limit) + " shared by everyone</span><span>Yours: " + size(s.mine) + "</span></div>" +
        '<div style="height:8px;border-radius:999px;background:var(--tool);overflow:hidden"><div style="height:100%;width:' + pct.toFixed(1) + "%;background:" + col + ';border-radius:999px"></div></div>';
    });
  }

  /* ---- the Files window ---- */
  function filesDialog(opts) {
    var dlg = document.getElementById("ofFiles");
    if (!dlg) {
      dlg = document.createElement("dialog"); dlg.id = "ofFiles"; dlg.className = "of-dialog";
      document.body.appendChild(dlg);
    }
    dlg.innerHTML =
      '<div class="of-dhead"><h2>' + esc(opts.heading) + '</h2>' +
      '<button type="button" class="of-x" data-act="close" aria-label="Close">✕</button></div>' +
      '<p class="of-where">' + esc(cloud ? "Saved to your Launchpad account - they follow you to any device."
        : (user ? "Saved in this browser for now." : "Saved in this browser. Sign in on the Launchpad to keep them in your account.")) + '</p>' +
      '<div class="of-new"><button type="button" class="of-btn" data-act="new">＋ New ' + esc(opts.noun) + '</button>' +
      '<label class="of-btn ghost">Open a file from your device<input type="file" hidden></label></div>' +
      '<ul class="of-list"><li class="of-empty">Loading…</li></ul><div class="of-store" style="margin-top:12px"></div>';
    storageBar(dlg.querySelector(".of-store"));
    var ul = dlg.querySelector(".of-list");
    function draw() {
      list().then(function (rows) {
        if (!rows.length) { ul.innerHTML = '<li class="of-empty">Nothing here yet.</li>'; return; }
        ul.innerHTML = rows.map(function (r) {
          return '<li' + (r.id === opts.current ? ' class="cur"' : '') + '><button type="button" class="of-open" data-id="' + esc(r.id) + '">' +
            '<span class="of-ic">' + opts.icon + '</span><span><strong>' + esc(r.title || "Untitled") + '</strong><small>' + (r.shared ? "👥 Shared with you · " : "") + esc(when(r.updated_at)) + '</small></span></button>' +
            (r.shared ? "" : '<button type="button" class="of-del" data-del="' + esc(r.id) + '" aria-label="Delete ' + esc(r.title || "Untitled") + '">🗑</button>') + '</li>';
        }).join("");
      }).catch(function (e) { ul.innerHTML = '<li class="of-empty">Couldn\'t load your files: ' + esc(e.message || e) + '</li>'; });
    }
    dlg.onclick = function (ev) {
      var t = ev.target.closest("[data-act],[data-id],[data-del]");
      if (ev.target === dlg) { dlg.close(); return; }
      if (!t) return;
      if (t.dataset.act === "close") dlg.close();
      else if (t.dataset.act === "new") { dlg.close(); opts.onNew(); }
      else if (t.dataset.id) { dlg.close(); opts.onOpen(t.dataset.id); }
      else if (t.dataset.del) {
        var li = t.closest("li"), name = li.querySelector("strong").textContent;
        if (!confirmInline(t, "Delete “" + name + "”?")) return;
        remove(t.dataset.del).then(function () {
          if (t.dataset.del === opts.current) opts.onDeletedCurrent();
          draw();
        }).catch(function (e) { toast("Couldn't delete: " + (e.message || e), true); });
      }
    };
    dlg.querySelector('input[type=file]').onchange = function () {
      var f = this.files && this.files[0]; this.value = ""; if (!f) return; dlg.close();
      /* Google Docs, Sheets and Slides aren't real files - Drive only hands
         over a shortcut - so say how to get a copy that can be opened. */
      var ext = (f.name.match(/\.[^.]+$/) || [""])[0].toLowerCase();
      if (/^\.(gdoc|gsheet|gslides|gdraw|gform)$/.test(ext) || /google-apps/.test(f.type) || (!ext && f.size < 2048)) { googleHelp(opts.noun); return; }
      var ok = String(opts.accept || "").split(",").map(function (x) { return x.trim().toLowerCase(); });
      if (ok.indexOf(ext) < 0) { toast("That's a " + (ext || "file without a type") + " - this app opens " + ok.join(", "), true); return; }
      opts.onImport(f);
    };
    draw();
    if (dlg.showModal) dlg.showModal(); else dlg.setAttribute("open", "");
  }
  function googleHelp(noun) {
    var fmt = { document: ["Word (.docx)", "Google Docs"], spreadsheet: ["Microsoft Excel (.xlsx)", "Google Sheets"], presentation: ["Microsoft PowerPoint (.pptx)", "Google Slides"], notebook: ["Word (.docx)", "Google Docs"] }[noun] || ["Word (.docx)", "Google Docs"];
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>Make a copy first</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
      '<p>' + esc(fmt[1]) + ' files live inside Google, so the Drive picker can only hand over a shortcut, not the writing itself. Save a copy, then open that:</p>' +
      '<ol style="line-height:1.6;padding-left:1.3em"><li>Open the file in the <b>' + esc(fmt[1]) + '</b> app (or docs.google.com).</li>' +
      '<li>Tap <b>⋯</b> → <b>Share and export</b> → <b>Save as</b> (on a computer: <b>File</b> → <b>Download</b>).</li>' +
      '<li>Pick <b>' + esc(fmt[0]) + '</b>, save it to <b>Files</b> (or your Downloads).</li>' +
      '<li>Come back here, tap <b>Open a file from your device</b> and pick the copy.</li></ol>' +
      '<div class="of-new" style="margin-top:10px"><button type="button" class="of-btn">Got it</button></div>';
    document.body.appendChild(d); d.showModal();
    d.onclick = function (e) { if (e.target === d || e.target.closest(".of-x,.of-btn")) { d.close(); d.remove(); } };
  }
  /* A second tap on the same bin confirms, so no browser pop-up is needed. */
  function confirmInline(btn, label) {
    if (btn.dataset.armed) return true;
    btn.dataset.armed = "1"; btn.textContent = "Delete?"; btn.classList.add("armed"); btn.setAttribute("aria-label", label);
    setTimeout(function () { if (btn.isConnected) { delete btn.dataset.armed; btn.textContent = "🗑"; btn.classList.remove("armed"); } }, 3000);
    return false;
  }

  /* ---- a debounced autosaver with a status line ---- */
  function autosaver(statusEl, getFile, o) {
    var t = null, busy = false, again = false;
    o = o || {};
    function label(s) { statusEl.textContent = s; }
    function run() {
      if (busy) { again = true; return; }
      if (o.skip && o.skip()) { if (o.done) o.done(); return; }   // view only: nothing to save
      busy = true; label("Saving…");
      (o.save || save)(getFile()).then(function () {
        label(cloud ? "✓ Saved to your account" : "✓ Saved on this device");
      }).catch(function (e) {
        label("⚠ Not saved"); toast("Couldn't save: " + (e.message || e), true);
      }).then(function () { busy = false; if (again) { again = false; run(); } else if (o.done) o.done(); });
    }
    return {
      poke: function () { if (o.skip && o.skip()) return; label("Editing…"); clearTimeout(t); t = setTimeout(run, 1200); },
      now: function () { clearTimeout(t); run(); }
    };
  }

  /* Ask the shared cookie for the theme, so these follow the Launchpad. */
  function themeToggle(btn) {
    function cur() { return document.documentElement.getAttribute("data-theme") ||
      (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light"); }
    function show() { btn.textContent = cur() === "dark" ? "☀️" : "🌙"; btn.setAttribute("aria-label", cur() === "dark" ? "Light mode" : "Dark mode"); }
    btn.onclick = function () {
      var t = cur() === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", t);
      try { document.cookie = "hl_theme=" + t + "; domain=.civilinitiatives.org; path=/; max-age=34560000; samesite=lax"; } catch (e) {}
      try { localStorage.setItem("hl.theme", JSON.stringify(t)); } catch (e) {}
      show();
    };
    show();
  }

  /* ---- rich text that came from outside (paste, a file, the database)
     is rebuilt from a short list of safe tags and styles, so nothing in
     it can run code. ---- */
  var TAGS = { P:1, DIV:"P", H1:1, H2:1, H3:1, BLOCKQUOTE:1, UL:1, OL:1, LI:1, B:1, STRONG:"B", I:1, EM:"I", U:1,
    S:1, STRIKE:"S", DEL:"S", SUP:1, SUB:1, A:1, BR:1, SPAN:1, FONT:"SPAN", TABLE:1, THEAD:1, TBODY:1, TR:1, TD:1, TH:1,
    IMG:1, HR:1, MARK:"SPAN", LABEL:"SPAN", INPUT:1, CODE:1, PRE:1, H4:"H3", H5:"H3", H6:"H3", SECTION:"P", ARTICLE:"P" };
  var STYLES = ["color", "background-color", "font-size", "font-family", "text-align", "font-weight", "font-style",
    "text-decoration", "text-decoration-line", "margin-left", "margin", "border", "padding", "vertical-align"];
  function okURL(u, img) {
    u = String(u || "").trim();
    if (/^https?:\/\//i.test(u)) return u;
    if (!img && /^mailto:/i.test(u)) return u;
    if (img && /^data:image\/(png|jpe?g|gif|webp);base64,[a-z0-9+\/=\s]+$/i.test(u)) return u;
    return "";
  }
  function cleanStyle(s) {
    var out = [];
    String(s || "").split(";").forEach(function (d) {
      var i = d.indexOf(":"); if (i < 0) return;
      var k = d.slice(0, i).trim().toLowerCase(), v = d.slice(i + 1).trim();
      if (STYLES.indexOf(k) < 0) return;
      if (/url\s*\(|expression|javascript:|[<>{}\\]/i.test(v) || v.length > 80) return;
      out.push(k + ":" + v);
    });
    return out.join(";");
  }
  function cleanHTML(html) {
    var doc = new DOMParser().parseFromString("<body>" + String(html || "") + "</body>", "text/html");
    var out = document.createElement("div");
    function walk(src, dst) {
      for (var n = src.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3) { dst.appendChild(document.createTextNode(n.nodeValue)); continue; }
        if (n.nodeType !== 1) continue;
        var tag = n.tagName, map = TAGS[tag];
        if (!map) { if (!/^(SCRIPT|STYLE|IFRAME|OBJECT|EMBED|TEMPLATE|SVG|MATH|HEAD|TITLE|META|LINK|NOSCRIPT|BUTTON|SELECT|TEXTAREA|FORM)$/.test(tag)) walk(n, dst); continue; }
        var name = map === 1 ? tag : map;
        if (name === "INPUT" && (n.getAttribute("type") || "").toLowerCase() !== "checkbox") continue;
        var el = document.createElement(name);
        var st = cleanStyle(n.getAttribute("style"));
        if (tag === "FONT") {
          if (n.getAttribute("color")) st += (st ? ";" : "") + "color:" + n.getAttribute("color").replace(/[^#a-z0-9]/gi, "");
          if (n.getAttribute("face")) st += (st ? ";" : "") + "font-family:" + n.getAttribute("face").replace(/[^a-z0-9 ,'-]/gi, "");
        }
        if (st) el.setAttribute("style", st);
        if (name === "A") { var h = okURL(n.getAttribute("href")); if (h) { el.setAttribute("href", h); el.setAttribute("target", "_blank"); el.setAttribute("rel", "noopener noreferrer"); } }
        if (name === "IMG") { var s = okURL(n.getAttribute("src"), true); if (!s) continue; el.setAttribute("src", s); el.setAttribute("alt", (n.getAttribute("alt") || "").slice(0, 200));
          var w = parseInt(n.getAttribute("width"), 10); if (w > 0 && w < 4000) el.setAttribute("width", w); }
        if (name === "INPUT") { el.setAttribute("type", "checkbox"); if (n.hasAttribute("checked") || n.checked) el.setAttribute("checked", ""); }
        if ((name === "TD" || name === "TH") && +n.getAttribute("colspan") > 1) el.setAttribute("colspan", Math.min(+n.getAttribute("colspan"), 50));
        var CLS = /\b(doc-title|doc-subtitle|todo|done|cmt|pb|toc|toch)\b/g;
        if (n.className && typeof n.className === "string" && n.className.match(CLS)) el.className = n.className.match(CLS).join(" ");
        var dc = n.getAttribute("data-c"); if (dc && /^[a-z0-9]{1,16}$/.test(dc)) el.setAttribute("data-c", dc);
        var db = n.getAttribute("data-b"); if (db && /^[a-z0-9]{1,16}$/.test(db)) el.setAttribute("data-b", db);   // paragraph ids, for working together
        if (name !== "BR" && name !== "HR" && name !== "IMG" && name !== "INPUT") walk(n, el);
        dst.appendChild(el);
      }
    }
    walk(doc.body, out);
    return out.innerHTML;
  }

  /* Pictures are shrunk before they go in a file, so files stay small. */
  function shrinkImage(file, max) {
    max = max || 1400;
    return new Promise(function (ok, no) {
      if (!/^image\//.test(file.type)) { no(new Error("That isn't a picture.")); return; }
      var r = new FileReader();
      r.onload = function () {
        var img = new Image();
        img.onload = function () {
          var k = Math.min(1, max / Math.max(img.width, img.height)), c = document.createElement("canvas");
          c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
          var g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
          ok(c.toDataURL(/png|gif/.test(file.type) ? "image/png" : "image/jpeg", 0.85));
        };
        img.onerror = function () { no(new Error("Couldn't read that picture.")); };
        img.src = r.result;
      };
      r.onerror = function () { no(new Error("Couldn't read that file.")); };
      r.readAsDataURL(file);
    });
  }

  /* ---- installing the Suite as an app, and offline ---- */
  if ("serviceWorker" in navigator && location.protocol === "https:") {
    window.addEventListener("load", function () { navigator.serviceWorker.register("/sw.js").catch(function () {}); });
  }
  var installEvt = null, importer = null, waiting = [];
  window.addEventListener("beforeinstallprompt", function (e) { e.preventDefault(); installEvt = e; document.dispatchEvent(new Event("suite-installable")); });
  function canInstall() { return !!installEvt; }
  function isApp() { return matchMedia("(display-mode: standalone)").matches || navigator.standalone === true || /LaunchpadSuite\//.test(navigator.userAgent); }
  function install() {
    if (installEvt) { installEvt.prompt(); return installEvt.userChoice.then(function (c) { installEvt = null; return c.outcome === "accepted"; }); }
    var ios = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
    var how = ios ? "In Safari, tap <b>Share</b> (the square with an arrow), then <b>Add to Home Screen</b>, then <b>Add</b>."
      : /Mac/.test(navigator.platform) ? "In Safari, choose <b>File → Add to Dock</b>. In Chrome or Edge, use the <b>install</b> icon at the right of the address bar."
      : "In Chrome or Edge, use the <b>install</b> icon at the right of the address bar, or the <b>⋮</b> menu → <b>Install Launchpad Suite</b>. On a Chromebook it's the same.";
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>Install the Suite</h2><button type="button" class="of-x" aria-label="Close">✕</button></div><p>' + how +
      '</p><p class="of-where">It opens in its own window, has its own icon, and works without the internet.</p><div class="of-new"><button type="button" class="of-btn">Got it</button></div>';
    document.body.appendChild(d); d.showModal();
    d.onclick = function (e) { if (e.target === d || e.target.closest(".of-x,.of-btn")) { d.close(); d.remove(); } };
    return Promise.resolve(false);
  }
  /* files opened with the installed app (double-click a .docx, say) */
  if ("launchQueue" in window) window.launchQueue.setConsumer(function (p) {
    (p.files || []).slice(0, 1).forEach(function (h) { h.getFile().then(function (f) { if (importer) importer(f); else waiting.push(f); }); });
  });
  function setImporter(fn) { importer = fn; waiting.splice(0).forEach(fn); }

  /* ---- Office-style ribbon tabs and drop-downs ----
     <nav class="tabsbar"> buttons with data-tab, panels .ribbon[data-panel].
     A button with data-dd="someId" opens the .dd with that id. */
  function ribbon(opts) {
    opts = opts || {};
    var cur = null;
    function show(t) {
      cur = t;
      document.querySelectorAll(".tabsbar [data-tab]").forEach(function (b) { b.setAttribute("aria-selected", String(b.dataset.tab === t)); });
      document.querySelectorAll(".ribbon[data-panel]").forEach(function (p) { p.hidden = p.dataset.panel !== t; });
      closeDds(); if (opts.onTab) opts.onTab(t);
    }
    function closeDds(except) { document.querySelectorAll(".dd").forEach(function (d) { if (d !== except) d.hidden = true; }); }
    document.querySelectorAll(".tabsbar [data-tab]").forEach(function (b) {
      b.addEventListener("mousedown", function (e) { e.preventDefault(); });
      b.onclick = function () { show(b.dataset.tab); };
    });
    document.querySelectorAll("[data-dd]").forEach(function (b) {
      var d = document.getElementById(b.dataset.dd); if (!d) return;
      b.addEventListener("mousedown", function (e) { e.preventDefault(); if (opts.keep) opts.keep(); });
      b.addEventListener("click", function (e) { e.stopPropagation(); closeDds(d); d.hidden = !d.hidden; });
      d.addEventListener("click", function (e) { if (e.target.closest("button")) setTimeout(function () { d.hidden = true; }, 0); });
    });
    document.addEventListener("click", function (e) { if (!e.target.closest(".dd")) closeDds(); });
    var first = document.querySelector('.tabsbar [data-tab][aria-selected="true"]') || document.querySelector(".tabsbar [data-tab]");
    if (first) show(first.dataset.tab);
    return { show: show, current: function () { return cur; }, close: closeDds };
  }

  /* ---- the File screen (like Office's "backstage") ----
     pages: { home: {title, items:[{a,icon,title,sub}]}, save: {...}, info: fn -> html }
     onPick(action) gets the data-a of the tile that was chosen. */
  function backstage(opts) {
    var old = document.getElementById("backstage"); if (old) old.remove();
    var bs = document.createElement("div"); bs.id = "backstage"; bs.setAttribute("role", "dialog"); bs.setAttribute("aria-label", "File");
    var names = Object.keys(opts.pages);
    bs.innerHTML = '<nav><button type="button" data-bk="back">← Back</button>' + names.map(function (n, i) {
      return '<button type="button" data-bk="' + n + '"' + (i === 0 ? ' aria-selected="true"' : "") + ">" + esc(opts.pages[n].label || n) + "</button>";
    }).join("") + (opts.print ? '<button type="button" data-bk="print">Print</button>' : "") + '</nav><section id="bkBody"></section>';
    document.body.appendChild(bs);
    function close() { bs.remove(); document.removeEventListener("keydown", key); if (opts.onClose) opts.onClose(); }
    function key(e) { if (e.key === "Escape") close(); }
    document.addEventListener("keydown", key);
    function show(n) {
      bs.querySelectorAll("[data-bk]").forEach(function (b) { b.setAttribute("aria-selected", String(b.dataset.bk === n)); });
      var pg = opts.pages[n], s = bs.querySelector("#bkBody");
      if (typeof pg.html === "function") { s.innerHTML = pg.html(); return; }
      s.innerHTML = "<h2>" + esc(pg.title || pg.label) + '</h2><div class="bgrid">' + (pg.items || []).map(function (it) {
        return '<button type="button" data-a="' + esc(it.a) + '"><i>' + it.icon + "</i><span><b>" + esc(it.title) + "</b><small>" + esc(it.sub || "") + "</small></span></button>";
      }).join("") + "</div>";
    }
    bs.onclick = function (e) {
      var nb = e.target.closest("[data-bk]");
      if (nb) { if (nb.dataset.bk === "back") close(); else if (nb.dataset.bk === "print") { close(); setTimeout(function () { (opts.print === true ? window.print : opts.print)(); }, 120); } else show(nb.dataset.bk); return; }
      var a = e.target.closest("[data-a]"); if (a) { close(); opts.onPick(a.dataset.a); }
    };
    show(names[0]);
    return { close: close };
  }

  window.Office = {
    ribbon: ribbon, backstage: backstage,
    storage: storage, storageBar: storageBar, size: size,
    install: install, canInstall: canInstall, isApp: isApp, setImporter: setImporter, sb: function () { return sb; },
    cleanHTML: cleanHTML, shrinkImage: shrinkImage,
    init: init, list: list, get: get, save: save, remove: remove, uuid: uuid, esc: esc,
    download: download, safeName: safeName, toast: toast, filesDialog: filesDialog,
    autosaver: autosaver, themeToggle: themeToggle,
    isCloud: function () { return cloud; }, user: function () { return user; }
  };
})();
