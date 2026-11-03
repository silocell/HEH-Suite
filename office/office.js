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
      return sb.from("office_files").select("id,title,updated_at").eq("kind", kind)
        .order("updated_at", { ascending: false }).limit(300).then(function (r) {
          if (r.error) throw r.error; return r.data || [];
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
      '<label class="of-btn ghost">Open a file from your device<input type="file" accept="' + esc(opts.accept) + '" hidden></label></div>' +
      '<ul class="of-list"><li class="of-empty">Loading…</li></ul>';
    var ul = dlg.querySelector(".of-list");
    function draw() {
      list().then(function (rows) {
        if (!rows.length) { ul.innerHTML = '<li class="of-empty">Nothing here yet.</li>'; return; }
        ul.innerHTML = rows.map(function (r) {
          return '<li' + (r.id === opts.current ? ' class="cur"' : '') + '><button type="button" class="of-open" data-id="' + esc(r.id) + '">' +
            '<span class="of-ic">' + opts.icon + '</span><span><strong>' + esc(r.title || "Untitled") + '</strong><small>' + esc(when(r.updated_at)) + '</small></span></button>' +
            '<button type="button" class="of-del" data-del="' + esc(r.id) + '" aria-label="Delete ' + esc(r.title || "Untitled") + '">🗑</button></li>';
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
      var f = this.files && this.files[0]; if (!f) return; dlg.close(); opts.onImport(f);
    };
    draw();
    if (dlg.showModal) dlg.showModal(); else dlg.setAttribute("open", "");
  }
  /* A second tap on the same bin confirms, so no browser pop-up is needed. */
  function confirmInline(btn, label) {
    if (btn.dataset.armed) return true;
    btn.dataset.armed = "1"; btn.textContent = "Delete?"; btn.classList.add("armed"); btn.setAttribute("aria-label", label);
    setTimeout(function () { if (btn.isConnected) { delete btn.dataset.armed; btn.textContent = "🗑"; btn.classList.remove("armed"); } }, 3000);
    return false;
  }

  /* ---- a debounced autosaver with a status line ---- */
  function autosaver(statusEl, getFile) {
    var t = null, busy = false, again = false;
    function label(s) { statusEl.textContent = s; }
    function run() {
      if (busy) { again = true; return; }
      busy = true; label("Saving…");
      save(getFile()).then(function () {
        label(cloud ? "✓ Saved to your account" : "✓ Saved on this device");
      }).catch(function (e) {
        label("⚠ Not saved"); toast("Couldn't save: " + (e.message || e), true);
      }).then(function () { busy = false; if (again) { again = false; run(); } });
    }
    return {
      poke: function () { label("Editing…"); clearTimeout(t); t = setTimeout(run, 1200); },
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
        if (n.className && typeof n.className === "string" && /\b(doc-title|todo|done)\b/.test(n.className))
          el.className = n.className.match(/\b(doc-title|todo|done)\b/g).join(" ");
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

  window.Office = {
    cleanHTML: cleanHTML, shrinkImage: shrinkImage,
    init: init, list: list, get: get, save: save, remove: remove, uuid: uuid, esc: esc,
    download: download, safeName: safeName, toast: toast, filesDialog: filesDialog,
    autosaver: autosaver, themeToggle: themeToggle,
    isCloud: function () { return cloud; }, user: function () { return user; }
  };
})();
