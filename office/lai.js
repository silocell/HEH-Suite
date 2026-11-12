/* Launchpad AI inside the Suite: one side panel every app can use.

   An app tells it how to read the document and how to put an answer in:

     Office.lai({
       app: "Write",
       doc: function () { return { title: "...", text: "..." }; },   // the whole document as text
       selection: function () { return "..."; },                      // what's selected ("" if nothing)
       insert: function (md, html) { ... },                           // put an answer in at the cursor
       replace: function (md, html) { ... },                          // put an answer in place of the selection (optional)
       actions: [{ label: "✏️ Improve", ask: "Improve this writing" }, ...]
     });

   Answers come from Launchpad AI (our own AI, with its own allowance) through the same server
   as ai.civilinitiatives.org, so they get the same personal-information filter and safety
   check. The chat is saved like any other Launchpad AI chat, so it can be carried on there.
   HWLAI web search can be switched on too (it uses HWLAI's allowance). */
(function () {
  "use strict";
  var cfg = null, thread = null, busy = false, stopper = null, web = false, built = false;
  try { web = localStorage.getItem("hl.lai.web") === "1"; } catch (e) {}
  var esc = function (s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); };
  var $ = function (id) { return document.getElementById(id); };

  /* a small, safe Markdown: headings, lists, bold, italics, code, links to http(s) only */
  function inline(t) {
    return esc(t).replace(/`([^`]+)`/g, "<code>$1</code>").replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>").replace(/(^|[^*])\*([^*\n]+)\*/g, "$1<em>$2</em>")
      .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>');
  }
  function md(t) {
    var out = [], list = null, code = null, tbl = null;
    var endTable = function () { if (!tbl) return; out.push("<table>" + tbl.map(function (r, i) { var c = i === 0 ? "th" : "td"; return "<tr>" + r.map(function (v) { return "<" + c + ">" + inline(v) + "</" + c + ">"; }).join("") + "</tr>"; }).join("") + "</table>"); tbl = null; };
    String(t || "").replace(/\r/g, "").split("\n").forEach(function (ln) {
      if (code === null && /^\s*\|.*\|\s*$/.test(ln)) {   // a Markdown table row
        if (list) { out.push("</" + list + ">"); list = null; }
        if (/^\s*\|[\s:|-]+\|\s*$/.test(ln)) return;    // the |---|---| line
        (tbl = tbl || []).push(ln.trim().slice(1, -1).split("|").map(function (v) { return v.trim(); })); return;
      }
      endTable();
      if (code !== null) { if (/^```/.test(ln)) { out.push("<pre><code>" + esc(code.join("\n")) + "</code></pre>"); code = null; } else code.push(ln); return; }
      if (/^```/.test(ln)) { if (list) { out.push("</" + list + ">"); list = null; } code = []; return; }
      var h = /^(#{1,4})\s+(.*)$/.exec(ln), ul = /^\s*[-*•]\s+(.*)$/.exec(ln), ol = /^\s*\d+[.)]\s+(.*)$/.exec(ln);
      if ((ul && list !== "ul") || (ol && list !== "ol") || (!ul && !ol && list)) { if (list) out.push("</" + list + ">"); list = null; }
      if (ul || ol) { if (!list) { list = ul ? "ul" : "ol"; out.push("<" + list + ">"); } out.push("<li>" + inline((ul || ol)[1]) + "</li>"); return; }
      if (h) { var n = Math.min(4, h[1].length + 1); out.push("<h" + n + ">" + inline(h[2]) + "</h" + n + ">"); return; }
      if (ln.trim()) out.push("<p>" + inline(ln) + "</p>");
    });
    endTable();
    if (code !== null) out.push("<pre><code>" + esc(code.join("\n")) + "</code></pre>");
    if (list) out.push("</" + list + ">");
    return out.join("");
  }

  function sbClient() { return window.Office && Office.sb ? Office.sb() : null; }
  function css() {
    var s = document.createElement("style");
    s.textContent = [
      ".lai{position:fixed;top:0;right:0;bottom:0;width:min(400px,100vw);background:var(--card);color:var(--ink);border-left:1px solid var(--rule);box-shadow:-12px 0 40px -18px rgba(0,0,0,.4);z-index:80;display:flex;flex-direction:column;font:14px/1.5 var(--ui)}",
      ".lai header{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid var(--rule)}",
      ".lai header b{font-size:1rem}.lai header .sp{flex:1}",
      ".lai .x{border:0;background:transparent;border-radius:8px;padding:5px 8px;color:var(--soft)}.lai .x:hover{background:var(--hover);color:var(--ink)}",
      ".lai .mk{display:grid;place-items:center;width:26px;height:26px;border-radius:8px;background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#fff;font-size:.9rem}",
      ".lai .chips{display:flex;flex-wrap:wrap;gap:6px;padding:10px 12px;border-bottom:1px solid var(--rule)}",
      ".lai .chips button{border:1px solid var(--rule);background:var(--tool);border-radius:999px;padding:5px 10px;font-size:.8rem}.lai .chips button:hover{border-color:#7c3aed}",
      ".lai .log{flex:1 1 auto;min-height:0;overflow:auto;overscroll-behavior:contain;padding:12px;display:flex;flex-direction:column;gap:12px}",
      ".lai .log>*{flex:none}",
      ".lai .u{align-self:flex-end;background:var(--on);border-radius:14px;padding:7px 11px;max-width:88%;white-space:pre-wrap}",
      ".lai .a{max-width:100%}.lai .a .body p{margin:0 0 .5em}.lai .a .body :is(ul,ol){margin:0 0 .5em;padding-left:1.3em}.lai .a .body pre{background:var(--tool);padding:8px;border-radius:8px;overflow:auto;font-size:.82rem}",
      ".lai .a .body h2,.lai .a .body h3,.lai .a .body h4{font-size:1rem;margin:.4em 0}",
      ".lai .a .body table{border-collapse:collapse;margin:0 0 .5em;font-size:.84rem}.lai .a .body :is(th,td){border:1px solid var(--rule);padding:3px 8px;text-align:left}.lai .a .body th{background:var(--tool)}",
      ".lai .acts{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px}.lai .acts button{border:1px solid var(--rule);background:var(--card);border-radius:8px;padding:4px 9px;font-size:.8rem;font-weight:600}.lai .acts button.go{background:#4f46e5;border-color:#4f46e5;color:#fff}",
      ".lai .by{font-size:.74rem;color:var(--soft);margin-top:4px}.lai .err{color:var(--bad)}.lai .note{font-size:.8rem;color:var(--soft);text-align:center}",
      ".lai .think{color:var(--soft);display:flex;gap:8px;align-items:center}.lai .think i{width:8px;height:8px;border-radius:50%;background:#7c3aed;animation:laip 1s ease-in-out infinite}@keyframes laip{50%{opacity:.25;transform:scale(.7)}}",
      ".lai form{border-top:1px solid var(--rule);padding:10px 12px;display:flex;flex-direction:column;gap:8px}",
      ".lai textarea{width:100%;resize:none;min-height:42px;max-height:160px;border:1px solid var(--rule);border-radius:12px;padding:8px 10px;background:var(--tool)}.lai textarea:focus{outline:2px solid #7c3aed;outline-offset:-1px}",
      ".lai .row{display:flex;align-items:center;gap:8px}.lai .row .sp{flex:1}",
      ".lai .sw{display:inline-flex;align-items:center;gap:6px;font-size:.8rem;color:var(--soft);cursor:pointer}.lai .sw input{accent-color:#7c3aed}",
      ".lai .send{border:0;border-radius:10px;width:36px;height:34px;background:var(--ink);color:var(--card);font-size:1rem}.lai .send:disabled{opacity:.35}",
      ".lai .fine{font-size:.72rem;color:var(--soft);margin:0}",
      ".lai .srcs{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px}.lai .srcs a{font-size:.74rem;border:1px solid var(--rule);border-radius:999px;padding:2px 8px;color:inherit;text-decoration:none}",
      "@media (max-width:640px){.lai .chips{flex-wrap:nowrap;overflow-x:auto;scrollbar-width:none}.lai .chips button{flex:none}.lai{top:auto;left:0;right:0;width:auto;max-width:100%;height:82vh;border-left:0;border-top:1px solid var(--rule);border-radius:16px 16px 0 0}}"
    ].join("\n");
    document.head.appendChild(s);
  }
  function build() {
    if (built) return; built = true; css();
    var p = document.createElement("aside"); p.className = "lai"; p.id = "laiPanel"; p.hidden = true; p.setAttribute("aria-label", "Launchpad AI");
    p.innerHTML = '<header><span class="mk" aria-hidden="true">✦</span><b>Launchpad AI</b><span class="sp"></span>' +
      '<button type="button" class="x" id="laiNew" title="Start again">New</button><button type="button" class="x" id="laiClose" aria-label="Close Launchpad AI">✕</button></header>' +
      '<div class="chips" id="laiChips"></div><div class="log" id="laiLog" aria-live="polite"></div>' +
      '<form id="laiForm"><textarea id="laiText" rows="1" maxlength="4000" placeholder="Ask about your ' + esc((cfg.noun || "document")) + '…" aria-label="Ask Launchpad AI"></textarea>' +
      '<div class="row"><label class="sw" title="HWLAI searches the web first and answers with sources (uses HWLAI\'s allowance)"><input type="checkbox" id="laiWeb"> 🔎 HWLAI web search</label><span class="sp"></span><button type="submit" class="send" id="laiSend" aria-label="Send">↑</button></div>' +
      '<p class="fine">It reads your ' + esc(cfg.noun || "document") + ' to answer. Don\'t share personal information. AI can make mistakes.</p></form>';
    document.body.appendChild(p);
    $("laiChips").innerHTML = (cfg.actions || []).map(function (a, i) { return '<button type="button" data-i="' + i + '">' + esc(a.label) + "</button>"; }).join("");
    $("laiChips").onclick = function (e) { var b = e.target.closest("[data-i]"); if (b) ask(cfg.actions[+b.dataset.i].ask, cfg.actions[+b.dataset.i].label); };
    $("laiClose").onclick = function () { open(false); };
    $("laiNew").onclick = function () { if (busy) return; thread = null; $("laiLog").innerHTML = ""; hello(); };
    $("laiWeb").checked = web; $("laiWeb").onchange = function () { web = this.checked; try { localStorage.setItem("hl.lai.web", web ? "1" : "0"); } catch (e) {} };
    var ta = $("laiText"); ta.addEventListener("input", function () { ta.style.height = "auto"; ta.style.height = Math.min(160, ta.scrollHeight) + "px"; });
    ta.addEventListener("keydown", function (e) { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); $("laiForm").requestSubmit ? $("laiForm").requestSubmit() : $("laiSend").click(); } });
    $("laiForm").onsubmit = function (e) { e.preventDefault(); if (busy) { if (stopper) stopper.abort(); return; } var t = ta.value.trim(); if (!t) return; ta.value = ""; ta.style.height = ""; ask(t); };
    document.addEventListener("keydown", function (e) { if (e.key === "Escape" && !p.hidden && !busy) open(false); });
    hello();
  }
  function hello() { $("laiLog").innerHTML = '<p class="note">Ask anything about your ' + esc(cfg.noun || "document") + ', or pick a button above. Select some text first to work on just that part.</p>'; }
  function open(on) {
    build(); var p = $("laiPanel"); if (on == null) on = p.hidden;
    p.hidden = !on; var b = $("laiBtn"); if (b) b.setAttribute("aria-expanded", on ? "true" : "false");
    if (on) setTimeout(function () { $("laiText").focus(); }, 30); else if (b) b.focus();
  }
  function setBusy(on) { busy = on; var s = $("laiSend"); s.textContent = on ? "■" : "↑"; s.setAttribute("aria-label", on ? "Stop" : "Send"); }

  function ask(question, label) {
    if (busy) return;
    var sb = sbClient(); if (!sb) { Office.toast("Sign in on the Launchpad to use Launchpad AI.", true); return; }
    var sel = ""; try { sel = (cfg.selection && cfg.selection()) || ""; } catch (e) {}
    var d = { title: "", text: "" }; try { d = cfg.doc() || d; } catch (e) {}
    var files = [];
    if (sel.trim()) files.push({ name: "Selected text", text: sel.slice(0, 12000) });
    if (String(d.text || "").trim()) files.push({ name: (d.title || "Untitled") + " (the whole " + (cfg.noun || "document") + ")", text: String(d.text).slice(0, 26000) });
    var log = $("laiLog"); if (log.querySelector(".note") && log.children.length === 1) log.innerHTML = "";
    var u = document.createElement("div"); u.className = "u"; u.textContent = label ? label + (sel.trim() ? " (selected text)" : "") : question; log.appendChild(u);
    var a = document.createElement("div"); a.className = "a"; a.innerHTML = '<div class="think"><i></i><span>Launchpad AI is thinking…</span></div>'; log.appendChild(a); log.scrollTop = log.scrollHeight;
    setBusy(true); stopper = window.AbortController ? new AbortController() : null;
    var live = "", body = null;
    var payload = { action: "send", stream: true, model: "inhouse", app: cfg.app, text: question, files: files, thread: thread };
    if (web) payload.search = true;
    sb.auth.getSession().then(function (s) {
      var tok = s.data && s.data.session && s.data.session.access_token;
      if (!tok) throw new Error("signin");
      // the same server and public key the Suite already uses
      var api = (Office.api && Office.api()) || { url: sb.supabaseUrl, key: sb.supabaseKey };
      return fetch(String(api.url).replace(/\/+$/, "") + "/functions/v1/hwchat", { method: "POST", headers: { Authorization: "Bearer " + tok, apikey: api.key, "content-type": "application/json" }, body: JSON.stringify(payload), signal: stopper ? stopper.signal : undefined });
    }).then(function (r) {
      if ((r.headers.get("content-type") || "").indexOf("text/event-stream") < 0 || !r.body) return r.json();
      var rd = r.body.getReader(), dec = new TextDecoder(), buf = "", done = null;
      function pump() {
        return rd.read().then(function (x) {
          if (x.done) return done || { error: "The answer stopped early - try again." };
          buf += dec.decode(x.value, { stream: true }); var i;
          while ((i = buf.indexOf("\n\n")) >= 0) {
            var ev = buf.slice(0, i); buf = buf.slice(i + 2);
            ev.split("\n").forEach(function (ln) {
              if (ln.indexOf("data:") !== 0) return;
              try {
                var o = JSON.parse(ln.slice(5));
                if (o.thread && !thread) thread = o.thread;
                if (o.t) { live += o.t; if (!body) { a.innerHTML = '<div class="body"></div>'; body = a.querySelector(".body"); } body.innerHTML = md(live); log.scrollTop = log.scrollHeight; }
                else if (o.done) done = o;
                else if (o.s && !body) { var st = a.querySelector(".think span"); if (st) st.textContent = o.s === "searching" ? "Searching the web…" : o.s === "reading" ? "Reading " + o.n + " pages…" : o.s === "thinking" ? "Thinking it through…" : o.s === "next" ? "Another AI is answering instead…" : st.textContent; }
              } catch (e) {}
            });
          }
          return pump();
        });
      }
      return pump();
    }).then(function (res) {
      if (res && res.thread) thread = res.thread;
      if (!res || res.error) { a.innerHTML = '<p class="err">' + esc((res && res.error) || "Something went wrong - try again.") + "</p>"; return; }
      show(a, res.reply, res.by, res.sources, !!sel.trim(), res.note);
    }).catch(function (e) {
      if (stopper && stopper.signal.aborted) { if (live.trim()) show(a, live, "Stopped", null, !!sel.trim()); else a.innerHTML = '<p class="note">Stopped.</p>'; return; }
      a.innerHTML = '<p class="err">' + (e && e.message === "signin" ? "Sign in on the Launchpad to use Launchpad AI." : "Couldn't reach Launchpad AI - check your internet and try again.") + "</p>";
    }).then(function () { setBusy(false); stopper = null; });
  }
  function show(a, text, by, sources, hadSel, note) {
    a.innerHTML = '<div class="body">' + md(text) + "</div>" +
      (sources && sources.length ? '<div class="srcs">' + sources.map(function (s) { return /^https?:/.test(s.url) ? '<a href="' + esc(s.url) + '" target="_blank" rel="noopener noreferrer nofollow">' + (+s.n) + " " + esc(s.site) + "</a>" : ""; }).join("") + "</div>" : "") +
      '<div class="acts">' + (cfg.insert ? '<button type="button" class="go" data-act="insert">' + esc(cfg.insertLabel || "Insert") + "</button>" : "") +
      (cfg.replace && hadSel ? '<button type="button" data-act="replace">Replace selection</button>' : "") +
      '<button type="button" data-act="copy">Copy</button>' + (thread ? '<a class="x" style="font-size:.78rem;text-decoration:none" href="https://ai.civilinitiatives.org/?c=' + encodeURIComponent(thread) + '" target="_blank" rel="noopener">Open in Launchpad AI ↗</a>' : "") + "</div>" +
      (by ? '<div class="by">' + esc(by) + "</div>" : "") + (note ? '<p class="note">' + esc(note) + "</p>" : "");
    a.querySelector(".acts").onclick = function (e) {
      var b = e.target.closest("[data-act]"); if (!b) return;
      var clean = text.replace(/\s*\[\d{1,2}\]/g, "");   // web search markers don't belong in the document
      var h = md(clean), one = /^<p>([\s\S]*)<\/p>$/.exec(h);
      if (one && one[1].indexOf("<p>") < 0) h = one[1];   // a single paragraph goes in as words, not a new paragraph
      if (b.dataset.act === "insert") { if (cfg.insert(clean, h) !== false) Office.toast(cfg.insertDone || "Put into your " + (cfg.noun || "document") + "."); }
      else if (b.dataset.act === "replace") { cfg.replace(clean, h); Office.toast("Replaced the selected text."); }
      else if (navigator.clipboard) navigator.clipboard.writeText(clean).then(function () { Office.toast("Copied."); });
    };
    var log = $("laiLog"); log.scrollTop = log.scrollHeight;
  }

  /* the ✦ button in the top bar; hidden for anyone whose Launchpad AI is switched off */
  function button() {
    var top = document.querySelector("header.top"); if (!top || $("laiBtn")) return;
    var b = document.createElement("button"); b.type = "button"; b.className = "tbtn"; b.id = "laiBtn"; b.setAttribute("aria-expanded", "false"); b.setAttribute("aria-controls", "laiPanel");
    b.innerHTML = '✦<span class="lbl"> Launchpad AI</span>'; b.title = "Launchpad AI: help with this " + (cfg.noun || "document");
    b.onclick = function () { open(); };
    var before = $("filesBtn"); if (before) top.insertBefore(b, before); else top.appendChild(b);
    var sb = sbClient();
    if (sb) sb.auth.getSession().then(function (s) {
      if (!s.data || !s.data.session) return;
      return sb.rpc("my_features").then(function (f) { if (f && f.data && f.data.hwchat === false) b.hidden = true; });
    }).catch(function () {});
  }
  window.Office = window.Office || {};
  Office.lai = function (c) {
    cfg = c;
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", button); else button();
  };
  Office.laiAsk = function (q) { open(true); ask(q); };
  Office.laiMd = md;
})();
