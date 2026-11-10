/* Launchpad Forms - make a form or a quiz, share the link, see the answers.
   Only Launchpad accounts can answer. Quizzes are marked on the server, so
   the right answers never reach the people answering. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var esc = Office.esc, main = $("main"), sb = null, user = null;
  var TYPES = { short: "Short answer", long: "Paragraph", choice: "Multiple choice", checks: "Tick boxes", dropdown: "Drop-down", scale: "Scale 1-5", date: "Date" };
  var form = null, key = {}, tab = "q", selQ = null, saveT = null;
  Office.themeToggle($("themeBtn"));
  function uid() { return Math.random().toString(36).slice(2, 10); }
  function status(t) { $("status").textContent = t || ""; }
  var SITE = "https://hosting.civilinitiatives.org/f/";

  /* Saving, and working on a form together with other teachers live. */
  var saver = Office.autosaver($("status"), function () {
    return form ? { id: form.id, title: form.title, body: { body: form.body, key: key, settings: form.settings, open: form.open } } : { id: null, title: "", body: {} };
  }, {
    shareLabel: "Editors", editOnly: true, surface: main,
    save: function (f) {
      if (!form || f.id !== form.id) return Promise.resolve(f);
      return sb.from("forms").update({ title: form.title, body: form.body, answer_key: key, settings: form.settings, open: form.open }).eq("id", form.id)
        .then(function (r) { if (r.error) throw r.error; return f; });
    },
    where: function () { return selQ ? '.q[data-q="' + selQ + '"]' : ""; },
    apply: function (b, title) {
      if (!form) return;
      var nb = b.body && typeof b.body === "object" ? b.body : {}; nb.questions = Array.isArray(nb.questions) ? nb.questions : [];
      nb.questions.forEach(function (q) { q.options = Array.isArray(q.options) ? q.options : []; });
      form.title = String(title || "Untitled").slice(0, 150); form.body = nb; key = b.key && typeof b.key === "object" ? b.key : {};
      form.settings = b.settings && typeof b.settings === "object" ? b.settings : {}; form.open = b.open !== false;
      if (tab === "r") return;
      var fs = focusNow(); render(); focusBack(fs);
    }
  });
  /* where the cursor was, so a change from someone else doesn't move it */
  function focusNow() {
    var a = document.activeElement; if (!a || !main.contains(a) || !/INPUT|TEXTAREA/.test(a.tagName)) return null;
    var card = a.closest("[data-q]"), what = a.id ? "#" + a.id : a.matches("input.qt") ? "input.qt" : a.matches("input[data-key]") ? "input[data-key]" : a.matches("input[data-pts]") ? "input[data-pts]" : a.matches("input.ot") ? "input.ot" : null;
    if (!what) return null;
    var idx = what === "input.ot" ? Array.prototype.indexOf.call(card.querySelectorAll("input.ot"), a) : 0;
    var st = null; try { st = a.selectionStart; } catch (e) {}
    return { q: card && card.dataset.q, what: what, idx: idx, at: st, val: a.value };
  }
  function focusBack(f) {
    if (!f) return;
    var root = f.q ? main.querySelector('[data-q="' + f.q + '"]') : main; if (!root) return;
    var el = f.what === "input.ot" ? root.querySelectorAll("input.ot")[f.idx] : root.querySelector(f.what); if (!el) return;
    el.focus();
    if (f.at == null) return;
    var o = f.val, nt = el.value, at = f.at, p = 0, q = 0, mn = Math.min(o.length, nt.length);
    while (p < mn && o.charCodeAt(p) === nt.charCodeAt(p)) p++;
    while (q < mn - p && o.charCodeAt(o.length - 1 - q) === nt.charCodeAt(nt.length - 1 - q)) q++;
    at = at <= p ? at : at >= o.length - q ? at + nt.length - o.length : nt.length - q;
    try { el.selectionStart = el.selectionEnd = at; } catch (e) {}
  }

  /* ---------- start ---------- */
  Office.init("form").then(function (st) {
    sb = Office.sb(); user = st.user;
    var f = new URLSearchParams(location.search).get("f"), edit = new URLSearchParams(location.search).get("edit");
    if (!user) { signIn(f); return; }
    if (f) return answer(f);
    if (edit) return openForm(edit);
    return myForms();
  });
  $("mineBtn").onclick = function () { history.pushState(null, "", "/forms/"); myForms(); };
  function signIn(f) {
    main.innerHTML = '<div class="card center"><div style="font-size:3rem">🔐</div><h2>Sign in to ' + (f ? "answer this form" : "make forms and quizzes") + "</h2>" +
      '<p style="color:var(--soft)">Forms are only for Launchpad accounts, so you know who answered.</p>' +
      '<p><a class="of-btn" href="https://civilinitiatives.org/">Sign in on the Launchpad</a></p><p style="color:var(--soft);font-size:.88rem">Then come back to this page.</p></div>';
  }
  function serverMissing() {
    main.innerHTML = '<div class="card center"><div style="font-size:3rem">🛠️</div><h2>Forms need one server update</h2><p style="color:var(--soft)">The Launchpad\'s server doesn\'t have the forms tables yet. Once the update is run, this page works.</p></div>';
  }

  /* ---------- my forms ---------- */
  /* the File button, tabs and ribbon, as in Write */
  var ICON = { short: "✏️", long: "📝", choice: "🔘", checks: "☑️", dropdown: "🔽", scale: "⭐", date: "📅" };
  function drawBar(mode) {
    var fb = $("fbar"); if (!fb) return;
    var tabs = mode === "edit" ? [["q", "Questions"], ["r", "Answers"], ["s", "Share"], ["p", "Preview"]] : [["home", "Home"]];
    var on = mode === "edit" ? tab : "home";
    var h = '<nav class="tabsbar" role="tablist" aria-label="Ribbon"><button type="button" class="file" id="fileTab">File</button>' +
      tabs.map(function (t) { return '<button type="button" role="tab" data-ftab="' + t[0] + '" aria-selected="' + (t[0] === on) + '">' + t[1] + "</button>"; }).join("") + "</nav>";
    if (mode !== "edit") h += '<div class="ribbon" data-panel="home" role="tabpanel" aria-label="Home"><div class="rg"><div class="row">' +
      '<button class="big" data-fnew="form" title="New form"><i>📋</i>New form</button><button class="big" data-fnew="quiz" title="New quiz"><i>🧠</i>New quiz</button></div><div class="gl">New</div></div></div>';
    else if (tab === "q") h += '<div class="ribbon" data-panel="q" role="tabpanel" aria-label="Questions"><div class="rg"><div class="row">' +
      Object.keys(TYPES).map(function (t) { return '<button class="big" data-fadd="' + t + '" title="Add a ' + TYPES[t].toLowerCase() + ' question"><i>' + ICON[t] + "</i>" + esc(TYPES[t]) + "</button>"; }).join("") +
      '</div><div class="gl">Add a question</div></div><div class="rg"><div class="row"><button class="big" data-ftab-go="p" title="See it as people will"><i>👁</i>Preview</button><button class="big" data-ftab-go="s" title="Share the link"><i>🔗</i>Share</button></div><div class="gl">Send</div></div></div>';
    fb.innerHTML = h;
    fb.querySelectorAll("[data-ftab]").forEach(function (b) { b.onclick = function () { if (mode === "edit") { tab = b.dataset.ftab; render(); } }; });
    fb.querySelectorAll("[data-ftab-go]").forEach(function (b) { b.onclick = function () { tab = b.dataset.ftabGo; render(); }; });
    fb.querySelectorAll("[data-fnew]").forEach(function (b) { b.onclick = function () { newForm(b.dataset.fnew === "quiz"); }; });
    fb.querySelectorAll("[data-fadd]").forEach(function (b) { b.onclick = function () { var a = document.querySelector('#pane [data-add="' + b.dataset.fadd + '"]'); if (a) a.click(); }; });
    $("fileTab").onclick = function () {
      Office.backstage({
        pages: {
          home: { label: "Home", title: form ? form.title : "Forms", items: [
            { a: "new-form", icon: "📋", title: "New form", sub: "A survey - see everyone's answers" },
            { a: "new-quiz", icon: "🧠", title: "New quiz", sub: "Questions with right answers and a score" },
            { a: "mine", icon: "📂", title: "Open", sub: "Your forms and quizzes" }] },
          info: { label: "Info", html: function () { return form ? '<h2>Info</h2><dl class="info"><dt>Name</dt><dd>' + esc(form.title) + "</dd><dt>Kind</dt><dd>" + (form.settings.quiz ? "Quiz" : "Form") + "</dd><dt>Questions</dt><dd>" + form.body.questions.length + "</dd><dt>Open</dt><dd>" + (form.open ? "Yes - taking answers" : "No") + "</dd></dl>" : "<h2>Info</h2><p>Open a form to see its details.</p>"; } }
        },
        onPick: function (a) { if (a === "new-form") newForm(false); else if (a === "new-quiz") newForm(true); else if (a === "mine") { history.pushState(null, "", location.pathname); myForms(); } }
      });
    };
  }
  function myForms() {
    form = null; status(""); saver.opened();
    drawBar("list");
    sb.from("forms").select("id,title,updated_at,open,settings,owner").order("updated_at", { ascending: false }).limit(200).then(function (r) {
      if (r.error) { serverMissing(); return; }
      var rows = r.data || [];
      main.innerHTML = '<div class="card headcard"><h1 style="margin:0 0 6px">Forms &amp; Quizzes</h1><p style="margin:0;color:var(--soft)">Make a survey or a quiz, share the link with your class, and see the answers come in.</p>' +
        '<div class="of-new" style="margin-top:14px"><button type="button" class="of-btn" data-new="form">＋ New form</button><button type="button" class="of-btn" data-new="quiz" style="background:#0d9488">＋ New quiz</button></div></div>' +
        (rows.length ? '<h2 style="font-size:1.05rem">Your forms</h2><div class="mylist">' + rows.map(function (f) {
          return '<a href="?edit=' + f.id + '"><span><b>' + esc(f.title) + "</b><br><small>" + (user && f.owner && f.owner !== user.id ? "👥 Shared with you · " : "") + ((f.settings || {}).quiz ? "Quiz" : "Form") + " · " + (f.open ? "open" : "closed") + "</small></span><small>" + new Date(f.updated_at).toLocaleDateString("en-GB", { day: "numeric", month: "short" }) + "</small></a>";
        }).join("") + "</div>" : '<p style="color:var(--soft);text-align:center">You haven\'t made any yet.</p>');
      main.querySelectorAll("[data-new]").forEach(function (b) { b.onclick = function () { newForm(b.dataset.new === "quiz"); }; });
      main.querySelectorAll(".mylist a").forEach(function (a) { a.onclick = function (e) { e.preventDefault(); var id = a.getAttribute("href").split("=")[1]; history.pushState(null, "", "?edit=" + id); openForm(id); }; });
    });
  }
  function newForm(quiz) {
    var q1 = { id: uid(), type: "choice", q: quiz ? "What is 7 × 8?" : "How did you find the homework?", options: quiz ? ["54", "56", "64"] : ["Easy", "Just right", "Hard"], required: true, points: 1 };
    var body = { desc: "", questions: [q1] }, k = {}; if (quiz) k[q1.id] = "56";
    sb.from("forms").insert({ title: quiz ? "Untitled quiz" : "Untitled form", body: body, answer_key: k, settings: { quiz: quiz, one: quiz, showScore: true } }).select("id").single().then(function (r) {
      if (r.error) { if (/relation|does not exist|schema cache/i.test(r.error.message)) serverMissing(); else Office.toast(r.error.message, true); return; }
      history.pushState(null, "", "?edit=" + r.data.id); openForm(r.data.id);
    });
  }

  /* ---------- editing ---------- */
  function openForm(id) {
    sb.from("forms").select("*").eq("id", id).maybeSingle().then(function (r) {
      if (r.error || !r.data) { Office.toast("Couldn't open that form.", true); myForms(); return; }
      form = r.data; form.body = form.body || {}; form.body.questions = form.body.questions || []; form.settings = form.settings || {};
      key = form.answer_key || {}; tab = "q"; render(); saver.opened();
    });
  }
  function save() { saver.poke(); }
  function render() {
    var s = form.settings;
    drawBar("edit");
    main.innerHTML = '<div id="pane"></div>';
    var pane = $("pane");
    if (tab === "q") questionsPane(pane);
    else if (tab === "r") answersPane(pane);
    else if (tab === "s") sharePane(pane);
    else { pane.innerHTML = '<div class="center"><span class="preview-pill">Preview - this is what people see</span></div>'; fillPane(pane, { id: form.id, title: form.title, body: form.body, settings: s, open: true }, true); }
  }
  function questionsPane(pane) {
    var s = form.settings, qs = form.body.questions;
    pane.innerHTML = '<div class="card headcard"><input class="ft" id="ft" maxlength="150" aria-label="Form title"><textarea id="fd" rows="2" placeholder="Description (optional)" aria-label="Description"></textarea></div>' +
      qs.map(function (q, i) { return qCard(q, i, s.quiz); }).join("") +
      '<div class="addbar">' + Object.keys(TYPES).map(function (t) { return '<button type="button" class="of-btn ghost" data-add="' + t + '">＋ ' + TYPES[t] + "</button>"; }).join("") + "</div>";
    $("ft").value = form.title; $("fd").value = form.body.desc || "";
    $("ft").oninput = function () { form.title = this.value.slice(0, 150) || "Untitled"; save(); };
    $("fd").oninput = function () { form.body.desc = this.value.slice(0, 2000); save(); };
    pane.querySelectorAll("[data-add]").forEach(function (b) {
      b.onclick = function () {
        var t = b.dataset.add, q = { id: uid(), type: t, q: "", options: /choice|checks|dropdown/.test(t) ? ["Option 1", "Option 2"] : [], required: false, points: 1 };
        qs.push(q); selQ = q.id; save(); render();
        var el = pane.querySelector('[data-q="' + q.id + '"] input.qt'); if (el) { el.focus(); el.scrollIntoView({ block: "center" }); }
      };
    });
    pane.querySelectorAll(".q").forEach(function (card) {
      var q = qs.filter(function (x) { return x.id === card.dataset.q; })[0];
      card.addEventListener("focusin", function () { pane.querySelectorAll(".q").forEach(function (c) { c.classList.toggle("sel", c === card); }); selQ = q.id; });
      card.querySelector("input.qt").oninput = function () { q.q = this.value.slice(0, 500); save(); };
      card.querySelector("select").onchange = function () {
        q.type = this.value; if (/choice|checks|dropdown/.test(q.type) && !q.options.length) q.options = ["Option 1", "Option 2"];
        delete key[q.id]; save(); render();
      };
      card.querySelectorAll("input.ot").forEach(function (inp, j) { inp.oninput = function () { var old = q.options[j]; q.options[j] = this.value.slice(0, 200); fixKey(q, old, q.options[j]); save(); }; });
      card.querySelectorAll("[data-delopt]").forEach(function (b) { b.onclick = function () { var j = +b.dataset.delopt, old = q.options[j]; q.options.splice(j, 1); fixKey(q, old, null); save(); render(); }; });
      var ao = card.querySelector("[data-addopt]"); if (ao) ao.onclick = function () { q.options.push("Option " + (q.options.length + 1)); save(); render(); var all = card.parentNode.querySelectorAll('[data-q="' + q.id + '"] input.ot'); };
      card.querySelectorAll("input[data-right]").forEach(function (c) {
        c.onchange = function () {
          var o = q.options[+c.dataset.right];
          if (q.type === "checks") { var k = Array.isArray(key[q.id]) ? key[q.id] : []; k = k.filter(function (x) { return x !== o; }); if (c.checked) k.push(o); if (k.length) key[q.id] = k; else delete key[q.id]; }
          else { if (c.checked) key[q.id] = o; else delete key[q.id]; }
          save();
        };
      });
      var kt = card.querySelector("input[data-key]"); if (kt) kt.oninput = function () { var v = this.value.trim(); if (v) key[q.id] = v.slice(0, 300); else delete key[q.id]; save(); };
      var pts = card.querySelector("input[data-pts]"); if (pts) pts.oninput = function () { q.points = Math.max(0, Math.min(100, +this.value || 0)); save(); };
      card.querySelector("input[data-req]").onchange = function () { q.required = this.checked; save(); };
      card.querySelector("[data-up]").onclick = function () { var i = qs.indexOf(q); if (i > 0) { qs.splice(i, 1); qs.splice(i - 1, 0, q); save(); render(); } };
      card.querySelector("[data-down]").onclick = function () { var i = qs.indexOf(q); if (i < qs.length - 1) { qs.splice(i, 1); qs.splice(i + 1, 0, q); save(); render(); } };
      card.querySelector("[data-dup]").onclick = function () { var c = JSON.parse(JSON.stringify(q)); c.id = uid(); if (key[q.id] !== undefined) key[c.id] = JSON.parse(JSON.stringify(key[q.id])); qs.splice(qs.indexOf(q) + 1, 0, c); save(); render(); };
      card.querySelector("[data-del]").onclick = function () { if (qs.length < 2) { Office.toast("A form needs at least one question."); return; } qs.splice(qs.indexOf(q), 1); delete key[q.id]; save(); render(); };
    });
  }
  function fixKey(q, old, now) {
    var k = key[q.id]; if (k === undefined || old === null) return;
    if (Array.isArray(k)) { key[q.id] = k.map(function (x) { return x === old ? now : x; }).filter(function (x) { return x !== null; }); }
    else if (k === old) { if (now === null) delete key[q.id]; else key[q.id] = now; }
  }
  function qCard(q, i, quiz) {
    var opts = "";
    if (/choice|checks|dropdown/.test(q.type)) {
      opts = '<div class="opts">' + q.options.map(function (o, j) {
        var right = q.type === "checks" ? (Array.isArray(key[q.id]) && key[q.id].indexOf(o) >= 0) : key[q.id] === o;
        return '<div class="opt"><span class="mark ' + (q.type === "checks" ? "sq" : "round") + '"></span><input class="ot" value="' + esc(o) + '" aria-label="Option ' + (j + 1) + '">' +
          (quiz ? '<label class="correct"><input type="' + (q.type === "checks" ? "checkbox" : "radio") + '" name="k' + q.id + '" data-right="' + j + '"' + (right ? " checked" : "") + "> right answer</label>" : "") +
          '<button type="button" data-delopt="' + j + '" aria-label="Remove option">✕</button></div>';
      }).join("") + '<div class="opt"><span class="mark ' + (q.type === "checks" ? "sq" : "round") + '" style="opacity:.4"></span><button type="button" data-addopt="1" style="color:var(--a1);font-weight:700">Add option</button></div></div>';
    } else if (q.type === "scale") opts = '<div class="scale" style="margin-top:12px;opacity:.6">' + [1, 2, 3, 4, 5].map(function (n) { return "<label>" + n + "</label>"; }).join("") + "</div>";
    else opts = '<div style="margin-top:12px;color:var(--soft);border-bottom:1px dashed var(--rule);padding:8px 2px">' + (q.type === "long" ? "Long answer" : q.type === "date" ? "dd/mm/yyyy" : "Short answer") + "</div>";
    var keyRow = "";
    if (quiz && /short|scale/.test(q.type)) keyRow = '<div class="keyrow">✔ Right answer <input data-key="1" value="' + esc(key[q.id] || "") + '" placeholder="' + (q.type === "scale" ? "e.g. 5" : "e.g. Paris|paris city (| between answers that all count)") + '"></div>';
    return '<div class="card q' + (selQ === q.id ? " sel" : "") + '" data-q="' + q.id + '"><div class="qrow"><input class="qt" value="' + esc(q.q) + '" placeholder="Question ' + (i + 1) + '" aria-label="Question ' + (i + 1) + '">' +
      '<select aria-label="Question type">' + Object.keys(TYPES).map(function (t) { return '<option value="' + t + '"' + (t === q.type ? " selected" : "") + ">" + TYPES[t] + "</option>"; }).join("") + "</select></div>" + opts + keyRow +
      '<div class="qfoot"><label><input type="checkbox" data-req="1"' + (q.required ? " checked" : "") + "> Must answer</label>" +
      (quiz && q.type !== "long" && q.type !== "date" ? '<label style="margin-right:0">Points <input type="number" data-pts="1" min="0" max="100" value="' + (q.points === undefined ? 1 : q.points) + '" style="width:56px;height:30px;border:1px solid var(--rule);border-radius:8px;padding:0 6px;background:var(--tool)"></label>' : "") +
      '<button type="button" data-up="1" title="Move up" aria-label="Move up">↑</button><button type="button" data-down="1" title="Move down" aria-label="Move down">↓</button><button type="button" data-dup="1" title="Copy" aria-label="Copy">⧉</button><button type="button" data-del="1" title="Delete" aria-label="Delete">🗑</button></div></div>';
  }

  /* ---------- sharing ---------- */
  function sharePane(pane) {
    var s = form.settings, link = SITE + "?id=" + form.id;
    pane.innerHTML = '<div class="card"><h2 style="margin-top:0">Share the link</h2><p style="color:var(--soft)">Anyone with a Launchpad account can answer. Post it in the year chat or send it to your class.</p>' +
      '<div class="sharebox"><input id="lnk" readonly value="' + esc(link) + '" aria-label="Link"><button type="button" class="of-btn" id="cp">Copy link</button>' +
      (navigator.share ? '<button type="button" class="of-btn ghost" id="sh">Share…</button>' : "") + "</div></div>" +
      '<div class="card"><h2 style="margin-top:0">Settings</h2><div class="setgrid">' +
      '<label><input type="checkbox" id="sOpen"' + (form.open ? " checked" : "") + "> Taking answers (untick to close the form)</label>" +
      '<label><input type="checkbox" id="sQuiz"' + (s.quiz ? " checked" : "") + "> It's a quiz - mark the answers and give a score</label>" +
      '<label><input type="checkbox" id="sOne"' + (s.one ? " checked" : "") + "> Only one go each</label>" +
      '<label><input type="checkbox" id="sShow"' + (s.showScore !== false ? " checked" : "") + "> Show people their score straight away</label></div></div>" +
      '<div class="card"><h2 style="margin-top:0;color:#dc2626">Delete this form</h2><p style="color:var(--soft)">This also deletes every answer.</p><button type="button" class="of-btn" id="delF" style="background:#dc2626">Delete form</button></div>';
    $("cp").onclick = function () { (navigator.clipboard ? navigator.clipboard.writeText(link) : Promise.reject()).then(function () { Office.toast("Link copied."); }, function () { $("lnk").select(); document.execCommand("copy"); Office.toast("Link copied."); }); };
    if ($("sh")) $("sh").onclick = function () { navigator.share({ title: form.title, url: link }).catch(function () {}); };
    $("sOpen").onchange = function () { form.open = this.checked; save(); };
    $("sQuiz").onchange = function () { s.quiz = this.checked; save(); };
    $("sOne").onchange = function () { s.one = this.checked; save(); };
    $("sShow").onchange = function () { s.showScore = this.checked; save(); };
    $("delF").onclick = function () {
      if (!this.dataset.armed) { this.dataset.armed = 1; this.textContent = "Really delete it and all the answers?"; return; }
      sb.from("forms").delete().eq("id", form.id).then(function (r) { if (r.error) Office.toast(r.error.message, true); else { Office.toast("Deleted."); history.pushState(null, "", "/forms/"); myForms(); } });
    };
  }

  /* ---------- answers ---------- */
  function answersPane(pane) {
    pane.innerHTML = '<p class="center" style="color:var(--soft)">Loading answers…</p>';
    sb.from("form_responses").select("id,name,answers,score,max_score,created_at").eq("form_id", form.id).order("created_at", { ascending: true }).limit(2000).then(function (r) {
      if (r.error) { pane.innerHTML = '<p class="center bad">' + esc(r.error.message) + "</p>"; return; }
      var rows = r.data || [], qs = form.body.questions, quiz = form.settings.quiz;
      if (!rows.length) { pane.innerHTML = '<div class="card center"><div style="font-size:3rem">📭</div><h2>No answers yet</h2><p style="color:var(--soft)">Share the link (Share tab) and answers appear here.</p></div>'; return; }
      var h = '<div class="card"><h2 style="margin:0">' + rows.length + (rows.length === 1 ? " answer" : " answers") + "</h2>";
      if (quiz) { var sc = rows.filter(function (x) { return x.max_score; }); if (sc.length) { var avg = sc.reduce(function (a, x) { return a + x.score / x.max_score; }, 0) / sc.length; h += '<p style="margin:6px 0 0">Average score <b>' + Math.round(avg * 100) + "%</b></p>"; } }
      h += '<div class="of-new" style="margin-top:12px"><button type="button" class="of-btn" id="toGrid">📊 Open in Grid</button><button type="button" class="of-btn ghost" id="toCsv">⬇ CSV</button></div></div>';
      qs.forEach(function (q) {
        var vals = rows.map(function (x) { return x.answers[q.id]; }).filter(function (v) { return v !== undefined && v !== null && v !== "" && !(Array.isArray(v) && !v.length); });
        h += '<div class="card"><p style="font-weight:700;margin:0">' + esc(q.q || "(no question)") + '</p><p style="color:var(--soft);margin:2px 0 0;font-size:.85rem">' + vals.length + " answered" + (quiz && key[q.id] !== undefined ? " · right answer: " + esc(Array.isArray(key[q.id]) ? key[q.id].join(", ") : String(key[q.id]).split("|")[0]) : "") + "</p>";
        if (/choice|checks|dropdown|scale/.test(q.type)) {
          var opts = q.type === "scale" ? ["1", "2", "3", "4", "5"] : q.options, count = {};
          vals.forEach(function (v) { (Array.isArray(v) ? v : [v]).forEach(function (o) { count[o] = (count[o] || 0) + 1; }); });
          var max = Math.max.apply(null, opts.map(function (o) { return count[o] || 0; }).concat([1]));
          h += '<div class="bars">' + opts.map(function (o) { var n = count[o] || 0; return '<div class="bar"><span class="t">' + esc(o) + '</span><span class="track"><span class="fillb" style="display:block;width:' + (n / max * 100) + '%"></span></span><span>' + n + "</span></div>"; }).join("") + "</div>";
        } else h += '<div class="texts">' + vals.slice(-200).map(function (v) { return "<div>" + esc(String(v)) + "</div>"; }).join("") + "</div>";
        h += "</div>";
      });
      h += '<div class="card"><h2 style="margin-top:0;font-size:1.05rem">Everyone</h2><div style="overflow:auto"><table class="resp"><thead><tr><th>Name</th><th>When</th>' + (quiz ? "<th>Score</th>" : "") + "<th></th></tr></thead><tbody>" +
        rows.map(function (x) { return "<tr><td>" + esc(x.name || "?") + "</td><td>" + new Date(x.created_at).toLocaleString("en-GB", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) + "</td>" + (quiz ? "<td>" + (x.max_score ? x.score + " / " + x.max_score : "") + "</td>" : "") + '<td><button type="button" class="of-x" data-rm="' + x.id + '" aria-label="Delete this answer" title="Delete this answer" style="width:28px;height:28px">✕</button></td></tr>'; }).join("") + "</tbody></table></div></div>";
      pane.innerHTML = h;
      pane.querySelectorAll("[data-rm]").forEach(function (b) {
        b.onclick = function () { if (!b.dataset.armed) { b.dataset.armed = 1; b.textContent = "?"; return; } sb.from("form_responses").delete().eq("id", b.dataset.rm).then(function () { answersPane(pane); }); };
      });
      function table() {
        var head = ["Name", "When"].concat(quiz ? ["Score", "Out of"] : []).concat(qs.map(function (q) { return q.q || "Question"; }));
        return [head].concat(rows.map(function (x) {
          return [x.name || "", new Date(x.created_at).toLocaleString("en-GB")].concat(quiz ? [x.score === null ? "" : x.score, x.max_score === null ? "" : x.max_score] : [])
            .concat(qs.map(function (q) { var v = x.answers[q.id]; return Array.isArray(v) ? v.join(", ") : v === undefined || v === null ? "" : String(v); }));
        }));
      }
      $("toCsv").onclick = function () {
        var csv = table().map(function (r) { return r.map(function (c) { c = String(c); return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c; }).join(","); }).join("\r\n");
        Office.download(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }), Office.safeName(form.title + " - answers", ".csv"));
      };
      $("toGrid").onclick = function () {
        var t = table(), cells = {}, col = function (c) { var s = ""; c++; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
        t.forEach(function (row, i) { row.slice(0, 52).forEach(function (v, j) { if (v !== "") cells[col(j) + (i + 1)] = i === 0 ? { v: String(v), b: 1, bg: "#ede9fe" } : { v: String(v) }; }); });
        Office.init("sheet").then(function () { return Office.save({ title: form.title + " - answers", body: { sheets: [{ name: "Answers", cells: cells, w: { A: 160, B: 150 }, rows: Math.max(100, t.length + 20), cols: Math.max(26, Math.min(52, t[0].length)) }], active: 0 } }); })
          .then(function (f) { location.href = "/grid/?id=" + encodeURIComponent(f.id); })
          .catch(function (e) { Office.toast("Couldn't open Grid: " + (e.message || e), true); });
      };
    });
  }

  /* ---------- answering ---------- */
  function answer(fid) {
    form = null; saver.opened();
    sb.rpc("get_form", { fid: fid }).then(function (r) {
      if (r.error) { if (/function|schema cache/i.test(r.error.message)) serverMissing(); else main.innerHTML = '<p class="center bad">' + esc(r.error.message) + "</p>"; return; }
      var f = r.data;
      if (!f) { main.innerHTML = '<div class="card center"><div style="font-size:3rem">🤷</div><h2>That form isn\'t here</h2><p style="color:var(--soft)">It may have been deleted, or the link is wrong.</p></div>'; return; }
      document.title = f.title + " - Forms";
      if (!f.open) { main.innerHTML = '<div class="card center headcard"><h1>' + esc(f.title) + '</h1><p style="color:var(--soft)">This form has closed and isn\'t taking answers any more.</p></div>'; return; }
      if (f.settings && f.settings.one && f.mine > 0) { main.innerHTML = '<div class="card center headcard"><h1>' + esc(f.title) + '</h1><p>✅ You\'ve already answered this one.</p></div>'; return; }
      fillPane(main, f, false);
      if (f.owner) main.insertAdjacentHTML("afterbegin", '<div class="center"><a class="preview-pill" href="?edit=' + f.id + '">You made this - edit it</a></div>');
    });
  }
  function fillPane(box, f, preview) {
    var qs = (f.body && f.body.questions) || [], s = f.settings || {};
    var h = '<div class="card headcard"><h1 style="margin:0 0 6px">' + esc(f.title) + "</h1>" + (f.body.desc ? '<p style="margin:0;color:var(--soft);white-space:pre-wrap">' + esc(f.body.desc) + "</p>" : "") +
      (qs.some(function (q) { return q.required; }) ? '<p style="margin:10px 0 0;font-size:.85rem" class="req">* must be answered</p>' : "") + '</div><form id="ansForm" novalidate>';
    qs.forEach(function (q, i) {
      var name = "a_" + q.id, inner = "";
      if (q.type === "choice") inner = q.options.map(function (o) { return '<label class="choice"><input type="radio" name="' + name + '" value="' + esc(o) + '"> ' + esc(o) + "</label>"; }).join("");
      else if (q.type === "checks") inner = q.options.map(function (o) { return '<label class="choice"><input type="checkbox" name="' + name + '" value="' + esc(o) + '"> ' + esc(o) + "</label>"; }).join("");
      else if (q.type === "dropdown") inner = '<select name="' + name + '"><option value="">Choose…</option>' + q.options.map(function (o) { return "<option>" + esc(o) + "</option>"; }).join("") + "</select>";
      else if (q.type === "scale") inner = '<div class="scale">' + [1, 2, 3, 4, 5].map(function (n) { return '<label><input type="radio" name="' + name + '" value="' + n + '">' + n + "</label>"; }).join("") + "</div>";
      else if (q.type === "long") inner = '<textarea name="' + name + '" rows="4" maxlength="4000"></textarea>';
      else if (q.type === "date") inner = '<input type="date" name="' + name + '">';
      else inner = '<input type="text" name="' + name + '" maxlength="500" autocomplete="off">';
      h += '<div class="card ans" data-q="' + q.id + '"><p class="qtext">' + (i + 1) + ". " + esc(q.q || "") + (q.required ? ' <span class="req">*</span>' : "") + (s.quiz && q.points > 1 ? ' <small style="color:var(--soft);font-weight:500">(' + q.points + " points)</small>" : "") + "</p>" + inner + '<p class="fb" hidden></p></div>';
    });
    h += '<div class="center"><button type="submit" class="of-btn" id="sendBtn" style="font-size:1.05rem;padding:12px 26px">' + (preview ? "Send (preview - nothing is saved)" : "Send") + "</button></div></form>";
    box.insertAdjacentHTML("beforeend", h);
    var fm = box.querySelector("#ansForm");
    fm.onsubmit = function (e) {
      e.preventDefault();
      var ans = {}, missing = null;
      qs.forEach(function (q) {
        var els = fm.querySelectorAll('[name="a_' + q.id + '"]'), v;
        if (q.type === "checks") v = Array.prototype.filter.call(els, function (x) { return x.checked; }).map(function (x) { return x.value; });
        else if (q.type === "choice" || q.type === "scale") { var c = Array.prototype.filter.call(els, function (x) { return x.checked; })[0]; v = c ? c.value : ""; }
        else v = els[0] ? els[0].value.trim() : "";
        if (v !== "" && !(Array.isArray(v) && !v.length)) ans[q.id] = v;
        else if (q.required && !missing) missing = q;
      });
      if (missing) { var c = fm.querySelector('[data-q="' + missing.id + '"]'); c.scrollIntoView({ block: "center", behavior: "smooth" }); c.style.borderColor = "#dc2626"; Office.toast("Please answer: " + (missing.q || "this question"), true); return; }
      if (preview) { Office.toast("This is a preview - nothing was sent."); return; }
      var btn = $("sendBtn"); btn.disabled = true; btn.textContent = "Sending…";
      sb.rpc("submit_form", { fid: f.id, ans: ans }).then(function (r) {
        if (r.error) { btn.disabled = false; btn.textContent = "Send"; Office.toast(r.error.message, true); return; }
        var res = r.data || {};
        if (res.max) {
          qs.forEach(function (q) {
            if (!(q.id in (res.results || {}))) return;
            var p = fm.querySelector('[data-q="' + q.id + '"] .fb'); p.hidden = false; p.className = "fb " + (res.results[q.id] ? "good" : "bad"); p.textContent = res.results[q.id] ? "✓ Right" : "✗ Not quite";
          });
          fm.querySelectorAll("input,select,textarea").forEach(function (x) { x.disabled = true; });
          btn.remove();
          box.querySelector(".headcard").insertAdjacentHTML("beforeend", '<div class="score">' + res.score + " / " + res.max + '</div><p class="center" style="margin:0">' + (res.score === res.max ? "🎉 Full marks!" : res.score / res.max >= 0.7 ? "👏 Great work!" : "Keep practising - you've got this.") + "</p>");
          box.querySelector(".headcard").scrollIntoView({ behavior: "smooth" });
        } else {
          box.innerHTML = '<div class="card center headcard"><div style="font-size:3rem">✅</div><h1>' + esc(f.title) + "</h1><p>Thanks - your answers have been sent.</p></div>";
        }
      });
    };
  }
  window.addEventListener("popstate", function () { location.reload(); });
})();
