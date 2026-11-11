/* Homework Chat: longer conversations about homework. Free AIs up to a daily
   limit ("Auto" picks the best one for each question), Claude for pence. The
   server function "hwchat" does the safety and personal-information checks. */
(function () {
  var $ = function (id) { return document.getElementById(id); }, esc = Office.esc, sb = null, user = null;
  var thread = null, info = null, choice = "auto", busy = false, choices = [];
  var FREE_ICON = { auto: "✨", fast: "⚡", maths: "🧠", writing: "✍️", code: "💻" }, PAID_ICON = { haiku: "🪶", sonnet: "🎼" };
  try { choice = localStorage.getItem("hl.chat.ai") || "auto"; } catch (e) {}

  Office.themeToggle($("themeBtn"));
  Office.ribbon();

  /* a little safe Markdown for the answers: everything is escaped first */
  function md(t) {
    var parts = String(t || "").split(/```/), out = "";
    parts.forEach(function (p, i) {
      if (i % 2) { out += "<pre><code>" + esc(p.replace(/^[a-z0-9+-]*\n/i, "")) + "</code></pre>"; return; }
      var lines = esc(p).split("\n"), list = null, para = [];
      function flushP() { if (para.length) { out += "<p>" + para.join("<br>") + "</p>"; para = []; } }
      function flushL() { if (list) { out += "<" + list.t + ">" + list.items.map(function (x) { return "<li>" + x + "</li>"; }).join("") + "</" + list.t + ">"; list = null; } }
      lines.forEach(function (l) {
        var inl = function (s) { return s.replace(/`([^`]+)`/g, "<code>$1</code>").replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>").replace(/(^|[\s(])\*([^*\s][^*]*)\*/g, "$1<i>$2</i>"); };
        var m;
        if (!l.trim()) { flushP(); flushL(); return; }
        if ((m = l.match(/^#{1,4}\s+(.*)/))) { flushP(); flushL(); out += "<h4>" + inl(m[1]) + "</h4>"; return; }
        if ((m = l.match(/^\s*[-*•]\s+(.*)/))) { flushP(); if (!list || list.t !== "ul") { flushL(); list = { t: "ul", items: [] }; } list.items.push(inl(m[1])); return; }
        if ((m = l.match(/^\s*\d+[.)]\s+(.*)/))) { flushP(); if (!list || list.t !== "ol") { flushL(); list = { t: "ol", items: [] }; } list.items.push(inl(m[1])); return; }
        flushL(); para.push(inl(l));
      });
      flushP(); flushL();
    });
    return out;
  }

  function bubble(role, text, by) {
    var log = $("log"), h = log.querySelector(".hello"); if (h) h.remove();
    var d = document.createElement("div");
    if (role === "note") { d.className = "note"; d.textContent = text; }
    else if (role === "user") { d.className = "msg me"; d.textContent = text; }
    else { d.className = "msg ai"; d.innerHTML = md(text) + (by ? '<small class="by">' + esc(by) + "</small>" : ""); }
    log.appendChild(d); log.scrollTop = log.scrollHeight; return d;
  }
  function hello() {
    $("log").innerHTML = '<div class="hello"><div style="font-size:2.6rem">💬</div><h2>What are you working on?</h2>' +
      "<p>Ask about any homework and keep the conversation going. Auto picks the best free AI for each question.</p>" +
      '<div class="ideas">' + ["Explain photosynthesis simply", "How do I solve 3x + 7 = 22?", "Help me plan a persuasive essay", "What's a for loop in Python?"].map(function (q) {
        return '<button type="button" data-q="' + esc(q) + '">' + esc(q) + "</button>"; }).join("") + "</div></div>";
  }
  $("log").addEventListener("click", function (e) { var b = e.target.closest("[data-q]"); if (b) { $("text").value = b.dataset.q; send(); } });

  function meter() {
    if (!info) { $("meter").textContent = ""; return; }
    var left = Math.max(0, info.free_per_day - info.free_used);
    $("meter").textContent = (info.unlimited ? "Free messages: unlimited (admin)" : left + " of " + info.free_per_day + " free messages left today") + " · You have " + Number(info.pence).toLocaleString("en-GB") + "p";
  }
  function loadInfo() { return sb.rpc("hw_info").then(function (r) { if (!r.error && r.data) { info = r.data; meter(); drawChoices(); } }); }

  function drawChoices() {
    var f = $("freeAIs"), p = $("paidAIs"); f.textContent = ""; p.textContent = "";
    choices.forEach(function (c) {
      var b = document.createElement("button"); b.type = "button"; b.className = "big"; b.dataset.ai = c.id; b.title = c.sub;
      var price = c.free ? "" : info ? (c.id === "haiku" ? info.haiku_pence : info.sonnet_pence) + "p a message" : "";
      b.innerHTML = "<i>" + (FREE_ICON[c.id] || PAID_ICON[c.id] || "🤖") + "</i>" + esc(c.label) + (price ? '<small class="price">' + esc(price) + "</small>" : "");
      b.setAttribute("aria-pressed", String(c.id === choice));
      b.onclick = function () { choice = c.id; try { localStorage.setItem("hl.chat.ai", choice); } catch (e) {} drawChoices(); $("text").focus(); };
      (c.free ? f : p).appendChild(b);
    });
  }

  function listThreads() {
    return sb.from("hw_threads").select("id,title,updated_at").order("updated_at", { ascending: false }).limit(100).then(function (r) {
      var ul = $("threads"), rows = r.data || [];
      ul.innerHTML = rows.map(function (t) {
        var d = new Date(t.updated_at).toLocaleDateString("en-GB", { day: "numeric", month: "short" });
        return '<li data-id="' + esc(t.id) + '"' + (t.id === thread ? ' aria-current="true"' : "") + '><a href="?c=' + esc(t.id) + '">' + esc(t.title) + "<small>" + d + '</small></a><button type="button" data-del="' + esc(t.id) + '" aria-label="Delete this chat" title="Delete">🗑</button></li>';
      }).join("") || '<li style="padding:8px 12px;color:var(--soft);font-size:.86rem">No chats yet.</li>';
    });
  }
  $("threads").addEventListener("click", function (e) {
    var del = e.target.closest("[data-del]");
    if (del) {
      e.preventDefault();
      if (!del.dataset.armed) { del.dataset.armed = "1"; del.textContent = "Delete?"; setTimeout(function () { if (del.isConnected) { delete del.dataset.armed; del.textContent = "🗑"; } }, 3000); return; }
      sb.from("hw_threads").delete().eq("id", del.dataset.del).then(function () { if (del.dataset.del === thread) newChat(); listThreads(); });
      return;
    }
    var a = e.target.closest("a"); if (!a) return;
    e.preventDefault(); openThread(a.closest("li").dataset.id); $("side").classList.remove("open");
  });

  function openThread(id) {
    thread = id; history.replaceState(null, "", "?c=" + encodeURIComponent(id));
    $("log").innerHTML = '<p class="note">Loading…</p>';
    sb.from("hw_messages").select("role,body,model,pence").eq("thread_id", id).order("id").limit(400).then(function (r) {
      $("log").textContent = "";
      if (r.error || !r.data || !r.data.length) { hello(); return; }
      r.data.forEach(function (m) { bubble(m.role === "user" ? "user" : m.role === "note" ? "note" : "ai", m.body, m.role === "ai" ? (m.model || "") + (m.pence ? " · " + m.pence + "p" : "") : ""); });
      listThreads();
    });
  }
  function newChat() { thread = null; history.replaceState(null, "", location.pathname); hello(); listThreads(); $("text").focus(); }
  $("newBtn").onclick = $("newBtn2").onclick = newChat;
  $("chatsBtn").onclick = function () { $("side").classList.toggle("open"); };

  function send() {
    var t = $("text").value.trim(); if (!t || busy) return;
    busy = true; $("sendBtn").disabled = true; $("text").value = ""; grow();
    bubble("user", t);
    var wait = bubble("ai", ""); wait.classList.add("wait"); wait.textContent = choice === "auto" ? "Picking the best AI" : "Thinking";
    sb.functions.invoke("hwchat", { body: { action: "send", thread: thread, text: t, model: choice } }).then(function (r) {
      var d = r.data;
      if (r.error && !d) {
        // the function's own message comes back in the error's context
        return (r.error.context && r.error.context.json ? r.error.context.json() : Promise.resolve(null)).then(function (x) { d = x || { error: r.error.message || "Something went wrong." }; return d; });
      }
      return d;
    }).then(function (d) {
      wait.remove();
      if (d && d.thread && d.thread !== thread) { thread = d.thread; history.replaceState(null, "", "?c=" + encodeURIComponent(thread)); }
      if (!d || d.error) { bubble("note", (d && d.error) || "Something went wrong - try again."); if (d && d.outOfFree) { choice = "haiku"; drawChoices(); } return; }
      if (d.removed && d.removed.length) bubble("note", "🔒 I took out " + d.removed.join(", ") + " before sending - please don't share personal information.");
      bubble("ai", d.reply, (d.by || "") + (d.pence ? " · " + d.pence + "p" : ""));
    }).catch(function () { wait.remove(); bubble("note", "Couldn't reach the server - check your internet and try again."); })
      .then(function () { busy = false; $("sendBtn").disabled = false; loadInfo(); listThreads(); $("text").focus(); });
  }
  $("form").onsubmit = function (e) { e.preventDefault(); send(); };
  function grow() { var ta = $("text"); ta.style.height = "auto"; ta.style.height = Math.min(180, ta.scrollHeight) + "px"; }
  $("text").addEventListener("input", grow);
  $("text").addEventListener("keydown", function (e) { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } });
  document.addEventListener("keydown", function (e) { if (e.ctrlKey && e.shiftKey && (e.key === "O" || e.key === "o")) { e.preventDefault(); newChat(); } });

  $("fileTab").onclick = function () {
    Office.backstage({
      pages: {
        home: { label: "Home", title: "Homework Chat", items: [
          { a: "new", icon: "💬", title: "New chat", sub: "Start a fresh conversation" },
          { a: "copy", icon: "📋", title: "Copy this chat", sub: "As text, to paste into Write or your notes" }] },
        info: { label: "About", html: function () {
          return "<h2>About Homework Chat</h2><dl class=\"info\"><dt>Free AIs</dt><dd>" + (info ? (info.unlimited ? "Unlimited (admin)" : info.free_per_day + " messages a day") : "-") +
            "</dd><dt>Claude</dt><dd>" + (info ? "Haiku " + info.haiku_pence + "p, Sonnet " + info.sonnet_pence + "p a message, from your pence" : "-") +
            "</dd><dt>Auto</dt><dd>A quick AI reads each question and picks the best free AI for it - maths, writing, coding or anything else.</dd><dt>Privacy</dt><dd>Emails, phone numbers, postcodes, addresses, passwords and your name are taken out before anything is sent or saved. Only you can see your chats.</dd></dl>"; } }
      },
      onPick: function (a) {
        if (a === "new") newChat();
        else if (a === "copy") {
          var txt = Array.prototype.map.call($("log").querySelectorAll(".msg"), function (m) { return (m.classList.contains("me") ? "Me: " : "AI: ") + m.innerText; }).join("\n\n");
          (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).then(function () { Office.toast("Copied."); }, function () { Office.toast("Couldn't copy - select the text instead.", true); });
        }
      }
    });
  };

  Office.init("doc").then(function (st) {
    sb = Office.sb(); user = st.user;
    if (!user) { $("main").innerHTML = '<div class="gate"><h2>💬 Homework Chat</h2><p>Sign in on the Launchpad to chat about your homework.</p><a class="of-btn" href="https://civilinitiatives.org/">Sign in on the Launchpad</a></div>'; return; }
    sb.rpc("my_features").then(function (f) {
      if (f.data && f.data.hwchat === false) { $("main").innerHTML = '<div class="gate"><h2>💬 Homework Chat</h2><p>Homework Chat is switched off for your account just now.</p></div>'; return; }
      sb.functions.invoke("hwchat", { body: { action: "models" } }).then(function (r) {
        choices = (r.data && r.data.choices) || [];
        if (!choices.length) { $("log").innerHTML = '<div class="gate"><h2>💬 Homework Chat</h2><p>Homework Chat needs the latest Launchpad update on the server.</p></div>'; return; }
        if (!choices.some(function (c) { return c.id === choice; })) choice = "auto";
        loadInfo(); drawChoices();
        var c = new URLSearchParams(location.search).get("c");
        listThreads(); if (c) openThread(c); else hello();
      });
    });
  });
})();
