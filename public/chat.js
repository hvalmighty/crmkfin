// Assistant chat UI. History lives only in this page (sessionStorage); the server keeps nothing.
(function () {
  var box = document.getElementById("chat");
  if (!box) return;
  var log = document.getElementById("chat-log"), form = document.getElementById("chat-form"),
      input = document.getElementById("chat-input"), send = document.getElementById("chat-send"),
      chips = document.getElementById("chat-chips");
  var history = [];
  try { history = JSON.parse(sessionStorage.getItem("crm_chat") || "[]"); } catch (e) {}

  function save() { try { sessionStorage.setItem("crm_chat", JSON.stringify(history.slice(-12))); } catch (e) {} }

  // tiny safe renderer: **bold**, "- " bullets, paragraphs, and | tables | — built with DOM nodes, never innerHTML
  function inline(parent, text) {
    text.split(/(\*\*[^*]+\*\*)/g).forEach(function (part) {
      if (/^\*\*[^*]+\*\*$/.test(part)) { var b = document.createElement("strong"); b.textContent = part.slice(2, -2); parent.appendChild(b); }
      else if (part) parent.appendChild(document.createTextNode(part));
    });
  }
  function render(el, text) {
    el.textContent = "";
    var lines = text.replace(/\r/g, "").split("\n"), i = 0;
    while (i < lines.length) {
      var ln = lines[i];
      if (/^\s*[-*•]\s+/.test(ln)) {
        var ul = document.createElement("ul");
        while (i < lines.length && /^\s*[-*•]\s+/.test(lines[i])) { var li = document.createElement("li"); inline(li, lines[i].replace(/^\s*[-*•]\s+/, "")); ul.appendChild(li); i++; }
        el.appendChild(ul);
      } else if (/^\s*\|.*\|\s*$/.test(ln)) {
        var tbl = document.createElement("table"), first = true;
        while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) {
          if (!/^\s*\|[\s:|-]+\|\s*$/.test(lines[i])) {
            var tr = document.createElement("tr");
            lines[i].trim().replace(/^\||\|$/g, "").split("|").forEach(function (c) { var td = document.createElement(first ? "th" : "td"); inline(td, c.trim()); tr.appendChild(td); });
            tbl.appendChild(tr); first = false;
          }
          i++;
        }
        var wrap = document.createElement("div"); wrap.className = "scroll"; wrap.appendChild(tbl); el.appendChild(wrap);
      } else if (ln.trim()) {
        var p = document.createElement("p"); inline(p, ln.trim()); el.appendChild(p); i++;
      } else i++;
    }
  }
  function add(role, text, cls) {
    var m = document.createElement("div"); m.className = "msg " + (role === "user" ? "you" : "bot") + (cls ? " " + cls : "");
    var b = document.createElement("div"); b.className = "bub";
    if (role === "user") b.textContent = text; else render(b, text);
    m.appendChild(b); log.appendChild(m); log.scrollTop = log.scrollHeight; return m;
  }
  history.forEach(function (h) { add(h.role, h.content); });
  if (history.length) chips.style.display = "none";

  function ask(q) {
    q = (q || "").trim(); if (!q) return;
    chips.style.display = "none";
    add("user", q); history.push({ role: "user", content: q }); save();
    input.value = ""; send.disabled = true; input.disabled = true;
    var wait = add("bot", "Thinking…", "wait");
    var fd = new FormData(); fd.append("_csrf", box.dataset.csrf); fd.append("messages", JSON.stringify(history.slice(-12)));
    fetch(box.dataset.url, { method: "POST", body: fd, credentials: "same-origin" })
      .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }, function () { throw new Error("session"); }); })
      .then(function (x) {
        wait.remove();
        if (x.ok && x.j.reply) { add("bot", x.j.reply); history.push({ role: "assistant", content: x.j.reply }); save(); }
        else { history.pop(); save(); add("bot", x.j.error || "Something went wrong.", "err"); }
      })
      .catch(function (e) {
        wait.remove(); history.pop(); save();
        add("bot", e.message === "session" ? "Your session may have expired – reload the page and log in again." : "Couldn't reach the server. Please try again.", "err");
      })
      .then(function () { send.disabled = false; input.disabled = false; input.focus(); });
  }
  form.addEventListener("submit", function (e) { e.preventDefault(); ask(input.value); });
  chips.addEventListener("click", function (e) { if (e.target.tagName === "BUTTON") ask(e.target.textContent); });
  var clr = document.getElementById("chat-clear");
  if (clr) clr.addEventListener("click", function () {
    history = []; save(); log.textContent = ""; add("bot", "New chat started. What would you like to know?"); chips.style.display = "";
  });
})();
