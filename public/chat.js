// IRIS - floating assistant. History lives only in this browser tab (sessionStorage); the server keeps nothing.
(function () {
  var root = document.getElementById("iris");
  if (!root) return;
  var fab = document.getElementById("iris-fab"), panel = document.getElementById("iris-panel");
  function setOpen(on) {
    panel.hidden = !on; fab.setAttribute("aria-expanded", on ? "true" : "false"); fab.classList.toggle("open", on);
    try { sessionStorage.setItem("iris_open", on ? "1" : "0"); } catch (e) {}
    if (on) { var i = document.getElementById("chat-input"); if (i) setTimeout(function () { i.focus(); }, 50); scrollDown(); }
  }
  function scrollDown() { var l = document.getElementById("chat-log"); if (l) l.scrollTop = l.scrollHeight; }
  fab.addEventListener("click", function () { setOpen(panel.hidden); });
  document.getElementById("iris-close").addEventListener("click", function () { setOpen(false); fab.focus(); });
  document.addEventListener("keydown", function (e) { if (e.key === "Escape" && !panel.hidden) { setOpen(false); fab.focus(); } });

  var log = document.getElementById("chat-log"), form = document.getElementById("chat-form");
  if (!form) { try { if (sessionStorage.getItem("iris_open") === "1") setOpen(true); } catch (e) {} return; }   // not set up
  var input = document.getElementById("chat-input"), send = document.getElementById("chat-send"), chips = document.getElementById("chat-chips");
  var history = [];
  try { history = JSON.parse(sessionStorage.getItem("crm_chat") || "[]"); } catch (e) {}
  function save() { try { sessionStorage.setItem("crm_chat", JSON.stringify(history.slice(-12))); } catch (e) {} }

  // tiny safe renderer: **bold**, "- " bullets, paragraphs, | tables | — built from DOM nodes, never innerHTML
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
    m.appendChild(b); log.appendChild(m); scrollDown(); return m;
  }
  function greet() { add("bot", "Hi " + root.dataset.user + ", I'm **IRIS**. Ask me about your prospects, clients, products or regions – I can look things up but I can't change any data."); }

  greet();
  history.forEach(function (h) { add(h.role, h.content); });
  if (history.length) chips.style.display = "none";

  function ask(q) {
    q = (q || "").trim(); if (!q) return;
    chips.style.display = "none";
    add("user", q); history.push({ role: "user", content: q }); save();
    input.value = ""; send.disabled = true; input.disabled = true;
    var wait = add("bot", "Thinking…", "wait");
    var fd = new FormData(); fd.append("_csrf", root.dataset.csrf); fd.append("messages", JSON.stringify(history.slice(-12)));
    fetch(root.dataset.url, { method: "POST", body: fd, credentials: "same-origin" })
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
  document.getElementById("iris-new").addEventListener("click", function () {
    history = []; save(); log.textContent = ""; greet(); chips.style.display = ""; input.focus();
  });
  try { if (sessionStorage.getItem("iris_open") === "1") setOpen(true); } catch (e) {}
})();
