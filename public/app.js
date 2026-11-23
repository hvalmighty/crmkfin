// Confirm before destructive actions: <form data-confirm="Are you sure?">
document.addEventListener("submit", function (e) {
  var msg = e.target.getAttribute("data-confirm");
  if (msg && !window.confirm(msg)) e.preventDefault();
});

// Sidebar: on phones it slides over the page; on desktop it collapses to an icon-only rail (remembered).
(function () {
  var b = document.getElementById("menu-btn"), sc = document.getElementById("scrim");
  var mobile = window.matchMedia("(max-width: 960px)");
  try { if (localStorage.getItem("crm_side") === "1") document.body.classList.add("side-collapsed"); } catch (e) {}
  function sync() {
    if (!b) return;
    var open = mobile.matches ? document.body.classList.contains("nav-open") : !document.body.classList.contains("side-collapsed");
    b.setAttribute("aria-expanded", open ? "true" : "false");
    b.title = mobile.matches ? "Menu" : (open ? "Collapse menu" : "Expand menu");
    b.setAttribute("aria-label", b.title);
  }
  if (b) b.addEventListener("click", function () {
    if (mobile.matches) document.body.classList.toggle("nav-open");
    else {
      var c = document.body.classList.toggle("side-collapsed");
      try { localStorage.setItem("crm_side", c ? "1" : "0"); } catch (e) {}
      window.dispatchEvent(new Event("resize"));      // charts re-fit to the new width
    }
    sync(); document.dispatchEvent(new Event("crm:layout"));
  });
  if (sc) sc.addEventListener("click", function () { document.body.classList.remove("nav-open"); sync(); });
  if (mobile.addEventListener) mobile.addEventListener("change", sync);
  sync();
})();

// Prospect / client table: choose which columns to show (remembered per browser, separately for prospects and clients).
(function () {
  var pick = document.getElementById("colpick"), table = document.getElementById("rec-table");
  if (!pick) return;
  var kind = pick.getAttribute("data-kind"), key = "crm_cols_" + kind;
  var defaults = pick.getAttribute("data-default").split(",");
  var boxes = Array.prototype.slice.call(pick.querySelectorAll("input[type=checkbox]"));
  var hint = document.getElementById("col-hint"), wrap = document.getElementById("rec-scroll");
  var on = defaults.slice();
  try { var saved = JSON.parse(localStorage.getItem(key) || "null"); if (Array.isArray(saved)) on = saved; } catch (e) {}
  function apply(save) {
    boxes.forEach(function (b) { b.checked = on.indexOf(b.value) >= 0; });
    if (table) table.querySelectorAll("[data-col]").forEach(function (el) { el.classList.toggle("col-hidden", on.indexOf(el.getAttribute("data-col")) < 0); });
    if (save) { try { localStorage.setItem(key, JSON.stringify(on)); } catch (e) {} }
    if (hint && wrap) hint.classList.toggle("show", wrap.scrollWidth > wrap.clientWidth + 2);
  }
  boxes.forEach(function (b) {
    b.addEventListener("change", function () { on = boxes.filter(function (x) { return x.checked; }).map(function (x) { return x.value; }); apply(true); });
  });
  document.getElementById("col-reset").addEventListener("click", function () { on = defaults.slice(); apply(true); });
  document.getElementById("col-all").addEventListener("click", function () { on = boxes.map(function (x) { return x.value; }); apply(true); });
  document.addEventListener("click", function (e) { if (pick.open && !pick.contains(e.target)) pick.removeAttribute("open"); });
  document.addEventListener("keydown", function (e) { if (e.key === "Escape") pick.removeAttribute("open"); });
  window.addEventListener("resize", function () { apply(false); });
  document.addEventListener("crm:layout", function () { setTimeout(function () { apply(false); }, 260); });
  apply(false);
})();

// Currency switch (₹ / $) - remembered per browser, shared by KPIs, tables and charts.
(function () {
  var wrap = document.querySelector(".cur-wrap");
  if (!wrap) return;
  var cur = "inr";
  try { cur = localStorage.getItem("crm_cur") || "inr"; } catch (e) {}
  if (cur !== "inr" && cur !== "usd") cur = "inr";
  function apply(c, fromClick) {
    cur = c; wrap.setAttribute("data-cur", c);
    document.querySelectorAll(".seg button[data-cur]").forEach(function (b) { b.classList.toggle("on", b.getAttribute("data-cur") === c); });
    if (fromClick) { try { localStorage.setItem("crm_cur", c); } catch (e) {} }
    document.dispatchEvent(new CustomEvent("crm:currency", { detail: c }));
  }
  document.querySelectorAll(".seg button[data-cur]").forEach(function (b) {
    b.addEventListener("click", function () { apply(b.getAttribute("data-cur"), true); });
  });
  window.CRM_CUR = function () { return cur; };
  apply(cur, false);
})();

// Record form: show prospect-only / client-only fields depending on the selected type.
(function () {
  var sel = document.getElementById("type");
  if (!sel) return;
  function apply() {
    var t = sel.value;
    document.querySelectorAll("[data-only]").forEach(function (el) {
      var show = el.getAttribute("data-only") === t;
      el.hidden = !show;
      el.querySelectorAll("input,select").forEach(function (i) { i.disabled = !show; });
    });
  }
  sel.addEventListener("change", apply);
  apply();

  // Suggest the default probability when the stage changes (only if untouched).
  var stage = document.getElementById("stage"), prob = document.getElementById("probability");
  if (stage && prob) {
    var defaults = JSON.parse(stage.getAttribute("data-defaults") || "{}");
    stage.addEventListener("change", function () {
      if (defaults[stage.value] !== undefined) prob.value = defaults[stage.value];
    });
  }
})();


// Sales targets grid: live yearly totals per row.
(function () {
  var f = document.getElementById("tgt-form");
  if (!f) return;
  function fmt(c, n) { return n ? (c === "inr" ? "₹" : "$") + Math.round(n).toLocaleString(c === "inr" ? "en-IN" : "en-US") : "–"; }
  function recalc() {
    f.querySelectorAll("tr").forEach(function (tr) {
      ["inr", "usd"].forEach(function (c) {
        var sum = 0, any = false;
        tr.querySelectorAll('input[data-sum="' + c + '"]').forEach(function (i) { var v = parseFloat(String(i.value).replace(/[^\d.]/g, "")); if (!isNaN(v)) { sum += v; any = true; } });
        var out = tr.querySelector('[data-tot="' + c + '"]'); if (out) out.textContent = any ? fmt(c, sum) : "–";
      });
    });
  }
  f.addEventListener("input", recalc); recalc();
})();


// Global search with predictive suggestions.
(function () {
  var f = document.getElementById("gsearch");
  if (!f) return;
  var inp = document.getElementById("gs-input"), list = document.getElementById("gs-list");
  var items = [], act = -1, timer = null, seq = 0;
  var tag = { record: "", rep: "Rep", product: "Product" };
  function close() { list.hidden = true; inp.setAttribute("aria-expanded", "false"); act = -1; }
  function mark(txt, q) {                      // bold the matched part, safely (no innerHTML)
    var frag = document.createDocumentFragment(), i = txt.toLowerCase().indexOf(q.toLowerCase());
    if (i < 0) { frag.appendChild(document.createTextNode(txt)); return frag; }
    frag.appendChild(document.createTextNode(txt.slice(0, i)));
    var b = document.createElement("b"); b.textContent = txt.slice(i, i + q.length); frag.appendChild(b);
    frag.appendChild(document.createTextNode(txt.slice(i + q.length))); return frag;
  }
  function render(q) {
    list.textContent = "";
    if (!items.length) { var e = document.createElement("li"); e.className = "gs-empty"; e.textContent = "No matches - press Enter to search everything"; list.appendChild(e); }
    items.forEach(function (it, n) {
      var li = document.createElement("li"); li.setAttribute("role", "option"); li.id = "gs-" + n; li.className = "gs-it " + it.kind;
      var a = document.createElement("a"); a.href = it.url;
      var t = document.createElement("span"); t.className = "gs-t"; t.appendChild(mark(it.label, q));
      var s = document.createElement("span"); s.className = "gs-s"; s.textContent = it.sub;
      a.appendChild(t); a.appendChild(s); li.appendChild(a);
      li.addEventListener("mousedown", function (ev) { ev.preventDefault(); });   // keep focus so the click lands
      list.appendChild(li);
    });
    list.hidden = false; inp.setAttribute("aria-expanded", "true");
  }
  function setAct(n) {
    var els = list.querySelectorAll(".gs-it");
    els.forEach(function (e) { e.classList.remove("on"); e.removeAttribute("aria-selected"); });
    act = els.length ? (n + els.length) % els.length : -1;
    if (act >= 0) { els[act].classList.add("on"); els[act].setAttribute("aria-selected", "true"); inp.setAttribute("aria-activedescendant", els[act].id); els[act].scrollIntoView({ block: "nearest" }); }
  }
  function fetchNow() {
    var q = inp.value.trim();
    if (q.length < 2) { close(); return; }
    var my = ++seq;
    fetch(f.getAttribute("data-url") + "?q=" + encodeURIComponent(q), { credentials: "same-origin", headers: { Accept: "application/json" } })
      .then(function (r) { return r.ok ? r.json() : { items: [] }; })
      .then(function (d) { if (my !== seq) return; items = d.items || []; render(q); })
      .catch(function () { close(); });
  }
  inp.addEventListener("input", function () { clearTimeout(timer); timer = setTimeout(fetchNow, 180); });
  inp.addEventListener("focus", function () { if (inp.value.trim().length >= 2 && items.length) list.hidden = false; });
  inp.addEventListener("blur", function () { setTimeout(close, 120); });
  inp.addEventListener("keydown", function (ev) {
    if (ev.key === "ArrowDown") { ev.preventDefault(); if (list.hidden) fetchNow(); else setAct(act + 1); }
    else if (ev.key === "ArrowUp") { ev.preventDefault(); setAct(act - 1); }
    else if (ev.key === "Escape") { close(); inp.blur(); }
    else if (ev.key === "Enter" && act >= 0 && !list.hidden) { ev.preventDefault(); window.location.href = items[act].url; }
  });
  document.addEventListener("keydown", function (ev) {   // "/" focuses the search box
    if (ev.key === "/" && !/^(INPUT|TEXTAREA|SELECT)$/.test((document.activeElement || {}).tagName || "")) { ev.preventDefault(); inp.focus(); inp.select(); }
  });
})();


// Notification bell: overdue and upcoming (30 days) prospect closures.
(function () {
  var box = document.getElementById("bell");
  if (!box) return;
  var btn = document.getElementById("bell-btn"), pop = document.getElementById("bell-pop"), badge = document.getElementById("bell-n");
  function el(tag, cls, txt) { var e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; }
  function when(d) { return d === 0 ? "Due today" : d === 1 ? "Tomorrow" : d > 1 ? "In " + d + " days" : (-d === 1 ? "1 day overdue" : -d + " days overdue"); }
  function section(title, cls, n, list, url) {
    var s = el("div", "bell-sec " + cls); s.appendChild(el("div", "bell-h", title + " (" + n + ")"));
    list.forEach(function (it) {
      var a = el("a", "bell-it"); a.href = it.url;
      a.appendChild(el("span", "bell-t", it.label)); a.appendChild(el("span", "bell-w", when(it.days)));
      a.appendChild(el("span", "bell-s", it.sub)); s.appendChild(a);
    });
    if (n > list.length) { var m = el("a", "bell-more", "View all " + n + " \u2192"); m.href = url; s.appendChild(m); }
    else if (n) { var m2 = el("a", "bell-more", "Open in list \u2192"); m2.href = url; s.appendChild(m2); }
    return s;
  }
  function paint(d) {
    var n = d.overdue_n + d.upcoming_n;
    badge.hidden = !n; badge.textContent = n > 99 ? "99+" : n; badge.className = "bell-n" + (d.overdue_n ? " red" : "");
    btn.setAttribute("aria-label", "Notifications: " + d.overdue_n + " overdue, " + d.upcoming_n + " upcoming");
    pop.textContent = "";
    if (!n) { pop.appendChild(el("div", "bell-empty", "All clear - nothing overdue or closing in the next 30 days.")); return; }
    if (d.overdue_n) pop.appendChild(section("Overdue", "over", d.overdue_n, d.overdue, d.overdue_url));
    if (d.upcoming_n) pop.appendChild(section("Closing in next 30 days", "soon", d.upcoming_n, d.upcoming, d.upcoming_url));
  }
  function load() {
    return fetch(box.getAttribute("data-url"), { credentials: "same-origin", headers: { Accept: "application/json" } })
      .then(function (r) { return r.ok ? r.json() : null; }).then(function (d) { if (d) paint(d); }).catch(function () {});
  }
  btn.addEventListener("click", function (ev) {
    ev.stopPropagation();
    var open = pop.hidden; pop.hidden = !open; btn.setAttribute("aria-expanded", String(open));
    if (open) load();
  });
  document.addEventListener("click", function (ev) { if (!box.contains(ev.target)) { pop.hidden = true; btn.setAttribute("aria-expanded", "false"); } });
  document.addEventListener("keydown", function (ev) { if (ev.key === "Escape" && !pop.hidden) { pop.hidden = true; btn.focus(); } });
  setTimeout(load, 400);               // after the page is up, so it never delays the first paint
})();
