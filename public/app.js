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
