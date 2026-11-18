// Confirm before destructive actions: <form data-confirm="Are you sure?">
document.addEventListener("submit", function (e) {
  var msg = e.target.getAttribute("data-confirm");
  if (msg && !window.confirm(msg)) e.preventDefault();
});

// Mobile sidebar
(function () {
  var b = document.getElementById("menu-btn"), s = document.getElementById("scrim");
  if (b) b.addEventListener("click", function () { document.body.classList.toggle("nav-open"); });
  if (s) s.addEventListener("click", function () { document.body.classList.remove("nav-open"); });
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
