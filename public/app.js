// Confirm before destructive actions: <form data-confirm="Are you sure?">
document.addEventListener("submit", function (e) {
  var msg = e.target.getAttribute("data-confirm");
  if (msg && !window.confirm(msg)) e.preventDefault();
});

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
