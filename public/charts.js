// Draws the charts on Dashboard / Analytics from the JSON in <script id="chart-data">.
(function () {
  var el = document.getElementById("chart-data");
  if (!el || !window.Chart) return;
  var D = JSON.parse(el.textContent);

  var INK = "#231a1c", MUTED = "#85787b", GRID = "#f3e9ea";
  var C = { brand: "#d0202f", violet: "#7f1d1d", sky: "#6b7280", teal: "#f87171", amber: "#f59e0b", rose: "#ef4444", green: "#22a06b", slate: "#d9cfd1" };
  var PAL = [C.brand, C.sky, C.teal, C.amber, C.violet, C.rose, C.green, "#fb923c", "#94a3b8"];
  Chart.defaults.font.family = "'Bai Jamjuree', system-ui, sans-serif";
  Chart.defaults.font.size = 12.5;
  Chart.defaults.color = MUTED;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = INK;
  Chart.defaults.plugins.tooltip.padding = 12;
  Chart.defaults.plugins.tooltip.cornerRadius = 10;
  Chart.defaults.plugins.tooltip.titleFont = { weight: "600" };
  Chart.defaults.maintainAspectRatio = false;

  function trim(x) { return x.toFixed(2).replace(/0+$/, "").replace(/\.$/, ""); }
  function inrGroup(n) {                       // Indian digit grouping: 12,34,567
    var s = String(Math.round(Math.abs(n)));
    if (s.length > 3) { var h = s.slice(0, -3), t = s.slice(-3), p = [];
      while (h.length > 2) { p.unshift(h.slice(-2)); h = h.slice(0, -2); }
      if (h) p.unshift(h); s = p.concat([t]).join(","); }
    return (n < 0 ? "-" : "") + s;
  }
  function full(c, v) { return c === "inr" ? "₹" + inrGroup(v) : "$" + Math.round(v).toLocaleString("en-US"); }
  function compact(c, v) {
    if (c === "inr") { if (v >= 1e7) return "₹" + trim(v / 1e7) + " Cr"; if (v >= 1e5) return "₹" + trim(v / 1e5) + " L"; return full(c, v); }
    if (v >= 1e9) return "$" + trim(v / 1e9) + "B"; if (v >= 1e6) return "$" + trim(v / 1e6) + "M"; if (v >= 1e4) return "$" + trim(v / 1e3) + "K"; return full(c, v);
  }
  function axisMoney(c) { return { grid: { color: GRID }, border: { display: false }, beginAtZero: true, ticks: { maxTicksLimit: 6, maxRotation: 0, autoSkip: true, callback: function (v) { return compact(c, v); } } }; }
  var noGrid = { grid: { display: false }, border: { display: false } };
  function moneyTip(c, extra) {
    return { callbacks: { label: function (ctx) { return " " + ctx.dataset.label + ": " + full(c, ctx.parsed.y !== undefined && ctx.chart.options.indexAxis !== "y" ? ctx.parsed.y : ctx.parsed.x); }, afterBody: extra } };
  }
  function hMoneyTip(c, extra) {
    return { callbacks: { label: function (ctx) { return " " + ctx.dataset.label + ": " + full(c, ctx.parsed.x); }, afterBody: extra } };
  }
  function bars(label, data, color) { return { label: label, data: data, backgroundColor: color, borderRadius: 8, borderSkipped: false, maxBarThickness: 44 }; }

  var defs = {
    "c-stages": function (c) {
      var S = D.stages;
      return { type: "bar", data: { labels: S.labels, datasets: [Object.assign(bars("Pipeline", S[c], [C.slate, C.sky, C.amber, C.violet]), { maxBarThickness: 64 })] },
        options: { scales: { x: noGrid, y: axisMoney(c) }, plugins: { legend: { display: false },
          tooltip: moneyTip(c, function (items) { var i = items[0].dataIndex; return ["Prospects: " + S.count[i], "Weighted: " + full(c, S["w_" + c][i])]; }) } } };
    },
    "c-forecast": function (c) {
      var F = D.forecast;
      return { type: "bar", data: { labels: F.labels, datasets: [bars("Pipeline", F[c], "#f6c3c8"), bars("Weighted", F["w_" + c], C.brand)] },
        options: { scales: { x: noGrid, y: axisMoney(c) }, plugins: { legend: { position: "bottom" },
          tooltip: moneyTip(c, function (items) { return ["Prospects: " + F.count[items[0].dataIndex]]; }) } } };
    },
    "c-won": function (c) {
      var W = D.won;
      return { type: "bar", data: { labels: W.labels, datasets: [bars("Won value", W[c], C.teal),
          { type: "line", label: "Clients", data: W.count, yAxisID: "y2", borderColor: C.brand, backgroundColor: C.brand, tension: .35, pointRadius: 4, borderWidth: 2.5 }] },
        options: { scales: { x: noGrid, y: axisMoney(c), y2: { position: "right", beginAtZero: true, grid: { display: false }, border: { display: false }, ticks: { precision: 0 } } },
          plugins: { legend: { position: "bottom" }, tooltip: { callbacks: { label: function (ctx) { return " " + ctx.dataset.label + ": " + (ctx.dataset.yAxisID === "y2" ? ctx.parsed.y : full(c, ctx.parsed.y)); } } } } } };
    },
    "c-region": function (c) { return doughnut(D.region, c); },
    "c-source": function (c) { return doughnut(D.source, c); },
    "c-presales": function (c) { return hbar(D.presales, c); },
    "c-owner": function (c) { return hbar(D.owner, c); },
    "c-prob": function (c) {
      var P = D.prob;
      return { type: "bar", data: { labels: P.labels, datasets: [bars("Pipeline", P[c], [C.slate, C.sky, C.violet, C.brand])] },
        options: { scales: { x: noGrid, y: axisMoney(c) }, plugins: { legend: { display: false },
          tooltip: moneyTip(c, function (items) { return ["Prospects: " + P.count[items[0].dataIndex]]; }) } } };
    },
    "c-size": function (c) {
      var Z = D.size[c];
      return { type: "bar", data: { labels: Z.labels, datasets: [bars("Open deals", Z.count, C.violet)] },
        options: { scales: { x: noGrid, y: { grid: { color: GRID }, border: { display: false }, beginAtZero: true, ticks: { precision: 0 } } },
          plugins: { legend: { display: false } } } };
    },
    "c-targets": function (c) {
      var T = D.targets;
      return { type: "bar", data: { labels: T.labels, datasets: [bars("Target", T["t_" + c], "#e9d8da"), bars("Achieved", T["a_" + c], C.brand), bars("Pipeline (weighted)", T["w_" + c], "#f6a5ac")] },
        options: { scales: { x: noGrid, y: axisMoney(c) }, plugins: { legend: { position: "bottom" }, tooltip: moneyTip(c, function (items) {
          var i = items[0].dataIndex, t = T["t_" + c][i], a = T["a_" + c][i];
          return t ? ["Attainment: " + Math.round(100 * a / t) + "%"] : []; }) } } };
    },
    "c-product": function (c) { return hbar(D.product, c); },
    "c-offering": function (c) { return doughnut(D.offering, c); },
    "c-prodwon": function (c) { return doughnut(D.prodwon, c); },
    "c-prodwin": function () { return winrate(D.prodwin); },
    "c-winrate": function () { return winrate(D.winrate); },
    "c-added": function () {
      var A = D.added;
      function line(l, d, col) { return { label: l, data: d, borderColor: col, backgroundColor: col + "22", fill: true, tension: .35, pointRadius: 4, borderWidth: 2.5 }; }
      return { type: "line", data: { labels: A.labels, datasets: [line("Prospects added", A.prospects, C.brand), line("Clients added", A.clients, C.teal)] },
        options: { scales: { x: noGrid, y: { grid: { color: GRID }, border: { display: false }, beginAtZero: true, ticks: { precision: 0 } } }, plugins: { legend: { position: "bottom" } } } };
    }
  };

  function winrate(R) {
    return { type: "bar", data: { labels: R.labels, datasets: [bars("Win rate", R.rate, C.green)] },
      options: { indexAxis: "y", scales: { x: { min: 0, max: 100, grid: { color: GRID }, border: { display: false }, ticks: { callback: function (v) { return v + "%"; } } }, y: noGrid },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: function (ctx) { var i = ctx.dataIndex; return " " + R.rate[i] + "% (" + R.won[i] + " won, " + R.lost[i] + " lost)"; } } } } } };
  }
  function hbar(S, c) {
    return { type: "bar", data: { labels: S.labels, datasets: [bars("Pipeline", S[c], "#f6c3c8"), bars("Weighted", S["w_" + c], C.brand)] },
      options: { indexAxis: "y", scales: { x: axisMoney(c), y: noGrid }, plugins: { legend: { position: "bottom" },
        tooltip: hMoneyTip(c, function (items) { return ["Prospects: " + S.n[items[0].dataIndex]]; }) } } };
  }
  function doughnut(S, c) {
    var total = S[c].reduce(function (a, b) { return a + b; }, 0);
    return { type: "doughnut", data: { labels: S.labels, datasets: [{ data: S[c], backgroundColor: PAL, borderWidth: 3, borderColor: "#fff", hoverOffset: 6 }] },
      options: { cutout: "66%", plugins: { legend: { position: "right", labels: { padding: 14 } },
        tooltip: { callbacks: { label: function (ctx) { var v = ctx.parsed; return " " + ctx.label + ": " + full(c, v) + (total ? " (" + Math.round(100 * v / total) + "%)" : ""); } } } } } };
  }

  var live = {};
  function render() {
    var c = (window.CRM_CUR && window.CRM_CUR()) || "inr";
    Object.keys(defs).forEach(function (id) {
      var cv = document.getElementById(id);
      if (!cv) return;
      if (live[id]) live[id].destroy();
      live[id] = new Chart(cv, defs[id](c));
    });
  }
  render();
  document.addEventListener("crm:currency", render);
})();
