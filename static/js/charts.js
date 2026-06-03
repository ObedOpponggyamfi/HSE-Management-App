/* Build dashboard / page charts from window.HSE_CHARTS using Chart.js (vendored). */
(function () {
  const D = window.HSE_CHARTS;
  if (typeof Chart === "undefined" || !D) return;

  const NAVY = "#143352", GOLD = "#c9a227", RED = "#b3261e",
        GREEN = "#1a7a3d", BLUE = "#1f4e79", TEAL = "#2aa6a0";
  const PALETTE = ["#143352", GOLD, GREEN, RED, "#6a5acd", TEAL, "#e8772e", "#8aa1b4"];

  Chart.defaults.font.family = "'Segoe UI',Roboto,Arial,sans-serif";
  Chart.defaults.color = "#52617a";
  Chart.defaults.plugins.legend.labels.boxWidth = 12;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.maintainAspectRatio = false;

  const $ = (id) => document.getElementById(id);
  const has = (id, data) => $(id) && data;

  function line(id, labels, datasets) {
    new Chart($(id), { type: "line",
      data: { labels, datasets },
      options: { responsive: true, interaction: { mode: "index", intersect: false },
        scales: { y: { beginAtZero: true, grid: { color: "#eef1f6" } },
                  x: { grid: { display: false } } } } });
  }
  function bar(id, labels, datasets, horizontal) {
    new Chart($(id), { type: "bar",
      data: { labels, datasets },
      options: { responsive: true, indexAxis: horizontal ? "y" : "x",
        plugins: { legend: { display: datasets.length > 1 } },
        scales: { y: { beginAtZero: true, grid: { color: "#eef1f6" } },
                  x: { grid: { display: false } } } } });
  }
  function doughnut(id, labels, data) {
    new Chart($(id), { type: "doughnut",
      data: { labels, datasets: [{ data, backgroundColor: PALETTE, borderWidth: 2, borderColor: "#fff" }] },
      options: { responsive: true, cutout: "58%", plugins: { legend: { position: "right" } } } });
  }
  const limitLine = (n, val, label) => ({
    type: "line", label, data: Array(n).fill(val), borderColor: RED,
    borderDash: [6, 5], borderWidth: 2, pointRadius: 0, fill: false });

  /* ---- Dashboard ---- */
  if (has("trendChart", D.trend)) {
    line("trendChart", D.trend.labels, [
      { label: "Incidents", data: D.trend.incidents, borderColor: NAVY,
        backgroundColor: "rgba(20,51,82,.10)", borderWidth: 2.5, fill: true, tension: .3 },
      { label: "Recordables", data: D.trend.recordables, borderColor: RED,
        borderWidth: 2.5, tension: .3 },
      { label: "Near misses", data: D.trend.near_misses, borderColor: GOLD,
        borderWidth: 2, tension: .3, hidden: true } ]);
  }
  if (has("typeChart", D.type)) doughnut("typeChart", D.type.labels, D.type.data);
  if (has("locationChart", D.location))
    bar("locationChart", D.location.labels,
        [{ label: "Incidents", data: D.location.data, backgroundColor: GOLD }], true);
  if (has("actionsChart", D.actions)) {
    const colors = D.actions.labels.map(s =>
      ({ "Closed": GREEN, "Open": RED, "In Progress": BLUE, "Due Soon": GOLD }[s] || NAVY));
    new Chart($("actionsChart"), { type: "bar",
      data: { labels: D.actions.labels, datasets: [{ data: D.actions.data, backgroundColor: colors }] },
      options: { plugins: { legend: { display: false } },
        scales: { y: { beginAtZero: true }, x: { grid: { display: false } } } } });
  }
  if (has("trainingChart", D.training))
    bar("trainingChart", D.training.labels, [
      { label: "Completion %", data: D.training.data, backgroundColor: NAVY },
      limitLine(D.training.labels.length, D.training.target, "Target") ]);

  /* ---- Environmental ---- */
  if (D.env) {
    const e = D.env;
    if ($("pm10Chart")) line("pm10Chart", e.labels, [
      { label: "PM10 (µg/m³)", data: e.pm10, borderColor: GOLD,
        backgroundColor: "rgba(201,162,39,.12)", fill: true, borderWidth: 2.5, tension: .3 },
      limitLine(e.labels.length, e.pm10_limit, "EPA limit") ]);
    if ($("cnChart")) line("cnChart", e.labels, [
      { label: "WAD CN (mg/L)", data: e.cn, borderColor: GREEN,
        backgroundColor: "rgba(26,122,61,.12)", fill: true, borderWidth: 2.5, tension: .3 },
      limitLine(e.labels.length, e.cn_limit, "ICMC limit") ]);
    if ($("energyChart")) bar("energyChart", e.labels,
        [{ label: "Energy (MWh)", data: e.energy, backgroundColor: NAVY }]);
    if ($("waterChart")) bar("waterChart", e.labels,
        [{ label: "Water (m³)", data: e.water, backgroundColor: TEAL }]);
  }

  /* ---- Event reports ---- */
  if (D.events) {
    if ($("eventsCatChart")) doughnut("eventsCatChart", D.events.cats.labels, D.events.cats.data);
    if ($("eventsTrendChart")) line("eventsTrendChart", D.events.trend.labels, [
      { label: "Reports", data: D.events.trend.data, borderColor: GREEN,
        backgroundColor: "rgba(26,122,61,.12)", fill: true, borderWidth: 2.5, tension: .3 }]);
  }

  /* ---- Contractors ---- */
  if (D.contractor && $("contractorChart")) {
    const c = D.contractor;
    const colors = c.trifr.map(v => v > c.target ? RED : GREEN);
    new Chart($("contractorChart"), {
      type: "bar",
      data: { labels: c.labels, datasets: [
        { type: "bar", label: "TRIFR", data: c.trifr, backgroundColor: colors },
        limitLine(c.labels.length, c.target, "Target TRIFR") ] },
      options: { plugins: { legend: { display: true } },
        scales: { y: { beginAtZero: true }, x: { grid: { display: false } } } } });
  }

  /* ---- Rolling frequency rates ---- */
  if (D.rates && $("ratesChart")) {
    const r = D.rates;
    line("ratesChart", r.labels, [
      { label: "TRIFR (12-mo)", data: r.trifr, borderColor: NAVY, borderWidth: 2.6, tension: .3 },
      { label: "LTIFR (12-mo)", data: r.ltifr, borderColor: GREEN, borderWidth: 2.6, tension: .3 },
      { label: "AIFR (12-mo)", data: r.aifr, borderColor: GOLD, borderWidth: 2, tension: .3 },
      limitLine(r.labels.length, r.target_trifr, "TRIFR target") ]);
  }
})();
