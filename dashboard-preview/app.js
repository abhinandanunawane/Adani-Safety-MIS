/**
 * Adani Safety MIS — fixed 1280×720 preview (no page scroll)
 */
(function () {
  "use strict";

  const root = document.getElementById("app-root");
  const liveRegion = document.getElementById("sr-live");
  const DATA = window.__DASHBOARD_DATA__;

  /** Sync primary nav with route for screen readers (IA: Home vs Categories scope). */
  function updateHeaderNavState() {
    const raw = location.hash || "";
    const h = raw === "" || raw === "#" ? "#landing" : raw;
    const home = document.getElementById("nav-home");
    const cats = document.getElementById("nav-categories");
    const onLanding = h === "#landing";
    const onCategories = h === "#categories";
    const onCategoryDrill = /^#cat=\d+/.test(h);
    if (home) {
      if (onLanding) home.setAttribute("aria-current", "page");
      else home.removeAttribute("aria-current");
    }
    if (cats) {
      if (onCategories || onCategoryDrill) {
        cats.setAttribute("aria-current", "page");
      } else {
        cats.removeAttribute("aria-current");
      }
    }
  }

  function setHeaderTimestamp(iso) {
    const hu = document.getElementById("header-updated");
    if (!hu) return;
    if (!iso) {
      hu.textContent = "—";
      return;
    }
    try {
      hu.textContent = new Date(iso).toLocaleString(undefined, {
        dateStyle: "medium",
        timeStyle: "short",
      });
    } catch {
      hu.textContent = String(iso);
    }
  }

  if (!DATA) {
    setHeaderTimestamp(null);
    if (root) {
      root.innerHTML =
        '<div class="boot-error" style="padding:12px;font-size:13px;color:#0f172a">' +
        "<strong>Data not loaded.</strong> The header and title above should still be visible. " +
        "Ensure <code>embedded-data.js</code> is in the same folder as <code>index.html</code> " +
        "and open the page via a local server (e.g. <code>npx serve</code>) if scripts are blocked. " +
        "Then run <code>Refresh-PreviewData.ps1</code> to regenerate data.</div>";
    }
    return;
  }

  if (!DATA.factRows) DATA.factRows = [];
  if (!DATA.months) DATA.months = [];
  if (!DATA.states) DATA.states = [];

  /** States & UTs for filter dropdown (merged with any names present in data). */
  const INDIA_STATES_UT = [
    "Andaman and Nicobar Islands",
    "Andhra Pradesh",
    "Arunachal Pradesh",
    "Assam",
    "Bihar",
    "Chandigarh",
    "Chhattisgarh",
    "Dadra and Nagar Haveli and Daman and Diu",
    "Delhi",
    "Goa",
    "Gujarat",
    "Haryana",
    "Himachal Pradesh",
    "Jammu and Kashmir",
    "Jharkhand",
    "Karnataka",
    "Kerala",
    "Ladakh",
    "Lakshadweep",
    "Madhya Pradesh",
    "Maharashtra",
    "Manipur",
    "Meghalaya",
    "Mizoram",
    "Nagaland",
    "Odisha",
    "Puducherry",
    "Punjab",
    "Rajasthan",
    "Sikkim",
    "Tamil Nadu",
    "Telangana",
    "Tripura",
    "Uttar Pradesh",
    "Uttarakhand",
    "West Bengal",
  ];

  function mergedIndiaStateList(fromData) {
    const set = new Set(INDIA_STATES_UT);
    (fromData || []).forEach((s) => {
      const t = s == null ? "" : String(s).trim();
      if (t) set.add(t);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }

  const meta = DATA.meta || {};

  /** Detail table rows per page (keep in sync with styles.css --detail-table-body-rows) */
  const PAGE_SIZE = 5;
  let tableState = {
    sortKey: "yearMonth",
    asc: false,
    page: 0,
  };

  let currentCategoryKey = null;
  let catSearchAnnounceTimer = null;

  /** Category keys with full KPI drill-down: 1 Incident Management, 2 Hazard & Observation Management (Leading). */
  const ACTIVE_PREVIEW_CATEGORY_KEYS = new Set([1, 2]);

  function announce(msg) {
    if (liveRegion) liveRegion.textContent = msg;
  }

  /** Vs comparison (monthly facts). Labels read as day/week/month/quarter/year; math uses the closest monthly windows available. */
  const DEFAULT_VS_MODE = "vs_last_month";
  const VS_OPTIONS = [
    { id: "vs_yesterday", label: "Vs Yesterday" },
    { id: "vs_last_week", label: "Vs Last Week" },
    { id: "vs_last_month", label: "Vs Last Month" },
    { id: "vs_last_quarter", label: "Vs Last Quarter" },
    { id: "vs_last_year", label: "Vs Last Year" },
  ];

  const VS_PERIOD_CAPTION = {
    vs_yesterday: "Today vs yesterday",
    vs_last_week: "Current week vs last week",
    vs_last_month: "Current month vs last month",
    vs_last_quarter: "Current quarter vs last quarter",
    vs_last_year: "Current year vs last year",
  };

  function getRefMonth() {
    if (DATA.months && DATA.months.length) {
      return DATA.months[DATA.months.length - 1].yearMonth;
    }
    return meta.lastDataMonth || "2024-01";
  }

  function vsOptionLabel(id) {
    const o = VS_OPTIONS.find((x) => x.id === id);
    return o ? o.label : id;
  }

  /** Short tag on KPI tiles only (dropdown keeps full labels). */
  function vsTagShort(id) {
    switch (id) {
      case "vs_yesterday":
        return "y'day";
      case "vs_last_week":
        return "LW";
      case "vs_last_month":
        return "LM";
      case "vs_last_quarter":
        return "LQ";
      case "vs_last_year":
        return "LY";
      default:
        return vsOptionLabel(id);
    }
  }

  function formatMonthYear(ym) {
    if (!ym || ym.length < 7) return ym || "—";
    const parts = ym.split("-");
    const y = Number(parts[0]);
    const m = Number(parts[1]);
    if (Number.isNaN(y) || Number.isNaN(m)) return ym;
    const d = new Date(y, m - 1, 1);
    try {
      return d.toLocaleString(undefined, { month: "short", year: "numeric" });
    } catch {
      return ym;
    }
  }

  function isAdditiveUnit(ut) {
    return ut === "Count" || ut === "Hours" || ut === "Days";
  }

  /** ref, ref-1, … (length months). */
  function rollingMonthsFrom(endMonth, length) {
    const a = [];
    let m = endMonth;
    for (let i = 0; i < length; i++) {
      a.push(m);
      m = monthAdd(m, -1);
    }
    return a;
  }

  function tilePeriodForKpi(refMonth, mode, unitType) {
    void refMonth;
    void unitType;
    return VS_PERIOD_CAPTION[mode] || VS_PERIOD_CAPTION.vs_last_month;
  }

  /** IA: visible journey — aligns with header nav (Home / Categories); no duplicate controls. */
  function journeyStepsHtml(step) {
    function item(n, label, isActive) {
      const cls = isActive ? " route-steps__item--active" : "";
      const cur = isActive ? ' aria-current="step"' : "";
      return (
        '<span class="route-steps__item' +
        cls +
        '"' +
        cur +
        ">" +
        n +
        " " +
        escapeHtml(label) +
        "</span>"
      );
    }
    return (
      '<nav class="route-steps" aria-label="Journey: Home, Categories, Explore KPIs">' +
      item(1, "Home", step === 1) +
      '<span class="route-steps__sep" aria-hidden="true">→</span>' +
      item(2, "Categories", step === 2) +
      '<span class="route-steps__sep" aria-hidden="true">→</span>' +
      item(3, "Explore KPIs", step === 3) +
      "</nav>"
    );
  }

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function escapeAttr(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;");
  }

  const LS_VARIABLE_FILTER = "mis-variable-checkpoints";
  const CHECKPOINT_LABELS = ["Field Force", "O and M", "Office", "Projects"];

  function getRowCheckpoint(r) {
    if (r.checkpoint != null && String(r.checkpoint).trim() !== "") {
      return String(r.checkpoint);
    }
    const s =
      String(r.businessKey) +
      "\0" +
      String(r.kpiKey) +
      "\0" +
      String(r.yearMonth);
    let h = 2166136261;
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return CHECKPOINT_LABELS[Math.abs(h) % 4];
  }

  function rowMatchesVariable(r, variableSel) {
    if (variableSel == null || variableSel.length === 0) return true;
    return variableSel.includes(getRowCheckpoint(r));
  }

  function variableFilterFieldHtml() {
    return (
      '<div class="field field--variable field--var-inline">' +
      '<span class="field-label" id="f-var-lbl">Variable</span>' +
      '<details class="var-scope var-scope--toolbar" id="f-var-details">' +
      '<summary class="var-scope__summary" aria-labelledby="f-var-lbl">' +
      '<span class="var-scope__summary-text">' +
      '<span class="var-scope__hint" id="f-var-hint">All</span>' +
      "</span>" +
      '<span class="var-scope__chev" aria-hidden="true"></span>' +
      "</summary>" +
      '<div class="var-scope__panel" role="group" aria-label="Checkpoint options">' +
      '<div class="var-scope__panel-head">Checkpoints</div>' +
      '<div class="var-scope__menu">' +
      '<label class="field-variable-check field-variable-check--row field-variable-check--all">' +
      '<span class="field-variable-check__text">All checkpoints</span>' +
      '<input type="checkbox" id="f-var-all" checked />' +
      "</label>" +
      '<div class="var-scope__divider" aria-hidden="true"></div>' +
      '<div class="var-scope__options">' +
      '<label class="field-variable-check field-variable-check--row">' +
      '<span class="field-variable-check__text">Field Force</span>' +
      '<input type="checkbox" class="f-var-cb" value="Field Force" />' +
      "</label>" +
      '<label class="field-variable-check field-variable-check--row">' +
      '<span class="field-variable-check__text">O and M</span>' +
      '<input type="checkbox" class="f-var-cb" value="O and M" />' +
      "</label>" +
      '<label class="field-variable-check field-variable-check--row">' +
      '<span class="field-variable-check__text">Office</span>' +
      '<input type="checkbox" class="f-var-cb" value="Office" />' +
      "</label>" +
      '<label class="field-variable-check field-variable-check--row">' +
      '<span class="field-variable-check__text">Projects</span>' +
      '<input type="checkbox" class="f-var-cb" value="Projects" />' +
      "</label>" +
      "</div></div></div></details></div>"
    );
  }

  function updateVariableSummary() {
    const hint = document.getElementById("f-var-hint");
    const all = document.getElementById("f-var-all");
    const cbs = document.querySelectorAll("input.f-var-cb");
    if (!hint || !all) return;
    if (all.checked) {
      hint.textContent = "All";
      return;
    }
    const sel = [...cbs].filter((cb) => cb.checked).map((cb) => cb.value);
    if (sel.length === 0) {
      hint.textContent = "All";
      return;
    }
    if (sel.length === 1) {
      hint.textContent = sel[0];
      return;
    }
    if (sel.length === 2) {
      hint.textContent = sel.join(", ");
      return;
    }
    hint.textContent = sel.length + " selected";
  }

  function readVariableSelectionFromDom() {
    const all = document.getElementById("f-var-all");
    const cbs = document.querySelectorAll("input.f-var-cb");
    if (!all && !cbs.length) return null;
    if (all && all.checked) return null;
    const sel = [];
    cbs.forEach((cb) => {
      if (cb.checked) sel.push(cb.value);
    });
    return sel.length ? sel : null;
  }

  function saveVariableFilterToStorage() {
    try {
      localStorage.setItem(
        LS_VARIABLE_FILTER,
        JSON.stringify(readVariableSelectionFromDom())
      );
    } catch {
      /* ignore */
    }
  }

  function applyVariableFilterFromStorage() {
    const all = document.getElementById("f-var-all");
    const cbs = document.querySelectorAll("input.f-var-cb");
    if (!all || !cbs.length) return;
    let raw = null;
    try {
      raw = localStorage.getItem(LS_VARIABLE_FILTER);
    } catch {
      return;
    }
    if (raw == null || raw === "") {
      all.checked = true;
      cbs.forEach((cb) => {
        cb.checked = false;
      });
      return;
    }
    let sel = null;
    try {
      sel = JSON.parse(raw);
    } catch {
      return;
    }
    if (!Array.isArray(sel) || sel.length === 0) {
      all.checked = true;
      cbs.forEach((cb) => {
        cb.checked = false;
      });
      return;
    }
    all.checked = false;
    const set = new Set(sel);
    cbs.forEach((cb) => {
      cb.checked = set.has(cb.value);
    });
    if (![...cbs].some((cb) => cb.checked)) {
      all.checked = true;
      cbs.forEach((cb) => {
        cb.checked = false;
      });
    }
    updateVariableSummary();
  }

  function wireVariableFilterControls(onChange) {
    const all = document.getElementById("f-var-all");
    const cbs = document.querySelectorAll("input.f-var-cb");
    if (!all) return;
    function subChange() {
      const any = [...cbs].some((cb) => cb.checked);
      if (any) all.checked = false;
      else all.checked = true;
      saveVariableFilterToStorage();
      updateVariableSummary();
      if (onChange) onChange();
    }
    function allChange() {
      if (all.checked) {
        cbs.forEach((cb) => {
          cb.checked = false;
        });
      }
      saveVariableFilterToStorage();
      updateVariableSummary();
      if (onChange) onChange();
    }
    all.addEventListener("change", allChange);
    cbs.forEach((cb) => cb.addEventListener("change", subChange));
    updateVariableSummary();
  }

  /** Decorative icons for category cards (list context; button has aria-label). */
  function categoryIconSvg(categoryKey) {
    const svg = (paths) =>
      '<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      paths +
      "</svg>";
    switch (categoryKey) {
      case 1:
        return svg(
          '<path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>'
        );
      case 2:
        return svg(
          '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>'
        );
      case 3:
        return svg('<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>');
      case 4:
        return svg(
          '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>'
        );
      case 5:
        return svg(
          '<path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>'
        );
      case 6:
        return svg(
          '<path d="M4 19.5A2.5 2.5 0 016.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z"/>'
        );
      case 7:
        return svg(
          '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>'
        );
      case 8:
        return svg(
          '<path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/>'
        );
      case 9:
        return svg(
          '<rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><line x1="9" y1="1" x2="9" y2="4"/><line x1="15" y1="1" x2="15" y2="4"/><line x1="9" y1="20" x2="9" y2="23"/><line x1="15" y1="20" x2="15" y2="23"/><line x1="20" y1="9" x2="23" y2="9"/><line x1="20" y1="14" x2="23" y2="14"/><line x1="1" y1="9" x2="4" y2="9"/><line x1="1" y1="14" x2="4" y2="14"/>'
        );
      default:
        return svg('<circle cx="12" cy="12" r="10"/>');
    }
  }

  function formatValue(v, unitType) {
    if (v == null || v === "") return "—";
    const n = Number(v);
    if (Number.isNaN(n)) return "—";
    if (unitType === "Hours") {
      return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
    }
    if (unitType === "Days") {
      return n.toLocaleString(undefined, { maximumFractionDigits: 1 });
    }
    if (unitType === "PercentOrRate") {
      if (Math.abs(n) < 5 && !Number.isInteger(n)) {
        return n.toLocaleString(undefined, { maximumFractionDigits: 4 });
      }
      return n.toLocaleString(undefined, { maximumFractionDigits: 1 }) + "%";
    }
    return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }

  function getCategory(catKey) {
    return DATA.categories.find((c) => c.categoryKey === catKey);
  }

  function getKpis(catKey) {
    const k = DATA.kpiDetailByCategory.find((x) => x.categoryKey === catKey);
    return k ? k.kpis : [];
  }

  /** Category map search: match category name or any KPI name in that category. */
  function categoryMatchesSearchQuery(cat, qLower) {
    if (!qLower) return true;
    if ((cat.categoryName || "").toLowerCase().includes(qLower)) return true;
    return getKpis(cat.categoryKey).some((k) =>
      (k.kpiName || "").toLowerCase().includes(qLower)
    );
  }

  function kpiUnitTypeForFilter(catKey, f) {
    if (f.kpi === "all") return null;
    const kpis = getKpis(catKey);
    const k = kpis.find((x) => String(x.kpiKey) === String(f.kpi));
    return k ? k.unitType : null;
  }

  function destroyCharts() {
    ["chart-line", "chart-area-state", "chart-biz"].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.destroy();
      }
    });
  }

  /** Approved brand colors only — cycle for charts */
  const CHART_BRAND_HEX = ["#00B16B", "#006DB6", "#8E278F", "#F04C23"];

  /**
   * By business: doughnut chart — share of total (|value|) per business; top slices + Other.
   */
  function renderBizBreakdown(bizLabels, bizData) {
    const el = document.getElementById("chart-biz");
    const emptyEl = document.getElementById("chart-biz-empty");
    if (!el || typeof Chart === "undefined") return;
    const prev = Chart.getChart(el);
    if (prev) prev.destroy();

    if (!bizLabels.length) {
      el.style.display = "none";
      if (emptyEl) {
        emptyEl.hidden = false;
        emptyEl.textContent = "No business data for current filters.";
      }
      return;
    }

    const pairs = bizLabels
      .map((name, i) => ({
        name,
        v: Math.abs(Number(bizData[i]) || 0),
      }))
      .filter((p) => p.v > 0);
    if (!pairs.length) {
      el.style.display = "none";
      if (emptyEl) {
        emptyEl.hidden = false;
        emptyEl.textContent = "No business data for current filters.";
      }
      return;
    }

    el.style.display = "block";
    if (emptyEl) emptyEl.hidden = true;

    pairs.sort((a, b) => b.v - a.v);
    const maxSlices = 7;
    let fullNames = [];
    let labels = [];
    let values = [];
    function shortLabel(name) {
      const s = String(name);
      return s.length > 12 ? s.slice(0, 11) + "…" : s;
    }
    if (pairs.length <= maxSlices) {
      fullNames = pairs.map((p) => p.name);
      labels = fullNames.map(shortLabel);
      values = pairs.map((p) => p.v);
    } else {
      const top = pairs.slice(0, maxSlices);
      const rest = pairs.slice(maxSlices);
      const other = rest.reduce((s, p) => s + p.v, 0);
      fullNames = top.map((p) => p.name);
      labels = fullNames.map(shortLabel);
      values = top.map((p) => p.v);
      if (other > 0) {
        fullNames.push("Other (" + rest.length + " businesses)");
        labels.push("Other");
        values.push(other);
      }
    }

    const total = values.reduce((a, b) => a + b, 0) || 1;

    new Chart(el, {
      type: "doughnut",
      data: {
        labels: labels,
        datasets: [
          {
            data: values,
            backgroundColor: labels.map(
              (_, i) => CHART_BRAND_HEX[i % CHART_BRAND_HEX.length]
            ),
            borderColor: "#ffffff",
            borderWidth: 1.5,
            hoverOffset: 3,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "54%",
        plugins: {
          legend: {
            position: "right",
            align: "center",
            labels: {
              font: { size: 7 },
              boxWidth: 8,
              boxHeight: 8,
              padding: 4,
              color: "#231F20",
            },
          },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0].dataIndex;
                return fullNames[i] != null ? String(fullNames[i]) : "";
              },
              label(ctx) {
                const raw = ctx.raw;
                const pct = ((Number(raw) / total) * 100).toFixed(1);
                const v = Number(raw);
                const num =
                  v >= 1e6
                    ? v.toLocaleString(undefined, { maximumFractionDigits: 0 })
                    : v.toLocaleString(undefined, { maximumFractionDigits: 2 });
                return " " + num + " (" + pct + "% of total)";
              },
            },
          },
        },
      },
    });
  }

  function readFilters(catKey) {
    const elKpi = document.getElementById("f-kpi");
    const elVs = document.getElementById("f-vs");
    const elSt = document.getElementById("f-state");
    if (!elSt) return null;
    const elBiz = document.getElementById("f-biz");
    return {
      catKey,
      kpi: elKpi ? elKpi.value : "all",
      vsMode: elVs ? elVs.value : DEFAULT_VS_MODE,
      refMonth: getRefMonth(),
      state: elSt.value,
      business: elBiz ? elBiz.value : "all",
      unitType: "all",
      variable: readVariableSelectionFromDom(),
    };
  }

  /** Table / KPI tiles: current reference month only (plus non-month filters). */
  function applyRowFilter(rows, f) {
    return rows.filter((r) => {
      if (f.kpi !== "all" && String(r.kpiKey) !== f.kpi) return false;
      if (r.yearMonth !== f.refMonth) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      return true;
    });
  }

  /** Charts: last 12 months ending at ref month. */
  function chartMonthKeys(refMonth, count) {
    const keys = [];
    let m = refMonth;
    for (let i = 0; i < count; i++) {
      keys.push(m);
      m = monthAdd(m, -1);
    }
    return keys.reverse();
  }

  function applyChartFilter(rows, f) {
    const range = new Set(
      chartMonthsForVsMode(f.vsMode || DEFAULT_VS_MODE, f.refMonth)
    );
    return rows.filter((r) => {
      if (f.kpi !== "all" && String(r.kpiKey) !== f.kpi) return false;
      if (!range.has(r.yearMonth)) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      return true;
    });
  }


  function getRowsForCategory(catKey) {
    return DATA.factRows.filter((r) => r.categoryKey === catKey);
  }

  function distinctSorted(rows, pick) {
    const set = new Set();
    rows.forEach((r) => {
      const v = pick(r);
      if (v != null && v !== "") set.add(String(v));
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }

  function getFilterConfig(_catKey) {
    return {
      showKpi: true,
      showState: true,
      showBusiness: true,
    };
  }

  function avg(nums) {
    if (!nums.length) return 0;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  }

  /** Incident Management: wireframe order — chunk 4 = ΔRepeat, ΔFatal, Man-days, Vehicle. */
  const INCIDENT_KPI_ORDER = [
    1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 14, 15, 19, 22, 28, 44,
  ];

  function sortKpisForDisplay(catKey, kpisMeta) {
    if (catKey === 1) {
      const map = new Map(kpisMeta.map((k) => [k.kpiKey, k]));
      return INCIDENT_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    return kpisMeta.slice();
  }

  function monthAdd(ym, delta) {
    if (!ym || ym.length < 7) return null;
    const parts = ym.split("-");
    const y = Number(parts[0]);
    const m = Number(parts[1]);
    const d = new Date(y, m - 1 + delta, 1);
    return (
      d.getFullYear() +
      "-" +
      String(d.getMonth() + 1).padStart(2, "0")
    );
  }

  /** Jan … ref month (calendar YTD), same month count in prior year. */
  function monthsYtdThrough(ym) {
    if (!ym || ym.length < 7) return [];
    const y = Number(ym.split("-")[0]);
    const mo = Number(ym.split("-")[1]);
    if (Number.isNaN(y) || Number.isNaN(mo)) return [];
    const out = [];
    for (let m = 1; m <= mo; m++) {
      out.push(y + "-" + String(m).padStart(2, "0"));
    }
    return out;
  }

  function priorYearYtdMonths(ym) {
    if (!ym || ym.length < 7) return [];
    const y = Number(ym.split("-")[0]);
    const mo = Number(ym.split("-")[1]);
    if (Number.isNaN(y) || Number.isNaN(mo)) return [];
    const out = [];
    for (let m = 1; m <= mo; m++) {
      out.push(y - 1 + "-" + String(m).padStart(2, "0"));
    }
    return out;
  }

  /** Months shown on the trend chart for the active Vs mode. */
  function chartMonthsForVsMode(mode, ref) {
    if (!ref) return chartMonthKeys(getRefMonth(), 12);
    const m1 = monthAdd(ref, -1);
    const m2 = monthAdd(ref, -2);
    if (mode === "vs_yesterday") {
      return [m1, ref].filter(Boolean);
    }
    if (mode === "vs_last_week") {
      return [monthAdd(ref, -3), monthAdd(ref, -2), m1, ref].filter(Boolean);
    }
    if (mode === "vs_last_month") {
      return [m2, m1, ref].filter(Boolean);
    }
    if (mode === "vs_last_quarter") {
      const a = [];
      for (let i = 5; i >= 0; i--) a.push(monthAdd(ref, -i));
      return a.filter(Boolean);
    }
    return chartMonthKeys(ref, 12);
  }

  /** Months rolled into By business / Unit mix for the active Vs mode. */
  function bizUnitWindowMonths(mode, ref) {
    if (!ref) return [getRefMonth()];
    const m1 = monthAdd(ref, -1);
    const m2 = monthAdd(ref, -2);
    if (mode === "vs_yesterday" || mode === "vs_last_month") {
      return [ref];
    }
    if (mode === "vs_last_week") {
      return [m1, ref].filter(Boolean);
    }
    if (mode === "vs_last_quarter") {
      return [m2, m1, ref].filter(Boolean);
    }
    if (mode === "vs_last_year") {
      return monthsYtdThrough(ref);
    }
    return [ref];
  }

  function applyNonMonthFilters(rows, f) {
    return rows.filter((r) => {
      if (f.kpi !== "all" && String(r.kpiKey) !== f.kpi) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      return true;
    });
  }

  function pctChange(cur, prev) {
    if (prev == null || cur == null) return null;
    const p = Number(prev);
    const c = Number(cur);
    if (Number.isNaN(p) || Number.isNaN(c)) return null;
    if (Math.abs(p) < 1e-12) return null;
    return ((c - p) / p) * 100;
  }

  function vsDir(cur, base) {
    if (cur == null || base == null) return "neutral";
    if (cur > base) return "up";
    if (cur < base) return "down";
    return "same";
  }

  function formatSignedPct(p) {
    if (p == null || Number.isNaN(p)) return "—";
    const sign = p >= 0 ? "+" : "";
    return sign + p.toFixed(1) + "%";
  }

  function buildKpiDetailMetrics(catKey, kpisMeta, f) {
    let sorted = sortKpisForDisplay(catKey, kpisMeta);
    if (f.kpi !== "all") {
      sorted = sorted.filter((k) => String(k.kpiKey) === String(f.kpi));
    }
    const base = applyNonMonthFilters(getRowsForCategory(catKey), f);
    const ref = f.refMonth;
    const mode = f.vsMode || DEFAULT_VS_MODE;
    const m1 = monthAdd(ref, -1);
    const m2 = monthAdd(ref, -2);
    const m3 = monthAdd(ref, -3);
    const m4 = monthAdd(ref, -4);
    const m5 = monthAdd(ref, -5);

    return sorted.map((k) => {
      const kk = k.kpiKey;
      const ut = k.unitType;
      const add = isAdditiveUnit(ut);

      function aggMonth(ym) {
        if (!ym) return null;
        const rows = base.filter(
          (r) => r.kpiKey === kk && r.yearMonth === ym
        );
        if (!rows.length) return null;
        if (add) {
          return rows.reduce((a, r) => a + Number(r.value), 0);
        }
        return avg(rows.map((r) => r.value));
      }

      function aggMonthsRolling(endYm, numMonths) {
        const yms = rollingMonthsFrom(endYm, numMonths);
        const vals = yms.map(aggMonth).filter((v) => v != null && !Number.isNaN(v));
        if (!vals.length) return null;
        if (add) return vals.reduce((a, b) => a + b, 0);
        return avg(vals);
      }

      function avgMonthsList(yms) {
        const vals = yms.map(aggMonth).filter((v) => v != null && !Number.isNaN(v));
        if (!vals.length) return null;
        return avg(vals);
      }

      let cur;
      let baseVal;
      if (mode === "vs_last_month" || mode === "vs_yesterday") {
        cur = aggMonth(ref);
        baseVal = aggMonth(m1);
      } else if (mode === "vs_last_year") {
        const ytdM = monthsYtdThrough(ref);
        const pyM = priorYearYtdMonths(ref);
        if (add) {
          const cv = ytdM.map(aggMonth).filter((v) => v != null && !Number.isNaN(v));
          const bv = pyM.map(aggMonth).filter((v) => v != null && !Number.isNaN(v));
          cur = cv.length ? cv.reduce((a, b) => a + b, 0) : null;
          baseVal = bv.length ? bv.reduce((a, b) => a + b, 0) : null;
        } else {
          const cv = ytdM.map(aggMonth).filter((v) => v != null && !Number.isNaN(v));
          const bv = pyM.map(aggMonth).filter((v) => v != null && !Number.isNaN(v));
          cur = cv.length ? avg(cv) : null;
          baseVal = bv.length ? avg(bv) : null;
        }
      } else if (mode === "vs_last_week") {
        if (add) {
          cur = aggMonthsRolling(ref, 2);
          baseVal = aggMonthsRolling(m2, 2);
        } else {
          cur = avgMonthsList([ref, m1]);
          baseVal = avgMonthsList([m2, m3]);
        }
      } else if (mode === "vs_last_quarter") {
        if (add) {
          cur = aggMonthsRolling(ref, 3);
          baseVal = aggMonthsRolling(m3, 3);
        } else {
          cur = avgMonthsList([ref, m1, m2]);
          baseVal = avgMonthsList([m3, m4, m5]);
        }
      } else {
        cur = aggMonth(ref);
        baseVal = aggMonth(m1);
      }

      const vsPct = pctChange(cur, baseVal);
      const vsD = vsDir(cur, baseVal);
      return {
        kpiKey: kk,
        kpiName: k.kpiName,
        unitType: ut,
        value: cur,
        vsPct: vsPct,
        vsDir: vsD,
        vsMode: mode,
        periodCaption: tilePeriodForKpi(ref, mode, ut),
      };
    });
  }

  function chunkArray(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) {
      out.push(arr.slice(i, i + size));
    }
    return out;
  }

  function cmpArrowClass(dir) {
    if (dir === "up") return "multi-kpi-tile__arrow multi-kpi-tile__arrow--up";
    if (dir === "down") return "multi-kpi-tile__arrow multi-kpi-tile__arrow--down";
    if (dir === "same") return "multi-kpi-tile__arrow multi-kpi-tile__arrow--same";
    return "multi-kpi-tile__arrow multi-kpi-tile__arrow--na";
  }

  function arrowGlyph(dir) {
    if (dir === "up") return "▲";
    if (dir === "down") return "▼";
    if (dir === "same") return "◆";
    return "—";
  }

  /** Short KPI names first; long names (likely two lines in the tile) sort to the bottom. */
  function sortKpiTilesForDisplay(items) {
    return items.slice().sort((a, b) => {
      const la = (a.kpiName || "").length;
      const lb = (b.kpiName || "").length;
      const longA = la > 26 ? 1 : 0;
      const longB = lb > 26 ? 1 : 0;
      if (longA !== longB) return longA - longB;
      return la - lb;
    });
  }

  function appendKpiTileEl(grid, item, refMonth, vsMode, tileClass) {
    const tile = document.createElement("div");
    tile.className = tileClass || "multi-kpi-tile";
    const vsPct = formatSignedPct(item.vsPct);
    const vDir = item.vsDir || "neutral";
    const vsLbl = vsOptionLabel(item.vsMode);
    const vsShort = vsTagShort(item.vsMode);
    const periodLine =
      item.periodCaption ||
      tilePeriodForKpi(refMonth, vsMode || DEFAULT_VS_MODE, item.unitType);
    tile.setAttribute(
      "aria-label",
      item.kpiName +
        ". Data " +
        periodLine +
        ". Value: " +
        formatValue(item.value, item.unitType) +
        ". " +
        vsLbl +
        " " +
        vsPct +
        "."
    );
    tile.innerHTML =
      '<div class="multi-kpi-tile__name">' +
      escapeHtml(item.kpiName) +
      "</div>" +
      '<div class="multi-kpi-tile__value-row">' +
      '<span class="multi-kpi-tile__val">' +
      formatValue(item.value, item.unitType) +
      "</span>" +
      "</div>" +
      '<div class="multi-kpi-tile__unit">' +
      escapeHtml(item.unitType) +
      "</div>" +
      '<div class="multi-kpi-tile__period">' +
      escapeHtml(periodLine) +
      "</div>" +
      '<div class="multi-kpi-tile__cmp multi-kpi-tile__cmp--vs">' +
      '<span class="multi-kpi-tile__cmp-tag multi-kpi-tile__cmp-tag--short">' +
      escapeHtml(vsShort) +
      "</span>" +
      '<span class="' +
      cmpArrowClass(vDir) +
      '" aria-hidden="true">' +
      arrowGlyph(vDir) +
      "</span>" +
      '<span class="multi-kpi-tile__cmp-pct vs-pct-num">' +
      escapeHtml(vsPct) +
      "</span>" +
      "</div>";
    grid.appendChild(tile);
  }

  function renderMultiKpiCards(container, aggregatesList, refMonth, vsMode) {
    container.innerHTML = "";
    const sorted = sortKpiTilesForDisplay(aggregatesList);
    const headItems = sorted.slice(0, 12);
    const tailItems = sorted.slice(12);

    const rowMain = document.createElement("div");
    rowMain.className = "multi-kpi-row";

    function appendMetricCard(startMetric, items, useTailStyle) {
      if (!items.length) return;
      const card = document.createElement("div");
      card.className = useTailStyle
        ? "multi-kpi-card multi-kpi-card--tail"
        : "multi-kpi-card";
      const head = document.createElement("div");
      head.className = useTailStyle
        ? "multi-kpi-card__head multi-kpi-card__head--tail"
        : "multi-kpi-card__head";
      const endMetric = startMetric + items.length - 1;
      head.textContent =
        items.length === 1
          ? "Metric " + startMetric
          : "Metrics " + startMetric + "–" + endMetric;
      const grid = document.createElement("div");
      grid.className = "multi-kpi-card__grid";
      const tileClass = useTailStyle
        ? "multi-kpi-tile multi-kpi-tile--tail"
        : "multi-kpi-tile";
      items.forEach((item) => {
        appendKpiTileEl(grid, item, refMonth, vsMode, tileClass);
      });
      while (grid.children.length < 4) {
        const ph = document.createElement("div");
        ph.className = tileClass;
        ph.style.visibility = "hidden";
        ph.innerHTML = "&nbsp;";
        grid.appendChild(ph);
      }
      card.appendChild(head);
      card.appendChild(grid);
      rowMain.appendChild(card);
    }

    const chunks = chunkArray(headItems, 4);
    chunks.forEach((chunk, idx) => {
      appendMetricCard(idx * 4 + 1, chunk, false);
    });

    const tailA = tailItems.slice(0, 4);
    const tailB = tailItems.slice(4);
    if (tailA.length) {
      appendMetricCard(13, tailA, true);
    }
    if (tailB.length) {
      appendMetricCard(13 + tailA.length, tailB, true);
    }

    container.appendChild(rowMain);
  }

  /**
   * Trend + business + unit charts respect Vs mode windows (monthly data).
   */
  function buildCharts(catKey, f) {
    destroyCharts();

    const poolAll = getRowsForCategory(catKey);
    const filteredTrend = applyChartFilter(poolAll, f);
    const snapPool = applyNonMonthFilters(poolAll, f);
    const winMonths = new Set(
      bizUnitWindowMonths(f.vsMode || DEFAULT_VS_MODE, f.refMonth)
    );
    const snapRows = snapPool.filter((r) => winMonths.has(r.yearMonth));

    const byMonth = {};
    filteredTrend.forEach((r) => {
      if (!byMonth[r.yearMonth]) byMonth[r.yearMonth] = [];
      byMonth[r.yearMonth].push(r.value);
    });
    const monthKeys = Object.keys(byMonth).sort();
    const lineLabels = monthKeys;
    const lineData = monthKeys.map((m) => avg(byMonth[m]));

    const byBizVals = {};
    snapRows.forEach((r) => {
      const b = r.businessName || "—";
      if (!byBizVals[b]) byBizVals[b] = [];
      byBizVals[b].push(Number(r.value));
    });
    const bizLabels = Object.keys(byBizVals);
    const utChart = kpiUnitTypeForFilter(catKey, f);
    const bizData = bizLabels.map((b) => {
      const vals = byBizVals[b];
      if (utChart && isAdditiveUnit(utChart)) {
        return vals.reduce((a, x) => a + Number(x), 0);
      }
      return avg(vals);
    });

    /** Location / state mix (same roll-up window as former unit doughnut). */
    const stateMap = {};
    snapRows.forEach((r) => {
      const s = r.state || "—";
      if (f.kpi === "all") {
        stateMap[s] = (stateMap[s] || 0) + 1;
      } else {
        const v = Number(r.value);
        stateMap[s] = (stateMap[s] || 0) + (Number.isFinite(v) ? v : 0);
      }
    });
    const stateEntries = Object.keys(stateMap).map((k) => ({
      state: k,
      v: stateMap[k],
    }));
    stateEntries.sort((a, b) => b.v - a.v);
    const topStates = stateEntries.slice(0, 14);
    const stateLabels = topStates.map((e) => e.state);
    const stateData = topStates.map((e) => e.v);

    const elLine = document.getElementById("chart-line");
    if (elLine && lineLabels.length) {
      new Chart(elLine, {
        type: "line",
        data: {
          labels: lineLabels,
          datasets: [
            {
              label: "Avg",
              data: lineData,
              borderColor: "#006DB6",
              backgroundColor: "rgba(0, 177, 107, 0.12)",
              fill: true,
              tension: 0.2,
              borderWidth: 1.5,
              pointRadius: 2,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            y: {
              beginAtZero: false,
              ticks: { font: { size: 9 }, color: "#231F20" },
              grid: { color: "rgba(109, 110, 113, 0.2)" },
            },
            x: {
              ticks: { font: { size: 8 }, maxRotation: 45, color: "#231F20" },
              grid: { color: "rgba(109, 110, 113, 0.15)" },
            },
          },
        },
      });
    }

    renderBizBreakdown(bizLabels, bizData);

    const elArea = document.getElementById("chart-area-state");
    if (elArea && stateLabels.length) {
      new Chart(elArea, {
        type: "line",
        data: {
          labels: stateLabels,
          datasets: [
            {
              label: "By state",
              data: stateData,
              borderColor: "#006DB6",
              backgroundColor: "rgba(0, 109, 182, 0.18)",
              fill: true,
              tension: 0.35,
              borderWidth: 2,
              pointRadius: 3,
              pointBackgroundColor: "#006DB6",
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                title(items) {
                  const i = items[0].dataIndex;
                  return stateLabels[i] || "";
                },
              },
            },
          },
          scales: {
            y: {
              beginAtZero: true,
              ticks: { font: { size: 9 }, color: "#231F20" },
              grid: { color: "rgba(109, 110, 113, 0.2)" },
              title: {
                display: true,
                text: f.kpi === "all" ? "Row count" : "Value",
                font: { size: 9 },
                color: "#6D6E71",
              },
            },
            x: {
              ticks: {
                font: { size: 7 },
                maxRotation: 55,
                minRotation: 35,
                color: "#231F20",
              },
              grid: { color: "rgba(109, 110, 113, 0.12)" },
            },
          },
        },
      });
    }

    updateChartHints(f);
  }

  function updateChartHints(f) {
    const mode = f.vsMode || DEFAULT_VS_MODE;
    const ref = f.refMonth;
    const trendEl = document.getElementById("chart-trend-hint");
    const bizEl = document.getElementById("chart-biz-hint");
    const stateLocEl = document.getElementById("chart-state-hint");
    const n = chartMonthsForVsMode(mode, ref).length;
    if (trendEl) {
      trendEl.textContent =
        mode === "vs_last_year" ? "(12 months)" : "(" + n + " months)";
    }
    const bizHint = {
      vs_yesterday: "(latest month)",
      vs_last_month: "(latest month)",
      vs_last_week: "(last 2 months)",
      vs_last_quarter: "(last 3 months)",
      vs_last_year: "(YTD)",
    };
    const bh = bizHint[mode] || "(latest month)";
    if (bizEl) bizEl.textContent = bh;
    if (stateLocEl) stateLocEl.textContent = bh;
  }

  function sortRows(rows, key, asc, pool, f) {
    const copy = rows.slice();
    const mul = asc ? 1 : -1;
    copy.sort((a, b) => {
      let va = a[key];
      let vb = b[key];
      if (key === "value" || key === "target") {
        if (key === "value" && pool && f) {
          va = rowCompareForMode(a, pool, f).cur;
          vb = rowCompareForMode(b, pool, f).cur;
        }
        va =
          va == null || va === ""
            ? Number.NEGATIVE_INFINITY
            : Number(va);
        vb =
          vb == null || vb === ""
            ? Number.NEGATIVE_INFINITY
            : Number(vb);
        if (Number.isNaN(va)) va = Number.NEGATIVE_INFINITY;
        if (Number.isNaN(vb)) vb = Number.NEGATIVE_INFINITY;
        return (va - vb) * mul;
      }
      va = va == null ? "" : String(va);
      vb = vb == null ? "" : String(vb);
      return va.localeCompare(vb, undefined, { numeric: true }) * mul;
    });
    return copy;
  }

  function rowTupleKey(r) {
    return (
      String(r.kpiKey) +
      "|" +
      r.state +
      "|" +
      r.businessName +
      "|" +
      r.unitType
    );
  }

  function rowValueAt(pool, tupleKey, ym) {
    const hit = pool.find(
      (x) => rowTupleKey(x) === tupleKey && x.yearMonth === ym
    );
    return hit == null ? null : Number(hit.value);
  }

  function rowWindowAgg(pool, tupleKey, endMonth, numMonths, unitType) {
    if (numMonths <= 0) return null;
    const nums = [];
    let m = endMonth;
    for (let i = 0; i < numMonths; i++) {
      const v = rowValueAt(pool, tupleKey, m);
      if (v != null && !Number.isNaN(v)) nums.push(v);
      m = monthAdd(m, -1);
    }
    if (!nums.length) return null;
    if (isAdditiveUnit(unitType)) return nums.reduce((a, b) => a + b, 0);
    return avg(nums);
  }

  /**
   * Row-level current vs base (matches KPI tile windows). Used for Vs % and Value column.
   */
  function rowCompareForMode(row, pool, f) {
    const mode = f.vsMode || DEFAULT_VS_MODE;
    const ref = row.yearMonth;
    const tk = rowTupleKey(row);
    const ut = row.unitType || "";
    const add = isAdditiveUnit(ut);
    const m1 = monthAdd(ref, -1);
    const m2 = monthAdd(ref, -2);
    const m3 = monthAdd(ref, -3);
    const m4 = monthAdd(ref, -4);
    const m5 = monthAdd(ref, -5);

    function avgPair(yms) {
      const nums = yms
        .map((ym) => rowValueAt(pool, tk, ym))
        .filter((v) => v != null && !Number.isNaN(v));
      if (!nums.length) return null;
      return avg(nums);
    }

    let cur;
    let baseVal;
    if (mode === "vs_last_month" || mode === "vs_yesterday") {
      cur = rowValueAt(pool, tk, ref);
      baseVal = rowValueAt(pool, tk, m1);
    } else if (mode === "vs_last_year") {
      const ytd = monthsYtdThrough(ref);
      const py = priorYearYtdMonths(ref);
      if (add) {
        const cv = ytd
          .map((ym) => rowValueAt(pool, tk, ym))
          .filter((v) => v != null && !Number.isNaN(v));
        const bv = py
          .map((ym) => rowValueAt(pool, tk, ym))
          .filter((v) => v != null && !Number.isNaN(v));
        cur = cv.length ? cv.reduce((a, b) => a + b, 0) : null;
        baseVal = bv.length ? bv.reduce((a, b) => a + b, 0) : null;
      } else {
        cur = avgPair(ytd);
        baseVal = avgPair(py);
      }
    } else if (mode === "vs_last_week") {
      if (add) {
        cur = rowWindowAgg(pool, tk, ref, 2, ut);
        baseVal = rowWindowAgg(pool, tk, m2, 2, ut);
      } else {
        cur = avgPair([ref, m1]);
        baseVal = avgPair([m2, m3]);
      }
    } else if (mode === "vs_last_quarter") {
      if (add) {
        cur = rowWindowAgg(pool, tk, ref, 3, ut);
        baseVal = rowWindowAgg(pool, tk, m3, 3, ut);
      } else {
        cur = avgPair([ref, m1, m2]);
        baseVal = avgPair([m3, m4, m5]);
      }
    } else {
      cur = rowValueAt(pool, tk, ref);
      baseVal = rowValueAt(pool, tk, m1);
    }
    const p = pctChange(cur, baseVal);
    return {
      cur: cur,
      baseVal: baseVal,
      pct: p,
      dir: vsDir(cur, baseVal),
    };
  }

  function rowVsForMode(row, pool, f) {
    const x = rowCompareForMode(row, pool, f);
    return { pct: x.pct, dir: x.dir };
  }

  /** Table value column: “current” side of the active Vs window (not only the row month when window is wider). */
  function tableValueCell(r, pool, f) {
    const unit = r.unitType || "";
    const cmp = rowCompareForMode(r, pool, f);
    const actualStr = formatValue(cmp.cur, unit);
    return (
      '<td class="col-num">' +
      '<span class="data-table__value-num">' +
      escapeHtml(actualStr) +
      "</span></td>"
    );
  }

  function renderTableBody(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    const pool = applyNonMonthFilters(getRowsForCategory(catKey), f);
    let rows = applyRowFilter(pool, f);
    rows = sortRows(rows, tableState.sortKey, tableState.asc, pool, f);

    const total = rows.length;
    const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
    if (tableState.page >= pages) tableState.page = pages - 1;
    if (tableState.page < 0) tableState.page = 0;

    const start = tableState.page * PAGE_SIZE;
    const pageRows = rows.slice(start, start + PAGE_SIZE);

    const tbody = document.getElementById("tbl-body");
    const pageInfo = document.getElementById("tbl-pageinfo");
    const btnP = document.getElementById("tbl-prev");
    const btnN = document.getElementById("tbl-next");

    if (!tbody) return;

    const summaryEl = document.getElementById("tbl-summary");
    if (summaryEl) {
      if (total === 0) {
        summaryEl.textContent =
          "No rows match for the reference month. Set State to All or reset filters.";
      } else {
        summaryEl.textContent =
          "Showing " +
          pageRows.length +
          " of " +
          total +
          " matching rows on this page. Click column headers to sort.";
      }
    }

    if (!pageRows.length) {
      tbody.innerHTML =
        '<tr><td colspan="8" class="empty-msg">No rows for current filters.</td></tr>';
    } else {
      tbody.innerHTML = pageRows
        .map((r) => {
          const rv = rowVsForMode(r, pool, f);
          const pctStr = formatSignedPct(rv.pct);
          const arrow =
            rv.dir === "up"
              ? "▲"
              : rv.dir === "down"
                ? "▼"
                : rv.dir === "same"
                  ? "◆"
                  : "—";
          return (
            '<tr aria-label="' +
            escapeAttr(
              r.kpiName + " — " + r.businessName + " — " + r.yearMonth
            ) +
            '">' +
            "<td>" +
            escapeHtml(r.yearMonth) +
            "</td>" +
            "<td>" +
            escapeHtml(r.state) +
            "</td>" +
            "<td>" +
            escapeHtml(r.businessName) +
            "</td>" +
            '<td class="col-kpi">' +
            escapeHtml(r.kpiName) +
            "</td>" +
            "<td>" +
            escapeHtml(r.unitType) +
            "</td>" +
            tableValueCell(r, pool, f) +
            '<td class="col-num data-table__vs">' +
            '<span class="multi-kpi-tile__arrow data-table__vs-arrow ' +
            (rv.dir === "up"
              ? "multi-kpi-tile__arrow--up"
              : rv.dir === "down"
                ? "multi-kpi-tile__arrow--down"
                : "multi-kpi-tile__arrow--na") +
            '" aria-hidden="true">' +
            arrow +
            '</span> <span class="vs-pct-num">' +
            escapeHtml(pctStr) +
            "</span></td>" +
            '<td class="col-num">' +
            (r.target == null || r.target === ""
              ? "—"
              : formatValue(r.target, r.unitType)) +
            "</td>" +
            "</tr>"
          );
        })
        .join("");
    }

    if (pageInfo) {
      pageInfo.textContent =
        total === 0
          ? "0 rows"
          : "Rows " +
            (start + 1) +
            "–" +
            Math.min(start + PAGE_SIZE, total) +
            " of " +
            total +
            " (page " +
            (tableState.page + 1) +
            "/" +
            pages +
            ")";
    }
    if (btnP) {
      btnP.disabled = tableState.page <= 0;
    }
    if (btnN) {
      btnN.disabled = tableState.page >= pages - 1;
    }

    document.querySelectorAll("#tbl-detail th[data-sort]").forEach((th) => {
      const k = th.getAttribute("data-sort");
      th.setAttribute(
        "aria-sort",
        k === tableState.sortKey
          ? tableState.asc
            ? "ascending"
            : "descending"
          : "none"
      );
    });
  }

  function wireTableHeaders(catKey) {
    document.querySelectorAll("#tbl-detail th[data-sort]").forEach((th) => {
      th.addEventListener("click", () => {
        const k = th.getAttribute("data-sort");
        if (tableState.sortKey === k) {
          tableState.asc = !tableState.asc;
        } else {
          tableState.sortKey = k;
          tableState.asc = true;
        }
        tableState.page = 0;
        renderTableBody(catKey);
      });
    });
    const btnP = document.getElementById("tbl-prev");
    const btnN = document.getElementById("tbl-next");
    if (btnP) {
      btnP.addEventListener("click", () => {
        tableState.page--;
        renderTableBody(catKey);
      });
    }
    if (btnN) {
      btnN.addEventListener("click", () => {
        tableState.page++;
        renderTableBody(catKey);
      });
    }
  }

  function announceFilterSummary(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    const n = applyRowFilter(
      applyNonMonthFilters(getRowsForCategory(catKey), f),
      f
    ).length;
    const cat = getCategory(catKey);
    const label = cat ? cat.categoryName + ". " : "";
    announce(
      label +
        n +
        " detail rows for " +
        formatMonthYear(f.refMonth) +
        ". Sort with column headers; Prev and Next move pages."
    );
  }

  function refreshCategoryView(catKey) {
    const kpisMeta = getKpis(catKey);
    const f = readFilters(catKey);
    if (!f) return;

    const aggList = buildKpiDetailMetrics(catKey, kpisMeta, f);

    const multiWrap = document.getElementById("multi-kpi-wrap");
    if (multiWrap) {
      if (!aggList.length) {
        multiWrap.innerHTML =
          '<div class="empty-msg" style="padding:8px">No KPI data for this selection. Clear filters or choose All KPIs.</div>';
      } else {
        renderMultiKpiCards(multiWrap, aggList, f.refMonth, f.vsMode);
      }
    }

    buildCharts(catKey, f);
    renderTableBody(catKey);

    const cat = getCategory(catKey);
    const ctx = document.getElementById("cat-context");
    if (ctx && cat) {
      ctx.textContent = cat.uxNote || "";
    }

    announceFilterSummary(catKey);
  }

  function renderCategory(catKey) {
    currentCategoryKey = catKey;
    tableState = { sortKey: "yearMonth", asc: false, page: 0 };
    destroyCharts();

    const cat = getCategory(catKey);
    if (!cat) {
      renderCategories();
      return;
    }

    if (!ACTIVE_PREVIEW_CATEGORY_KEYS.has(catKey)) {
      history.replaceState(null, "", "#categories");
      renderCategories();
      announce(
        "This preview includes Incident Management and Hazard and Observation Management, Leading. Choose one on the category list."
      );
      return;
    }

    const kpisMeta = getKpis(catKey);
    const rowsForCat = getRowsForCategory(catKey);
    const cfg = getFilterConfig(catKey);

    const kpiOpts =
      '<option value="all">All KPIs</option>' +
      kpisMeta
        .map(
          (k) =>
            '<option value="' +
            k.kpiKey +
            '">' +
            escapeHtml(k.kpiName).replace(/"/g, "&quot;") +
            "</option>"
        )
        .join("");

    const vsOpts = VS_OPTIONS.map(
      (o) =>
        '<option value="' +
        o.id +
        '"' +
        (o.id === DEFAULT_VS_MODE ? " selected" : "") +
        ">" +
        escapeHtml(o.label) +
        "</option>"
    ).join("");

    const stateOpts =
      '<option value="all">All states</option>' +
      mergedIndiaStateList(DATA.states)
        .map(
          (s) =>
            '<option value="' +
            escapeHtml(s) +
            '">' +
            escapeHtml(s) +
            "</option>"
        )
        .join("");

    const bizList = distinctSorted(rowsForCat, (r) => r.businessName);
    const bizOpts =
      '<option value="all">All businesses</option>' +
      bizList
        .map(
          (b) =>
            '<option value="' +
            escapeHtml(b) +
            '">' +
            escapeHtml(b) +
            "</option>"
        )
        .join("");

    const wrap = document.createElement("div");
    wrap.className = "cat-view";
    /* Horizontal scroll only for native selects; Variable <details> panel is position:absolute and must not sit inside overflow-y:hidden (clips dropdown / blocks clicks). */
    const coreFieldsScroll =
      (cfg.showKpi
        ? '<div class="field"><label class="field-label" for="f-kpi">KPI</label>' +
          '<select id="f-kpi">' +
          kpiOpts +
          "</select></div>"
        : "") +
      '<div class="field"><label class="field-label" for="f-vs">Vs</label>' +
      '<select id="f-vs" aria-describedby="filter-hint">' +
      vsOpts +
      "</select></div>" +
      (cfg.showState
        ? '<div class="field"><label class="field-label" for="f-state">State</label>' +
          '<select id="f-state">' +
          stateOpts +
          "</select></div>"
        : "") +
      (cfg.showBusiness
        ? '<div class="field"><label class="field-label" for="f-biz">Business</label>' +
          '<select id="f-biz">' +
          bizOpts +
          "</select></div>"
        : "") +
      "";
    const variableFieldHtml = variableFilterFieldHtml();
    wrap.innerHTML =
      '<div class="cat-top-bar">' +
      '<div class="cat-top-bar__lead">' +
      journeyStepsHtml(3) +
      '<div class="cat-top-bar__title">' +
      '<nav class="breadcrumb" aria-label="Breadcrumb">' +
      '<ol><li><a href="#categories" id="bc-cats">Categories</a></li><li aria-current="page">' +
      '<h2 class="cat-heading" id="cat-heading" tabindex="-1">' +
      escapeHtml(cat.categoryName) +
      "</h2></li></ol></nav>" +
      "</div></div>" +
      '<fieldset class="cat-toolbar cat-toolbar--compact" aria-label="Refine results">' +
      '<legend class="visually-hidden">Refine results</legend>' +
      '<div class="cat-toolbar__inner" role="group" aria-describedby="filter-hint">' +
      '<p id="filter-hint" class="visually-hidden">User research, IA, usability testing, accessibility, consistency, and a user-centered approach: filters, KPI scope, and the detail table stay aligned so sessions are comparable—supporting UCD, HCI, and CX review of usability, desirability, and usefulness. Narrow by KPI, Versus, state, and business when shown. Data is monthly: Today vs yesterday and Current month vs last month both use latest month versus prior month. Current week versus last week uses the last two calendar months versus the two before that. Current quarter vs last quarter uses three-month windows. Current year vs last year uses calendar year-to-date versus the same months in the prior year. Charts and KPI tiles use the same windows. Choose All states and All businesses to see the full preview slice.</p>' +
      '<div class="cat-toolbar__filters-scroll">' +
      '<div class="cat-toolbar__filters-core">' +
      coreFieldsScroll +
      "</div>" +
      variableFieldHtml +
      "</div>" +
      '<div class="toolbar-actions">' +
      '<button type="button" class="btn" id="f-reset">Reset</button>' +
      "</div></div></fieldset>" +
      "</div>" +
      '<fieldset class="kpi-summary-region">' +
      '<legend class="visually-hidden">KPI summary for current filters</legend>' +
      '<div class="multi-kpi-row" id="multi-kpi-wrap"></div>' +
      "</fieldset>" +
      '<div class="cat-charts" role="group" aria-label="Charts for filtered data">' +
      '<div class="chart-box"><h3>Trend <span id="chart-trend-hint" class="chart-box__hint">(12 months)</span></h3><div class="chart-canvas-wrap"><canvas id="chart-line" role="img" aria-label="Line chart: monthly average for filtered KPIs in the trend window"></canvas></div></div>' +
      '<div class="chart-box chart-box--biz"><h3>By business <span id="chart-biz-hint" class="chart-box__hint">(latest month)</span></h3><div class="chart-canvas-wrap chart-canvas-wrap--biz"><canvas id="chart-biz" role="img" aria-label="Share of total by business"></canvas><p id="chart-biz-empty" class="chart-biz-empty" hidden></p></div></div>' +
      '<div class="chart-box"><h3>By state (location) <span id="chart-state-hint" class="chart-box__hint">(latest month)</span></h3><div class="chart-canvas-wrap"><canvas id="chart-area-state" role="img" aria-label="Values by state for the active Versus window"></canvas></div></div>' +
      "</div>" +
      '<div class="table-zone">' +
      '<div class="table-zone__head">' +
      '<div class="table-zone__title">' +
      '<span class="table-zone__label">Detail data</span>' +
      '<span id="tbl-summary" class="table-zone__summary" aria-live="polite"></span>' +
      "</div>" +
      '<div class="table-pager">' +
      '<button type="button" id="tbl-prev" aria-label="Previous page">Prev</button>' +
      '<span id="tbl-pageinfo"></span>' +
      '<button type="button" id="tbl-next" aria-label="Next page">Next</button>' +
      "</div></div>" +
      '<div class="table-scroll">' +
      '<table class="data-table" id="tbl-detail">' +
      '<caption class="visually-hidden">Detailed rows matching filters. Use column headers to sort.</caption>' +
      "<colgroup>" +
      '<col class="col-m" /><col class="col-s" /><col class="col-s" /><col class="col-kpi" />' +
      '<col class="col-s" /><col class="col-num" /><col class="col-num" /><col class="col-num" />' +
      "</colgroup>" +
      "<thead><tr>" +
      '<th scope="col" data-sort="yearMonth" class="col-m">Month</th>' +
      '<th scope="col" data-sort="state" class="col-s">State</th>' +
      '<th scope="col" data-sort="businessName" class="col-s">Business</th>' +
      '<th scope="col" data-sort="kpiName" class="col-kpi">KPI</th>' +
      '<th scope="col" data-sort="unitType" class="col-s">Unit</th>' +
      '<th scope="col" data-sort="value" class="col-num" title="Actual value for the row">Value</th>' +
      '<th scope="col" class="col-num">Vs %</th>' +
      '<th scope="col" data-sort="target" class="col-num">Target</th>' +
      "</tr></thead>" +
      '<tbody id="tbl-body"></tbody></table></div></div>' +
      '<p class="cat-context" id="cat-context">' +
      escapeHtml(cat.uxNote) +
      "</p>";

    root.innerHTML = "";
    root.appendChild(wrap);

    document.getElementById("f-vs").value = DEFAULT_VS_MODE;
    if (cfg.showState) document.getElementById("f-state").value = "all";
    if (cfg.showBusiness) document.getElementById("f-biz").value = "all";
    if (cfg.showKpi) document.getElementById("f-kpi").value = "all";
    applyVariableFilterFromStorage();

    function onFilterChange() {
      tableState.page = 0;
      refreshCategoryView(catKey);
    }

    ["f-kpi", "f-vs", "f-state", "f-biz"].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", onFilterChange);
    });
    wireVariableFilterControls(onFilterChange);

    document.getElementById("f-reset").addEventListener("click", () => {
      document.getElementById("f-vs").value = DEFAULT_VS_MODE;
      if (cfg.showKpi) document.getElementById("f-kpi").value = "all";
      if (cfg.showState) document.getElementById("f-state").value = "all";
      if (cfg.showBusiness) document.getElementById("f-biz").value = "all";
      try {
        localStorage.removeItem(LS_VARIABLE_FILTER);
      } catch {
        /* ignore */
      }
      const vAll = document.getElementById("f-var-all");
      const vCbs = document.querySelectorAll("input.f-var-cb");
      if (vAll) {
        vAll.checked = true;
        vCbs.forEach((cb) => {
          cb.checked = false;
        });
        updateVariableSummary();
      }
      tableState.page = 0;
      refreshCategoryView(catKey);
    });

    wireTableHeaders(catKey);
    refreshCategoryView(catKey);
    const h = document.getElementById("cat-heading");
    if (h) h.focus();
    updateHeaderNavState();
  }

  function renderLanding() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#landing");

    const box = document.createElement("div");
    box.className = "landing";
    box.innerHTML =
      journeyStepsHtml(1) +
      '<div class="landing__hero" role="region" aria-label="About this dashboard">' +
      '<div class="landing__copy">' +
      '<h2 class="landing__title" id="landing-h">Adani Safety MIS</h2>' +
      '<p class="landing__subtitle">Group-level safety KPI dashboard for all Adani businesses — monitor trends, drill by category, and compare performance across periods.</p>' +
      '<div class="landing__bullets" role="list">' +
      '<div class="landing__bullet" role="listitem"><strong>What’s included</strong><span>Safety KPIs by category, Versus comparison, filters, multi KPI cards with Vs % change, trend and business charts, by-state (location) area chart, and a sortable detail table.</span></div>' +
      '<div class="landing__bullet" role="listitem"><strong>Who it’s for</strong><span>Anyone tracking group-wide safety performance across Adani businesses — from quick scans to deeper drill-down.</span></div>' +
      "</div>" +
      '<div class="landing__companies" role="region" aria-label="Adani group companies represented">' +
      "<strong>Group companies</strong>" +
      "<span>Adani Enterprises Ltd (flagship incubator), Adani Ports &amp; SEZ Ltd, Adani Green Energy Ltd, Adani Energy Solutions Ltd, Adani Power Ltd, Adani Total Gas Ltd, Adani Wilmar Ltd, Ambuja Cements Ltd, and ACC Ltd.</span>" +
      "</div>" +
      '<div class="landing__actions">' +
      '<button type="button" class="btn btn-primary" id="btn-start">Start now</button>' +
      "</div>" +
      "</div>" +
      '<div class="landing__graphic landing__graphic--photo" aria-hidden="true">' +
      '<img class="landing__banner-img" src="./assets/Adani-Group.jpg" width="640" height="400" alt="" decoding="async" loading="eager" onerror="if(!this.dataset.fb){this.dataset.fb=\'1\';this.src=\'./assets/Adani-Group-home.png\';}else if(this.dataset.fb===\'1\'){this.dataset.fb=\'2\';this.src=\'./assets/adani-logo.png\';}else{this.onerror=null;}" />' +
      "</div>" +
      "</div>";

    root.innerHTML = "";
    root.appendChild(box);
    const start = document.getElementById("btn-start");
    function goCats() {
      history.replaceState(null, "", "#categories");
      renderCategories();
    }
    if (start) start.addEventListener("click", goCats);
    const hh = document.getElementById("landing-h");
    if (hh) hh.focus();
    announce("Adani Safety MIS home. Use Start now to open categories.");
    updateHeaderNavState();
  }

  function renderCategories() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#categories");

    const box = document.createElement("div");
    box.className = "home-body home-body--launchpad";
    box.innerHTML =
      journeyStepsHtml(2) +
      '<div class="home-intro home-intro--launchpad">' +
      '<h2 id="home-h">Categories</h2>' +
      '<p class="home-lede"><strong>Incident Management</strong> and <strong>Hazard &amp; Observation Management (Leading)</strong> are interactive in this preview; other categories are for context. Use <strong>Home</strong> in the header to return here.</p>' +
      '<div class="home-tools" role="search">' +
      '<label class="home-search-label" for="cat-q">Find category or KPI</label>' +
      '<input id="cat-q" class="home-search" type="search" placeholder="Category or KPI name (e.g. Incident, LTI…)" autocomplete="off" />' +
      "</div>" +
      "</div>" +
      '<div class="home-grid home-grid--launchpad" role="list" aria-labelledby="home-h"></div>';

    const grid = box.querySelector(".home-grid");
    function scheduleCategorySearchAnnounce(opts) {
      const o = opts || {};
      clearTimeout(catSearchAnnounceTimer);
      const run = () => {
        if (o.noMatch) {
          announce("No matches for category or KPI name.");
          return;
        }
        announce(
          o.count +
            " categor" +
            (o.count === 1 ? "y" : "ies") +
            " shown." +
            (o.filtered ? " Search filter applied." : "")
        );
      };
      if (o.immediate) run();
      else catSearchAnnounceTimer = setTimeout(run, 380);
    }

    function buildCards(query, opts) {
      const o = opts || {};
      grid.innerHTML = "";
      const q = (query || "").trim().toLowerCase();
      const cats = DATA.categories.filter((c) =>
        categoryMatchesSearchQuery(c, q)
      );
      if (!cats.length) {
        grid.removeAttribute("role");
        grid.removeAttribute("aria-labelledby");
        const empty = document.createElement("div");
        empty.className = "home-empty";
        empty.setAttribute("role", "status");
        empty.setAttribute("aria-live", "polite");
        empty.innerHTML =
          "<p><strong>No matches.</strong> Try a different category name or KPI keyword.</p>";
        grid.appendChild(empty);
        scheduleCategorySearchAnnounce({ noMatch: true, immediate: o.immediate });
        return;
      }
      grid.setAttribute("role", "list");
      grid.setAttribute("aria-labelledby", "home-h");
      cats.forEach((cat) => {
        const active = ACTIVE_PREVIEW_CATEGORY_KEYS.has(cat.categoryKey);
        const el = active
          ? document.createElement("button")
          : document.createElement("div");
        if (active) el.type = "button";
        el.className =
          "category-card category-card--launchpad" +
          (active ? "" : " category-card--disabled");
        el.setAttribute("role", "listitem");
        el.setAttribute(
          "aria-label",
          active
            ? cat.categoryName + ", " + cat.kpiCount + " KPIs"
            : cat.categoryName +
                ", " +
                cat.kpiCount +
                " KPIs. Not available in this preview; open Incident Management or Hazard and Observation Management, Leading."
        );
        if (!active) {
          el.setAttribute("aria-disabled", "true");
          el.setAttribute("tabindex", "-1");
        }
        const desc =
          cat.uxNote && String(cat.uxNote).trim()
            ? escapeHtml(cat.uxNote)
            : escapeHtml(String(cat.kpiCount)) +
              " KPIs · drill into trends, charts, and detail data.";
        el.innerHTML =
          '<div class="category-card__lp-top">' +
          '<span class="category-card__icon">' +
          categoryIconSvg(cat.categoryKey) +
          "</span>" +
          '<span class="category-card__lp-badge ' +
          (active ? "category-card__lp-badge--live" : "category-card__lp-badge--muted") +
          '">' +
          (active ? "Interactive" : "Preview") +
          "</span></div>" +
          '<div class="category-card__lp-body">' +
          '<span class="category-card__name">' +
          escapeHtml(cat.categoryName) +
          "</span>" +
          '<span class="category-card__meta">' +
          desc +
          "</span></div>" +
          '<div class="category-card__lp-footer">' +
          '<span class="category-card__lp-cta' +
          (active ? "" : " category-card__lp-cta--muted") +
          '">' +
          (active ? "Explore" : "Not in preview") +
          "</span></div>";
        if (active) {
          el.addEventListener("click", () => {
            history.replaceState({}, "", "#cat=" + cat.categoryKey);
            renderCategory(cat.categoryKey);
          });
          el.addEventListener("keydown", (e) => {
            if (e.key === "Enter" || e.key === " ") {
              e.preventDefault();
              history.replaceState({}, "", "#cat=" + cat.categoryKey);
              renderCategory(cat.categoryKey);
            }
          });
        }
        grid.appendChild(el);
      });
      scheduleCategorySearchAnnounce({
        count: cats.length,
        filtered: !!q,
        immediate: o.immediate,
      });
    }

    root.innerHTML = "";
    root.appendChild(box);
    buildCards("", { immediate: true });
    const q = document.getElementById("cat-q");
    if (q) {
      q.addEventListener("input", () => buildCards(q.value, {}));
      q.focus();
    }
    updateHeaderNavState();
  }

  window.addEventListener("hashchange", () => {
    const m = location.hash.match(/cat=(\d+)/);
    if (m) renderCategory(parseInt(m[1], 10));
    else if (location.hash === "#categories") renderCategories();
    else renderLanding();
  });

  setHeaderTimestamp(meta.lastUpdateISO);

  const m0 = location.hash.match(/cat=(\d+)/);
  if (m0) renderCategory(parseInt(m0[1], 10));
  else if (location.hash === "#categories") renderCategories();
  else renderLanding();
})();
