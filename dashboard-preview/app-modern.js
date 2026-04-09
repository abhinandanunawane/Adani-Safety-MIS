/**
 * Adani Safety MIS — Insights layout: same data as classic; copy foregrounds user research, IA, usability testing,
 * accessibility, consistency, user-centered approach, consistency & hierarchy, iterative process, UCD, HCI,
 * customer experience (CX) design, usability, desirability, usefulness. Maps to Power BI bookmarks + pages.
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
        "Ensure <code>embedded-data.js</code> is in the same folder as <code>insights.html</code> " +
        "and open the page via a local server (e.g. <code>npx serve</code>) if scripts are blocked. " +
        "Then run <code>Refresh-PreviewData.ps1</code> to regenerate data.</div>";
    }
    return;
  }

  if (!DATA.factRows) DATA.factRows = [];
  if (!DATA.months) DATA.months = [];
  if (!DATA.states) DATA.states = [];

  if (typeof Chart !== "undefined") {
    Chart.defaults.font.family =
      '"Adani", system-ui, -apple-system, "Segoe UI", sans-serif';
    Chart.defaults.font.size = 10;
    Chart.defaults.color = "#231f20";
  }

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

  /** Evidence table: rows per page (table layout fits more than cards) */
  const EVIDENCE_PAGE_SIZES = [20, 40, 60];
  const DEFAULT_EVIDENCE_PAGE_SIZE = 40;

  function loadEvidencePageSize() {
    try {
      const raw = localStorage.getItem("insights-evidence-page-size");
      const n = parseInt(raw, 10);
      if (EVIDENCE_PAGE_SIZES.includes(n)) return n;
    } catch {
      /* ignore */
    }
    return DEFAULT_EVIDENCE_PAGE_SIZE;
  }

  let tableState = {
    sortKey: "yearMonth",
    asc: false,
    page: 0,
    pageSize: DEFAULT_EVIDENCE_PAGE_SIZE,
  };

  let currentCategoryKey = null;
  let catSearchAnnounceTimer = null;

  /** Category keys: 1 Incident, 2 Hazard & Observation (Leading), 3 Safety Performance Indices. */
  const ACTIVE_PREVIEW_CATEGORY_KEYS = new Set([1, 2, 3]);
  const SPI_CATEGORY_KEY = 3;
  const EVENT_LEVEL_LABELS = [
    "0 Near Miss",
    "1 Minor",
    "2 Moderate",
    "3 Serious",
    "4 Major",
    "5 Catastrophic",
    "No Level",
  ];
  const SPI_KPI_ORDER = [21, 16, 17, 18, 20, 29];
  const STATE_CENTROID_BY_STATE = {
    "Andaman and Nicobar Islands": [11.7401, 92.6586],
    "Andhra Pradesh": [15.9129, 79.74],
    "Arunachal Pradesh": [28.218, 94.7278],
    Assam: [26.2006, 92.9376],
    Bihar: [25.0961, 85.3131],
    Chandigarh: [30.7333, 76.7794],
    Chhattisgarh: [21.2787, 81.8661],
    "Dadra and Nagar Haveli and Daman and Diu": [20.2734, 73.0169],
    Delhi: [28.7041, 77.1025],
    Goa: [15.2993, 74.124],
    Gujarat: [22.2587, 71.1924],
    Haryana: [29.0588, 76.0856],
    "Himachal Pradesh": [31.1048, 77.1734],
    "Jammu and Kashmir": [33.7782, 76.5762],
    Jharkhand: [23.6102, 85.2799],
    Karnataka: [15.3173, 75.7139],
    Kerala: [10.8505, 76.2711],
    Ladakh: [34.2268, 77.5619],
    Lakshadweep: [10.5667, 72.6417],
    "Madhya Pradesh": [22.9734, 78.6569],
    Maharashtra: [19.7515, 75.7139],
    Manipur: [24.6637, 93.9063],
    Meghalaya: [25.467, 91.3662],
    Mizoram: [23.1645, 92.9376],
    Nagaland: [26.1584, 94.5624],
    Odisha: [20.9517, 85.0985],
    Puducherry: [11.9416, 79.8083],
    Punjab: [31.1471, 75.3412],
    Rajasthan: [27.0238, 74.2179],
    Sikkim: [27.533, 88.5122],
    "Tamil Nadu": [11.1271, 78.6569],
    Telangana: [18.1124, 79.0193],
    Tripura: [23.9408, 91.9882],
    "Uttar Pradesh": [26.8467, 80.9462],
    Uttarakhand: [30.0668, 79.0193],
    "West Bengal": [22.9868, 87.855],
  };

  function hash32(str) {
    let h = 2166136261;
    const s = String(str);
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return h >>> 0;
  }

  function eventLevelIndexForRow(r) {
    const key =
      String(r.kpiKey) +
      "|" +
      (r.state || "") +
      "|" +
      (r.businessName || "") +
      "|" +
      (r.yearMonth || "");
    return hash32(key) % 7;
  }

  const LS_KPI_PREFIX = "insights-kpi-keys-";
  const LS_TABLE_SCOPE = "insights-table-kpi-scope";
  /** Chart view vs detail table — shared with index preview (`app.js`). */
  const LS_CAT_MAIN_VIEW = "adani_cat_main_view";

  function loadKpiSelection(catKey, kpisMeta) {
    try {
      const raw = localStorage.getItem(LS_KPI_PREFIX + catKey);
      if (raw) {
        const arr = JSON.parse(raw);
        if (Array.isArray(arr) && arr.length) {
          const valid = new Set(kpisMeta.map((k) => String(k.kpiKey)));
          const kept = arr.map(String).filter((id) => valid.has(id));
          if (kept.length) return kept;
        }
      }
    } catch {
      /* ignore */
    }
    return [defaultKpiKeyForCategory(catKey, kpisMeta)];
  }

  function saveKpiSelection(catKey, keys) {
    try {
      localStorage.setItem(LS_KPI_PREFIX + catKey, JSON.stringify(keys));
    } catch {
      /* ignore */
    }
  }

  function readSelectedKpiKeysFromDom() {
    const boxes = document.querySelectorAll(
      '#f-kpi-panel input[name="f-kpi-cb"]:checked'
    );
    return Array.from(boxes).map((cb) => String(cb.value));
  }

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

  /** IA: Discover → Orient → Insights (aligns with user mental model; same routes as classic). */
  function journeyStepsHtml(step) {
    function item(n, label, isActive) {
      const cls =
        " route-steps__item" +
        (isActive ? " route-steps__item--active" : "") +
        " route-steps__item--modern";
      const cur = isActive ? ' aria-current="step"' : "";
      return (
        '<span class="' +
        cls.trim() +
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
      '<nav class="route-steps route-steps--modern" aria-label="Journey: Discover, Orient, Insights">' +
      item(1, "Discover", step === 1) +
      '<span class="route-steps__sep" aria-hidden="true">→</span>' +
      item(2, "Orient", step === 2) +
      '<span class="route-steps__sep" aria-hidden="true">→</span>' +
      item(3, "Insights", step === 3) +
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

  /** Synthetic checkpoint when fact rows have no checkpoint column (preview data). */
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
      '<span class="field-label" id="f-var-lbl">Verticals</span>' +
      '<details class="var-scope var-scope--toolbar" id="f-var-details">' +
      '<summary class="var-scope__summary" aria-labelledby="f-var-lbl">' +
      '<span class="var-scope__summary-text">' +
      '<span class="var-scope__hint" id="f-var-hint">All checkpoints</span>' +
      "</span>" +
      '<span class="var-scope__chev" aria-hidden="true"></span>' +
      "</summary>" +
      '<div class="var-scope__panel" role="group" aria-label="Vertical options">' +
      '<div class="var-scope__menu">' +
      '<label class="field-variable-check field-variable-check--row field-variable-check--all">' +
      '<input type="checkbox" id="f-var-all" checked />' +
      '<span class="field-variable-check__text">All checkpoints</span>' +
      "</label>" +
      '<div class="var-scope__divider" aria-hidden="true"></div>' +
      '<div class="var-scope__options">' +
      '<label class="field-variable-check field-variable-check--row">' +
      '<input type="checkbox" class="f-var-cb" value="Field Force" />' +
      '<span class="field-variable-check__text">Field Force</span>' +
      "</label>" +
      '<label class="field-variable-check field-variable-check--row">' +
      '<input type="checkbox" class="f-var-cb" value="O and M" />' +
      '<span class="field-variable-check__text">O and M</span>' +
      "</label>" +
      '<label class="field-variable-check field-variable-check--row">' +
      '<input type="checkbox" class="f-var-cb" value="Office" />' +
      '<span class="field-variable-check__text">Office</span>' +
      "</label>" +
      '<label class="field-variable-check field-variable-check--row">' +
      '<input type="checkbox" class="f-var-cb" value="Projects" />' +
      '<span class="field-variable-check__text">Projects</span>' +
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
      hint.textContent = "All checkpoints";
      return;
    }
    const sel = [...cbs].filter((cb) => cb.checked).map((cb) => cb.value);
    if (sel.length === 0) {
      hint.textContent = "All checkpoints";
      return;
    }
    hint.textContent = sel.length + " Selected";
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

  /** Icons for landing “How to use” / category screen guide. */
  function guideIconSvg(kind) {
    const svg = (paths) =>
      '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      paths +
      "</svg>";
    switch (kind) {
      case "play":
        return svg(
          '<circle cx="12" cy="12" r="10"/><polygon points="10 8 16 12 10 16 10 8"/>'
        );
      case "grid":
        return svg(
          '<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/>'
        );
      case "filter":
        return svg(
          '<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>'
        );
      case "layers":
        return svg(
          '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>'
        );
      case "tiles":
        return svg(
          '<rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="11" width="7" height="10" rx="1"/><rect x="3" y="15" width="7" height="6" rx="1"/>'
        );
      case "chart":
        return svg(
          '<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>'
        );
      case "table":
        return svg(
          '<path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 1-2-2V9m0 0h18"/>'
        );
      default:
        return svg('<circle cx="12" cy="12" r="10"/>');
    }
  }

  function landingGuideRow(iconKind, strongText, restSentence) {
    return (
      '<li class="landing__guide-item">' +
      '<span class="landing__guide-icon">' +
      guideIconSvg(iconKind) +
      "</span>" +
      '<span class="landing__guide-item__body"><strong>' +
      escapeHtml(strongText) +
      "</strong> " +
      escapeHtml(restSentence) +
      "</span></li>"
    );
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
    const keys = f.kpiKeys || [];
    if (!keys.length) return null;
    const kpis = getKpis(catKey);
    if (keys.length === 1) {
      const k = kpis.find((x) => String(x.kpiKey) === String(keys[0]));
      return k ? k.unitType : null;
    }
    const types = keys
      .map((id) => kpis.find((x) => String(x.kpiKey) === String(id)))
      .filter(Boolean)
      .map((x) => x.unitType);
    const uniq = [...new Set(types)];
    return uniq.length === 1 ? uniq[0] : null;
  }

  function destroySpiLeafletMap() {
    if (window.__adaniSpiMap) {
      try {
        window.__adaniSpiMap.remove();
      } catch {
        /* ignore */
      }
      window.__adaniSpiMap = null;
    }
    const host = document.getElementById("chart-spi-map");
    if (host) host.innerHTML = "";
  }

  function destroyCharts() {
    ["chart-trend", "chart-verticals", "chart-biz", "chart-spi-bubble"].forEach(
      (id) => {
        const el = document.getElementById(id);
        if (el && typeof Chart !== "undefined") {
          const c = Chart.getChart(el);
          if (c) c.destroy();
        }
      }
    );
    destroySpiLeafletMap();
  }

  function resizeSpiMapIfAny() {
    if (window.__adaniSpiMap && typeof window.__adaniSpiMap.invalidateSize === "function") {
      try {
        window.__adaniSpiMap.invalidateSize();
      } catch {
        /* ignore */
      }
    }
  }

  function resizeAllChartsModern() {
    ["chart-trend", "chart-verticals", "chart-biz", "chart-spi-bubble"].forEach(
      (id) => {
        const el = document.getElementById(id);
        if (el && typeof Chart !== "undefined") {
          const c = Chart.getChart(el);
          if (c) c.resize();
        }
      }
    );
    resizeSpiMapIfAny();
  }

  function renderSpiEventBubbleChart(snapRows) {
    const el = document.getElementById("chart-spi-bubble");
    if (!el || typeof Chart === "undefined") return;
    const prev = Chart.getChart(el);
    if (prev) prev.destroy();
    const counts = [0, 0, 0, 0, 0, 0, 0];
    snapRows.forEach((r) => {
      counts[eventLevelIndexForRow(r)]++;
    });
    const bubbleData = counts.map((cnt, xi) => ({
      x: xi,
      y: cnt,
      r: Math.max(6, Math.sqrt(cnt + 1) * 7),
    }));
    new Chart(el, {
      type: "bubble",
      data: {
        datasets: [
          {
            label: "Rows",
            data: bubbleData,
            backgroundColor: "rgba(0, 109, 182, 0.45)",
            borderColor: "#006DB6",
            borderWidth: 1.5,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 10, right: 12, bottom: 8, left: 8 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(ctx) {
                const raw = ctx.raw;
                const xi = Math.round(Number(raw.x));
                const lbl = EVENT_LEVEL_LABELS[xi] || "";
                return " " + lbl + ": " + raw.y + " rows";
              },
            },
          },
        },
        scales: {
          x: {
            type: "linear",
            min: -0.5,
            max: 6.5,
            ticks: {
              stepSize: 1,
              font: { size: 8 },
              maxRotation: 55,
              minRotation: 35,
              autoSkip: false,
              callback(v) {
                const i = Math.round(Number(v));
                return i >= 0 && i < EVENT_LEVEL_LABELS.length
                  ? EVENT_LEVEL_LABELS[i]
                  : "";
              },
            },
            title: {
              display: true,
              text: "Event level",
              font: { size: 10 },
              color: "#6D6E71",
            },
            grid: { color: "rgba(109, 110, 113, 0.12)" },
          },
          y: {
            beginAtZero: true,
            ticks: { font: { size: 10 }, color: "#231f20" },
            title: {
              display: true,
              text: "Count (rows)",
              font: { size: 10 },
              color: "#6D6E71",
            },
            grid: { color: "rgba(109, 110, 113, 0.2)" },
          },
        },
      },
    });
  }

  function renderSpiPerformanceMap(snapRows) {
    const el = document.getElementById("chart-spi-map");
    if (!el) return;
    destroySpiLeafletMap();
    if (typeof L === "undefined") {
      el.innerHTML =
        '<p class="chart-spi-map-fallback">Map unavailable. Load Leaflet to see state circles.</p>';
      return;
    }
    const byState = {};
    snapRows.forEach((r) => {
      const st = String(r.state || "").trim() || "—";
      byState[st] = (byState[st] || 0) + 1;
    });
    const map = L.map(el, { scrollWheelZoom: false }).setView([22.6, 79.0], 4.5);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: "&copy; OpenStreetMap contributors",
    }).addTo(map);
    const latLngs = [];
    Object.keys(byState).forEach((st) => {
      const ll = STATE_CENTROID_BY_STATE[st];
      if (!ll) return;
      const n = byState[st];
      latLngs.push(ll);
      const rad = Math.max(6, Math.min(32, Math.sqrt(n) * 3.2));
      L.circleMarker(ll, {
        radius: rad,
        color: "#006DB6",
        fillColor: "#00B16B",
        fillOpacity: 0.52,
        weight: 2,
      })
        .bindPopup(escapeHtml(st) + ": " + n + " rows (filtered)")
        .addTo(map);
    });
    if (latLngs.length === 1) {
      map.setView(latLngs[0], 6);
    } else if (latLngs.length > 1) {
      map.fitBounds(L.latLngBounds(latLngs), { padding: [28, 28], maxZoom: 7 });
    }
    window.__adaniSpiMap = map;
    setTimeout(() => {
      try {
        map.invalidateSize();
      } catch {
        /* ignore */
      }
    }, 80);
  }

  function wireCatMainView() {
    const main = document.getElementById("cat-main-view");
    if (!main) return;
    const tabCharts = document.getElementById("view-tab-charts");
    const tabTable = document.getElementById("view-tab-table");
    const panelCharts = document.getElementById("view-panel-charts");
    const panelTable = document.getElementById("view-panel-table");
    function apply(mode) {
      const charts = mode === "charts";
      main.classList.toggle("m2-cat-main--charts", charts);
      main.classList.toggle("m2-cat-main--table", !charts);
      main.dataset.view = mode;
      if (tabCharts) {
        tabCharts.setAttribute("aria-selected", charts ? "true" : "false");
        tabCharts.classList.toggle("view-tabs__btn--active", charts);
        tabCharts.tabIndex = charts ? 0 : -1;
      }
      if (tabTable) {
        tabTable.setAttribute("aria-selected", charts ? "false" : "true");
        tabTable.classList.toggle("view-tabs__btn--active", !charts);
        tabTable.tabIndex = charts ? -1 : 0;
      }
      if (panelCharts) panelCharts.hidden = !charts;
      if (panelTable) panelTable.hidden = charts;
      try {
        localStorage.setItem(LS_CAT_MAIN_VIEW, mode);
      } catch {
        /* ignore */
      }
      if (charts) {
        requestAnimationFrame(() => {
          requestAnimationFrame(() => resizeAllChartsModern());
        });
      }
    }
    let initial = "charts";
    try {
      const s = localStorage.getItem(LS_CAT_MAIN_VIEW);
      if (s === "charts" || s === "table") initial = s;
    } catch {
      /* ignore */
    }
    apply(initial);
    tabCharts?.addEventListener("click", () => apply("charts"));
    tabTable?.addEventListener("click", () => apply("table"));
  }

  /* Approved palette only: #00B16B #006DB6 #8E278F #F04C23 (+ white/grey backgrounds) */
  const chartPaletteColors = [
    "rgba(0, 177, 107, 0.85)",
    "rgba(0, 109, 182, 0.85)",
    "rgba(142, 39, 143, 0.85)",
    "rgba(240, 76, 35, 0.85)",
    "rgba(0, 177, 107, 0.55)",
    "rgba(0, 109, 182, 0.55)",
    "rgba(142, 39, 143, 0.55)",
    "rgba(240, 76, 35, 0.55)",
  ];

  /**
   * By business: donut chart — share of total (|value|) per business; top slices + Other.
   * Matches classic preview: legend on the right, same window as filters.
   */
  function renderBizBreakdown(bizLabels, bizData, opts) {
    const el = document.getElementById("chart-biz");
    const emptyEl = document.getElementById("chart-biz-empty");
    if (!el || typeof Chart === "undefined") return;
    const emptyMsg =
      opts && opts.emptyMessage
        ? opts.emptyMessage
        : "No business data for current filters.";
    const existing = Chart.getChart(el);
    if (existing) existing.destroy();
    if (!bizLabels.length) {
      el.hidden = true;
      if (emptyEl) {
        emptyEl.hidden = false;
        emptyEl.textContent = emptyMsg;
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
      el.hidden = true;
      if (emptyEl) {
        emptyEl.hidden = false;
        emptyEl.textContent = emptyMsg;
      }
      return;
    }

    el.hidden = false;
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
              (_, i) => chartPaletteColors[i % chartPaletteColors.length]
            ),
            borderColor: "#ffffff",
            borderWidth: 1.5,
            hoverOffset: 4,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        radius: "70%",
        cutout: "55%",
        layout: {
          padding: { top: 6, right: 4, bottom: 6, left: 4 },
        },
        plugins: {
          legend: {
            display: true,
            position: "right",
            align: "center",
            labels: {
              font: {
                size: 9,
                family: '"Adani", system-ui, "Segoe UI", sans-serif',
              },
              color: "#231f20",
              boxWidth: 10,
              boxHeight: 10,
              padding: 6,
              maxWidth: 118,
              usePointStyle: true,
              generateLabels: function (chart) {
                const ds = chart.data.datasets[0];
                const labs = chart.data.labels || [];
                if (!labs.length || !ds) return [];
                return labs.map(function (label, i) {
                  const raw = Number(ds.data[i]) || 0;
                  const pct =
                    total > 0
                      ? ((raw / total) * 100).toFixed(1)
                      : "0.0";
                  const bg = Array.isArray(ds.backgroundColor)
                    ? ds.backgroundColor[i]
                    : ds.backgroundColor;
                  return {
                    text: String(label) + " (" + pct + "%)",
                    fillStyle: bg,
                    strokeStyle: "#ffffff",
                    lineWidth: 1,
                    hidden: !chart.getDataVisibility(i),
                    index: i,
                    datasetIndex: 0,
                  };
                });
              },
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
                    ? v.toLocaleString(undefined, {
                        maximumFractionDigits: 0,
                      })
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
    const elVs = document.getElementById("f-vs");
    const elSt = document.getElementById("f-state");
    if (!elSt) return null;
    const elBiz = document.getElementById("f-biz");
    const elTblScope = document.getElementById("tbl-kpi-scope");
    const kpiKeys = readSelectedKpiKeysFromDom();
    let tableKpiScope = "selected";
    if (elTblScope && elTblScope.value === "all") tableKpiScope = "all";
    return {
      catKey,
      kpiKeys: kpiKeys,
      vsMode: elVs ? elVs.value : DEFAULT_VS_MODE,
      refMonth: getRefMonth(),
      state: elSt.value,
      business: elBiz ? elBiz.value : "all",
      unitType: "all",
      variable: readVariableSelectionFromDom(),
      tableKpiScope: tableKpiScope,
    };
  }

  /** Table / KPI tiles: current reference month only (plus non-month filters). */
  function applyRowFilter(rows, f, opts) {
    const ignoreKpi = opts && opts.ignoreKpi;
    return rows.filter((r) => {
      if (
        !ignoreKpi &&
        f.kpiKeys &&
        f.kpiKeys.length &&
        !f.kpiKeys.includes(String(r.kpiKey))
      ) {
        return false;
      }
      if (r.yearMonth !== f.refMonth) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
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
      if (
        !f.kpiKeys ||
        !f.kpiKeys.length ||
        !f.kpiKeys.includes(String(r.kpiKey))
      ) {
        return false;
      }
      if (!range.has(r.yearMonth)) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
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

  function getFilterConfig(catKey) {
    // Core (neutral) filters for all category pages.
    // Category-specific extras can be enabled where they add insight.
    const cfg = {
      showKpi: true,
      showState: true,
      showBusiness: true,
    };
    return cfg;
  }

  /** Checkboxes + summary: which KPIs drive tiles, charts, and evidence cards. */
  function kpiScopePanelHtml(catKey) {
    const kpisMeta = kpiListForFilterDropdown(catKey);
    const sel = loadKpiSelection(catKey, kpisMeta);
    const boxes = kpisMeta
      .map((k) => {
        const id = "f-kpi-cb-" + k.kpiKey;
        const checked = sel.includes(String(k.kpiKey)) ? " checked" : "";
        return (
          '<label class="m2-kpi-cb" for="' +
          id +
          '">' +
          '<input type="checkbox" name="f-kpi-cb" id="' +
          id +
          '" value="' +
          k.kpiKey +
          '"' +
          checked +
          "/>" +
          '<span class="m2-kpi-cb__text">' +
          escapeHtml(kpiDropdownLabel(k)) +
          "</span></label>"
        );
      })
      .join("");
    return (
      '<details class="m2-kpi-scope m2-kpi-scope--toolbar">' +
      '<summary class="m2-kpi-scope__summary" id="f-kpi-scope-label">' +
      '<span class="m2-kpi-scope__summary-text">' +
      '<span class="m2-kpi-scope__title">KPIs</span> ' +
      '<span class="m2-kpi-scope__count" id="f-kpi-count"></span>' +
      "</span>" +
      '<span class="m2-kpi-scope__chev" aria-hidden="true"></span>' +
      "</summary>" +
      '<div class="m2-kpi-panel" id="f-kpi-panel" role="group" aria-labelledby="f-kpi-scope-label">' +
      '<div class="m2-kpi-panel__bar">' +
      '<button type="button" class="m2-btn m2-btn--tiny" id="f-kpi-none">None</button>' +
      "</div>" +
      '<div class="m2-kpi-panel__list">' +
      boxes +
      "</div></div></details>"
    );
  }

  function updateKpiScopeCount() {
    const total = document.querySelectorAll(
      '#f-kpi-panel input[name="f-kpi-cb"]'
    ).length;
    const n = document.querySelectorAll(
      '#f-kpi-panel input[name="f-kpi-cb"]:checked'
    ).length;
    const el = document.getElementById("f-kpi-count");
    if (el) el.textContent = n + "/" + total;
  }

  function avg(nums) {
    if (!nums.length) return 0;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  }

  function isTriKpiMeta(k) {
    return (
      String(k.kpiKey) === "21" ||
      /Total Recordable|TRIR|\bTRI\b|Recordable Incident Rate/i.test(
        k.kpiName || ""
      )
    );
  }

  /** Incident Management: TRI first, then wireframe order (ΔRepeat, ΔFatal, Man-days, Vehicle). */
  const INCIDENT_KPI_ORDER = [
    21, 1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 14, 15, 19, 22, 28, 44,
  ];

  function sortKpisForDisplay(catKey, kpisMeta) {
    const list = kpisMeta.slice();
    if (catKey === 1) {
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      return INCIDENT_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (catKey === SPI_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      const ordered = SPI_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
      const rest = list.filter(
        (k) => !SPI_KPI_ORDER.includes(Number(k.kpiKey))
      );
      return ordered.concat(rest);
    }
    const tri = list.find((k) => isTriKpiMeta(k));
    if (!tri) return list;
    const rest = list.filter((k) => k !== tri);
    return [tri, ...rest];
  }

  const TRI_LABEL_FULL = "Total Recordable Incident Rate (TRI)";

  function defaultKpiKeyForCategory(catKey, kpisMetaForUi) {
    const merged =
      kpisMetaForUi && kpisMetaForUi.length
        ? kpisMetaForUi
        : kpiListForFilterDropdown(catKey);
    const tri = merged.find((k) => isTriKpiMeta(k));
    if (tri) return String(tri.kpiKey);
    const sorted = sortKpisForDisplay(catKey, merged);
    return sorted.length ? String(sorted[0].kpiKey) : "21";
  }

  function kpiDropdownLabel(k) {
    const name = String(k.kpiName || "").trim();
    if (/^TRIR$/i.test(name) || String(k.kpiKey) === "21") return TRI_LABEL_FULL;
    return k.kpiName;
  }

  /** KPI scope / storage: always offer TRI when missing; TRI listed first. */
  function kpiListForFilterDropdown(catKey) {
    const m = getKpis(catKey);
    const hasTri = m.some(isTriKpiMeta);
    const base = hasTri
      ? m.slice()
      : [
          {
            kpiKey: 21,
            kpiName: "TRIR",
            unitType: "PercentOrRate",
            latestValue: null,
          },
        ].concat(m);
    return sortKpisForDisplay(catKey, base);
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

  /** Months rolled into By business / Vs-by-KPI charts for the active Vs mode. */
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
      if (
        !f.kpiKeys ||
        !f.kpiKeys.length ||
        !f.kpiKeys.includes(String(r.kpiKey))
      ) {
        return false;
      }
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      return true;
    });
  }

  /** Multi-KPI tiles: all KPIs; geography / business / vertical filters only. */
  function applyNonMonthFiltersAllKpis(rows, f) {
    return rows.filter((r) => {
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
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
    const sorted = sortKpisForDisplay(catKey, kpisMeta);
    const base = applyNonMonthFiltersAllKpis(getRowsForCategory(catKey), f);
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

  function cmpArrowClass(dir) {
    if (dir === "up") return "m2-kpi-tile__arrow m2-kpi-tile__arrow--up";
    if (dir === "down") return "m2-kpi-tile__arrow m2-kpi-tile__arrow--down";
    if (dir === "same") return "m2-kpi-tile__arrow m2-kpi-tile__arrow--same";
    return "m2-kpi-tile__arrow m2-kpi-tile__arrow--na";
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

  function chunkArray(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) {
      out.push(arr.slice(i, i + size));
    }
    return out;
  }

  function appendM2KpiPlaceholder(board) {
    const ph = document.createElement("div");
    ph.className = "m2-kpi-tile m2-kpi-tile--card";
    ph.style.visibility = "hidden";
    ph.innerHTML = "&nbsp;";
    ph.setAttribute("aria-hidden", "true");
    board.appendChild(ph);
  }

  function appendM2KpiTile(
    board,
    item,
    refMonth,
    vsMode,
    extraClass
  ) {
    const tile = document.createElement("div");
    const vDir = item.vsDir || "neutral";
    const dirMod =
      vDir === "up"
        ? "m2-kpi-tile--dir-up"
        : vDir === "down"
          ? "m2-kpi-tile--dir-down"
          : vDir === "same"
            ? "m2-kpi-tile--dir-flat"
            : "m2-kpi-tile--dir-na";
    tile.className =
      "m2-kpi-tile m2-kpi-tile--card " +
      dirMod +
      (extraClass ? " " + extraClass : "");
    tile.setAttribute("role", "listitem");
    const vsPct = formatSignedPct(item.vsPct);
    const vsLbl = vsOptionLabel(item.vsMode);
    const vsShort = vsTagShort(item.vsMode);
    const periodLine =
      item.periodCaption ||
      tilePeriodForKpi(refMonth, vsMode || DEFAULT_VS_MODE, item.unitType);
    tile.setAttribute(
      "aria-label",
      item.kpiName +
        ". " +
        periodLine +
        ". Value " +
        formatValue(item.value, item.unitType) +
        ". " +
        vsLbl +
        " " +
        vsPct +
        "."
    );
    tile.innerHTML =
      '<div class="m2-kpi-tile__ribbon" aria-hidden="true"></div>' +
      '<div class="m2-kpi-tile__chips">' +
      '<span class="m2-kpi-tile__chip m2-kpi-tile__chip--unit">' +
      escapeHtml(item.unitType) +
      "</span>" +
      '<span class="m2-kpi-tile__chip m2-kpi-tile__chip--vs" title="' +
      escapeAttr(vsLbl) +
      '">' +
      escapeHtml(vsShort) +
      "</span></div>" +
      '<h4 class="m2-kpi-tile__name">' +
      escapeHtml(item.kpiName) +
      "</h4>" +
      '<div class="m2-kpi-tile__hero">' +
      '<span class="m2-kpi-tile__value">' +
      formatValue(item.value, item.unitType) +
      "</span>" +
      '<span class="m2-kpi-tile__delta">' +
      '<span class="' +
      cmpArrowClass(vDir) +
      '" aria-hidden="true">' +
      arrowGlyph(vDir) +
      "</span>" +
      '<span class="m2-kpi-tile__pct vs-pct-num">' +
      escapeHtml(vsPct) +
      "</span></span></div>" +
      '<p class="m2-kpi-tile__period">' +
      escapeHtml(periodLine) +
      "</p>";
    board.appendChild(tile);
  }

  function renderMultiKpiCards(
    container,
    aggregatesList,
    refMonth,
    vsMode
  ) {
    container.innerHTML = "";
    const sorted = sortKpiTilesForDisplay(aggregatesList);
    if (!sorted.length) return;

    const row = document.createElement("div");
    row.className = "m2-kpi-metrics-row";
    row.setAttribute("role", "group");
    row.setAttribute(
      "aria-label",
      "KPI metrics for this category"
    );

    function appendMetricCardBlock(startMetric, items, cardMod) {
      if (!items.length) return;
      const card = document.createElement("section");
      card.className =
        "m2-kpi-metric-card" +
        (cardMod ? " " + cardMod : "") +
        (startMetric >= 13 ? " m2-kpi-metric-card--tail" : "");
      const head = document.createElement("div");
      head.className = "m2-kpi-metric-card__head";
      const endMetric = startMetric + items.length - 1;
      head.textContent =
        items.length === 1
          ? "Metric " + startMetric
          : "Metrics " + startMetric + "–" + endMetric;
      const grid = document.createElement("div");
      grid.className = "m2-kpi-board m2-kpi-board--quad";
      grid.setAttribute("role", "list");
      const tailTile = startMetric >= 13 ? "m2-kpi-tile--tail" : "";
      items.forEach((item) => {
        appendM2KpiTile(grid, item, refMonth, vsMode, tailTile);
      });
      while (grid.children.length < 4) {
        appendM2KpiPlaceholder(grid);
      }
      card.appendChild(head);
      card.appendChild(grid);
      row.appendChild(card);
    }

    const headItems = sorted.slice(0, 12);
    const tailItems = sorted.slice(12);
    const headChunks = chunkArray(headItems, 4);
    headChunks.forEach((chunk, idx) => {
      appendMetricCardBlock(idx * 4 + 1, chunk, "");
    });

    const tailA = tailItems.slice(0, 4);
    const tailB = tailItems.slice(4);
    if (tailA.length) {
      appendMetricCardBlock(13, tailA, "");
    }
    if (tailB.length) {
      appendMetricCardBlock(13 + tailA.length, tailB, "");
    }

    container.appendChild(row);
  }

  /**
   * Charts: line trend (+ for SPI, event-level bubble + India map); else donut + by vertical.
   */
  function buildCharts(catKey, f) {
    destroyCharts();
    const kpisMeta = getKpis(catKey);
    const kpiKeys = f.kpiKeys || [];
    if (typeof Chart === "undefined") {
      updateChartHints(f);
      return;
    }
    if (!kpiKeys.length) {
      renderBizBreakdown([], [], {
        emptyMessage:
          "Select one or more KPIs in KPI scope to load charts and business share.",
      });
      updateChartHints(f);
      return;
    }

    const poolAll = getRowsForCategory(catKey);
    const filteredTrend = applyChartFilter(poolAll, f);
    const snapPool = applyNonMonthFilters(poolAll, f);
    const winMonths = new Set(
      bizUnitWindowMonths(f.vsMode || DEFAULT_VS_MODE, f.refMonth)
    );
    const snapRows = snapPool.filter((r) => winMonths.has(r.yearMonth));

    const monthSet = new Set();
    filteredTrend.forEach((r) => monthSet.add(r.yearMonth));
    const monthKeys = Array.from(monthSet).sort();
    const lineDatasets = kpiKeys.map((kk, idx) => {
      const meta = kpisMeta.find((x) => String(x.kpiKey) === String(kk));
      const shortName = meta
        ? meta.kpiName.length > 20
          ? meta.kpiName.slice(0, 18) + "…"
          : meta.kpiName
        : String(kk);
      const ut = meta ? meta.unitType : "";
      const data = monthKeys.map((ym) => {
        const slice = filteredTrend.filter(
          (r) => r.yearMonth === ym && String(r.kpiKey) === String(kk)
        );
        if (!slice.length) return null;
        const vals = slice.map((r) => Number(r.value));
        if (isAdditiveUnit(ut)) return vals.reduce((a, b) => a + b, 0);
        return avg(vals);
      });
      const base = chartPaletteColors[idx % chartPaletteColors.length];
      const lineCol = base.replace("0.45", "0.9");
      return {
        label: shortName,
        data: data,
        borderColor: lineCol,
        backgroundColor: base.replace("0.45", "0.12"),
        borderWidth: 1.5,
        tension: 0.25,
        fill: false,
        pointRadius: 2,
        pointHoverRadius: 3,
        spanGaps: true,
      };
    });

    const utChart = kpiUnitTypeForFilter(catKey, f);

    const elTrend = document.getElementById("chart-trend");
    if (elTrend && monthKeys.length) {
      new Chart(elTrend, {
        type: "line",
        data: {
          labels: monthKeys,
          datasets: lineDatasets,
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          layout: {
            padding: { top: 12, right: 16, bottom: 14, left: 22 },
          },
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: {
              display: lineDatasets.length > 1,
              position: "bottom",
              labels: { font: { size: 9 }, boxWidth: 10 },
            },
            tooltip: {
              callbacks: {
                title: function (items) {
                  return items.length ? String(items[0].label) : "";
                },
              },
            },
          },
          scales: {
            x: {
              ticks: { font: { size: 10 }, maxRotation: 40, padding: 4 },
              grid: { color: "rgba(15,23,42,0.06)" },
            },
            y: {
              beginAtZero: false,
              grace: "12%",
              ticks: { font: { size: 10 }, padding: 6 },
              grid: { color: "rgba(15,23,42,0.06)" },
            },
          },
        },
      });
    }

    if (catKey === SPI_CATEGORY_KEY) {
      renderSpiEventBubbleChart(snapRows);
      renderSpiPerformanceMap(snapRows);
      updateChartHints(f);
      return;
    }

    const byBizVals = {};
    snapRows.forEach((r) => {
      const b = r.businessName || "—";
      if (!byBizVals[b]) byBizVals[b] = [];
      byBizVals[b].push({
        v: Number(r.value),
        ut: r.unitType || "",
      });
    });
    const bizLabels = Object.keys(byBizVals);
    const bizData = bizLabels.map((b) => {
      const items = byBizVals[b];
      if (utChart && isAdditiveUnit(utChart)) {
        return items.reduce(
          (a, x) => a + (Number.isFinite(x.v) ? x.v : 0),
          0
        );
      }
      const nums = items.map((x) => x.v).filter((v) => Number.isFinite(v));
      return nums.length ? avg(nums) : 0;
    });

    renderBizBreakdown(bizLabels, bizData);

    const vertMap = {};
    snapRows.forEach((r) => {
      const vx = getRowCheckpoint(r);
      if (!vertMap[vx]) vertMap[vx] = [];
      vertMap[vx].push(Number(r.value));
    });
    const vertLabelsFull = Object.keys(vertMap).sort((a, b) =>
      a.localeCompare(b)
    );
    function shortVertLabelModern(name) {
      const s = String(name);
      return s.length > 14 ? s.slice(0, 13) + "…" : s;
    }
    const vertLabels = vertLabelsFull.map(shortVertLabelModern);
    const vertData = vertLabelsFull.map((label) => {
      const vals = vertMap[label];
      if (utChart && isAdditiveUnit(utChart)) {
        return vals.reduce(
          (a, x) => a + (Number.isFinite(x) ? x : 0),
          0
        );
      }
      const nums = vals.filter((x) => Number.isFinite(x));
      return nums.length ? avg(nums) : 0;
    });

    const elVert = document.getElementById("chart-verticals");
    if (elVert && vertLabels.length) {
      new Chart(elVert, {
        type: "line",
        data: {
          labels: vertLabels,
          datasets: [
            {
              label: "By vertical",
              data: vertData,
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
          layout: {
            padding: { top: 12, right: 20, bottom: 16, left: 28 },
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                title(items) {
                  const i = items[0].dataIndex;
                  return vertLabelsFull[i] || "";
                },
              },
            },
          },
          scales: {
            y: {
              beginAtZero: true,
              grace: "6%",
              ticks: {
                font: { size: 10 },
                color: "#231F20",
                padding: 6,
              },
              grid: { color: "rgba(109, 110, 113, 0.2)" },
              title: {
                display: true,
                text: "Value",
                font: { size: 10, family: '"Adani", system-ui, sans-serif' },
                color: "#6D6E71",
                padding: { top: 0, bottom: 8, left: 0, right: 0 },
              },
            },
            x: {
              ticks: {
                font: { size: 9 },
                maxRotation: 40,
                minRotation: 0,
                autoSkip: true,
                maxTicksLimit: 10,
                color: "#231F20",
                padding: 4,
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
    const trendTitleEl = document.getElementById("chart-trend-title");
    const bizEl = document.getElementById("chart-biz-hint");
    const vertHintEl = document.getElementById("chart-vertical-hint");
    const n = chartMonthsForVsMode(mode, ref).length;
    if (trendTitleEl && f && f.catKey != null) {
      const keys = f.kpiKeys || [];
      const list = kpiListForFilterDropdown(f.catKey);
      if (keys.length === 0) {
        trendTitleEl.textContent = TRI_LABEL_FULL;
      } else if (keys.length === 1) {
        const k = list.find((x) => String(x.kpiKey) === String(keys[0]));
        trendTitleEl.textContent = k ? kpiDropdownLabel(k) : TRI_LABEL_FULL;
      } else {
        trendTitleEl.textContent = "Time series";
      }
    }
    if (trendEl) {
      trendEl.textContent =
        mode === "vs_last_year"
          ? "(lines · 12 mo)"
          : "(lines · " + n + " mo)";
    }
    const bizHint = {
      vs_yesterday: "(share · latest mo)",
      vs_last_month: "(share · latest mo)",
      vs_last_week: "(share · 2 mo)",
      vs_last_quarter: "(share · 3 mo)",
      vs_last_year: "(share · YTD window)",
    };
    const bh = bizHint[mode] || "(share)";
    if (bizEl) bizEl.textContent = bh;
    const vertHint = {
      vs_yesterday: "(vertical · latest mo)",
      vs_last_month: "(vertical · latest mo)",
      vs_last_week: "(vertical · 2 mo)",
      vs_last_quarter: "(vertical · 3 mo)",
      vs_last_year: "(vertical · YTD window)",
    };
    if (vertHintEl) vertHintEl.textContent = vertHint[mode] || "(vertical)";
    const spiBubbleHint = document.getElementById("chart-spi-bubble-hint");
    const spiMapHint = document.getElementById("chart-spi-map-hint");
    const spiRoll =
      mode === "vs_last_year"
        ? "(bubble · YTD window)"
        : "(bubble · same window)";
    const spiMapRoll =
      mode === "vs_last_year"
        ? "(map · YTD window)"
        : "(map · same window)";
    if (spiBubbleHint) spiBubbleHint.textContent = spiRoll;
    if (spiMapHint) spiMapHint.textContent = spiMapRoll;
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

  function formatDetailValue(r, pool, f) {
    const unit = r.unitType || "";
    const cmp = rowCompareForMode(r, pool, f);
    return formatValue(cmp.cur, unit);
  }

  function distinctCount(rows, key) {
    const s = new Set();
    for (let i = 0; i < rows.length; i++) {
      const v = rows[i][key];
      if (v != null && v !== "") s.add(String(v));
    }
    return s.size;
  }

  function renderDetailCards(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    const poolAll = getRowsForCategory(catKey);
    const pool =
      f.tableKpiScope === "all"
        ? poolAll.filter((r) => {
            if (f.state !== "all" && r.state !== f.state) return false;
            if (f.business !== "all" && r.businessName !== f.business)
              return false;
            if (!rowMatchesVariable(r, f.variable)) return false;
            return true;
          })
        : applyNonMonthFilters(poolAll, f);
    let rows = applyRowFilter(pool, f, {
      ignoreKpi: f.tableKpiScope === "all",
    });
    rows = sortRows(rows, tableState.sortKey, tableState.asc, pool, f);

    const pageSize = tableState.pageSize || DEFAULT_EVIDENCE_PAGE_SIZE;
    const total = rows.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    if (tableState.page >= pages) tableState.page = pages - 1;
    if (tableState.page < 0) tableState.page = 0;

    const start = tableState.page * pageSize;
    const pageRows = rows.slice(start, start + pageSize);

    const host = document.getElementById("detail-cards");
    const pageInfo = document.getElementById("tbl-pageinfo");
    const btnP = document.getElementById("tbl-prev");
    const btnN = document.getElementById("tbl-next");

    if (!host) return;

    const summaryEl = document.getElementById("tbl-summary");
    const statTotal = document.getElementById("ev-st-total");
    const statKpi = document.getElementById("ev-st-kpi");
    const statSt = document.getElementById("ev-st-state");
    const statBiz = document.getElementById("ev-st-biz");
    if (summaryEl) {
      const summaryFull =
        total === 0
          ? "Adjust filters or KPI scope to see rows."
          : "Detail table for scan-and-compare (IA: same filters as tiles & charts). Per screen and page list support usability testing and accessibility-friendly paging.";
      summaryEl.textContent = summaryFull;
      summaryEl.title = summaryFull;
    }
    if (statTotal) statTotal.textContent = total === 0 ? "—" : String(total);
    if (statKpi)
      statKpi.textContent = total === 0 ? "—" : String(distinctCount(rows, "kpiName"));
    if (statSt)
      statSt.textContent = total === 0 ? "—" : String(distinctCount(rows, "state"));
    if (statBiz)
      statBiz.textContent =
        total === 0 ? "—" : String(distinctCount(rows, "businessName"));

    if (!pageRows.length) {
      host.innerHTML =
        '<p class="m2-detail-empty" role="status">No rows for these filters.</p>';
    } else {
      const head =
        '<table class="m2-evidence-table">' +
        "<thead><tr>" +
        '<th scope="col">Month</th>' +
        '<th scope="col" class="m2-evidence-table__col-kpi">KPI</th>' +
        '<th scope="col">State</th>' +
        '<th scope="col">Business</th>' +
        '<th scope="col">Unit</th>' +
        '<th scope="col" class="m2-evidence-table__num">Value</th>' +
        '<th scope="col" class="m2-evidence-table__num">Target</th>' +
        '<th scope="col" class="m2-evidence-table__num">Vs prior</th>' +
        "</tr></thead><tbody>";
      const body = pageRows
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
          const valStr = formatDetailValue(r, pool, f);
          const tgtStr =
            r.target == null || r.target === ""
              ? "—"
              : formatValue(r.target, r.unitType);
          const kpiFull = escapeHtml(r.kpiName);
          return (
            "<tr>" +
            "<td>" +
            escapeHtml(r.yearMonth) +
            "</td>" +
            '<td class="m2-evidence-table__kpi" title="' +
            escapeAttr(r.kpiName) +
            '"><span>' +
            kpiFull +
            "</span></td>" +
            "<td>" +
            escapeHtml(r.state) +
            "</td>" +
            "<td>" +
            escapeHtml(r.businessName) +
            "</td>" +
            "<td>" +
            escapeHtml(r.unitType) +
            "</td>" +
            '<td class="m2-evidence-table__num">' +
            escapeHtml(valStr) +
            "</td>" +
            '<td class="m2-evidence-table__num">' +
            escapeHtml(tgtStr) +
            "</td>" +
            '<td class="m2-evidence-table__num m2-evidence-table__vs">' +
            '<span class="m2-kpi-tile__arrow ' +
            (rv.dir === "up"
              ? "m2-kpi-tile__arrow--up"
              : rv.dir === "down"
                ? "m2-kpi-tile__arrow--down"
                : "m2-kpi-tile__arrow--na") +
            '">' +
            arrow +
            "</span> " +
            escapeHtml(pctStr) +
            "</td></tr>"
          );
        })
        .join("");
      host.innerHTML = head + body + "</tbody></table>";
    }

    if (pageInfo) {
      pageInfo.textContent =
        total === 0
          ? ""
          : "Rows " +
            (start + 1) +
            "–" +
            Math.min(start + pageSize, total) +
            " of " +
            total;
    }
    const pageSel = document.getElementById("tbl-page-select");
    if (pageSel) {
      pageSel.innerHTML = "";
      for (let p = 1; p <= pages; p++) {
        const opt = document.createElement("option");
        opt.value = String(p);
        opt.textContent = "Page " + p + " / " + pages;
        if (p === tableState.page + 1) opt.selected = true;
        pageSel.appendChild(opt);
      }
      pageSel.disabled = pages <= 1;
    }
    const pageSizeEl = document.getElementById("tbl-page-size");
    if (btnP) {
      btnP.disabled = tableState.page <= 0;
    }
    if (btnN) {
      btnN.disabled = tableState.page >= pages - 1;
    }
    if (pageSizeEl) {
      pageSizeEl.value = String(pageSize);
    }

    const sortField = document.getElementById("detail-sort-field");
    const sortDir = document.getElementById("detail-sort-dir");
    if (sortField) sortField.value = tableState.sortKey;
    if (sortDir) sortDir.value = tableState.asc ? "asc" : "desc";
  }

  function wireDetailControls(catKey) {
    const sortField = document.getElementById("detail-sort-field");
    const sortDir = document.getElementById("detail-sort-dir");
    function applySort() {
      if (sortField) tableState.sortKey = sortField.value;
      if (sortDir) tableState.asc = sortDir.value === "asc";
      tableState.page = 0;
      renderDetailCards(catKey);
    }
    if (sortField) sortField.addEventListener("change", applySort);
    if (sortDir) sortDir.addEventListener("change", applySort);

    const pageSizeEl = document.getElementById("tbl-page-size");
    if (pageSizeEl) {
      pageSizeEl.addEventListener("change", () => {
        const n = parseInt(pageSizeEl.value, 10);
        if (EVIDENCE_PAGE_SIZES.includes(n)) {
          tableState.pageSize = n;
          try {
            localStorage.setItem("insights-evidence-page-size", String(n));
          } catch {
            /* ignore */
          }
          tableState.page = 0;
          renderDetailCards(catKey);
        }
      });
    }

    const btnP = document.getElementById("tbl-prev");
    const btnN = document.getElementById("tbl-next");
    const pageSel = document.getElementById("tbl-page-select");
    if (btnP) {
      btnP.addEventListener("click", () => {
        tableState.page--;
        renderDetailCards(catKey);
      });
    }
    if (btnN) {
      btnN.addEventListener("click", () => {
        tableState.page++;
        renderDetailCards(catKey);
      });
    }
    if (pageSel) {
      pageSel.addEventListener("change", () => {
        const p = parseInt(pageSel.value, 10);
        if (Number.isFinite(p) && p >= 1) {
          tableState.page = p - 1;
          renderDetailCards(catKey);
        }
      });
    }
  }

  function announceFilterSummary(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    const poolAll = getRowsForCategory(catKey);
    let pool;
    if (f.tableKpiScope === "all") {
      pool = poolAll.filter((r) => {
        if (f.state !== "all" && r.state !== f.state) return false;
        if (f.business !== "all" && r.businessName !== f.business)
          return false;
        if (!rowMatchesVariable(r, f.variable)) return false;
        return true;
      });
    } else {
      pool = applyNonMonthFilters(poolAll, f);
    }
    const n = applyRowFilter(pool, f, {
      ignoreKpi: f.tableKpiScope === "all",
    }).length;
    const cat = getCategory(catKey);
    const label = cat ? cat.categoryName + ". " : "";
    announce(
      label +
        n +
        " rows in the detail table for " +
        formatMonthYear(f.refMonth) +
        ". Sort the list, then pick a page from the dropdown or use arrows."
    );
  }

  function refreshCategoryView(catKey) {
    const kpisMeta = getKpis(catKey);
    const f = readFilters(catKey);
    if (!f) return;

    const multiWrap = document.getElementById("multi-kpi-wrap");
    if (!f.kpiKeys || !f.kpiKeys.length) {
      if (multiWrap) {
        multiWrap.innerHTML =
          '<div class="m2-empty">Open <strong>KPI scope</strong> and tick one or more KPIs. Tiles, charts, and evidence cards use the same selection.</div>';
      }
      destroyCharts();
      renderDetailCards(catKey);
      const ctx = document.getElementById("cat-context");
      const cat = getCategory(catKey);
      if (ctx && cat) ctx.textContent = cat.uxNote || "";
      announceFilterSummary(catKey);
      return;
    }

    const aggList = buildKpiDetailMetrics(catKey, kpisMeta, f);

    if (multiWrap) {
      if (!aggList.length) {
        multiWrap.innerHTML =
          '<div class="m2-empty">No KPI data for this slice. Adjust Versus, geography, or KPI scope.</div>';
      } else {
        renderMultiKpiCards(multiWrap, aggList, f.refMonth, f.vsMode);
      }
    }

    buildCharts(catKey, f);
    renderDetailCards(catKey);
    requestAnimationFrame(() => {
      requestAnimationFrame(() => resizeAllChartsModern());
    });

    const cat = getCategory(catKey);
    const ctx = document.getElementById("cat-context");
    if (ctx && cat) {
      ctx.textContent = cat.uxNote || "";
    }

    announceFilterSummary(catKey);
  }

  function renderCategory(catKey) {
    currentCategoryKey = catKey;
    tableState = {
      sortKey: "yearMonth",
      asc: false,
      page: 0,
      pageSize: loadEvidencePageSize(),
    };
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
        "This preview includes Incident Management, Hazard and Observation Management (Leading), and Safety Performance Indices. Choose one on the category list."
      );
      return;
    }

    const kpisMeta = getKpis(catKey);
    const kpisForUi = kpiListForFilterDropdown(catKey);
    const rowsForCat = getRowsForCategory(catKey);
    const cfg = getFilterConfig(catKey);

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
    wrap.className = "cat-view cat-view--modern";
    const kpiFieldHtml =
      cfg.showKpi
        ? '<div class="field field--kpi-inline">' +
          kpiScopePanelHtml(catKey) +
          "</div>"
        : "";
    const filterCore =
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
    const m2LineChartBox =
      '<div class="chart-box m2-chart-card m2-chart-card--trend"><h3 class="m2-chart-title"><span class="m2-chart-title__label" id="chart-trend-title">' +
      escapeHtml(TRI_LABEL_FULL) +
      '</span> <span id="chart-trend-hint" class="chart-box__hint">(lines)</span></h3><div class="chart-canvas-wrap m2-chart-canvas"><canvas id="chart-trend" role="img" aria-label="Line chart: one series per KPI over months"></canvas></div></div>';
    const m2ChartsRowHtml =
      catKey === SPI_CATEGORY_KEY
        ? '<div class="m2-charts m2-charts--alt cat-charts cat-charts--spi" role="group" aria-label="Charts: time series, event level, state map">' +
          m2LineChartBox +
          '<div class="chart-box m2-chart-card m2-chart-card--spi-bubble"><h3 class="m2-chart-title"><span class="m2-chart-title__label">Event level</span> <span id="chart-spi-bubble-hint" class="chart-box__hint">(bubble · count)</span></h3><div class="chart-canvas-wrap m2-chart-canvas chart-canvas-wrap--spi-bubble"><canvas id="chart-spi-bubble" role="img" aria-label="Bubble chart: row counts by event severity level"></canvas></div></div>' +
          '<div class="chart-box m2-chart-card m2-chart-card--spi-map"><h3 class="m2-chart-title"><span class="m2-chart-title__label">Safety performance by state</span> <span id="chart-spi-map-hint" class="chart-box__hint">(map · row counts)</span></h3><div class="chart-spi-map-host m2-chart-spi-map" id="chart-spi-map" role="presentation" aria-label="Map of India: circle size by filtered row count per state"></div></div></div>'
        : '<div class="m2-charts m2-charts--alt cat-charts" role="group" aria-label="Three chart views: time series, business share, by vertical">' +
          m2LineChartBox +
          '<div class="chart-box m2-chart-card m2-chart-card--biz"><h3 class="m2-chart-title"><span class="m2-chart-title__label">By business</span> <span id="chart-biz-hint" class="chart-box__hint">(share · latest mo)</span></h3><div class="chart-canvas-wrap m2-chart-canvas m2-chart-canvas--wide m2-chart-wrap--biz-list"><canvas id="chart-biz" role="img" aria-label="Donut chart: share of total by business, legend shows each segment"></canvas><p id="chart-biz-empty" class="chart-biz-empty m2-chart-biz-empty" hidden></p></div></div>' +
          '<div class="chart-box m2-chart-card m2-chart-card--verticals"><h3 class="m2-chart-title"><span class="m2-chart-title__label">By vertical</span> <span id="chart-vertical-hint" class="chart-box__hint">(vertical · latest mo)</span></h3><div class="chart-canvas-wrap m2-chart-canvas"><canvas id="chart-verticals" role="img" aria-label="Values by vertical for KPI scope and Versus window"></canvas></div></div></div>';
    const toolbarInner =
      '<div class="cat-toolbar__inner" role="group" aria-describedby="filter-hint">' +
      '<p id="filter-hint" class="visually-hidden">User research, information architecture (IA), usability testing, accessibility, consistency, visual hierarchy, and an iterative user-centered process: filters, KPI scope, charts, and the detail table stay aligned for comparable sessions—supporting user-centered design (UCD), human-computer interaction (HCI), and customer experience (CX) review of usability, desirability, accessibility, and usefulness. Refine Versus, geography, business, and checkpoints. One slice across KPI tiles, charts, and detail table. Data is monthly: Today vs yesterday and Current month vs last month both use latest month versus prior month. Current week versus last week uses the last two calendar months versus the two before that. Current quarter vs last quarter uses three-month windows. Current year vs last year uses calendar year-to-date versus the same months in the prior year. Charts and KPI tiles use the same windows.</p>' +
      '<div class="cat-toolbar__filters-scroll">' +
      kpiFieldHtml +
      '<div class="cat-toolbar__filters-core">' +
      filterCore +
      "</div>" +
      variableFieldHtml +
      "</div>" +
      '<div class="toolbar-actions">' +
      '<button type="button" class="btn m2-btn m2-btn--ghost m2-btn--compact" id="f-reset">Reset</button>' +
      "</div></div>";
    wrap.innerHTML =
      '<div class="m2-cat-top">' +
      '<div class="m2-cat-top__lead">' +
      journeyStepsHtml(3) +
      '<div class="m2-cat-head__title">' +
      '<nav class="breadcrumb m2-breadcrumb" aria-label="Breadcrumb">' +
      '<ol><li><a href="#categories" id="bc-cats">Categories</a></li><li aria-current="page">' +
      '<h2 class="cat-heading m2-cat-heading" id="cat-heading" tabindex="-1">' +
      escapeHtml(cat.categoryName) +
      "</h2></li></ol></nav>" +
      "</div></div>" +
      '<div class="m2-toolbar-shell m2-toolbar-shell--cat-top">' +
      '<fieldset class="cat-toolbar cat-toolbar--compact" aria-label="Refine results">' +
      '<legend class="visually-hidden">Refine results</legend>' +
      toolbarInner +
      "</fieldset></div></div>" +
      '<section class="m2-kpi-block" aria-labelledby="m2-kpi-h">' +
      '<div class="m2-kpi-block__head">' +
      '<h3 id="m2-kpi-h" class="m2-section-title">KPI metrics</h3>' +
      "</div>" +
      '<div id="multi-kpi-wrap"></div>' +
      "</section>" +
      '<div class="m2-cat-main m2-cat-main--charts" id="cat-main-view" data-view="charts">' +
      '<div class="view-tabs" role="tablist" aria-label="Chart view or table view">' +
      '<button type="button" role="tab" id="view-tab-charts" class="view-tabs__btn view-tabs__btn--active" aria-selected="true" aria-controls="view-panel-charts" data-view="charts">Chart view</button>' +
      '<button type="button" role="tab" id="view-tab-table" class="view-tabs__btn" aria-selected="false" aria-controls="view-panel-table" tabindex="-1" data-view="table">Table view</button>' +
      "</div>" +
      '<div id="view-panel-charts" class="view-panel view-panel--charts" role="tabpanel" aria-labelledby="view-tab-charts">' +
      m2ChartsRowHtml +
      "</div>" +
      '<div id="view-panel-table" class="view-panel view-panel--table" role="tabpanel" aria-labelledby="view-tab-table" hidden>' +
      '<div class="table-zone m2-table-zone m2-evidence-zone">' +
      '<div class="table-zone__head m2-evidence-head">' +
      '<div class="m2-evidence-head__row">' +
      '<span class="table-zone__label m2-evidence-head__label">Detail data</span>' +
      '<span id="tbl-summary" class="m2-evidence-summary" aria-live="polite"></span>' +
      '<div class="m2-evidence-stats" id="evidence-stats" role="group" aria-label="Filtered list summary">' +
      '<div class="m2-evidence-stat"><span id="ev-st-total" class="m2-evidence-stat__val">—</span><span class="m2-evidence-stat__lbl">rows</span></div>' +
      '<div class="m2-evidence-stat"><span id="ev-st-kpi" class="m2-evidence-stat__val">—</span><span class="m2-evidence-stat__lbl">KPIs</span></div>' +
      '<div class="m2-evidence-stat"><span id="ev-st-state" class="m2-evidence-stat__val">—</span><span class="m2-evidence-stat__lbl">states</span></div>' +
      '<div class="m2-evidence-stat"><span id="ev-st-biz" class="m2-evidence-stat__val">—</span><span class="m2-evidence-stat__lbl">businesses</span></div>' +
      "</div></div>" +
      '<div class="m2-evidence-tools" role="toolbar" aria-label="Table controls">' +
      '<div class="m2-evidence-sort">' +
      '<label class="m2-table-scope" for="detail-sort-field">Sort by</label>' +
      '<select id="detail-sort-field">' +
      '<option value="yearMonth">Month</option>' +
      '<option value="state">State</option>' +
      '<option value="businessName">Business</option>' +
      '<option value="kpiName">KPI</option>' +
      '<option value="unitType">Unit</option>' +
      '<option value="value">Value</option>' +
      '<option value="target">Target</option>' +
      "</select>" +
      '<label class="m2-table-scope" for="detail-sort-dir">Order</label>' +
      '<select id="detail-sort-dir" title="Sort order">' +
      '<option value="desc">Descending</option>' +
      '<option value="asc">Ascending</option>' +
      "</select></div>" +
      '<div class="table-zone__modes">' +
      '<label class="m2-table-scope" for="tbl-kpi-scope">Rows</label>' +
      '<select id="tbl-kpi-scope" title="Which KPIs appear">' +
      '<option value="selected">Match KPI scope</option>' +
      '<option value="all">All KPIs (same filters)</option>' +
      "</select>" +
      "</div>" +
      '<div class="m2-evidence-page-size">' +
      '<label class="m2-table-scope" for="tbl-page-size">Per screen</label>' +
      '<select id="tbl-page-size" title="Rows per page">' +
      '<option value="20">20</option>' +
      '<option value="40">40</option>' +
      '<option value="60">60</option>' +
      "</select></div>" +
      '<div class="m2-evidence-pager">' +
      '<button type="button" id="tbl-prev" class="m2-pager-btn m2-pager-btn--arrow" aria-label="Previous page">‹</button>' +
      '<label class="visually-hidden" for="tbl-page-select">Jump to page</label>' +
      '<select id="tbl-page-select" class="m2-evidence-page-jump" title="Jump to page" aria-label="Jump to page"></select>' +
      '<span id="tbl-pageinfo" class="m2-evidence-range" aria-live="polite"></span>' +
      '<button type="button" id="tbl-next" class="m2-pager-btn m2-pager-btn--arrow" aria-label="Next page">›</button>' +
      "</div></div></div>" +
      '<div class="m2-detail-scroll m2-evidence-scroll" role="region" aria-label="Detail data table">' +
      '<div id="detail-cards" class="m2-evidence-table-host"></div></div></div></div></div>' +
      '<p class="cat-context m2-cat-context" id="cat-context">' +
      escapeHtml(cat.uxNote) +
      "</p>";

    root.innerHTML = "";
    root.appendChild(wrap);

    document.getElementById("f-vs").value = DEFAULT_VS_MODE;
    if (cfg.showState) document.getElementById("f-state").value = "all";
    if (cfg.showBusiness) document.getElementById("f-biz").value = "all";
    applyVariableFilterFromStorage();
    try {
      const ts = localStorage.getItem(LS_TABLE_SCOPE);
      if (ts === "all" || ts === "selected") {
        const sel = document.getElementById("tbl-kpi-scope");
        if (sel) sel.value = ts;
      }
    } catch {
      /* ignore */
    }

    function onFilterChange() {
      tableState.page = 0;
      refreshCategoryView(catKey);
    }

    ["f-vs", "f-state", "f-biz"].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", onFilterChange);
    });
    wireVariableFilterControls(onFilterChange);

    const tblScope = document.getElementById("tbl-kpi-scope");
    if (tblScope) {
      tblScope.addEventListener("change", () => {
        try {
          localStorage.setItem(LS_TABLE_SCOPE, tblScope.value);
        } catch {
          /* ignore */
        }
        tableState.page = 0;
        renderDetailCards(catKey);
        announceFilterSummary(catKey);
      });
    }

    if (cfg.showKpi) {
      updateKpiScopeCount();
      document.querySelectorAll('#f-kpi-panel input[name="f-kpi-cb"]').forEach((cb) => {
        cb.addEventListener("change", () => {
          saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
          updateKpiScopeCount();
          onFilterChange();
        });
      });
      const btnNone = document.getElementById("f-kpi-none");
      if (btnNone) {
        btnNone.addEventListener("click", () => {
          document
            .querySelectorAll('#f-kpi-panel input[name="f-kpi-cb"]')
            .forEach((cb) => {
              cb.checked = false;
            });
          saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
          updateKpiScopeCount();
          onFilterChange();
        });
      }
    }

    document.getElementById("f-reset").addEventListener("click", () => {
      document.getElementById("f-vs").value = DEFAULT_VS_MODE;
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
      if (cfg.showKpi) {
        try {
          localStorage.removeItem(LS_KPI_PREFIX + catKey);
        } catch {
          /* ignore */
        }
        const def = loadKpiSelection(catKey, kpisForUi);
        document.querySelectorAll('#f-kpi-panel input[name="f-kpi-cb"]').forEach((cb) => {
          cb.checked = def.includes(String(cb.value));
        });
        saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
        updateKpiScopeCount();
      }
      if (tblScope) tblScope.value = "selected";
      try {
        localStorage.setItem(LS_TABLE_SCOPE, "selected");
      } catch {
        /* ignore */
      }
      tableState.page = 0;
      refreshCategoryView(catKey);
    });

    wireDetailControls(catKey);
    wireCatMainView();
    refreshCategoryView(catKey);
    const h = document.getElementById("cat-heading");
    if (h) h.focus();
    updateHeaderNavState();
  }

  function renderLanding() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#landing");

    const lastMo = formatMonthYear(getRefMonth());
    const box = document.createElement("div");
    box.className = "landing landing--modern landing--v3 landing--v3--clean";
    box.innerHTML =
      '<div class="m2-v3 m2-v3--clean">' +
      '<aside class="m2-v3__rail" aria-hidden="true">' +
      journeyStepsHtml(1) +
      "</aside>" +
      '<div class="m2-v3__main m2-v3__main--clean">' +
      '<div class="m2-v3__hero-card">' +
      '<header class="m2-v3__mast m2-v3__mast--clean">' +
      '<p class="m2-v3__eyebrow" id="landing-h" tabindex="-1">Insights</p>' +
      '<h2 class="m2-v3__headline m2-v3__headline--clean">Explore safety KPIs by domain</h2>' +
      '<p class="m2-v3__sub m2-v3__sub--clean">Choose a category, set filters and KPI scope, then review tiles, charts, and the detail table. Built for user research, IA, usability testing, accessibility, consistency, hierarchy, iterative UX, UCD, HCI, and CX review of usability, desirability, and usefulness. Fixed layout <span class="m2-v3__dim">1280×720</span>. Sample data — confirm in Power BI.</p>' +
      '<div class="m2-v3__stats m2-v3__stats--inline" role="group" aria-label="Snapshot">' +
      '<div class="m2-v3__stat"><span class="m2-v3__stat-num">9</span><span class="m2-v3__stat-lbl">domains</span></div>' +
      '<div class="m2-v3__stat"><span class="m2-v3__stat-num">17</span><span class="m2-v3__stat-lbl">KPIs (Incident, full scope)</span></div>' +
      '<div class="m2-v3__stat"><span class="m2-v3__stat-num">' +
      escapeHtml(lastMo) +
      '</span><span class="m2-v3__stat-lbl">reference month</span></div>' +
      "</div>" +
      '<button type="button" class="m2-btn m2-btn--primary m2-v3__cta" id="btn-start">Browse domains</button>' +
      "</header>" +
      '<ul class="m2-v3__quick" aria-label="What you can do">' +
      "<li><strong>Filter</strong> — Versus window, state, business, checkpoints.</li>" +
      "<li><strong>Compare</strong> — Time series, by business, by vertical.</li>" +
      "<li><strong>Drill</strong> — Sortable detail table with paging.</li>" +
      "</ul>" +
      "</div>" +
      '<div class="m2-v3__ribbon m2-v3__ribbon--thin" role="presentation" aria-hidden="true"></div>' +
      "</div></div>";

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
    announce(
      "Insights home. User research, IA, usability testing, accessibility, and user-centered design: browse domains to open a category. Incident Management, Hazard and Observation Management (Leading), and Safety Performance Indices are live in this preview."
    );
    updateHeaderNavState();
  }

  function renderCategories() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#categories");

    const box = document.createElement("div");
    box.className =
      "home-body home-body--modern m2-cat-page m2-cat-page--directory";
    box.innerHTML =
      '<div class="m2-cat-dir-head">' +
      journeyStepsHtml(2) +
      '<div class="m2-cat-dir-intro">' +
      '<h2 id="home-h">Safety domains</h2>' +
      '<p class="m2-cat-dir-lede">Search by <strong>category or KPI name</strong>. <strong>Live preview</strong>: <strong>Incident Management</strong> and <strong>Hazard &amp; Observation Management (Leading)</strong>. Consistent hierarchy for user research, usability testing, accessibility review, and iterative UCD or CX walkthroughs.</p>' +
      '<div class="home-tools m2-cat-dir-search" role="search">' +
      '<label class="home-search-label" for="cat-q">Find</label>' +
      '<input id="cat-q" class="home-search home-search--modern" type="search" placeholder="Category or KPI name…" autocomplete="off" />' +
      "</div></div></div>" +
      '<div class="m2-cat-directory-wrap">' +
      '<p class="m2-cat-directory__label" id="cat-dir-lbl">All domains</p>' +
      '<div class="m2-cat-directory" role="list" aria-labelledby="cat-dir-lbl"></div>' +
      "</div>";

    const grid = box.querySelector(".m2-cat-directory");
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
        empty.className = "m2-cat-directory__empty";
        empty.setAttribute("role", "status");
        empty.setAttribute("aria-live", "polite");
        empty.innerHTML =
          "<p><strong>No matches.</strong> Try a different category name or KPI keyword.</p>";
        grid.appendChild(empty);
        scheduleCategorySearchAnnounce({ noMatch: true, immediate: o.immediate });
        return;
      }
      grid.setAttribute("role", "list");
      grid.setAttribute("aria-labelledby", "cat-dir-lbl");
      cats.forEach((cat) => {
        const active = ACTIVE_PREVIEW_CATEGORY_KEYS.has(cat.categoryKey);
        const el = active
          ? document.createElement("button")
          : document.createElement("div");
        if (active) el.type = "button";
        const order = String(cat.sortOrder || cat.categoryKey).padStart(2, "0");
        el.className =
          "m2-cat-row m2-cat-row--k" +
          cat.categoryKey +
          (active
            ? " m2-cat-row--open"
            : " m2-cat-row--locked category-card--disabled");
        el.setAttribute("role", "listitem");
        el.setAttribute(
          "aria-label",
          active
            ? cat.categoryName + ", " + cat.kpiCount + " KPIs, open dashboard"
            : cat.categoryName +
                ", " +
                cat.kpiCount +
                " KPIs. Not in live preview; open Incident Management or Hazard and Observation Management, Leading."
        );
        if (!active) {
          el.setAttribute("aria-disabled", "true");
          el.setAttribute("tabindex", "-1");
        }
        const note = (cat.uxNote || "").trim();
        el.innerHTML =
          '<span class="m2-cat-row__rail" aria-hidden="true"></span>' +
          '<span class="m2-cat-row__index">' +
          order +
          "</span>" +
          '<span class="m2-cat-row__icon" aria-hidden="true">' +
          categoryIconSvg(cat.categoryKey) +
          "</span>" +
          '<span class="m2-cat-row__main">' +
          (active
            ? '<span class="m2-cat-row__chip">Live preview</span>'
            : "") +
          '<span class="m2-cat-row__title">' +
          escapeHtml(cat.categoryName) +
          "</span>" +
          (note
            ? '<span class="m2-cat-row__note">' + escapeHtml(note) + "</span>"
            : "") +
          "</span>" +
          '<span class="m2-cat-row__aside">' +
          '<span class="m2-cat-row__pill">' +
          cat.kpiCount +
          " KPIs</span>" +
          '<span class="m2-cat-row__cta">' +
          (active ? "Open" : "—") +
          "</span>" +
          "</span>";
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
