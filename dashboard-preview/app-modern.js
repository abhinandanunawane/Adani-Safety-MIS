/**
 * Adani Safety Performance Dashboard — Insights layout: same data as classic; copy foregrounds user research, IA, usability testing,
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
    const navBack = document.getElementById("nav-back");
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
    if (navBack) {
      if (onLanding) {
        navBack.hidden = true;
        navBack.setAttribute("aria-hidden", "true");
      } else {
        navBack.hidden = false;
        navBack.removeAttribute("aria-hidden");
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
        '<h1 class="boot-error__title" style="margin:0 0 10px;font-size:1.05rem;font-weight:700;color:#8e278f">Adani Safety Performance Dashboard</h1>' +
        "<p style=\"margin:0 0 10px;line-height:1.45\"><strong>Data not loaded.</strong> The site header above should still be visible.</p>" +
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
      '"Segoe UI", "Segoe UI Variable", "Segoe UI Historic", system-ui, sans-serif';
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

  /** Category keys with full KPI drill-down: 1 Incident Management, 2 Hazard & Observation Management. */
  const ACTIVE_PREVIEW_CATEGORY_KEYS = new Set([1, 2]);

  const LS_KPI_PREFIX = "insights-kpi-keys-";
  const LS_TABLE_SCOPE = "insights-table-kpi-scope";

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
    if (catKey === 1 && kpisMeta.length) {
      return ["1", "2", "3", "4", "5"];
    }
    return kpisMeta
      .slice(0, Math.min(5, kpisMeta.length))
      .map((k) => String(k.kpiKey));
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

  function getRefMonth() {
    if (DATA.months && DATA.months.length) {
      return DATA.months[DATA.months.length - 1].yearMonth;
    }
    return meta.lastDataMonth || "2024-01";
  }

  /** Short tag on KPI tiles (Δ% for period comparison). */
  function vsTagShort(id) {
    switch (id) {
      case "vs_period":
        return "Δ%";
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
        return "Δ%";
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

  function formatMonthRangeShort(months) {
    if (!months || !months.length) return "—";
    const sorted = months.slice().filter(Boolean).sort();
    if (!sorted.length) return "—";
    const a = sorted[0];
    const b = sorted[sorted.length - 1];
    if (a === b) return formatMonthYear(a);
    return formatMonthYear(a) + " – " + formatMonthYear(b);
  }

  function dataMonthBounds() {
    const list = (DATA.months || [])
      .map((x) => x.yearMonth)
      .filter(Boolean)
      .sort();
    if (list.length) return { min: list[0], max: list[list.length - 1] };
    const m = meta.lastDataMonth || "2024-01";
    return { min: m, max: m };
  }

  function monthsInInclusiveRange(fromYm, toYm) {
    if (
      !fromYm ||
      !toYm ||
      String(fromYm).length < 7 ||
      String(toYm).length < 7
    )
      return [];
    let a = String(fromYm).slice(0, 7);
    let b = String(toYm).slice(0, 7);
    if (a > b) {
      const t = a;
      a = b;
      b = t;
    }
    const out = [];
    let m = a;
    let guard = 0;
    while (m <= b && guard++ < 240) {
      out.push(m);
      const next = monthAdd(m, 1);
      if (!next || next === m) break;
      m = next;
    }
    return out;
  }

  function defaultToolbarPeriodRanges() {
    const mb = dataMonthBounds();
    const curTo = mb.max;
    const curFrom = monthAdd(curTo, -11);
    const cmpTo = monthAdd(curTo, -12);
    const cmpFrom = monthAdd(curFrom, -12);
    return { curFrom, curTo, cmpFrom, cmpTo };
  }

  function currentPeriodMonthList(f) {
    if (!f) return [];
    const cf = f.currentFrom != null ? f.currentFrom : f.monthFrom;
    const ct = f.currentTo != null ? f.currentTo : f.monthTo || f.refMonth;
    if (cf && ct) return monthsInInclusiveRange(cf, ct);
    const ref = (f.monthTo || f.refMonth || "").slice(0, 7) || getRefMonth();
    return ref ? [ref] : [];
  }

  function comparisonPeriodMonthList(f) {
    if (!f) return [];
    const a = f.comparisonFrom;
    const b = f.comparisonTo;
    if (a && b) return monthsInInclusiveRange(a, b);
    return [];
  }

  function initPeriodRangeInputs() {
    const mb = dataMonthBounds();
    const ids = ["f-cur-from", "f-cur-to", "f-cmp-from", "f-cmp-to"];
    const els = ids
      .map((id) => document.getElementById(id))
      .filter(Boolean);
    if (els.length !== 4) return;

    function clampYm(ym) {
      if (!ym || ym.length < 7) return mb.max;
      const y = ym.slice(0, 7);
      if (y < mb.min) return mb.min;
      if (y > mb.max) return mb.max;
      return y;
    }

    ids.forEach((id) => {
      const el = document.getElementById(id);
      if (el) {
        el.min = mb.min;
        el.max = mb.max;
      }
    });

    const d = defaultToolbarPeriodRanges();
    const elCurFrom = document.getElementById("f-cur-from");
    const elCurTo = document.getElementById("f-cur-to");
    const elCmpFrom = document.getElementById("f-cmp-from");
    const elCmpTo = document.getElementById("f-cmp-to");

    let curTo =
      elCurTo && elCurTo.value && elCurTo.value.length >= 7
        ? clampYm(elCurTo.value.slice(0, 7))
        : d.curTo;
    let curFrom =
      elCurFrom && elCurFrom.value && elCurFrom.value.length >= 7
        ? clampYm(elCurFrom.value.slice(0, 7))
        : d.curFrom;
    let cmpTo =
      elCmpTo && elCmpTo.value && elCmpTo.value.length >= 7
        ? clampYm(elCmpTo.value.slice(0, 7))
        : d.cmpTo;
    let cmpFrom =
      elCmpFrom && elCmpFrom.value && elCmpFrom.value.length >= 7
        ? clampYm(elCmpFrom.value.slice(0, 7))
        : d.cmpFrom;

    if (curFrom > curTo) curFrom = curTo;
    if (cmpFrom > cmpTo) cmpFrom = cmpTo;

    if (elCurTo) elCurTo.value = curTo;
    if (elCurFrom) elCurFrom.value = curFrom;
    if (elCmpTo) elCmpTo.value = cmpTo;
    if (elCmpFrom) elCmpFrom.value = cmpFrom;
  }

  function isAdditiveUnit(ut) {
    return ut === "Count" || ut === "Hours" || ut === "Days";
  }

  function tilePeriodForKpi(_refMonth, _mode, _unitType) {
    void _refMonth;
    void _mode;
    void _unitType;
    return "Current vs comparison period";
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

  function businessFilterMatchesRow(r, f) {
    if (!f || f.business === "all") return true;
    const rb = String(r.businessName || "").trim();
    if (Array.isArray(f.business)) {
      if (!f.business.length) return false;
      return f.business.some((b) => rb === String(b || "").trim());
    }
    return rb === String(f.business || "").trim();
  }

  function readBusinessSelectionFromDom() {
    const cbs = document.querySelectorAll("input.f-biz-cb");
    if (!cbs.length) return "all";
    const checked = Array.from(cbs).filter((cb) => cb.checked);
    if (!checked.length) return [];
    if (checked.length === cbs.length) return "all";
    return checked.map((cb) => cb.value);
  }

  function updateBizFilterSummary() {
    const cbs = document.querySelectorAll("input.f-biz-cb");
    const el = document.getElementById("f-biz-hint");
    if (!el) return;
    const total = cbs.length;
    if (!total) {
      el.textContent = "";
      return;
    }
    const n = document.querySelectorAll("input.f-biz-cb:checked").length;
    el.textContent = n + "/" + total;
  }

  function initBusinessSiteCheckboxFilters() {
    document.querySelectorAll("input.f-biz-cb").forEach((cb) => {
      cb.checked = true;
    });
    updateBizFilterSummary();
  }

  function wireBusinessAndSiteFilterControls(onChange) {
    const panel = document.getElementById("f-biz-panel");
    const allBtn = document.getElementById("f-biz-btn-all");
    const noneBtn = document.getElementById("f-biz-btn-none");
    if (!panel) return;
    function listCbs() {
      return panel.querySelectorAll("input.f-biz-cb");
    }
    function refresh() {
      updateBizFilterSummary();
      if (onChange) onChange();
    }
    function closeIfAllSelected() {
      const cbs = listCbs();
      const tot = cbs.length;
      const n = panel.querySelectorAll("input.f-biz-cb:checked").length;
      if (tot && n === tot) {
        const det = document.getElementById("f-biz-details");
        if (det) det.open = false;
      }
    }
    if (allBtn) {
      allBtn.addEventListener("click", () => {
        listCbs().forEach((cb) => {
          cb.checked = true;
        });
        refresh();
        const det = document.getElementById("f-biz-details");
        if (det) det.open = false;
      });
    }
    if (noneBtn) {
      noneBtn.addEventListener("click", () => {
        listCbs().forEach((cb) => {
          cb.checked = false;
        });
        refresh();
      });
    }
    listCbs().forEach((cb) => {
      cb.addEventListener("change", () => {
        updateBizFilterSummary();
        closeIfAllSelected();
        if (onChange) onChange();
      });
    });
    updateBizFilterSummary();
  }

  function variableFilterFieldHtml() {
    return (
      '<div class="field field--variable field--var-inline">' +
      '<span class="field-label" id="f-var-lbl">Vertical</span>' +
      '<details class="var-scope var-scope--toolbar" id="f-var-details">' +
      '<summary class="var-scope__summary" aria-labelledby="f-var-lbl" title="Vertical">' +
      '<span class="var-scope__summary-text">' +
      '<span class="var-scope__hint" id="f-var-hint">All verticals</span>' +
      "</span>" +
      '<span class="var-scope__chev" aria-hidden="true"></span>' +
      "</summary>" +
      '<div class="var-scope__panel" role="group" aria-label="Vertical filter options">' +
      '<div class="var-scope__menu">' +
      '<label class="field-variable-check field-variable-check--row field-variable-check--all">' +
      '<input type="checkbox" id="f-var-all" checked />' +
      '<span class="field-variable-check__text">All verticals</span>' +
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
      hint.textContent = "All verticals";
      return;
    }
    const sel = [...cbs].filter((cb) => cb.checked).map((cb) => cb.value);
    if (sel.length === 0) {
      hint.textContent = "All verticals";
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
        const det = document.getElementById("f-var-details");
        if (det) det.open = false;
      }
      saveVariableFilterToStorage();
      updateVariableSummary();
      if (onChange) onChange();
    }
    all.addEventListener("change", allChange);
    cbs.forEach((cb) => cb.addEventListener("change", subChange));
    updateVariableSummary();
  }

  /** Single filter scroll strip: fixed-position <details> panels while open (avoids overflow clip). */
  function wireToolbarScopeScrollPanels(host) {
    const scope = host.querySelector(".cat-toolbar__filters-all-scroll");
    if (!scope) return;
    const pairs = [];
    [
      ["details.var-scope--toolbar", ".var-scope__panel"],
      ["details.m2-kpi-scope--toolbar", ".m2-kpi-panel"],
    ].forEach(([dSel, pSel]) => {
      scope.querySelectorAll(dSel).forEach((det) => {
        const panel = det.querySelector(pSel);
        if (panel) pairs.push([det, panel]);
      });
    });
    if (!pairs.length) return;

    function placePair(det, panel) {
      if (!det.open) {
        panel.style.position = "";
        panel.style.top = "";
        panel.style.left = "";
        panel.style.right = "";
        panel.style.width = "";
        panel.style.zIndex = "";
        return;
      }
      const summary = det.querySelector("summary");
      if (!summary) return;
      const r = summary.getBoundingClientRect();
      const estW = panel.offsetWidth || 280;
      let left = r.left;
      const maxL = window.innerWidth - estW - 8;
      if (left > maxL) left = Math.max(8, maxL);
      if (left < 8) left = 8;
      panel.style.position = "fixed";
      panel.style.left = left + "px";
      panel.style.top = r.bottom + 6 + "px";
      panel.style.right = "auto";
      panel.style.zIndex = "200";
    }

    function syncAll() {
      pairs.forEach(([d, p]) => placePair(d, p));
    }

    pairs.forEach(([det]) => {
      det.addEventListener("toggle", () => {
        requestAnimationFrame(() => {
          requestAnimationFrame(syncAll);
        });
      });
    });
    scope.addEventListener("scroll", syncAll, { passive: true });
    window.addEventListener("resize", syncAll);
    window.addEventListener("scroll", syncAll, true);
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
          '<line x1="12" y1="21" x2="12" y2="7"/><line x1="4" y1="7" x2="20" y2="7"/><path d="M7 7L4.5 14h5L7 7"/><path d="M17 7l2.5 7h-5L17 7"/>'
        );
      case 5:
        return svg(
          '<path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>'
        );
      case 6:
        return svg(
          '<circle cx="12" cy="8" r="6"/><path d="M15.477 12.89 17 22l-5-3-5 3 1.523-9.11"/>'
        );
      case 7:
        return svg(
          '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>'
        );
      case 8:
        /* Leadership & Safety Governance — large front leader, two teammates, up-chevron (clear at 22px) */
        return svg(
          '<path d="M3.8 22V16.2Q6.5 13.8 9.2 16.2V22"/>' +
            '<circle cx="6.5" cy="10.5" r="1.45"/>' +
            '<path d="M14.8 22V16.2Q17.5 13.8 20.2 16.2V22"/>' +
            '<circle cx="17.5" cy="10.5" r="1.45"/>' +
            '<path d="M6.2 22V14.8Q12 11.2 17.8 14.8V22"/>' +
            '<circle cx="12" cy="7.3" r="2.35"/>' +
            '<path d="M8.5 3.4L12 1L15.5 3.4"/>'
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

  function kpiUnitTypeLabel(unitType) {
    const u = String(unitType || "").trim();
    if (u === "Count") return "Count";
    if (u === "PercentOrRate") return "% / Rate";
    if (u === "Days") return "Days";
    if (u === "Hours") return "Hours";
    return u || "—";
  }

  function kpiUnitTypeTitle(unitType) {
    const u = String(unitType || "").trim();
    if (u === "Count") return "Measure type: count";
    if (u === "PercentOrRate") return "Measure type: percentage or rate";
    if (u === "Days") return "Measure type: days";
    if (u === "Hours") return "Measure type: hours";
    return u ? "Measure type: " + u : "Measure type";
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

  function destroyCharts() {
    ["chart-trend", "chart-kpi-vs", "chart-biz"].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.destroy();
      }
    });
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
   * By business: horizontal bars (Chart.js) — top businesses + Other; length = share of total.
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
    el.hidden = false;
    if (emptyEl) emptyEl.hidden = true;

    const ranked = bizLabels
      .map((name, i) => ({
        name: name.length > 18 ? name.slice(0, 16) + "…" : name,
        fullName: name,
        v: Math.max(0, Math.abs(Number(bizData[i]) || 0)),
      }))
      .sort((a, b) => b.v - a.v);
    const topN = 8;
    const top = ranked.slice(0, topN);
    const otherV = ranked.slice(topN).reduce((a, x) => a + x.v, 0);
    const rows = top.slice();
    if (otherV > 1e-12) {
      rows.push({
        name: "Other",
        fullName: "Other (remaining businesses)",
        v: otherV,
      });
    }
    const values = rows.map((r) => r.v);
    const fullNames = rows.map((r) => r.fullName);
    const labels = rows.map((r) => r.name);
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const bg = labels.map(
      (_, i) => chartPaletteColors[i % chartPaletteColors.length]
    );

    new Chart(el, {
      type: "bar",
      data: {
        labels: labels,
        datasets: [
          {
            data: values,
            backgroundColor: bg,
            borderColor: "#ffffff",
            borderWidth: 1,
            borderRadius: 4,
            maxBarThickness: 14,
          },
        ],
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0].dataIndex;
                return fullNames[i] != null ? String(fullNames[i]) : "";
              },
              label(ctx) {
                const raw = Number(ctx.raw);
                const pct = ((raw / total) * 100).toFixed(1);
                const num =
                  raw >= 1e6
                    ? raw.toLocaleString(undefined, {
                        maximumFractionDigits: 0,
                      })
                    : raw.toLocaleString(undefined, { maximumFractionDigits: 2 });
                return " " + num + " (" + pct + "% of total)";
              },
            },
          },
        },
        scales: {
          x: {
            beginAtZero: true,
            grid: { color: "rgba(15,23,42,0.08)" },
            ticks: { font: { size: 7 }, maxTicksLimit: 6 },
          },
          y: {
            grid: { display: false },
            ticks: { font: { size: 7 } },
          },
        },
      },
    });
  }

  function readFilters(catKey) {
    const elSt = document.getElementById("f-state");
    const elTblScope = document.getElementById("tbl-kpi-scope");
    const elCurFrom = document.getElementById("f-cur-from");
    const elCurTo = document.getElementById("f-cur-to");
    const elCmpFrom = document.getElementById("f-cmp-from");
    const elCmpTo = document.getElementById("f-cmp-to");
    const kpiKeys = readSelectedKpiKeysFromDom();
    let tableKpiScope = "selected";
    if (elTblScope && elTblScope.value === "all") tableKpiScope = "all";

    const mb = dataMonthBounds();
    const defs = defaultToolbarPeriodRanges();

    function readYm(el) {
      if (!el || !el.value || el.value.length < 7) return null;
      return el.value.slice(0, 7);
    }

    let currentFrom = readYm(elCurFrom);
    let currentTo = readYm(elCurTo);
    let comparisonFrom = readYm(elCmpFrom);
    let comparisonTo = readYm(elCmpTo);

    if (!currentTo) currentTo = mb.max;
    if (!currentFrom) currentFrom = defs.curFrom || monthAdd(currentTo, -11);
    if (!comparisonTo) comparisonTo = defs.cmpTo || monthAdd(currentTo, -12);
    if (!comparisonFrom)
      comparisonFrom = defs.cmpFrom || monthAdd(currentFrom, -12);

    if (currentFrom > currentTo) {
      const t = currentFrom;
      currentFrom = currentTo;
      currentTo = t;
    }
    if (comparisonFrom > comparisonTo) {
      const t = comparisonFrom;
      comparisonFrom = comparisonTo;
      comparisonTo = t;
    }

    return {
      catKey,
      kpiKeys: kpiKeys,
      vsMode: "vs_period",
      currentFrom: currentFrom,
      currentTo: currentTo,
      comparisonFrom: comparisonFrom,
      comparisonTo: comparisonTo,
      refMonth: currentTo,
      monthFrom: currentFrom,
      monthTo: currentTo,
      state: elSt ? elSt.value : "all",
      business: readBusinessSelectionFromDom(),
      unitType: "all",
      variable: readVariableSelectionFromDom(),
      tableKpiScope: tableKpiScope,
    };
  }

  /** Table / KPI tiles: months in Current Period (plus non-month filters). */
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
      if (f.monthFrom && f.monthTo) {
        if (r.yearMonth < f.monthFrom || r.yearMonth > f.monthTo) return false;
      } else if (r.yearMonth !== f.refMonth) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (!businessFilterMatchesRow(r, f)) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      return true;
    });
  }

  function applyChartFilter(rows, f) {
    const range = new Set(currentPeriodMonthList(f));
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
      if (!businessFilterMatchesRow(r, f)) return false;
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
  function kpiScopePanelHtml(catKey, kpisMeta) {
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
          escapeHtml(k.kpiName) +
          "</span></label>"
        );
      })
      .join("");
    return (
      '<details class="m2-kpi-scope m2-kpi-scope--toolbar">' +
      '<summary class="m2-kpi-scope__summary" id="f-kpi-scope-label">' +
      '<span class="m2-kpi-scope__summary-text">' +
      '<span class="m2-kpi-scope__title">KPIs</span>' +
      '<span class="m2-kpi-scope__count" id="f-kpi-count"></span>' +
      "</span></summary>" +
      '<div class="m2-kpi-panel" id="f-kpi-panel" role="group" aria-labelledby="f-kpi-scope-label">' +
      '<div class="m2-kpi-panel__bar">' +
      '<button type="button" class="m2-btn m2-btn--tiny" id="f-kpi-all">All</button>' +
      '<button type="button" class="m2-btn m2-btn--tiny" id="f-kpi-none">None</button>' +
      "</div>" +
      '<div class="m2-kpi-panel__list">' +
      boxes +
      "</div></div></details>"
    );
  }

  /** Business unit — KPI-style checkboxes (m2 chrome). */
  function businessFilterFieldHtml(bizList) {
    const list = bizList || [];
    const boxes = list
      .map((b, i) => {
        const id = "f-biz-cb-" + i;
        return (
          '<label class="m2-kpi-cb" for="' +
          id +
          '">' +
          '<input type="checkbox" class="f-biz-cb" name="f-biz-cb" id="' +
          id +
          '" value="' +
          escapeAttr(b) +
          '"/>' +
          '<span class="m2-kpi-cb__text">' +
          escapeHtml(b) +
          "</span></label>"
        );
      })
      .join("");
    return (
      '<div class="field field--toolbar-scope field--biz-scope-modern">' +
      '<span class="field-label" id="f-biz-field-lbl">Business unit</span>' +
      '<details class="m2-kpi-scope m2-kpi-scope--toolbar" id="f-biz-details">' +
      '<summary class="m2-kpi-scope__summary" aria-labelledby="f-biz-field-lbl" title="Business unit">' +
      '<span class="m2-kpi-scope__summary-text">' +
      '<span class="m2-kpi-scope__title">Business unit</span>' +
      '<span class="m2-kpi-scope__count" id="f-biz-hint"></span>' +
      "</span></summary>" +
      '<div class="m2-kpi-panel" id="f-biz-panel" role="group" aria-labelledby="f-biz-field-lbl">' +
      '<div class="m2-kpi-panel__bar">' +
      '<button type="button" class="m2-btn m2-btn--tiny" id="f-biz-btn-all">All</button>' +
      '<button type="button" class="m2-btn m2-btn--tiny" id="f-biz-btn-none">None</button>' +
      "</div>" +
      '<div class="m2-kpi-panel__list">' +
      boxes +
      "</div></div></details></div>"
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

  /**
   * Incident Management: same order as main app (Fatality → … → Investigation Closure %).
   * 56/57 = Fire incidents / Property Damage incidents when present in data.
   */
  const INCIDENT_KPI_ORDER = [
    8, 9, 10, 11, 12, 19, 14, 22, 4, 3, 7, 28, 56, 57, 1, 2, 5, 15, 44,
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
      if (!businessFilterMatchesRow(r, f)) return false;
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

  function aggregateKpiOverMonthList(baseRows, kk, months, ut) {
    const add = isAdditiveUnit(ut);
    function aggMonth(ym) {
      if (!ym) return null;
      const rows = baseRows.filter(
        (r) =>
          String(r.kpiKey) === String(kk) && String(r.yearMonth) === String(ym)
      );
      if (!rows.length) return null;
      if (add) {
        return rows.reduce((a, r) => a + Number(r.value), 0);
      }
      return avg(rows.map((r) => r.value));
    }
    const vals = months
      .map(aggMonth)
      .filter((v) => v != null && !Number.isNaN(v));
    if (!vals.length) return null;
    if (add) return vals.reduce((a, b) => a + b, 0);
    return avg(vals);
  }

  function buildKpiDetailMetrics(catKey, kpisMeta, f) {
    let sorted = sortKpisForDisplay(catKey, kpisMeta);
    if (f.kpiKeys && f.kpiKeys.length) {
      sorted = sorted.filter((k) =>
        f.kpiKeys.includes(String(k.kpiKey))
      );
    } else {
      sorted = [];
    }
    const base = applyNonMonthFilters(getRowsForCategory(catKey), f);
    const curM = currentPeriodMonthList(f);
    const cmpM = comparisonPeriodMonthList(f);

    return sorted.map((k) => {
      const kk = k.kpiKey;
      const ut = k.unitType;
      const cur = aggregateKpiOverMonthList(base, kk, curM, ut);
      const baseVal = aggregateKpiOverMonthList(base, kk, cmpM, ut);
      const vsPct = pctChange(cur, baseVal);
      const vsD = vsDir(cur, baseVal);
      return {
        kpiKey: kk,
        kpiName: k.kpiName,
        unitType: ut,
        value: cur,
        vsPct: vsPct,
        vsDir: vsD,
        vsMode: "vs_period",
        periodCaption: tilePeriodForKpi(f.refMonth, "vs_period", ut),
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

  function appendM2KpiTile(board, item, extraClass) {
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
    const vsLbl = "Current vs comparison period";
    const vsShort = vsTagShort(item.vsMode || "vs_period");
    const periodLine =
      item.periodCaption || tilePeriodForKpi(null, "vs_period", item.unitType);
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
      '<div class="m2-kpi-tile__head">' +
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

  function renderMultiKpiCards(container, aggregatesList) {
    container.innerHTML = "";
    const sorted = sortKpiTilesForDisplay(aggregatesList);
    if (!sorted.length) return;

    const row = document.createElement("div");
    row.className = "m2-kpi-metrics-row";
    row.setAttribute("role", "group");
    row.setAttribute(
      "aria-label",
      "KPI metrics for the KPIs selected in KPI scope"
    );

    function appendMetricCardBlock(items, cardMod, useTailTiles) {
      if (!items.length) return;
      const card = document.createElement("section");
      card.className =
        "m2-kpi-metric-card" + (cardMod ? " " + cardMod : "");
      const grid = document.createElement("div");
      grid.className = "m2-kpi-board m2-kpi-board--quad";
      grid.setAttribute("role", "list");
      const tailTile = useTailTiles ? "m2-kpi-tile--tail" : "";
      items.forEach((item) => {
        appendM2KpiTile(grid, item, tailTile);
      });
      while (grid.children.length < 4) {
        appendM2KpiPlaceholder(grid);
      }
      card.appendChild(grid);
      row.appendChild(card);
    }

    const headItems = sorted.slice(0, 12);
    const tailItems = sorted.slice(12);
    const headChunks = chunkArray(headItems, 4);
    headChunks.forEach((chunk) => {
      appendMetricCardBlock(chunk, "", false);
    });

    const tailA = tailItems.slice(0, 4);
    const tailB = tailItems.slice(4);
    if (tailA.length) {
      appendMetricCardBlock(tailA, "m2-kpi-metric-card--tail", true);
    }
    if (tailB.length) {
      appendMetricCardBlock(tailB, "m2-kpi-metric-card--tail", true);
    }

    container.appendChild(row);
  }

  /**
   * Three distinct stories:
   * 1) Line — KPI level vs time (same window as filters).
   * 2) Share list — business composition of value in that window.
   * 3) Horizontal bars — Vs % change by KPI (same logic as tiles).
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
    const winMonths = new Set(currentPeriodMonthList(f));
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
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: {
              display: lineDatasets.length > 1,
              position: "bottom",
              labels: { font: { size: 7 }, boxWidth: 8 },
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
              ticks: { font: { size: 8 }, maxRotation: 40 },
              grid: { color: "rgba(15,23,42,0.06)" },
            },
            y: {
              beginAtZero: false,
              ticks: { font: { size: 8 } },
              grid: { color: "rgba(15,23,42,0.06)" },
            },
          },
        },
      });
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
    const utChart = kpiUnitTypeForFilter(catKey, f);
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

    const aggForVs = buildKpiDetailMetrics(catKey, kpisMeta, f);
    const rankedVs = sortKpiTilesForDisplay(aggForVs);
    const elKpiVs = document.getElementById("chart-kpi-vs");
    if (elKpiVs && rankedVs.length) {
      const kpiLab = rankedVs.map((a) =>
        a.kpiName.length > 22 ? a.kpiName.slice(0, 20) + "…" : a.kpiName
      );
      const vsData = rankedVs.map((a) =>
        a.vsPct != null && !Number.isNaN(a.vsPct) ? a.vsPct : 0
      );
      const vsBg = rankedVs.map((a) => {
        const p = a.vsPct;
        if (p == null || Number.isNaN(p)) return "rgba(109, 110, 113, 0.35)";
        if (p > 0) return "rgba(0, 177, 107, 0.88)";
        if (p < 0) return "rgba(240, 76, 35, 0.88)";
        return "rgba(142, 39, 143, 0.55)";
      });
      new Chart(elKpiVs, {
        type: "bar",
        data: {
          labels: kpiLab,
          datasets: [
            {
              label: "Δ %",
              data: vsData,
              backgroundColor: vsBg,
              borderColor: "#FFFFFF",
              borderWidth: 1,
              borderRadius: 3,
              maxBarThickness: 12,
            },
          ],
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                title: function (items) {
                  const i = items[0].dataIndex;
                  return rankedVs[i] ? rankedVs[i].kpiName : "";
                },
                label: function (ctx) {
                  const i = ctx.dataIndex;
                  const p = rankedVs[i].vsPct;
                  if (p == null || Number.isNaN(p)) {
                    return " No Δ% (missing comparison)";
                  }
                  return " " + formatSignedPct(p) + " (same as KPI tiles)";
                },
              },
            },
          },
          scales: {
            x: {
              ticks: {
                font: { size: 7 },
                callback: (v) => v + "%",
              },
              grid: { color: "rgba(109, 110, 113, 0.2)" },
              title: {
                display: true,
                text: "Δ %",
                font: { size: 7 },
              },
            },
            y: {
              ticks: { font: { size: 6 } },
              grid: { display: false },
            },
          },
        },
      });
    }

    updateChartHints(f);
  }

  function updateChartHints(f) {
    const curM = currentPeriodMonthList(f);
    const n = curM.length || 1;
    const trendEl = document.getElementById("chart-trend-hint");
    const bizEl = document.getElementById("chart-biz-hint");
    const kpiVsEl = document.getElementById("chart-kpi-vs-hint");
    if (trendEl) {
      trendEl.textContent = "(lines · " + n + " mo)";
    }
    if (bizEl) {
      bizEl.textContent =
        n > 1 ? "(share · " + n + " mo)" : "(share · Current Period)";
    }
    if (kpiVsEl) kpiVsEl.textContent = "(Δ% vs comparison)";
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

  /**
   * Row-level current vs base (aligned Current vs Comparison periods when possible).
   */
  function rowCompareForMode(row, pool, f) {
    const ref = row.yearMonth;
    const tk = rowTupleKey(row);
    const curMonths = currentPeriodMonthList(f);
    const cmpMonths = comparisonPeriodMonthList(f);
    const cur = rowValueAt(pool, tk, ref);
    let baseVal = null;
    const ix = curMonths.indexOf(ref);
    if (ix !== -1 && cmpMonths.length > ix) {
      baseVal = rowValueAt(pool, tk, cmpMonths[ix]);
    }
    if (baseVal == null) {
      baseVal = rowValueAt(pool, tk, monthAdd(ref, -12));
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
            if (!businessFilterMatchesRow(r, f)) return false;
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
      summaryEl.textContent = "";
      summaryEl.removeAttribute("title");
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
        '<div class="m2-detail-empty" role="status"></div>';
    } else {
      const head =
        '<table class="m2-evidence-table">' +
        "<thead><tr>" +
        '<th scope="col">Month</th>' +
        '<th scope="col" class="m2-evidence-table__col-kpi">KPI</th>' +
        '<th scope="col">State</th>' +
        '<th scope="col">Business</th>' +
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
          : start +
            1 +
            "–" +
            Math.min(start + pageSize, total) +
            "/" +
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
        if (!businessFilterMatchesRow(r, f)) return false;
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
    announce(label.replace(/\.\s*$/, "") + (label ? " · " : "") + String(n));
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
          '<div class="m2-empty">No KPI data for this slice. Adjust periods, geography, or KPI scope.</div>';
      } else {
        renderMultiKpiCards(multiWrap, aggList);
      }
    }

    buildCharts(catKey, f);
    renderDetailCards(catKey);

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
        "This preview includes Incident Management and Hazard & Observation Management. Choose one on the category list."
      );
      return;
    }

    const kpisMeta = getKpis(catKey);
    const rowsForCat = getRowsForCategory(catKey);
    const cfg = getFilterConfig(catKey);

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

    const wrap = document.createElement("div");
    wrap.className = "cat-view cat-view--modern";
    const businessFieldHtml = cfg.showBusiness
      ? businessFilterFieldHtml(bizList)
      : "";
    const periodRangesFieldHtml =
      '<div class="field field--period-ranges field--calendar-compact">' +
      '<div class="field-period-pair">' +
      '<div class="field-period-row">' +
      '<span class="field-label field-label--inline" id="f-cur-lbl" title="Current period">Current</span>' +
      '<div class="field-calendar-range" role="group" aria-labelledby="f-cur-lbl">' +
      '<label class="visually-hidden" for="f-cur-from">Current period from</label>' +
      '<input type="month" id="f-cur-from" class="toolbar-date toolbar-month" />' +
      '<span class="field-calendar-sep" aria-hidden="true">→</span>' +
      '<label class="visually-hidden" for="f-cur-to">Current period to</label>' +
      '<input type="month" id="f-cur-to" class="toolbar-date toolbar-month" />' +
      "</div></div>" +
      '<div class="field-period-row">' +
      '<span class="field-label field-label--inline" id="f-cmp-lbl" title="Comparison period">Compare</span>' +
      '<div class="field-calendar-range" role="group" aria-labelledby="f-cmp-lbl">' +
      '<label class="visually-hidden" for="f-cmp-from">Comparison period from</label>' +
      '<input type="month" id="f-cmp-from" class="toolbar-date toolbar-month" />' +
      '<span class="field-calendar-sep" aria-hidden="true">→</span>' +
      '<label class="visually-hidden" for="f-cmp-to">Comparison period to</label>' +
      '<input type="month" id="f-cmp-to" class="toolbar-date toolbar-month" />' +
      "</div></div></div></div>";
    const stateFieldHtml = cfg.showState
      ? '<div class="field"><label class="field-label" for="f-state">State</label>' +
        '<select id="f-state">' +
        stateOpts +
        "</select></div>"
      : "";
    const filterCore = businessFieldHtml + periodRangesFieldHtml + stateFieldHtml;
    const variableFieldHtml = variableFilterFieldHtml();
    const kpiSurfaceHtml =
      cfg.showKpi
        ? '<div class="cat-toolbar__kpi-surface" role="group" aria-labelledby="f-kpi-surface-lbl">' +
          '<span id="f-kpi-surface-lbl" class="cat-toolbar__kpi-surface__lbl">KPI list</span>' +
          '<div class="cat-toolbar__kpi-surface__control">' +
          kpiScopePanelHtml(catKey, kpisMeta) +
          "</div></div>"
        : "";
    const filtersAllScrollHtml =
      '<div class="cat-toolbar__filters-all-scroll">' +
      variableFieldHtml +
      filterCore +
      kpiSurfaceHtml +
      "</div>";
    const toolbarInner =
      '<div class="cat-toolbar__inner" role="group">' +
      '<div class="cat-toolbar__filters-scroll">' +
      filtersAllScrollHtml +
      "</div>" +
      '<div class="toolbar-actions">' +
      '<button type="button" class="btn btn--reset-compact" id="f-reset">Reset</button>' +
      "</div></div>";
    wrap.innerHTML =
      '<div class="cat-top-bar cat-top-bar--filters-only">' +
      '<fieldset id="cat-heading" tabindex="-1" class="cat-toolbar cat-toolbar--compact" aria-label="Refine results">' +
      '<legend class="visually-hidden">Refine results</legend>' +
      toolbarInner +
      "</fieldset></div>" +
      '<section class="m2-kpi-block" aria-labelledby="m2-kpi-h">' +
      '<div class="m2-kpi-block__head">' +
      '<h3 id="m2-kpi-h" class="m2-section-title">KPI metrics</h3>' +
      "</div>" +
      '<div id="multi-kpi-wrap"></div>' +
      "</section>" +
      '<div class="m2-charts m2-charts--alt cat-charts" role="group" aria-label="Three chart views: time series, business share, change vs comparison period by KPI">' +
      '<div class="chart-box m2-chart-card m2-chart-card--trend"><h3 class="m2-chart-title">Time series <span id="chart-trend-hint" class="chart-box__hint">(lines)</span></h3><div class="chart-canvas-wrap m2-chart-canvas"><canvas id="chart-trend" role="img" aria-label="Line chart: one series per KPI over months"></canvas></div></div>' +
      '<div class="chart-box m2-chart-card m2-chart-card--biz"><h3 class="m2-chart-title">By business <span id="chart-biz-hint" class="chart-box__hint">(share · latest mo)</span></h3><div class="chart-canvas-wrap m2-chart-canvas m2-chart-canvas--wide m2-chart-wrap--biz-list"><canvas id="chart-biz" role="img" aria-label="Horizontal bars: share of total value by business"></canvas><p id="chart-biz-empty" class="chart-biz-empty m2-chart-biz-empty" hidden></p></div></div>' +
      '<div class="chart-box m2-chart-card m2-chart-card--kpi-vs"><h3 class="m2-chart-title">Change vs comparison <span id="chart-kpi-vs-hint" class="chart-box__hint">(Δ%)</span></h3><div class="chart-canvas-wrap m2-chart-canvas"><canvas id="chart-kpi-vs" role="img" aria-label="Horizontal bars: percent change versus comparison period for each KPI in scope"></canvas></div></div>' +
      "</div>" +
      '<div class="table-zone m2-table-zone m2-evidence-zone">' +
      '<div class="table-zone__head m2-evidence-head">' +
      '<div class="m2-evidence-head__row">' +
      '<span class="table-zone__label m2-evidence-head__label" id="m2-detail-heading">Detail data</span>' +
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
      '<div class="m2-detail-scroll m2-evidence-scroll" role="region" aria-labelledby="m2-detail-heading" tabindex="0">' +
      '<div id="detail-cards" class="m2-evidence-table-host"></div></div></div>' +
      '<p class="cat-context m2-cat-context" id="cat-context">' +
      escapeHtml(cat.uxNote) +
      "</p>";

    root.innerHTML = "";
    root.appendChild(wrap);

    if (cfg.showState) document.getElementById("f-state").value = "all";
    if (cfg.showBusiness) initBusinessSiteCheckboxFilters();
    applyVariableFilterFromStorage();
    initPeriodRangeInputs();
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

    [
      "f-state",
      "f-cur-from",
      "f-cur-to",
      "f-cmp-from",
      "f-cmp-to",
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", onFilterChange);
    });
    wireBusinessAndSiteFilterControls(onFilterChange);
    wireVariableFilterControls(onFilterChange);
    wireToolbarScopeScrollPanels(wrap);

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
      const btnAll = document.getElementById("f-kpi-all");
      const btnNone = document.getElementById("f-kpi-none");
      if (btnAll) {
        btnAll.addEventListener("click", () => {
          document
            .querySelectorAll('#f-kpi-panel input[name="f-kpi-cb"]')
            .forEach((cb) => {
              cb.checked = true;
            });
          saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
          updateKpiScopeCount();
          onFilterChange();
        });
      }
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
      ["f-cur-from", "f-cur-to", "f-cmp-from", "f-cmp-to"].forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.value = "";
      });
      initPeriodRangeInputs();
      if (cfg.showState) document.getElementById("f-state").value = "all";
      if (cfg.showBusiness) initBusinessSiteCheckboxFilters();
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
        const def = loadKpiSelection(catKey, kpisMeta);
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
      '<div class="m2-v3__main m2-v3__main--clean">' +
      '<div class="m2-v3__hero-card">' +
      '<header class="m2-v3__mast m2-v3__mast--clean">' +
      '<p class="m2-v3__eyebrow" id="landing-h" tabindex="-1">Insights</p>' +
      '<h2 class="m2-v3__headline m2-v3__headline--clean">Explore safety KPIs by domain</h2>' +
      '<p class="m2-v3__sub m2-v3__sub--clean">Choose a category, set filters and KPI scope, then review tiles, charts, and the detail table. Fixed layout <span class="m2-v3__dim">1280×720</span>. Sample data — confirm in Power BI.</p>' +
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
      "<li><strong>Filter</strong> — Versus window, state, business, verticals.</li>" +
      "<li><strong>Compare</strong> — Time series, by business, Vs % by KPI.</li>" +
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
      "Insights home. Browse domains to open a category; Incident Management and Hazard & Observation Management are live in this preview."
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
      '<div class="m2-cat-dir-intro">' +
      '<h2 id="home-h">Safety domains</h2>' +
      '<p class="m2-cat-dir-lede">Search by <strong>category or KPI name</strong>. <strong>Live preview</strong>: <strong>Incident Management</strong> and <strong>Hazard &amp; Observation Management</strong>.</p>' +
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
                " KPIs. Not in live preview; open Incident Management or Hazard & Observation Management."
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
