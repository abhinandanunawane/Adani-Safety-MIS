/**
 * Adani Safety MIS — fixed 1280×720 preview (no page scroll)
 */
(function () {
  "use strict";

  const root = document.getElementById("app-root");
  const liveRegion = document.getElementById("sr-live");
  const DATA = window.__DASHBOARD_DATA__;

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
  const PAGE_SIZE = 8;
  let tableState = {
    sortKey: "yearMonth",
    asc: false,
    page: 0,
  };

  let currentCategoryKey = null;
  let catSearchAnnounceTimer = null;

  /** Shared preview: only Incident Management (key 1) is navigable; other categories are visible but inactive. */
  const ACTIVE_PREVIEW_CATEGORY_KEY = 1;

  function announce(msg) {
    if (liveRegion) liveRegion.textContent = msg;
  }

  /** Vs comparison (monthly data; shorter labels use month-level proxies). */
  const DEFAULT_VS_MODE = "vs_last_month";
  const VS_OPTIONS = [
    { id: "vs_yesterday", label: "Vs Yesterday" },
    { id: "vs_last_week", label: "Vs Last Week" },
    { id: "vs_last_month", label: "Vs Last Month" },
    { id: "vs_last_quarter", label: "Vs Last Quarter" },
    { id: "vs_last_year", label: "Vs Last Year" },
  ];

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

  /** Honest comparison window for rates / averages (monthly facts). */
  function vsCompareShort(mode) {
    switch (mode) {
      case "vs_last_month":
      case "vs_yesterday":
        return "vs prior month";
      case "vs_last_year":
        return "vs year-ago month";
      case "vs_last_week":
        return "2-mo avg vs prior 2-mo";
      case "vs_last_quarter":
        return "3-mo avg vs prior 3-mo";
      default:
        return "vs prior month";
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
    const m = formatMonthYear(refMonth);
    const add = isAdditiveUnit(unitType);
    if (mode === "vs_last_year" && add) {
      return m + " · 12-mo total vs prior 12-mo";
    }
    if (mode === "vs_last_year" && !add) {
      return m + " · vs year-ago month";
    }
    if (mode === "vs_last_quarter" && add) {
      return m + " · 3-mo total vs prior 3-mo";
    }
    if (mode === "vs_last_week" && add) {
      return m + " · 2-mo total vs prior 2-mo";
    }
    if ((mode === "vs_last_month" || mode === "vs_yesterday") && add) {
      return m + " · month total vs prior month";
    }
    return m + " · " + vsCompareShort(mode);
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

  /** Table value column: formatted actual only (Target has its own column). */
  function tableValueCell(r) {
    const unit = r.unitType || "";
    const actualStr = formatValue(r.value, unit);
    return (
      '<td class="col-num">' +
      '<span class="data-table__value-num">' +
      escapeHtml(actualStr) +
      "</span></td>"
    );
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

  function destroyCharts() {
    ["chart-line", "chart-doughnut"].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.destroy();
      }
    });
    const bizHost = document.getElementById("chart-biz-breakdown");
    if (bizHost) bizHost.innerHTML = "";
  }

  /* Soft tints — readable on white without heavy saturation */
  const chartPaletteColors = [
    "rgba(0, 177, 107, 0.45)",
    "rgba(0, 109, 182, 0.45)",
    "rgba(142, 39, 143, 0.4)",
    "rgba(240, 76, 35, 0.42)",
    "rgba(0, 177, 107, 0.32)",
    "rgba(0, 109, 182, 0.32)",
    "rgba(142, 39, 143, 0.3)",
    "rgba(240, 76, 35, 0.32)",
  ];

  /**
   * By business: ranked list + proportional bars (readable labels; scroll inside panel).
   */
  function renderBizBreakdown(bizLabels, bizData) {
    const host = document.getElementById("chart-biz-breakdown");
    if (!host) return;
    if (!bizLabels.length) {
      host.innerHTML =
        '<p class="chart-biz-empty">No business data for current filters.</p>';
      return;
    }
    const pairs = bizLabels.map((name, i) => ({ name, v: Number(bizData[i]) }));
    pairs.sort((a, b) => Math.abs(b.v) - Math.abs(a.v));
    const maxV = Math.max(
      ...pairs.map((p) => Math.abs(p.v)),
      Number.EPSILON
    );
    host.innerHTML = pairs
      .map((p, i) => {
        const pct = (Math.abs(p.v) / maxV) * 100;
        const col = chartPaletteColors[i % chartPaletteColors.length];
        const num =
          Number.isFinite(p.v) && Math.abs(p.v) >= 1e6
            ? p.v.toLocaleString(undefined, { maximumFractionDigits: 0 })
            : Number.isFinite(p.v)
              ? p.v.toLocaleString(undefined, { maximumFractionDigits: 2 })
              : "—";
        return (
          '<div class="chart-biz-row" role="listitem">' +
          '<div class="chart-biz-row__line">' +
          '<span class="chart-biz-row__name">' +
          escapeHtml(p.name) +
          "</span>" +
          '<span class="chart-biz-row__val">' +
          escapeHtml(num) +
          "</span></div>" +
          '<div class="chart-biz-row__track" aria-hidden="true">' +
          '<div class="chart-biz-row__fill" style="width:' +
          pct.toFixed(1) +
          "%;background-color:" +
          col +
          '"></div></div></div>'
        );
      })
      .join("");
  }

  function readFilters(catKey) {
    const elKpi = document.getElementById("f-kpi");
    const elVs = document.getElementById("f-vs");
    const elSt = document.getElementById("f-state");
    if (!elSt) return null;
    const elBiz = document.getElementById("f-biz");
    const elUnit = document.getElementById("f-unit");
    return {
      catKey,
      kpi: elKpi ? elKpi.value : "all",
      vsMode: elVs ? elVs.value : DEFAULT_VS_MODE,
      refMonth: getRefMonth(),
      state: elSt.value,
      business: elBiz ? elBiz.value : "all",
      unitType: elUnit ? elUnit.value : "all",
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
    const range = new Set(chartMonthKeys(f.refMonth, 12));
    return rows.filter((r) => {
      if (f.kpi !== "all" && String(r.kpiKey) !== f.kpi) return false;
      if (!range.has(r.yearMonth)) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
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
      showUnitType: false,
    };
    if (catKey === 1) {
      // Incident Management: values span Count/Days/%; Unit filter helps isolate patterns.
      cfg.showUnitType = true;
    }
    return cfg;
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

  function applyNonMonthFilters(rows, f) {
    return rows.filter((r) => {
      if (f.kpi !== "all" && String(r.kpiKey) !== f.kpi) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (f.business !== "all" && r.businessName !== f.business) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
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
    const m12 = monthAdd(ref, -12);

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
        if (add) {
          cur = aggMonthsRolling(ref, 12);
          baseVal = aggMonthsRolling(monthAdd(ref, -12), 12);
        } else {
          cur = aggMonth(ref);
          baseVal = aggMonth(m12);
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

  function renderMultiKpiCards(container, aggregatesList, refMonth, vsMode) {
    container.innerHTML = "";
    const chunks = chunkArray(aggregatesList, 4);
    chunks.forEach((chunk, idx) => {
      const card = document.createElement("div");
      card.className = "multi-kpi-card";
      const head = document.createElement("div");
      head.className = "multi-kpi-card__head";
      const start = idx * 4 + 1;
      const end = idx * 4 + chunk.length;
      if (chunk.length === 1) {
        head.textContent = "Metric " + start;
      } else {
        head.textContent = "Metrics " + start + "–" + end;
      }
      const grid = document.createElement("div");
      grid.className = "multi-kpi-card__grid";
      chunk.forEach((item) => {
        const tile = document.createElement("div");
        tile.className = "multi-kpi-tile";
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
      });
      while (grid.children.length < 4) {
        const ph = document.createElement("div");
        ph.className = "multi-kpi-tile";
        ph.style.visibility = "hidden";
        ph.innerHTML = "&nbsp;";
        grid.appendChild(ph);
      }
      card.appendChild(head);
      card.appendChild(grid);
      container.appendChild(card);
    });
  }

  /**
   * Trend: last 12 months (filtered). By business + unit mix: reference month only so counts match the detail table.
   */
  function buildCharts(catKey, f) {
    destroyCharts();

    const poolAll = getRowsForCategory(catKey);
    const filteredTrend = applyChartFilter(poolAll, f);
    const snapPool = applyNonMonthFilters(poolAll, f);
    const snapRows = snapPool.filter((r) => r.yearMonth === f.refMonth);

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
    const bizData = bizLabels.map((b) => avg(byBizVals[b]));

    const unitMap = {};
    snapRows.forEach((r) => {
      const u = r.unitType || "—";
      if (f.kpi === "all") {
        unitMap[u] = (unitMap[u] || 0) + 1;
      } else {
        const v = Number(r.value);
        unitMap[u] = (unitMap[u] || 0) + (Number.isFinite(v) ? v : 0);
      }
    });
    const unitLabels = Object.keys(unitMap);
    const unitData = unitLabels.map((u) => unitMap[u]);

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
              borderColor: "rgba(0, 109, 182, 0.75)",
              backgroundColor: "rgba(0, 177, 107, 0.08)",
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
            y: { beginAtZero: false, ticks: { font: { size: 9 } } },
            x: { ticks: { font: { size: 8 }, maxRotation: 45 } },
          },
        },
      });
    }

    renderBizBreakdown(bizLabels, bizData);

    const elD = document.getElementById("chart-doughnut");
    if (elD && unitLabels.length) {
      new Chart(elD, {
        type: "doughnut",
        data: {
          labels: unitLabels,
          datasets: [
            {
              data: unitData,
              backgroundColor: chartPaletteColors,
              borderColor: "#ffffff",
              borderWidth: 1,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: "right", labels: { font: { size: 8 }, boxWidth: 10 } },
          },
        },
      });
    }
  }

  function sortRows(rows, key, asc) {
    const copy = rows.slice();
    const mul = asc ? 1 : -1;
    copy.sort((a, b) => {
      let va = a[key];
      let vb = b[key];
      if (key === "value" || key === "target") {
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

  function rowVsForMode(row, pool, f) {
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
    const m12 = monthAdd(ref, -12);

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
      if (add) {
        cur = rowWindowAgg(pool, tk, ref, 12, ut);
        baseVal = rowWindowAgg(pool, tk, monthAdd(ref, -12), 12, ut);
      } else {
        cur = rowValueAt(pool, tk, ref);
        baseVal = rowValueAt(pool, tk, m12);
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
    return {
      pct: pctChange(cur, baseVal),
      dir: vsDir(cur, baseVal),
    };
  }

  function renderTableBody(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    const pool = applyNonMonthFilters(getRowsForCategory(catKey), f);
    let rows = applyRowFilter(pool, f);
    rows = sortRows(rows, tableState.sortKey, tableState.asc);

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

    const cap = document.getElementById("tbl-caption");
    if (cap) {
      if (total === 0) {
        cap.textContent =
          "No rows match for the reference month. Set State to All or reset filters.";
      } else {
        cap.textContent =
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
            tableValueCell(r) +
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
    const n = applyRowFilter(getRowsForCategory(catKey), f).length;
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

    if (catKey !== ACTIVE_PREVIEW_CATEGORY_KEY) {
      history.replaceState(null, "", "#categories");
      renderCategories();
      announce(
        "This preview only includes Incident Management. Choose Incident Management on the category list."
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

    const unitList = distinctSorted(rowsForCat, (r) => r.unitType);
    const unitOpts =
      '<option value="all">All units</option>' +
      unitList
        .map(
          (u) =>
            '<option value="' +
            escapeHtml(u) +
            '">' +
            escapeHtml(u) +
            "</option>"
        )
        .join("");

    const wrap = document.createElement("div");
    wrap.className = "cat-view";
    const coreFields =
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
      (cfg.showUnitType
        ? '<div class="field"><label class="field-label" for="f-unit">Unit</label>' +
          '<select id="f-unit">' +
          unitOpts +
          "</select></div>"
        : "");
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
      '<p id="filter-hint" class="visually-hidden">Narrow by KPI, Versus comparison, state, business, or unit when shown. Trend uses twelve months; business and unit charts use the reference month. Count or Hours KPIs sum across months in wider Vs windows.</p>' +
      coreFields +
      '<div class="toolbar-actions">' +
      '<button type="button" class="btn" id="f-reset">Reset</button>' +
      "</div></div></fieldset>" +
      "</div>" +
      '<fieldset class="kpi-summary-region">' +
      '<legend class="visually-hidden">KPI summary for current filters</legend>' +
      '<div class="multi-kpi-row" id="multi-kpi-wrap"></div>' +
      "</fieldset>" +
      '<div class="cat-charts" role="group" aria-label="Charts for filtered data">' +
      '<div class="chart-box"><h3>Trend <span class="chart-box__hint">(12 months)</span></h3><div class="chart-canvas-wrap"><canvas id="chart-line" role="img" aria-label="Line chart: monthly average for filtered KPIs, last twelve months"></canvas></div></div>' +
      '<div class="chart-box chart-box--biz"><h3>By business <span class="chart-box__hint">(reference month)</span></h3><div class="chart-canvas-wrap chart-canvas-wrap--biz"><div id="chart-biz-breakdown" class="chart-biz-breakdown" role="list" aria-label="Average value by business for the reference month and current filters"></div></div></div>' +
      '<div class="chart-box"><h3>Unit mix <span class="chart-box__hint">(reference month)</span></h3><div class="chart-canvas-wrap"><canvas id="chart-doughnut" role="img" aria-label="Unit mix for the reference month: row counts if all KPIs, else summed values for the selected KPI"></canvas></div></div>' +
      "</div>" +
      '<div class="table-zone">' +
      '<div class="table-zone__head"><span>Detail data</span>' +
      '<div class="table-pager">' +
      '<button type="button" id="tbl-prev" aria-label="Previous page">Prev</button>' +
      '<span id="tbl-pageinfo"></span>' +
      '<button type="button" id="tbl-next" aria-label="Next page">Next</button>' +
      "</div></div>" +
      '<div class="table-scroll">' +
      '<table class="data-table" id="tbl-detail">' +
      '<caption id="tbl-caption">Detailed rows matching filters. Use column headers to sort.</caption>' +
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
    if (cfg.showUnitType) document.getElementById("f-unit").value = "all";
    if (cfg.showKpi) document.getElementById("f-kpi").value = "all";

    function onFilterChange() {
      tableState.page = 0;
      refreshCategoryView(catKey);
    }

    ["f-kpi", "f-vs", "f-state", "f-biz", "f-unit"].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", onFilterChange);
    });

    document.getElementById("f-reset").addEventListener("click", () => {
      document.getElementById("f-vs").value = DEFAULT_VS_MODE;
      if (cfg.showKpi) document.getElementById("f-kpi").value = "all";
      if (cfg.showState) document.getElementById("f-state").value = "all";
      if (cfg.showBusiness) document.getElementById("f-biz").value = "all";
      if (cfg.showUnitType) document.getElementById("f-unit").value = "all";
      tableState.page = 0;
      refreshCategoryView(catKey);
    });

    wireTableHeaders(catKey);
    refreshCategoryView(catKey);
    const h = document.getElementById("cat-heading");
    if (h) h.focus();
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
      '<p class="landing__subtitle">Group-level safety KPI dashboard for all Adani businesses — monitor trends, drill down by category, and compare performance month-over-month and year-over-year.</p>' +
      '<div class="landing__bullets" role="list">' +
      '<div class="landing__bullet" role="listitem"><strong>What’s included</strong><span>Safety KPIs by category, Versus comparison, filters, trend and business charts, unit mix, KPI tiles with Vs % change, and a sortable detail table.</span></div>' +
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
      "</div>" +
      '<section class="landing__guide" aria-labelledby="guide-h">' +
      '<div class="landing__guide-grid">' +
      '<div class="landing__guide-panel landing__guide-panel--how">' +
      '<h3 id="guide-h" class="landing__guide-panel__title">How to use this dashboard</h3>' +
      '<ul class="landing__guide-items">' +
      landingGuideRow(
        "play",
        "Start now",
        "to open the category list."
      ) +
      landingGuideRow(
        "grid",
        "Select a category",
        "to open that topic and explore its KPIs."
      ) +
      landingGuideRow(
        "filter",
        "Refine results",
        "using filters — core filters on every page, plus extra filters on some categories where they add insight."
      ) +
      landingGuideRow(
        "layers",
        "Review KPI tiles",
        "(Vs % change), then charts, then the detail table for underlying rows."
      ) +
      "</ul>" +
      "</div>" +
      '<div class="landing__guide-panel landing__guide-panel--screen">' +
      '<h4 id="guide-screen-h" class="landing__guide-panel__title landing__guide-panel__title--sub">What’s on each category screen</h4>' +
      '<ul class="landing__guide-items" aria-labelledby="guide-screen-h">' +
      landingGuideRow(
        "tiles",
        "KPI summary tiles",
        "— current values with the selected Versus comparison (% change with arrow)."
      ) +
      landingGuideRow(
        "chart",
        "Charts",
        "— trend over time, by business, and unit mix for the filtered data."
      ) +
      landingGuideRow(
        "table",
        "Detail data",
        "— sortable table; use column headers and pagination to scan rows."
      ) +
      "</ul>" +
      "</div></div></section>";

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
    announce("About this dashboard. Use Start now to continue.");
  }

  function renderCategories() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#categories");

    const box = document.createElement("div");
    box.className = "home-body";
    box.innerHTML =
      journeyStepsHtml(2) +
      '<div class="home-intro">' +
      '<h2 id="home-h">Categories</h2>' +
      '<p class="home-lede">In this shared preview, only <strong>Incident Management</strong> is interactive; other categories are shown for context. Open Incident Management to explore KPIs and filters. Use <strong>Home</strong> in the header to return to the overview.</p>' +
      '<div class="home-tools" role="search">' +
      '<label class="home-search-label" for="cat-q">Find category</label>' +
      '<input id="cat-q" class="home-search" type="search" placeholder="Type to filter (e.g., Incident, Training…)" autocomplete="off" />' +
      "</div>" +
      "</div>" +
      '<div class="home-grid" role="list" aria-labelledby="home-h"></div>';

    const grid = box.querySelector(".home-grid");
    function scheduleCategorySearchAnnounce(opts) {
      const o = opts || {};
      clearTimeout(catSearchAnnounceTimer);
      const run = () => {
        if (o.noMatch) {
          announce("No categories match your search.");
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
        (c.categoryName || "").toLowerCase().includes(q)
      );
      if (!cats.length) {
        grid.removeAttribute("role");
        grid.removeAttribute("aria-labelledby");
        const empty = document.createElement("div");
        empty.className = "home-empty";
        empty.setAttribute("role", "status");
        empty.setAttribute("aria-live", "polite");
        empty.innerHTML =
          "<p><strong>No categories match.</strong> Clear the search or try a shorter keyword.</p>";
        grid.appendChild(empty);
        scheduleCategorySearchAnnounce({ noMatch: true, immediate: o.immediate });
        return;
      }
      grid.setAttribute("role", "list");
      grid.setAttribute("aria-labelledby", "home-h");
      cats.forEach((cat) => {
        const active = cat.categoryKey === ACTIVE_PREVIEW_CATEGORY_KEY;
        const el = active
          ? document.createElement("button")
          : document.createElement("div");
        if (active) el.type = "button";
        el.className = "category-card" + (active ? "" : " category-card--disabled");
        el.setAttribute("role", "listitem");
        el.setAttribute(
          "aria-label",
          active
            ? cat.categoryName + ", " + cat.kpiCount + " KPIs"
            : cat.categoryName +
                ", " +
                cat.kpiCount +
                " KPIs. Not available in this preview; open Incident Management."
        );
        if (!active) {
          el.setAttribute("aria-disabled", "true");
          el.setAttribute("tabindex", "-1");
        }
        el.innerHTML =
          '<span class="category-card__icon">' +
          categoryIconSvg(cat.categoryKey) +
          "</span>" +
          '<span class="category-card__stack">' +
          '<span class="category-card__badge">' +
          cat.kpiCount +
          " KPIs</span>" +
          '<span class="category-card__name">' +
          escapeHtml(cat.categoryName) +
          "</span>" +
          '<span class="category-card__meta">' +
          escapeHtml(cat.uxNote) +
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
