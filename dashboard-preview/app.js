/**
 * Adani Safety Performance Profile — fixed 1280×720 preview (no page scroll). Copy and structure support
 * user research, IA, usability testing, accessibility, consistency, hierarchy, iterative UX,
 * UCD, HCI, and CX review (usability, desirability, accessibility, usefulness).
 */
(function () {
  "use strict";

  const root = document.getElementById("app-root");
  const liveRegion = document.getElementById("sr-live");
  const DATA = window.__DASHBOARD_DATA__;
  /** Matches styles.css --font (Segoe UI family + UI sans fallbacks). */
  const FONT_UI =
    '"Segoe UI", "Segoe UI Variable", "Segoe UI Historic", system-ui, sans-serif';

  function updateHeaderNavState() {
    const raw = location.hash || "";
    const h = raw === "" || raw === "#" ? "#landing" : raw;
    const home = document.getElementById("nav-home");
    const cats = document.getElementById("nav-categories");
    const nav = document.getElementById("top-header-nav");
    const onLanding = h === "#landing";
    const onCategories = h === "#categories";
    const onCategoryDrill = /^#cat=\d+/.test(h);
    if (nav) {
      nav.hidden = onLanding;
      nav.setAttribute("aria-hidden", onLanding ? "true" : "false");
    }
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

  /** Show today’s date in the header (per landing page requirement). */
  function setHeaderLastUpdatedToday() {
    const hu = document.getElementById("header-updated");
    if (!hu) return;
    try {
      hu.textContent = new Date().toLocaleDateString(undefined, {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
      });
    } catch {
      hu.textContent = new Date().toISOString().slice(0, 10);
    }
  }

  function setShellLandingMode(on) {
    const shell = document.querySelector("[data-app-shell]");
    if (!shell) return;
    shell.classList.toggle("shell--landing", !!on);
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
    setHeaderLastUpdatedToday();
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
  if (!DATA.categories) DATA.categories = [];
  if (!DATA.kpiDetailByCategory) DATA.kpiDetailByCategory = [];
  if (!DATA.monthlyByCategory) DATA.monthlyByCategory = [];
  if (!DATA.businessBreakdown) DATA.businessBreakdown = [];

  /** Preview-only category 10 (not in all CSV exports); survives Refresh-PreviewData.ps1 regeneration. */
  function ensureLocationVulnerabilityCategory() {
    if (!DATA || !Array.isArray(DATA.categories)) return;
    if (!Array.isArray(DATA.kpiDetailByCategory)) return;
    const hasCat = DATA.categories.some((c) => c.categoryKey === 10);
    const hasDetail = DATA.kpiDetailByCategory.some(
      (x) => x.categoryKey === 10
    );
    if (hasCat && hasDetail) return;
    if (!hasCat) {
      DATA.categories.push({
        categoryKey: 10,
        categoryName: "Location Vulnerability",
        sortOrder: 10,
        uxNote:
          "Location-based exposure and vulnerability — use Comparison view for current vs prior period by all BUs.",
        kpiCount: 8,
        latestMonthIndex: 65.5,
      });
      DATA.categories.sort(
        (a, b) => (a.sortOrder || 0) - (b.sortOrder || 0)
      );
    }
    if (!hasDetail) {
      DATA.kpiDetailByCategory.push({
        categoryKey: 10,
        kpis: [
          {
            kpiKey: 13,
            kpiName: "Near miss exposure by location",
            unitType: "Count",
            latestValue: 0,
          },
          {
            kpiKey: 38,
            kpiName: "Vulnerability observations (sites)",
            unitType: "Count",
            latestValue: 0,
          },
          {
            kpiKey: 39,
            kpiName: "Location risk closure %",
            unitType: "PercentOrRate",
            latestValue: 0,
          },
          {
            kpiKey: 40,
            kpiName: "SRFA completion (location work)",
            unitType: "PercentOrRate",
            latestValue: 0,
          },
          {
            kpiKey: 45,
            kpiName: "Inspection closure % (locations)",
            unitType: "PercentOrRate",
            latestValue: 0,
          },
          {
            kpiKey: 46,
            kpiName: "Action closure % (locations)",
            unitType: "PercentOrRate",
            latestValue: 0,
          },
          {
            kpiKey: 53,
            kpiName: "Unsafe acts per hour (site rounds)",
            unitType: "Count",
            latestValue: 0,
          },
        ],
      });
    }
  }

  ensureLocationVulnerabilityCategory();

  /** Insights shell (`insights.html`): modern chrome + directory home; same logic as dashboard. */
  const insightShell =
    typeof document !== "undefined" &&
    document.querySelector('.shell--modern[data-layout="insights"]') != null;

  if (typeof Chart !== "undefined") {
    Chart.defaults.font.family = FONT_UI;
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

  /** Group A (high risk) + Group B (moderate); merged with any business names present in data. */
  const PREVIEW_BUSINESS_NAMES = [
    "Power",
    "Ports",
    "Cements",
    "Copper",
    "Green PVC",
    "Mining",
    "Gas",
    "Airports",
    "MSPVL",
    "Defense & Aerospace",
    "ANIL-Wind",
    "ANIL-BESS",
    "AESL",
    "Green Energy",
    "Realty",
    "RMRW",
    "AEML",
    "Data Center",
    "Logistics",
    "NMDPL",
    "Smart Meters",
    "Hydro PSP",
    "Agri fresh",
    "Dredging",
  ];
  const PREVIEW_BUSINESS_SET = new Set(PREVIEW_BUSINESS_NAMES);

  /**
   * Map dashboard BU labels to fact `businessName` values (Dim_Business / embedded rows).
   * Preview lists 24 labels; seed data only has 10 businesses — aliases avoid blank series.
   */
  const PREVIEW_TO_FACT_BUSINESS = {
    Ports: "Port",
    Cements: "Cement",
    "Green PVC": "AGEL",
    Mining: "Copper",
    Gas: "ATGL",
    Airports: "Hazira",
    MSPVL: "Logistics",
    "Defense & Aerospace": "Hazira",
    "ANIL-Wind": "AGEL",
    "ANIL-BESS": "AGEL",
    "Green Energy": "AGEL",
    Realty: "Hazira",
    RMRW: "Logistics",
    "Data Center": "AEML",
    NMDPL: "Port",
    "Smart Meters": "AEML",
    "Hydro PSP": "Power",
    "Agri fresh": "Hazira",
    Dredging: "Port",
  };

  function factBusinessNameForPreview(buName) {
    const s = String(buName || "").trim();
    return PREVIEW_TO_FACT_BUSINESS[s] != null ? PREVIEW_TO_FACT_BUSINESS[s] : s;
  }

  /** Fact rows use Dim names; dropdown uses preview labels — treat both as the same slice. */
  function rowMatchesBusinessFilter(r, f) {
    if (f.business === "all") return true;
    const rb = String(r.businessName || "").trim();
    const sel = String(f.business || "").trim();
    if (rb === sel) return true;
    return factBusinessNameForPreview(sel) === rb;
  }

  function mergedBusinessList(distinctFromData) {
    const extra = (distinctFromData || []).filter((x) => {
      if (!x) return false;
      return !PREVIEW_BUSINESS_SET.has(String(x).trim());
    });
    extra.sort((a, b) => a.localeCompare(b));
    return PREVIEW_BUSINESS_NAMES.concat(extra);
  }

  const meta = DATA.meta || {};

  /** Synthetic preview category (replaces former Workforce cat 5 in bootstrap). */
  const ASSURANCE_CATEGORY_KEY = 5;
  /** Incident-only additions (Dim keys 18 / 19 are used elsewhere — 56 / 57 here). */
  const INCIDENT_FIRE_KPI_KEY = 56;
  const INCIDENT_PROPERTY_DAMAGE_KPI_KEY = 57;

  /** Align packaged JSON meta with header product name (avoid editing embedded-data.js). */
  function previewApplyDataMetaTitles() {
    if (!DATA || !DATA.meta) return;
    DATA.meta.dashboardTitle = "Adani Safety Performance Profile";
    DATA.meta.subtitle =
      "Safety performance indicators — Interactive Preview";
  }

  /**
   * Incident Management: add Fire + Property Damage (19 KPIs total). Seeded fact rows.
   */
  function previewAddIncidentFirePropertyKpis() {
    if (!DATA || !Array.isArray(DATA.factRows)) return;
    const INCIDENT_CAT = 1;
    const detail = DATA.kpiDetailByCategory.find(
      (x) => x.categoryKey === INCIDENT_CAT
    );
    if (!detail || !Array.isArray(detail.kpis)) return;
    const have = new Set(detail.kpis.map((k) => Number(k.kpiKey)));
    if (have.has(INCIDENT_FIRE_KPI_KEY)) return;

    function ymToDateKey(ym) {
      return Number(String(ym).replace("-", "").slice(0, 6) + "01");
    }
    function seedValue(seedStr, min, max) {
      let h = 2166136261;
      const s = String(seedStr);
      for (let i = 0; i < s.length; i++) {
        h ^= s.charCodeAt(i);
        h = Math.imul(h, 16777619);
      }
      const u = (h >>> 0) / 4294967296;
      return min + u * (max - min);
    }

    const extraMeta = [
      {
        kpiKey: INCIDENT_FIRE_KPI_KEY,
        kpiName: "Fire",
        unitType: "Count",
        latestValue: 0,
      },
      {
        kpiKey: INCIDENT_PROPERTY_DAMAGE_KPI_KEY,
        kpiName: "Property Damage",
        unitType: "Count",
        latestValue: 0,
      },
    ];
    detail.kpis.push(extraMeta[0], extraMeta[1]);
    const cat = DATA.categories.find((c) => c.categoryKey === INCIDENT_CAT);
    if (cat) cat.kpiCount = detail.kpis.length;

    const monthKeys = (DATA.months || [])
      .map((m) => m.yearMonth)
      .filter(Boolean)
      .sort();
    if (!monthKeys.length) return;

    const businesses = [];
    const seen = new Set();
    for (let i = 0; i < DATA.factRows.length; i++) {
      const r = DATA.factRows[i];
      const b = String(r.businessName || "").trim();
      if (!b || seen.has(b)) continue;
      seen.add(b);
      businesses.push({
        businessName: b,
        businessKey: r.businessKey,
        state: String(r.state || "Maharashtra").trim() || "Maharashtra",
      });
      if (businesses.length >= 14) break;
    }
    if (!businesses.length) {
      businesses.push({
        businessName: "Power",
        businessKey: 1,
        state: "Madhya Pradesh",
      });
    }

    const newRows = [];
    for (let mi = 0; mi < monthKeys.length; mi++) {
      const ym = monthKeys[mi];
      const dk = ymToDateKey(ym);
      for (let bi = 0; bi < businesses.length; bi++) {
        const biz = businesses[bi];
        for (let ki = 0; ki < extraMeta.length; ki++) {
          const kpi = extraMeta[ki];
          const v = Math.round(
            seedValue(ym + "|" + biz.businessName + "|" + kpi.kpiKey, 0, 12)
          );
          newRows.push({
            yearMonth: ym,
            dateKey: dk,
            businessKey: biz.businessKey,
            businessName: biz.businessName,
            state: biz.state,
            kpiKey: kpi.kpiKey,
            kpiName: kpi.kpiName,
            categoryKey: INCIDENT_CAT,
            unitType: kpi.unitType,
            value: v,
            target: null,
          });
        }
      }
    }
    DATA.factRows.push(...newRows);
  }

  /**
   * Remove Risk Control Programs (7): S-4 / S-5 SRFA → Hazard (2); other KPIs → Consequence (4).
   * Idempotent if category 7 is already absent.
   */
  function previewStripRiskControlPrograms() {
    const RISK_CONTROL_CATEGORY_KEY = 7;
    const HAZARD_CAT = 2;
    const CONSEQUENCE_CAT = 4;
    const KPIS_TO_HAZARD = new Set([41, 42]);
    const KPIS_TO_CONSEQUENCE = new Set([51, 52, 54]);
    if (
      !DATA ||
      !Array.isArray(DATA.categories) ||
      !DATA.categories.some((c) => c.categoryKey === RISK_CONTROL_CATEGORY_KEY)
    ) {
      return;
    }
    const riskDetail = DATA.kpiDetailByCategory.find(
      (x) => x.categoryKey === RISK_CONTROL_CATEGORY_KEY
    );
    DATA.categories = DATA.categories.filter(
      (c) => c.categoryKey !== RISK_CONTROL_CATEGORY_KEY
    );
    DATA.kpiDetailByCategory = DATA.kpiDetailByCategory.filter(
      (x) => x.categoryKey !== RISK_CONTROL_CATEGORY_KEY
    );
    if (DATA.monthlyByCategory) {
      DATA.monthlyByCategory = DATA.monthlyByCategory.filter(
        (x) => x.categoryKey !== RISK_CONTROL_CATEGORY_KEY
      );
    }
    if (DATA.businessBreakdown) {
      DATA.businessBreakdown = DATA.businessBreakdown.filter(
        (x) => x.categoryKey !== RISK_CONTROL_CATEGORY_KEY
      );
    }
    if (riskDetail && Array.isArray(riskDetail.kpis)) {
      const hazardBlock = DATA.kpiDetailByCategory.find(
        (x) => x.categoryKey === HAZARD_CAT
      );
      const consBlock = DATA.kpiDetailByCategory.find(
        (x) => x.categoryKey === CONSEQUENCE_CAT
      );
      const seenH = new Set(
        (hazardBlock && hazardBlock.kpis
          ? hazardBlock.kpis
          : []
        ).map((k) => k.kpiKey)
      );
      const seenC = new Set(
        (consBlock && consBlock.kpis ? consBlock.kpis : []).map((k) => k.kpiKey)
      );
      for (let i = 0; i < riskDetail.kpis.length; i++) {
        const k = riskDetail.kpis[i];
        if (KPIS_TO_HAZARD.has(k.kpiKey) && hazardBlock && !seenH.has(k.kpiKey)) {
          hazardBlock.kpis.push(k);
          seenH.add(k.kpiKey);
        } else if (
          KPIS_TO_CONSEQUENCE.has(k.kpiKey) &&
          consBlock &&
          !seenC.has(k.kpiKey)
        ) {
          consBlock.kpis.push(k);
          seenC.add(k.kpiKey);
        }
      }
    }
    if (Array.isArray(DATA.factRows)) {
      for (let i = 0; i < DATA.factRows.length; i++) {
        const r = DATA.factRows[i];
        if (r.categoryKey !== RISK_CONTROL_CATEGORY_KEY) continue;
        if (KPIS_TO_HAZARD.has(r.kpiKey)) r.categoryKey = HAZARD_CAT;
        else if (KPIS_TO_CONSEQUENCE.has(r.kpiKey)) r.categoryKey = CONSEQUENCE_CAT;
      }
    }
    function syncKpiCount(catKey) {
      const cat = DATA.categories.find((c) => c.categoryKey === catKey);
      const det = DATA.kpiDetailByCategory.find((x) => x.categoryKey === catKey);
      if (cat && det && Array.isArray(det.kpis)) {
        cat.kpiCount = det.kpis.length;
      }
    }
    syncKpiCount(HAZARD_CAT);
    syncKpiCount(CONSEQUENCE_CAT);
  }

  /**
   * Remove Workforce Exposure & Manhour Base (5); add Assurance (5) with seeded facts. Other categories unchanged.
   */
  function previewCategoryDataBootstrap() {
    if (!DATA || !Array.isArray(DATA.factRows)) return;

    function stripCategory(k) {
      DATA.categories = (DATA.categories || []).filter((c) => c.categoryKey !== k);
      DATA.factRows = (DATA.factRows || []).filter((r) => r.categoryKey !== k);
      if (DATA.monthlyByCategory) {
        DATA.monthlyByCategory = DATA.monthlyByCategory.filter(
          (x) => x.categoryKey !== k
        );
      }
      if (DATA.kpiDetailByCategory) {
        DATA.kpiDetailByCategory = DATA.kpiDetailByCategory.filter(
          (x) => x.categoryKey !== k
        );
      }
      if (DATA.businessBreakdown) {
        DATA.businessBreakdown = DATA.businessBreakdown.filter(
          (x) => x.categoryKey !== k
        );
      }
    }

    function ymToDateKey(ym) {
      return Number(String(ym).replace("-", "").slice(0, 6) + "01");
    }

    function seedValue(seedStr, min, max) {
      let h = 2166136261;
      const s = String(seedStr);
      for (let i = 0; i < s.length; i++) {
        h ^= s.charCodeAt(i);
        h = Math.imul(h, 16777619);
      }
      const u = (h >>> 0) / 4294967296;
      return min + u * (max - min);
    }

    stripCategory(ASSURANCE_CATEGORY_KEY);

    const monthKeys = (DATA.months || [])
      .map((m) => m.yearMonth)
      .filter(Boolean)
      .sort();
    if (!monthKeys.length) return;

    const businesses = [];
    const seen = new Set();
    for (let i = 0; i < DATA.factRows.length; i++) {
      const r = DATA.factRows[i];
      const b = String(r.businessName || "").trim();
      if (!b || seen.has(b)) continue;
      seen.add(b);
      businesses.push({
        businessName: b,
        businessKey: r.businessKey,
        state: String(r.state || "Maharashtra").trim() || "Maharashtra",
      });
      if (businesses.length >= 14) break;
    }
    if (!businesses.length) {
      businesses.push({
        businessName: "Power",
        businessKey: 1,
        state: "Madhya Pradesh",
      });
    }

    const assuranceKpis = [
      {
        kpiKey: 501,
        kpiName: "FRC Compliance Rate %",
        unitType: "PercentOrRate",
        latestValue: 88,
      },
      {
        kpiKey: 502,
        kpiName: "Standard Implementation %",
        unitType: "PercentOrRate",
        latestValue: 72,
      },
      {
        kpiKey: 503,
        kpiName: "Key Learning Implementation %",
        unitType: "PercentOrRate",
        latestValue: 81,
      },
    ];
    DATA.categories.push({
      categoryKey: ASSURANCE_CATEGORY_KEY,
      categoryName: "Assurance",
      sortOrder: 5,
      uxNote:
        "Governance assurance visits, audits, and critical finding closure.",
      kpiCount: assuranceKpis.length,
      latestMonthIndex: 1,
    });
    DATA.categories.sort(
      (a, b) => (a.sortOrder || 0) - (b.sortOrder || 0)
    );

    DATA.kpiDetailByCategory.push({
      categoryKey: ASSURANCE_CATEGORY_KEY,
      kpis: assuranceKpis,
    });

    const newRows = [];
    for (let mi = 0; mi < monthKeys.length; mi++) {
      const ym = monthKeys[mi];
      const dk = ymToDateKey(ym);
      for (let bi = 0; bi < businesses.length; bi++) {
        const biz = businesses[bi];
        for (let ki = 0; ki < assuranceKpis.length; ki++) {
          const kpi = assuranceKpis[ki];
          let v;
          if (kpi.unitType === "Count") {
            v = Math.round(
              seedValue(ym + "|" + biz.businessName + "|" + kpi.kpiKey, 8, 220)
            );
          } else {
            v = seedValue(
              ym + "|" + biz.businessName + "|" + kpi.kpiKey,
              55,
              98
            );
          }
          newRows.push({
            yearMonth: ym,
            dateKey: dk,
            businessKey: biz.businessKey,
            businessName: biz.businessName,
            state: biz.state,
            kpiKey: kpi.kpiKey,
            kpiName: kpi.kpiName,
            categoryKey: ASSURANCE_CATEGORY_KEY,
            unitType: kpi.unitType,
            value: v,
            target: null,
          });
        }
      }
    }
    DATA.factRows.push(...newRows);

    function rollupSeries(catKey, sampleKpiKey) {
      return monthKeys.map((ym) => {
        const nums = newRows
          .filter(
            (r) =>
              r.yearMonth === ym &&
              r.categoryKey === catKey &&
              r.kpiKey === sampleKpiKey
          )
          .map((r) => Number(r.value));
        const v = nums.length
          ? nums.reduce((a, b) => a + b, 0) / nums.length
          : 0;
        return { yearMonth: ym, value: Math.round(v * 100) / 100 };
      });
    }

    DATA.monthlyByCategory.push({
      categoryKey: ASSURANCE_CATEGORY_KEY,
      series: rollupSeries(ASSURANCE_CATEGORY_KEY, 501),
    });
  }

  previewApplyDataMetaTitles();
  previewStripRiskControlPrograms();
  previewAddIncidentFirePropertyKpis();
  previewCategoryDataBootstrap();

  /** Detail table rows per page (keep in sync with styles.css --detail-table-body-rows) */
  const PAGE_SIZE = 5;
  let tableState = {
    sortKey: "yearMonth",
    asc: false,
    page: 0,
  };

  let currentCategoryKey = null;
  let catSearchAnnounceTimer = null;

  const SPI_CATEGORY_KEY = 3;
  const HAZARD_CATEGORY_KEY = 2;
  /**
   * Consequence Management — stacked CMP accountability chart replaces the
   * trend line when primary KPI is LTI vs CMP (24) or Total CMP action Taken (25).
   */
  const CONSEQUENCE_MANAGEMENT_CATEGORY_KEY = 4;
  const CMP_ACCOUNTABILITY_CHART_PRIMARY_KPI_KEYS = new Set(["24", "25"]);
  const TRAINING_CATEGORY_KEY = 6;
  const SYSTEMS_ADOPTION_CATEGORY_KEY = 9;
  const LEADERSHIP_NOT_IN_PREVIEW_CATEGORY_KEY = 8;
  const CATEGORY_DISABLED_NOT_IN_PREVIEW_KEYS = new Set([
    LEADERSHIP_NOT_IN_PREVIEW_CATEGORY_KEY,
  ]);
  /** Percent speedometer (Adani_Safety_MIS_Theme.json: bad → good / teal). */
  const SPEEDOMETER_ARC_HEX = [
    "#DC2626",
    "#EA580C",
    "#F59E0B",
    "#00A3A3",
    "#0D9488",
  ];
  /** Assurance — all three preview KPIs use the speedometer. */
  const ASSURANCE_SPEEDOMETER_KPI_KEYS = new Set(["501", "502", "503"]);
  const TRAINING_SAKSHAM_SPEEDOMETER_KPI_KEY = "49";
  const SYSTEMS_SAFEX_SPEEDOMETER_KPI_KEY = "50";
  const LOCATION_VULN_CAT_KEY = 10;
  const LOCATION_VULN_SOURCE_CAT = 2;
  const LOCATION_VULN_KPI_KEYS = new Set([13, 38, 39, 40, 45, 46, 53]);
  const EVENT_LEVEL_LABELS = [
    "0 Near Miss",
    "1 Minor",
    "2 Moderate",
    "3 Serious",
    "4 Major",
    "5 Catastrophic",
    "No Level",
  ];
  /** Dim_KPI category 3 order — TRI first. */
  const SPI_KPI_ORDER = [21, 16, 17, 18, 20, 29];
  /** Risk quadrant bubble: X ≈ concern / reporting culture (Near Miss FR), Y = Fatality rate (%). */
  const SPI_QUADRANT_X_KPI = 20;
  const SPI_QUADRANT_Y_KPI = 16;
  /** Approximate centroids for India states/UTs (preview map). */
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
    /** J&K UT — ~between Jammu and Srinagar (lng was wrongly ~76.6°E, which sits in Himachal). */
    "Jammu and Kashmir": [33.85, 74.78],
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

  /**
   * Lat/lng corners to frame all of India on the preview map (incl. Jammu & Kashmir, NE states,
   * southern tip). BU markers alone would otherwise fitBounds too tight and crop the north.
   */
  const INDIA_MAP_BOUNDS_SW = [6.2, 68.0];
  const INDIA_MAP_BOUNDS_NE = [37.6, 97.6];

  /** Dim_Business home state — anchor BU markers on the India map (preview). */
  const FACT_BU_HOME_STATE = {
    AEML: "Gujarat",
    AESL: "Maharashtra",
    AGEL: "Karnataka",
    ATGL: "Gujarat",
    Cement: "Rajasthan",
    Copper: "Chhattisgarh",
    Port: "Gujarat",
    Power: "Madhya Pradesh",
    Logistics: "Haryana",
    Hazira: "Gujarat",
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

  /** Deterministic event-severity bucket per fact row (preview — replace with model field when available). */
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

  /** Info (i) hover text — follows the Vs filter selection in the toolbar. */
  function vsFilterComparisonTooltip(vsMode) {
    const m = vsMode || DEFAULT_VS_MODE;
    switch (m) {
      case "vs_yesterday":
        return "Today vs Yesterday";
      case "vs_last_week":
        return "Current week vs Last Week";
      case "vs_last_month":
        return "Current Month vs Last Month";
      case "vs_last_quarter":
        return "Current Quarter vs Last Quarter";
      case "vs_last_year":
        return "Current Year vs Last Year";
      default:
        return VS_PERIOD_CAPTION[m] || vsOptionLabel(m);
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

  const EXPORT_DL_ICON_SVG =
    '<svg class="export-dl-svg" width="16" height="16" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path fill="currentColor" d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>';

  const EXPORT_TABLE_CSV_SVG =
    '<svg class="export-dl-svg" width="16" height="16" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path fill="currentColor" d="M3 3h18v18H3V3zm2 2v5h5V5H5zm7 0v5h5V5h-5zm7 0v5h5V5h-5zM5 12v5h5v-5H5zm7 0v5h5v-5h-5zm7 0v5h5v-5h-5z"/></svg>';

  function chartDownloadButton(dataId, slug) {
    return (
      '<button type="button" class="chart-export-btn" data-chart-export="' +
      escapeAttr(dataId) +
      '" data-export-slug="' +
      escapeAttr(slug) +
      '" title="Download chart (JPEG)" aria-label="Download this chart as JPEG">' +
      EXPORT_DL_ICON_SVG +
      "</button>"
    );
  }

  function chartBlockTitleHtml(innerSpansHtml, dataId, slug) {
    return (
      '<h3 class="chart-analytics-title chart-analytics-title--toolbar">' +
      '<span class="chart-analytics-title__text">' +
      innerSpansHtml +
      "</span>" +
      chartDownloadButton(dataId, slug) +
      "</h3>"
    );
  }

  const LS_VARIABLE_FILTER = "mis-variable-checkpoints";
  const LS_KPI_PREFIX = insightShell
    ? "insights-kpi-keys-"
    : "preview-kpi-keys-";
  const LS_CAT_MAIN_VIEW = insightShell
    ? "insights_cat_main_view"
    : "adani_cat_main_view";

  const TREND_LINE_COLORS = [
    "#006DB6",
    "#00B16B",
    "#E87722",
    "#6B2D90",
    "#C41230",
    "#007A87",
    "#5C6BC0",
  ];

  function hexToRgba(hex, alpha) {
    const h = String(hex || "").replace("#", "");
    if (h.length !== 6) return "rgba(0, 109, 182, " + alpha + ")";
    const r = parseInt(h.slice(0, 2), 16);
    const g = parseInt(h.slice(2, 4), 16);
    const b = parseInt(h.slice(4, 6), 16);
    return "rgba(" + r + "," + g + "," + b + "," + alpha + ")";
  }

  /** Default KPI checkbox seed when localStorage is empty (aligns with insights: incident = first five TRI block). */
  function defaultPreviewKpiSeed(catKey, kpisMeta) {
    if (catKey === 1 && kpisMeta.length) {
      const want = ["1", "2", "3", "4", "5"];
      const valid = new Set(kpisMeta.map((k) => String(k.kpiKey)));
      const kept = want.filter((id) => valid.has(id));
      if (kept.length) return kept;
    }
    const dk = defaultKpiKeyForCategory(catKey, kpisMeta);
    return dk ? [String(dk)] : [];
  }

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
    return defaultPreviewKpiSeed(catKey, kpisMeta);
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

  /** KPI scope: checkboxes in display order (same ids as insights for consistency). */
  function kpiScopePanelHtml(catKey, kpisMeta) {
    const sel = loadKpiSelection(catKey, kpisMeta);
    const boxes = kpisMeta
      .map((k) => {
        const id = "f-kpi-cb-" + k.kpiKey;
        const checked = sel.includes(String(k.kpiKey)) ? " checked" : "";
        return (
          '<label class="kpi-cb" for="' +
          id +
          '">' +
          '<input type="checkbox" name="f-kpi-cb" id="' +
          id +
          '" value="' +
          k.kpiKey +
          '"' +
          checked +
          "/>" +
          '<span class="kpi-cb__text">' +
          escapeHtml(kpiDropdownLabel(k)) +
          "</span></label>"
        );
      })
      .join("");
    return (
      '<details class="kpi-scope kpi-scope--toolbar">' +
      '<summary class="kpi-scope__summary" id="f-kpi-scope-label">' +
      '<span class="kpi-scope__summary-text">' +
      '<span class="kpi-scope__title">KPI list</span>' +
      '<span class="kpi-scope__count" id="f-kpi-count"></span>' +
      "</span></summary>" +
      '<div class="kpi-panel" id="f-kpi-panel" role="group" aria-labelledby="f-kpi-scope-label">' +
      '<div class="kpi-panel__bar">' +
      '<button type="button" class="btn kpi-panel__btn" id="f-kpi-all">All</button>' +
      '<button type="button" class="btn kpi-panel__btn" id="f-kpi-none">None</button>' +
      "</div>" +
      '<div class="kpi-panel__list">' +
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

  /** Single KPI in toolbar + persist; refreshes charts/table (used by tiles & SPI trend chart). */
  function applyToolbarSingleKpiSelection(catKey, kpiKey) {
    const panel = document.getElementById("f-kpi-panel");
    if (!panel || catKey == null || kpiKey == null) return;
    panel.querySelectorAll('input[name="f-kpi-cb"]').forEach((cb) => {
      cb.checked = String(cb.value) === String(kpiKey);
    });
    saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
    updateKpiScopeCount();
    tableState.page = 0;
    refreshCategoryView(catKey);
  }

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
      '<span class="field-label" id="f-var-lbl">Vertical</span>' +
      '<details class="var-scope var-scope--toolbar" id="f-var-details">' +
      '<summary class="var-scope__summary" aria-labelledby="f-var-lbl" title="Vertical (checkpoints)">' +
      '<span class="var-scope__summary-text">' +
      '<span class="var-scope__hint" id="f-var-hint">All verticals</span>' +
      "</span>" +
      '<span class="var-scope__chev" aria-hidden="true"></span>' +
      "</summary>" +
      '<div class="var-scope__panel" role="group" aria-label="Vertical filter options">' +
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
      }
      saveVariableFilterToStorage();
      updateVariableSummary();
      if (onChange) onChange();
    }
    all.addEventListener("change", allChange);
    cbs.forEach((cb) => cb.addEventListener("change", subChange));
    updateVariableSummary();
  }

  /**
   * Filters share one horizontal scroll strip; <details> panels use fixed
   * positioning while open so overflow-x on the strip does not clip them.
   */
  function wireToolbarScopeScrollPanels(host) {
    const scope = host.querySelector(".cat-toolbar__filters-all-scroll");
    if (!scope) return;
    const pairs = [];
    [
      ["details.var-scope--toolbar", ".var-scope__panel"],
      ["details.kpi-scope--toolbar", ".kpi-panel"],
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
          '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>'
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
      case 10:
        return svg(
          '<path d="M12 21s-6-5.2-6-10a6 6 0 1112 0c0 4.8-6 10-6 10z"/><circle cx="12" cy="11" r="2.5"/>'
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
    const list = k ? k.kpis.slice() : [];
    if (catKey === SPI_CATEGORY_KEY) return list;
    return list.filter((kpi) => !isTriKpiMeta(kpi));
  }

  /** Preferred default KPI in dropdown (TRI / TRIR); fallback: first KPI in display order. */
  const TRI_LABEL_FULL = "Total Recordable Incident Rate (TRI)";

  function defaultKpiKeyForCategory(catKey, kpisMetaForUi) {
    const merged =
      kpisMetaForUi && kpisMetaForUi.length
        ? kpisMetaForUi
        : kpiListForFilterDropdown(catKey);
    if (catKey === HAZARD_CATEGORY_KEY) {
      const sorted = sortKpisForDisplay(catKey, merged);
      return sorted.length ? String(sorted[0].kpiKey) : "38";
    }
    if (catKey === SPI_CATEGORY_KEY) {
      const tri = merged.find((k) => isTriKpiMeta(k));
      if (tri) return String(tri.kpiKey);
    }
    const sorted = sortKpisForDisplay(catKey, merged);
    return sorted.length ? String(sorted[0].kpiKey) : "1";
  }

  function kpiDropdownLabel(k) {
    const name = String(k.kpiName || "").trim();
    if (/^TRIR$/i.test(name) || String(k.kpiKey) === "21") return TRI_LABEL_FULL;
    return k.kpiName;
  }

  function isTriKpiMeta(k) {
    return (
      String(k.kpiKey) === "21" ||
      /Total Recordable|TRIR|\bTRI\b|Recordable Incident Rate/i.test(
        k.kpiName || ""
      )
    );
  }

  /** KPI dropdown: TRIR / TRI only under Safety Performance Indices (category 3). */
  function kpiListForFilterDropdown(catKey) {
    let base = getKpis(catKey);
    if (catKey === SPI_CATEGORY_KEY && !base.some(isTriKpiMeta)) {
      base = [
        {
          kpiKey: 21,
          kpiName: "TRIR",
          unitType: "PercentOrRate",
          latestValue: null,
        },
      ].concat(base);
    }
    return sortKpisForDisplay(catKey, base);
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
    const kpis = getKpis(catKey);
    const k = kpis.find((x) => String(x.kpiKey) === String(f.kpi));
    return k ? k.unitType : null;
  }

  function destroySpiLeafletMap() {
    const spiInsightEl = document.getElementById("chart-spi-insights");
    if (spiInsightEl && typeof Chart !== "undefined") {
      const ch = Chart.getChart(spiInsightEl);
      if (ch) ch.destroy();
    }
    if (window.__adaniSpiRibbonResize) {
      try {
        window.removeEventListener("resize", window.__adaniSpiRibbonResize);
      } catch {
        /* ignore */
      }
      window.__adaniSpiRibbonResize = null;
    }
    window.__adaniSpiRibbonRedraw = null;
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

  function destroyLocBuLeafletMap() {
    if (window.__adaniLocBuMap) {
      try {
        window.__adaniLocBuMap.remove();
      } catch {
        /* ignore */
      }
      window.__adaniLocBuMap = null;
    }
    const host = document.getElementById("chart-loc-bu-map");
    if (host) host.innerHTML = "";
  }

  /** Single-flight loader for Google Maps JS + Leaflet.GridLayer.GoogleMutant (India BU map). */
  var __adaniGoogleMapsLoadPromise = null;

  function getGoogleMapsApiKey() {
    try {
      if (typeof window.__GOOGLE_MAPS_API_KEY__ === "string") {
        const wk = window.__GOOGLE_MAPS_API_KEY__.trim();
        if (wk) return wk;
      }
      const meta = document.querySelector('meta[name="google-maps-api-key"]');
      if (meta) {
        const mk = String(meta.getAttribute("content") || "").trim();
        if (mk) return mk;
      }
    } catch {
      /* ignore */
    }
    return "";
  }

  function loadGoogleMapsAndMutant() {
    if (
      typeof L !== "undefined" &&
      L.gridLayer &&
      L.gridLayer.googleMutant &&
      typeof window.google !== "undefined" &&
      window.google.maps
    ) {
      return Promise.resolve(true);
    }
    const key = getGoogleMapsApiKey();
    if (!key) return Promise.resolve(false);
    if (__adaniGoogleMapsLoadPromise) return __adaniGoogleMapsLoadPromise;

    __adaniGoogleMapsLoadPromise = new Promise(function (resolve) {
      function finishFail() {
        __adaniGoogleMapsLoadPromise = null;
        resolve(false);
      }

      function attachMutant() {
        if (L.gridLayer && L.gridLayer.googleMutant) {
          resolve(true);
          return;
        }
        const s2 = document.createElement("script");
        s2.src =
          "https://unpkg.com/leaflet.gridlayer.googlemutant@0.14.1/dist/Leaflet.GoogleMutant.js";
        s2.crossOrigin = "";
        s2.onload = function () {
          if (L.gridLayer && L.gridLayer.googleMutant) resolve(true);
          else finishFail();
        };
        s2.onerror = finishFail;
        document.head.appendChild(s2);
      }

      if (typeof window.google !== "undefined" && window.google.maps) {
        attachMutant();
        return;
      }

      const cbName = "__adaniGmapsApiReady";
      window[cbName] = function () {
        try {
          delete window[cbName];
        } catch {
          window[cbName] = undefined;
        }
        attachMutant();
      };

      const s = document.createElement("script");
      s.src =
        "https://maps.googleapis.com/maps/api/js?key=" +
        encodeURIComponent(key) +
        "&callback=" +
        cbName;
      s.async = true;
      s.defer = true;
      s.onerror = function () {
        try {
          delete window[cbName];
        } catch {
          window[cbName] = undefined;
        }
        finishFail();
      };
      document.head.appendChild(s);
    });

    return __adaniGoogleMapsLoadPromise;
  }

  function destroyCharts() {
    [
      "chart-line",
      "chart-verticals",
      "chart-biz",
      "chart-spi-bubble",
      "chart-spi-insights",
      "chart-bu-compare",
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.destroy();
      }
    });
    destroySpiLeafletMap();
    destroyLocBuLeafletMap();
    const spiHmHost = document.getElementById("spi-hazard-heatmap-host");
    if (spiHmHost) spiHmHost.innerHTML = "";
    try {
      window.__adaniSpeedometerGaugeRedraw = null;
    } catch {
      /* ignore */
    }
    const lineGaugeEl = document.getElementById("chart-line");
    if (
      lineGaugeEl &&
      lineGaugeEl.getAttribute("data-speedometer-gauge") === "1"
    ) {
      lineGaugeEl.removeAttribute("data-speedometer-gauge");
      const gctx = lineGaugeEl.getContext("2d");
      if (gctx) {
        gctx.clearRect(0, 0, lineGaugeEl.width, lineGaugeEl.height);
      }
    }
  }

  function resizeSpiMapIfAny() {
    const spiEl = document.getElementById("chart-spi-insights");
    if (spiEl && typeof Chart !== "undefined") {
      const c = Chart.getChart(spiEl);
      if (c) {
        try {
          c.resize();
        } catch {
          /* ignore */
        }
        return;
      }
    }
    if (typeof window.__adaniSpiRibbonRedraw === "function") {
      try {
        window.__adaniSpiRibbonRedraw();
      } catch {
        /* ignore */
      }
      return;
    }
    if (window.__adaniSpiMap && typeof window.__adaniSpiMap.invalidateSize === "function") {
      try {
        window.__adaniSpiMap.invalidateSize();
      } catch {
        /* ignore */
      }
    }
  }

  function resizeLocBuMapIfAny() {
    if (window.__adaniLocBuMap && typeof window.__adaniLocBuMap.invalidateSize === "function") {
      try {
        window.__adaniLocBuMap.invalidateSize();
      } catch {
        /* ignore */
      }
    }
  }

  function resizeAllChartsIndex() {
    [
      "chart-line",
      "chart-verticals",
      "chart-biz",
      "chart-spi-bubble",
      "chart-spi-insights",
      "chart-bu-compare",
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.resize();
      }
    });
    resizeSpiMapIfAny();
    resizeLocBuMapIfAny();
    if (typeof window.__adaniSpeedometerGaugeRedraw === "function") {
      try {
        window.__adaniSpeedometerGaugeRedraw();
      } catch {
        /* ignore */
      }
    }
  }

  const spiQuadrantZonePlugin = {
    id: "spiQuadrantZone",
    beforeDatasetsDraw(chart) {
      const mid = chart.options.plugins && chart.options.plugins.spiQuadrantMid;
      const sx = chart.scales.x;
      const sy = chart.scales.y;
      if (!mid || mid.x == null || mid.y == null || !sx || !sy) return;
      const { ctx, chartArea } = chart;
      const px = sx.getPixelForValue(mid.x);
      const py = sy.getPixelForValue(mid.y);
      const L = chartArea.left;
      const R = chartArea.right;
      const T = chartArea.top;
      const B = chartArea.bottom;
      const zones = [
        {
          x: L,
          y: T,
          w: px - L,
          h: py - T,
          fill: "rgba(254, 240, 138, 0.5)",
        },
        {
          x: px,
          y: T,
          w: R - px,
          h: py - T,
          fill: "rgba(253, 186, 116, 0.32)",
        },
        {
          x: L,
          y: py,
          w: px - L,
          h: B - py,
          fill: "rgba(229, 231, 235, 0.4)",
        },
        {
          x: px,
          y: py,
          w: R - px,
          h: B - py,
          fill: "rgba(187, 247, 208, 0.5)",
        },
      ];
      for (let i = 0; i < zones.length; i++) {
        const z = zones[i];
        ctx.save();
        ctx.fillStyle = z.fill;
        ctx.fillRect(z.x, z.y, z.w, z.h);
        ctx.restore();
      }
    },
  };

  const spiQuadrantLinePlugin = {
    id: "spiQuadrantLine",
    afterDatasetsDraw(chart) {
      const mid = chart.options.plugins && chart.options.plugins.spiQuadrantMid;
      const sx = chart.scales.x;
      const sy = chart.scales.y;
      if (!mid || mid.x == null || mid.y == null || !sx || !sy) return;
      const { ctx, chartArea } = chart;
      const px = sx.getPixelForValue(mid.x);
      const py = sy.getPixelForValue(mid.y);
      const L = chartArea.left;
      const R = chartArea.right;
      const T = chartArea.top;
      const B = chartArea.bottom;
      ctx.save();
      ctx.strokeStyle = "rgba(35, 31, 32, 0.55)";
      ctx.lineWidth = 1.25;
      ctx.setLineDash([5, 4]);
      ctx.beginPath();
      ctx.moveTo(px, T);
      ctx.lineTo(px, B);
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(L, py);
      ctx.lineTo(R, py);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = "rgba(35, 31, 32, 0.75)";
      ctx.font = "600 8px " + FONT_UI;
      ctx.textAlign = "left";
      ctx.fillText("Group avg", Math.min(px + 4, R - 52), T + 10);
      ctx.restore();
    },
  };

  /**
   * SPI cross-plot: Concern-style reporting (Near Miss FR %) vs Fatality rate (%),
   * bubble size ∝ SPI row volume for that BU; quadrants split at group averages.
   */
  function renderSpiRiskQuadrantChart(snapPool, f) {
    const el = document.getElementById("chart-spi-bubble");
    if (!el || typeof Chart === "undefined") return;
    const prev = Chart.getChart(el);
    if (prev) prev.destroy();

    const winMonths = new Set(
      bizUnitWindowMonths(f.vsMode || DEFAULT_VS_MODE, f.refMonth)
    );
    const rows = snapPool.filter((r) => winMonths.has(r.yearMonth));

    const BU_COLORS = [
      "rgba(0, 109, 182, 0.55)",
      "rgba(0, 177, 107, 0.55)",
      "rgba(142, 39, 143, 0.5)",
      "rgba(240, 76, 35, 0.5)",
      "rgba(59, 130, 246, 0.5)",
      "rgba(234, 179, 8, 0.55)",
      "rgba(99, 102, 241, 0.5)",
      "rgba(236, 72, 153, 0.45)",
    ];
    const BU_BORDERS = [
      "#006DB6",
      "#00a86b",
      "#8E278F",
      "#F04C23",
      "#2563eb",
      "#ca8a04",
      "#4f46e5",
      "#db2777",
    ];

    const points = [];
    PREVIEW_BUSINESS_NAMES.forEach((previewBu, bi) => {
      const buFact = factBusinessNameForPreview(previewBu);
      const sub = rows.filter(
        (r) => String(r.businessName || "").trim() === buFact
      );
      if (!sub.length) return;
      const xRows = sub.filter(
        (r) => String(r.kpiKey) === String(SPI_QUADRANT_X_KPI)
      );
      const yRows = sub.filter(
        (r) => String(r.kpiKey) === String(SPI_QUADRANT_Y_KPI)
      );
      const xVals = xRows
        .map((r) => Number(r.value))
        .filter((v) => !Number.isNaN(v));
      const yVals = yRows
        .map((r) => Number(r.value))
        .filter((v) => !Number.isNaN(v));
      const x = xVals.length ? avg(xVals) : null;
      const y = yVals.length ? avg(yVals) : null;
      if (x == null || y == null) return;
      const n = sub.length;
      const rPx = Math.max(10, Math.min(38, 6 + Math.sqrt(n) * 2.2));
      points.push({
        x,
        y,
        r: rPx,
        bu: previewBu,
        n,
        bg: BU_COLORS[bi % BU_COLORS.length],
        bd: BU_BORDERS[bi % BU_BORDERS.length],
      });
    });

    if (!points.length) {
      points.push({
        x: 2.5,
        y: 0.12,
        r: 18,
        bu: "Preview (no BU slice)",
        n: 0,
        bg: "rgba(0, 109, 182, 0.35)",
        bd: "#006DB6",
      });
    }

    const xs = points.map((p) => p.x);
    const ys = points.map((p) => p.y);
    const midX = avg(xs);
    const midY = avg(ys);
    const padX = Math.max(0.4, (Math.max(...xs) - Math.min(...xs)) * 0.12 || 0.5);
    const padY = Math.max(0.02, (Math.max(...ys) - Math.min(...ys)) * 0.15 || 0.03);

    new Chart(el, {
      type: "bubble",
      plugins: [spiQuadrantZonePlugin, spiQuadrantLinePlugin],
      data: {
        datasets: [
          {
            label: "Business units",
            data: points.map((p) => ({
              x: p.x,
              y: p.y,
              r: p.r,
              bu: p.bu,
              n: p.n,
            })),
            backgroundColor: points.map((p) => p.bg),
            borderColor: points.map((p) => p.bd),
            borderWidth: 1.5,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 14, right: 14, bottom: 10, left: 10 } },
        plugins: {
          legend: { display: false },
          spiQuadrantMid: { x: midX, y: midY },
          tooltip: {
            callbacks: {
              label(ctx) {
                const raw = ctx.raw;
                const bu = raw.bu || "—";
                const n = raw.n != null ? raw.n : 0;
                return (
                  " " +
                  bu +
                  " · Concern rate " +
                  Number(raw.x).toFixed(2) +
                  "% · Fatality " +
                  Number(raw.y).toFixed(3) +
                  "% · Rows " +
                  n
                );
              },
            },
          },
        },
        scales: {
          x: {
            type: "linear",
            suggestedMin: Math.max(0, Math.min(...xs) - padX),
            suggestedMax: Math.max(...xs) + padX,
            ticks: { font: { size: 9 }, color: "#231F20" },
            title: {
              display: true,
              text: "Concern reporting rate (Near Miss FR %)",
              font: { size: 10 },
              color: "#6D6E71",
            },
            grid: { color: "rgba(109, 110, 113, 0.14)" },
          },
          y: {
            type: "linear",
            beginAtZero: true,
            suggestedMax: Math.max(...ys) + padY,
            ticks: { font: { size: 9 }, color: "#231F20" },
            title: {
              display: true,
              text: "Fatality rate (%)",
              font: { size: 10 },
              color: "#6D6E71",
            },
            grid: { color: "rgba(109, 110, 113, 0.2)" },
          },
        },
      },
    });
    el.setAttribute(
      "aria-label",
      "Bubble quadrant chart: concern reporting rate versus fatality rate by business unit; bubble size reflects filtered SPI row volume; yellow top-left is higher risk, green bottom-right is lower risk; dashed lines are group averages."
    );
  }

  const SPI_INSIGHT_PALETTE = [
    { fill: "rgba(0, 109, 182, 0.38)", stroke: "#006DB6" },
    { fill: "rgba(0, 177, 107, 0.42)", stroke: "#00a86b" },
    { fill: "rgba(142, 39, 143, 0.38)", stroke: "#8E278F" },
    { fill: "rgba(240, 76, 35, 0.38)", stroke: "#F04C23" },
    { fill: "rgba(35, 31, 32, 0.22)", stroke: "#231F20" },
    { fill: "rgba(0, 109, 182, 0.22)", stroke: "#005a94" },
    { fill: "rgba(0, 177, 107, 0.26)", stroke: "#008f56" },
    { fill: "rgba(142, 39, 143, 0.24)", stroke: "#6a1f6b" },
  ];

  function spiChartMonthTick(ym) {
    if (!ym || ym.length < 7) return String(ym || "");
    const mons = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];
    const mo = Number(ym.slice(5, 7));
    if (Number.isNaN(mo) || mo < 1 || mo > 12) return ym;
    return mons[mo - 1] + " '" + ym.slice(2, 4);
  }

  /**
   * SPI: 100% stacked area (12 mo) — each month’s mix of KPI row counts (raw counts in tooltip).
   */
  function renderSpiInsightsChart(snapPool, f) {
    const host = document.getElementById("chart-spi-map");
    if (!host) return;
    destroySpiLeafletMap();

    const kpisMeta = sortKpisForDisplay(
      SPI_CATEGORY_KEY,
      getKpis(SPI_CATEGORY_KEY)
    ).slice(0, 8);
    if (!kpisMeta.length) {
      host.innerHTML =
        '<p class="chart-spi-map-fallback">No SPI KPIs for this chart.</p>';
      return;
    }
    if (typeof Chart === "undefined") {
      host.innerHTML =
        '<p class="chart-spi-map-fallback">Chart unavailable.</p>';
      return;
    }

    const months = calendarClippedMonthKeys(f, 12);
    const labels = months.map(spiChartMonthTick);

    host.innerHTML =
      '<div class="chart-canvas-wrap chart-canvas-wrap--spi-insights">' +
      '<canvas id="chart-spi-insights" role="img" aria-label="100 percent stacked SPI KPI mix by month"></canvas></div>';
    const canvas = document.getElementById("chart-spi-insights");
    if (!canvas) return;
    const existing = Chart.getChart(canvas);
    if (existing) existing.destroy();

    const rawMatrix = kpisMeta.map((k) =>
      months.map((ym) => {
        let c = 0;
        for (let i = 0; i < snapPool.length; i++) {
          const r = snapPool[i];
          if (r.yearMonth === ym && String(r.kpiKey) === String(k.kpiKey)) {
            c += 1;
          }
        }
        return c;
      })
    );

    const pctMatrix = rawMatrix.map((row) => row.slice());
    for (let j = 0; j < months.length; j++) {
      let T = 0;
      for (let si = 0; si < rawMatrix.length; si++) T += rawMatrix[si][j];
      if (T <= 0) continue;
      for (let si = 0; si < rawMatrix.length; si++) {
        pctMatrix[si][j] = (rawMatrix[si][j] / T) * 100;
      }
    }

    const datasets = kpisMeta.map((k, idx) => {
      const pal = SPI_INSIGHT_PALETTE[idx % SPI_INSIGHT_PALETTE.length];
      return {
        label: shortKpiHeaderLabel(k.kpiName),
        data: pctMatrix[idx],
        fill: true,
        stack: "spi",
        tension: 0.28,
        borderWidth: 1.25,
        pointRadius: 0,
        pointHoverRadius: 4,
        borderColor: pal.stroke,
        backgroundColor: pal.fill,
      };
    });

    new Chart(canvas, {
      type: "line",
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              boxWidth: 10,
              padding: 10,
              font: {
                size: 9,
                family: FONT_UI,
              },
            },
          },
          tooltip: {
            callbacks: {
              label(ctx) {
                const j = ctx.dataIndex;
                const si = ctx.datasetIndex;
                const raw = rawMatrix[si] ? rawMatrix[si][j] : 0;
                const pct = Number(ctx.parsed.y) || 0;
                return (
                  " " +
                  ctx.dataset.label +
                  ": " +
                  pct.toFixed(1) +
                  "% (" +
                  raw +
                  " rows)"
                );
              },
              footer(tooltipItems) {
                if (!tooltipItems.length) return "";
                const j = tooltipItems[0].dataIndex;
                let rawSum = 0;
                for (let si = 0; si < rawMatrix.length; si++) {
                  rawSum += rawMatrix[si][j] || 0;
                }
                return (
                  "Total rows: " +
                  rawSum +
                  " · height = share of rows per KPI (not KPI value)"
                );
              },
            },
          },
        },
        scales: {
          x: {
            stacked: true,
            ticks: {
              maxRotation: 50,
              font: { size: 9 },
              color: "#231F20",
            },
            grid: { color: "rgba(109, 110, 113, 0.1)" },
          },
          y: {
            stacked: true,
            min: 0,
            max: 100,
            ticks: {
              font: { size: 9 },
              color: "#231F20",
              callback(v) {
                return v + "%";
              },
            },
            title: {
              display: true,
              text: "% of rows by KPI (same filters)",
              font: { size: 10 },
              color: "#6D6E71",
            },
            grid: { color: "rgba(109, 110, 113, 0.16)" },
          },
        },
      },
    });

    host.setAttribute(
      "aria-label",
      "SPI mix chart: percent of filtered fact rows per KPI by month (not KPI numeric values), twelve months."
    );
    setTimeout(() => {
      const ch = Chart.getChart(canvas);
      if (ch) ch.resize();
    }, 80);
  }

  /**
   * SPI: multi-KPI monthly line trend (same averaging as Property Damage / incident trend).
   * Click a series or legend item to focus that KPI in the KPI list.
   */
  function renderSpiKpiTrendLineChart(poolAll, f) {
    const elLine = document.getElementById("chart-line");
    if (!elLine || typeof Chart === "undefined") return;

    const catKey = SPI_CATEGORY_KEY;
    const kpisMeta = sortKpisForDisplay(
      catKey,
      kpiListForFilterDropdown(catKey)
    );
    if (!kpisMeta.length) return;

    const allKeys = new Set(kpisMeta.map((k) => String(k.kpiKey)));
    const range = new Set(effectiveChartMonths(f));
    const base = applyNonMonthFiltersAllKpis(poolAll, f);
    const filtered = base.filter(
      (r) =>
        range.has(r.yearMonth) && allKeys.has(String(r.kpiKey))
    );

    const monthSet = new Set();
    filtered.forEach((r) => monthSet.add(r.yearMonth));
    const lineLabels = Array.from(monthSet).sort();
    if (!lineLabels.length) return;

    const utChart = kpiUnitTypeForFilter(catKey, f);
    const primaryKpi = String(
      f.kpiKeys && f.kpiKeys.length ? f.kpiKeys[0] : f.kpi
    );

    const datasets = kpisMeta.map((kMeta, idx) => {
      const kk = String(kMeta.kpiKey);
      const data = lineLabels.map((ym) => {
        const slice = filtered.filter(
          (r) => r.yearMonth === ym && String(r.kpiKey) === kk
        );
        if (!slice.length) return null;
        return avg(slice.map((r) => Number(r.value)));
      });
      const color = TREND_LINE_COLORS[idx % TREND_LINE_COLORS.length];
      const sel = kk === primaryKpi;
      return {
        label: kpiDropdownLabel(kMeta),
        data: data,
        unitType: kMeta.unitType,
        kpiKey: kk,
        borderColor: color,
        backgroundColor: hexToRgba(color, sel ? 0.2 : 0.1),
        fill: true,
        tension: 0.2,
        borderWidth: sel ? 2.75 : 1.35,
        pointRadius: sel ? 3.5 : 2,
        pointHoverRadius: 6,
        order: sel ? 99 : idx,
      };
    });

    elLine.setAttribute(
      "aria-label",
      "Line chart: monthly average for each Safety Performance Index KPI; click a series or legend entry to focus that KPI."
    );

    new Chart(elLine, {
      type: "line",
      data: {
        labels: lineLabels,
        datasets: datasets,
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
          padding: { top: 10, right: 14, bottom: 10, left: 18 },
        },
        interaction: { mode: "nearest", intersect: false, axis: "x" },
        onHover(evt, els) {
          const t = evt && evt.native && evt.native.target;
          if (t && t.style) t.style.cursor = els.length ? "pointer" : "default";
        },
        onClick(evt, elements, chart) {
          if (!elements.length) return;
          const dsi = elements[0].datasetIndex;
          const ds = chart.data.datasets[dsi];
          if (ds && ds.kpiKey) {
            applyToolbarSingleKpiSelection(catKey, ds.kpiKey);
          }
        },
        plugins: {
          legend: {
            display: datasets.length > 1,
            position: "bottom",
            labels: {
              boxWidth: 10,
              padding: 6,
              font: { size: 8, family: FONT_UI },
            },
            onClick(e, legendItem, legend) {
              const i = legendItem.datasetIndex;
              const ds = legend.chart.data.datasets[i];
              if (ds && ds.kpiKey) {
                if (e && typeof e.preventDefault === "function") {
                  e.preventDefault();
                }
                applyToolbarSingleKpiSelection(catKey, ds.kpiKey);
              }
            },
          },
          tooltip: {
            callbacks: {
              label(ctx) {
                const v =
                  ctx.parsed && ctx.parsed.y != null
                    ? ctx.parsed.y
                    : Number(ctx.raw) || 0;
                const ut = ctx.dataset.unitType || utChart;
                return " " + formatValue(v, ut);
              },
            },
          },
        },
        scales: {
          y: {
            beginAtZero: false,
            grace: "12%",
            ticks: {
              font: { size: 9 },
              color: "#231F20",
              padding: 4,
            },
            grid: { color: "rgba(109, 110, 113, 0.2)" },
          },
          x: {
            ticks: {
              font: { size: 9 },
              maxRotation: 45,
              color: "#231F20",
              padding: 2,
            },
            grid: { color: "rgba(109, 110, 113, 0.15)" },
          },
        },
      },
    });
  }

  /**
   * Location Vulnerability: India map with one marker per preview BU (24); circle size ∝ row count.
   */
  function renderLocationVulnerabilityBuMap(snapRows, f) {
    const el = document.getElementById("chart-loc-bu-map");
    if (!el) return;
    destroyLocBuLeafletMap();
    const vsShort = vsOptionLabel((f && f.vsMode) || DEFAULT_VS_MODE);
    const hint = document.getElementById("chart-loc-bu-map-hint");
    if (hint) hint.textContent = "(map · row counts · 24 BUs · " + vsShort + ")";
    if (typeof L === "undefined") {
      el.innerHTML =
        '<p class="chart-spi-map-fallback">Map unavailable. Load Leaflet to see business unit markers.</p>';
      return;
    }

    const bus = PREVIEW_BUSINESS_NAMES.slice();
    const counts = {};
    bus.forEach((bu) => {
      counts[bu] = 0;
    });
    snapRows.forEach((r) => {
      const b = String(r.businessName || "").trim();
      bus.forEach((bu) => {
        if (factBusinessNameForPreview(bu) === b) counts[bu] += 1;
      });
    });

    loadGoogleMapsAndMutant().then(function (useGoogle) {
      const host = document.getElementById("chart-loc-bu-map");
      if (!host || typeof L === "undefined") return;

      const map = L.map(host, { scrollWheelZoom: true }).setView(
        [22.6, 79.0],
        4.5
      );
      if (
        useGoogle &&
        L.gridLayer &&
        L.gridLayer.googleMutant
      ) {
        L.gridLayer
          .googleMutant({ type: "roadmap" })
          .addTo(map);
      } else {
        L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
          maxZoom: 19,
          attribution: "&copy; OpenStreetMap contributors",
        }).addTo(map);
      }

      const latLngs = [];
      const indiaBounds = L.latLngBounds(
        INDIA_MAP_BOUNDS_SW,
        INDIA_MAP_BOUNDS_NE
      );
      bus.forEach((bu) => {
        const n = counts[bu] || 0;
        const fact = factBusinessNameForPreview(bu);
        const state = FACT_BU_HOME_STATE[fact] || "Gujarat";
        const base = STATE_CENTROID_BY_STATE[state] || [22.6, 79.0];
        const h = hash32(bu + "|locbu");
        const ox = ((h % 200) / 200 - 0.5) * 1.4;
        const oy = (((h >> 9) % 200) / 200 - 0.5) * 1.4;
        const ll = [base[0] + ox * 0.42, base[1] + oy * 0.42];
        latLngs.push(ll);
        const rad =
          n === 0 ? 4 : Math.max(5, Math.min(28, Math.sqrt(n) * 2.8));
        const mk = L.circleMarker(ll, {
          radius: rad,
          color: "#8E278F",
          fillColor: "#006DB6",
          fillOpacity: n === 0 ? 0.22 : 0.52,
          weight: 2,
        }).addTo(map);
        const title =
          escapeHtml(bu) + ": " + n + " rows · " + escapeHtml(fact);
        mk.bindPopup(
          "<strong>" +
            escapeHtml(bu) +
            "</strong><br>" +
            n +
            " rows (filtered)<br><span style=\"color:#6d6e71;font-size:11px\">Fact BU: " +
            escapeHtml(fact) +
            "</span>"
        );
        mk.bindTooltip(title, {
          sticky: true,
          direction: "auto",
          opacity: 0.95,
          className: "loc-bu-map-tooltip",
        });
      });
      if (latLngs.length) {
        latLngs.forEach((ll) => {
          indiaBounds.extend(ll);
        });
        map.fitBounds(indiaBounds, { padding: [20, 20], maxZoom: 6 });
      }
      window.__adaniLocBuMap = map;
      setTimeout(() => {
        try {
          map.invalidateSize();
        } catch {
          /* ignore */
        }
      }, 80);
    });
  }

  function wireCatMainViewIndex() {
    const main = document.getElementById("cat-main-view");
    if (!main) return;
    const tabCharts = document.getElementById("view-tab-charts");
    const tabTable = document.getElementById("view-tab-table");
    const tabCompare = document.getElementById("view-tab-compare");
    const panelCharts = document.getElementById("view-panel-charts");
    const panelTable = document.getElementById("view-panel-table");
    const panelCompare = document.getElementById("view-panel-compare");
    const hasCompare = !!(tabCompare && panelCompare);

    function apply(mode) {
      if (hasCompare && mode === "compare") {
        main.classList.remove("cat-main-view--charts", "cat-main-view--table");
        main.classList.add("cat-main-view--compare");
        main.dataset.view = "compare";
        tabCharts?.setAttribute("aria-selected", "false");
        tabTable?.setAttribute("aria-selected", "false");
        tabCompare?.setAttribute("aria-selected", "true");
        tabCharts?.classList.remove("view-tabs__btn--active");
        tabTable?.classList.remove("view-tabs__btn--active");
        tabCompare?.classList.add("view-tabs__btn--active");
        if (tabCharts) tabCharts.tabIndex = -1;
        if (tabTable) tabTable.tabIndex = -1;
        if (tabCompare) tabCompare.tabIndex = 0;
        if (panelCharts) panelCharts.hidden = true;
        if (panelTable) panelTable.hidden = true;
        if (panelCompare) panelCompare.hidden = false;
        try {
          localStorage.setItem(LS_CAT_MAIN_VIEW, "compare");
        } catch {
          /* ignore */
        }
        requestAnimationFrame(() => {
          requestAnimationFrame(() => resizeAllChartsIndex());
        });
        return;
      }

      const charts = mode === "charts";
      main.classList.remove("cat-main-view--compare");
      main.classList.toggle("cat-main-view--charts", charts);
      main.classList.toggle("cat-main-view--table", !charts);
      main.dataset.view = charts ? "charts" : "table";
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
      if (tabCompare) {
        tabCompare.setAttribute("aria-selected", "false");
        tabCompare.classList.remove("view-tabs__btn--active");
        tabCompare.tabIndex = -1;
      }
      if (panelCharts) panelCharts.hidden = !charts;
      if (panelTable) panelTable.hidden = charts;
      if (panelCompare) panelCompare.hidden = true;
      try {
        localStorage.setItem(LS_CAT_MAIN_VIEW, charts ? "charts" : "table");
      } catch {
        /* ignore */
      }
      if (charts) {
        requestAnimationFrame(() => {
          requestAnimationFrame(() => resizeAllChartsIndex());
        });
      }
    }

    let initial = "charts";
    try {
      const s = localStorage.getItem(LS_CAT_MAIN_VIEW);
      if (s === "charts" || s === "table" || s === "compare") initial = s;
    } catch {
      /* ignore */
    }
    if (initial === "compare" && !tabCompare) initial = "charts";

    apply(initial);
    tabCharts?.addEventListener("click", () => apply("charts"));
    tabTable?.addEventListener("click", () => apply("table"));
    tabCompare?.addEventListener("click", () => apply("compare"));
  }

  /** Distinct hues for radar spokes (one color per business unit). */
  function radarSpokeColors(count) {
    const n = Math.max(1, Math.min(Number(count) || 1, 48));
    const out = [];
    for (let i = 0; i < n; i++) {
      const h = Math.round((360 * i) / n + 14) % 360;
      out.push({
        stroke: "hsl(" + h + " 72% 34%)",
        fill: "hsla(" + h + " 68% 46% / 0.9)",
        fillSoft: "hsla(" + h + " 62% 52% / 0.2)",
      });
    }
    return out;
  }

  /**
   * By business: radar chart — one spoke per preview BU (24); value from filtered rows or 0.
   */
  function renderBizBreakdown(fullNames, values, datasetLabel, unitType) {
    const el = document.getElementById("chart-biz");
    const emptyEl = document.getElementById("chart-biz-empty");
    if (!el || typeof Chart === "undefined") return;
    const prev = Chart.getChart(el);
    if (prev) prev.destroy();

    if (!fullNames || !fullNames.length) {
      el.style.display = "none";
      if (emptyEl) {
        emptyEl.hidden = false;
        emptyEl.textContent = "No business data for current filters.";
      }
      return;
    }

    const nums = values.map((v) =>
      Number.isFinite(Number(v)) ? Number(v) : 0
    );
    if (!nums.some((v) => v !== 0)) {
      el.style.display = "none";
      if (emptyEl) {
        emptyEl.hidden = false;
        emptyEl.textContent = "No business data for current filters.";
      }
      return;
    }

    el.style.display = "block";
    if (emptyEl) emptyEl.hidden = true;

    function radarBizLabel(name) {
      const s = String(name).trim();
      if (s.length <= 11) return s;
      return s.slice(0, 10) + "\u2026";
    }
    const shortLabels = fullNames.map(radarBizLabel);
    const palette = radarSpokeColors(nums.length);
    const pointStrokes = nums.map((_, i) => palette[i % palette.length].stroke);
    const pointFills = nums.map((_, i) => palette[i % palette.length].fill);

    const radarValueLabelPlugin = {
      id: "radarValueLabels",
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        const meta = chart.getDatasetMeta(0);
        if (!meta || !meta.data || !meta.data.length) return;
        const ds = chart.data.datasets[0].data;
        const { left, right, top, bottom } = chart.chartArea;
        const cx = (left + right) / 2;
        const cy = (top + bottom) / 2;
        ctx.save();
        ctx.font = "600 10px " + FONT_UI;
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        meta.data.forEach((pt, i) => {
          const v = Number(ds[i]);
          if (!Number.isFinite(v)) return;
          const xy = pt.getProps(["x", "y"], true);
          const x = xy.x;
          const y = xy.y;
          const dx = x - cx;
          const dy = y - cy;
          const len = Math.sqrt(dx * dx + dy * dy) || 1;
          const ox = (dx / len) * 18;
          const oy = (dy / len) * 18;
          const txt = formatValue(v, unitType);
          const ink = palette[i % palette.length].stroke;
          ctx.lineJoin = "round";
          ctx.strokeStyle = "rgba(255, 255, 255, 0.95)";
          ctx.lineWidth = 4;
          ctx.strokeText(txt, x + ox, y + oy);
          ctx.fillStyle = ink;
          ctx.fillText(txt, x + ox, y + oy);
        });
        ctx.restore();
      },
    };

    new Chart(el, {
      type: "radar",
      plugins: [radarValueLabelPlugin],
      data: {
        labels: shortLabels,
        datasets: [
          {
            label: datasetLabel || "By business",
            data: nums,
            borderColor: pointStrokes[0] || "#006DB6",
            backgroundColor: "transparent",
            borderWidth: 2,
            fill: true,
            segment: {
              borderWidth: 2,
              borderColor: (ctx) =>
                palette[ctx.p0DataIndex % palette.length].stroke,
              backgroundColor: (ctx) =>
                palette[ctx.p0DataIndex % palette.length].fillSoft,
            },
            pointRadius: 4,
            pointHoverRadius: 8,
            pointBorderWidth: 2,
            pointBackgroundColor: pointFills,
            pointBorderColor: pointStrokes,
            hoverBorderWidth: 2,
            hoverBackgroundColor: pointFills,
            hoverBorderColor: pointStrokes,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: "nearest",
          intersect: false,
        },
        layout: {
          padding: { top: 14, right: 18, bottom: 18, left: 18 },
        },
        scales: {
          r: {
            beginAtZero: true,
            suggestedMax:
              Math.max.apply(
                null,
                nums.map((n) => Number(n) || 0)
              ) * 1.12 || undefined,
            angleLines: { color: "rgba(109, 110, 113, 0.28)" },
            grid: { color: "rgba(109, 110, 113, 0.18)" },
            pointLabels: {
              font: {
                size: 8.5,
                weight: "500",
                family: FONT_UI,
              },
              color: (ctx) => {
                const i =
                  ctx.index != null
                    ? ctx.index
                    : ctx.tick && ctx.tick.index != null
                      ? ctx.tick.index
                      : 0;
                return palette[i % palette.length].stroke;
              },
              padding: 4,
            },
            ticks: {
              display: true,
              maxTicksLimit: 7,
              font: { size: 9, family: FONT_UI },
              backdropColor: "rgba(255,255,255,0.9)",
              color: "#5a5c5f",
            },
          },
        },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              boxWidth: 0,
              usePointStyle: false,
              generateLabels() {
                return [
                  {
                    text: String(datasetLabel || "By business"),
                    fillStyle: "transparent",
                    strokeStyle: "transparent",
                    lineWidth: 0,
                    hidden: false,
                    index: 0,
                    datasetIndex: 0,
                  },
                ];
              },
              font: {
                size: 10,
                family: FONT_UI,
              },
              color: "#231F20",
            },
          },
          tooltip: {
            enabled: true,
            backgroundColor: "rgba(35, 31, 32, 0.96)",
            titleColor: "#ffffff",
            bodyColor: "#f5f5f5",
            borderColor: "#006DB6",
            borderWidth: 1,
            padding: 12,
            cornerRadius: 6,
            caretSize: 7,
            titleFont: {
              size: 13,
              weight: "600",
              family: FONT_UI,
            },
            bodyFont: {
              size: 12,
              weight: "500",
              family: FONT_UI,
            },
            titleMarginBottom: 6,
            boxPadding: 6,
            displayColors: true,
            usePointStyle: true,
            callbacks: {
              title(items) {
                const i = items[0].dataIndex;
                return fullNames[i] != null ? String(fullNames[i]) : "";
              },
              label(ctx) {
                const v = Number(ctx.raw);
                const ut = unitType || "";
                const val = formatValue(v, unitType);
                const line1 = "Value: " + val;
                return ut ? [line1, "Unit: " + ut] : line1;
              },
            },
          },
        },
      },
    });
  }

  /** Selected KPI keys in category display order (for charts / table / highlights). */
  function orderKpiKeysForCharts(catKey, kpiKeysRaw) {
    const meta = kpiListForFilterDropdown(catKey);
    const valid = new Set(meta.map((k) => String(k.kpiKey)));
    const chosen = (kpiKeysRaw || [])
      .map(String)
      .filter((id) => valid.has(id));
    if (!chosen.length) return [];
    const order = sortKpisForDisplay(catKey, meta).map((k) => String(k.kpiKey));
    return order.filter((id) => chosen.includes(id));
  }

  /**
   * Keys for multi-series trend/hazard charts: selected checkboxes, else primary `f.kpi`,
   * else category default (matches applyChartFilter / applyNonMonthFilters fallback).
   */
  function effectiveKpiKeysForChartSeries(catKey, f) {
    const raw =
      f.kpiKeys && f.kpiKeys.length
        ? f.kpiKeys.map(String)
        : f.kpi != null && String(f.kpi) !== ""
          ? [String(f.kpi)]
          : [];
    let ordered = orderKpiKeysForCharts(catKey, raw);
    if (!ordered.length) {
      const dk = defaultKpiKeyForCategory(
        catKey,
        kpiListForFilterDropdown(catKey)
      );
      if (dk) ordered = orderKpiKeysForCharts(catKey, [String(dk)]);
    }
    return ordered;
  }

  function shouldShowCmpAccountabilityChart(catKey, f) {
    if (catKey !== CONSEQUENCE_MANAGEMENT_CATEGORY_KEY) return false;
    const keys = effectiveKpiKeysForChartSeries(catKey, f);
    return Boolean(
      keys.length && CMP_ACCOUNTABILITY_CHART_PRIMARY_KPI_KEYS.has(keys[0])
    );
  }

  function speedometerPrimaryKpiKey(catKey, f) {
    const keys = effectiveKpiKeysForChartSeries(catKey, f);
    if (!keys.length) return null;
    const primary = keys[0];
    if (
      catKey === ASSURANCE_CATEGORY_KEY &&
      ASSURANCE_SPEEDOMETER_KPI_KEYS.has(primary)
    ) {
      return primary;
    }
    if (
      catKey === TRAINING_CATEGORY_KEY &&
      primary === TRAINING_SAKSHAM_SPEEDOMETER_KPI_KEY
    ) {
      return primary;
    }
    if (
      catKey === SYSTEMS_ADOPTION_CATEGORY_KEY &&
      primary === SYSTEMS_SAFEX_SPEEDOMETER_KPI_KEY
    ) {
      return primary;
    }
    return null;
  }

  function shouldShowPercentSpeedometerChart(catKey, f) {
    return speedometerPrimaryKpiKey(catKey, f) != null;
  }

  function speedometerPercentFromSnapRows(catKey, snapRows, f) {
    const pk = speedometerPrimaryKpiKey(catKey, f);
    if (!pk) return 0;
    const rows = snapRows.filter((r) => String(r.kpiKey) === pk);
    const nums = rows.map((r) => Number(r.value)).filter(Number.isFinite);
    if (!nums.length) return 0;
    return avg(nums);
  }

  function readFilters(catKey) {
    const elVs = document.getElementById("f-vs");
    const elSt = document.getElementById("f-state");
    const elBiz = document.getElementById("f-biz");
    const elSite = document.getElementById("f-site");
    const elPersonal = document.getElementById("f-personal");
    const elFrom = document.getElementById("f-dt-from");
    const elTo = document.getElementById("f-dt-to");
    const kpisMeta = kpiListForFilterDropdown(catKey);
    const rawKeys = document.getElementById("f-kpi-panel")
      ? readSelectedKpiKeysFromDom()
      : [];
    let kpiKeys = orderKpiKeysForCharts(catKey, rawKeys);
    const dk = defaultKpiKeyForCategory(catKey, kpisMeta);
    const kpi = kpiKeys.length ? kpiKeys[0] : dk;
    const mb = dataMonthBounds();
    const monthTo =
      elTo && elTo.value && elTo.value.length >= 7
        ? elTo.value.slice(0, 7)
        : mb.max;
    const monthFrom =
      elFrom && elFrom.value && elFrom.value.length >= 7
        ? elFrom.value.slice(0, 7)
        : null;
    return {
      catKey,
      kpi: kpi,
      kpiKeys: kpiKeys,
      vsMode: elVs ? elVs.value : DEFAULT_VS_MODE,
      refMonth: monthTo,
      monthFrom: monthFrom,
      monthTo: monthTo,
      state: elSt ? elSt.value : "all",
      business: elBiz ? elBiz.value : "all",
      site: elSite ? elSite.value : "all",
      personalType: elPersonal ? elPersonal.value : "all",
      unitType: "all",
      variable: readVariableSelectionFromDom(),
    };
  }

  /** Table / KPI tiles: calendar end month (or full from–to when both set). */
  function applyRowFilter(rows, f) {
    const keys =
      f.kpiKeys && f.kpiKeys.length
        ? f.kpiKeys.map(String)
        : [String(f.kpi)];
    return rows.filter((r) => {
      if (!keys.includes(String(r.kpiKey))) return false;
      if (f.monthFrom && f.monthTo) {
        if (r.yearMonth < f.monthFrom || r.yearMonth > f.monthTo) return false;
      } else if (r.yearMonth !== f.refMonth) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (!rowMatchesBusinessFilter(r, f)) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      if (!rowMatchesSiteFilter(r, f)) return false;
      if (!rowMatchesPersonalTypeFilter(r, f)) return false;
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

  function dataMonthBounds() {
    const list = (DATA.months || [])
      .map((x) => x.yearMonth)
      .filter(Boolean)
      .sort();
    if (list.length) return { min: list[0], max: list[list.length - 1] };
    const m = meta.lastDataMonth || "2024-01";
    return { min: m, max: m };
  }

  function lastDayOfMonthYm(ym) {
    const p = String(ym || "").split("-");
    if (p.length < 2) return ym + "-28";
    const y = Number(p[0]);
    const mo = Number(p[1]);
    if (!y || !mo) return ym + "-28";
    const d = new Date(y, mo, 0).getDate();
    return ym + "-" + String(d).padStart(2, "0");
  }

  /** Vs window months, clipped to Calendar from–to when set (Power BI–style slicer). */
  function effectiveChartMonths(f) {
    const ref = (f && (f.monthTo || f.refMonth)) || getRefMonth();
    const base = chartMonthsForVsMode(f.vsMode || DEFAULT_VS_MODE, ref);
    const from = f && f.monthFrom;
    const to = (f && f.monthTo) || ref;
    if (!from) return base;
    return base.filter((ym) => ym >= from && ym <= to);
  }

  /** Fixed-length month keys ending at calendar “to”, clipped to from–to range. */
  function calendarClippedMonthKeys(f, count) {
    const end = (f && (f.monthTo || f.refMonth)) || getRefMonth();
    const months = chartMonthKeys(end, count);
    const from = f && f.monthFrom;
    const to = (f && f.monthTo) || end;
    if (!from) return months;
    return months.filter((ym) => ym >= from && ym <= to);
  }

  function rowDummySite(r) {
    const h = Math.abs(
      hash32(
        String(r.businessKey || "") +
          "|" +
          String(r.kpiKey || "") +
          "|" +
          String(r.yearMonth || "")
      )
    );
    return "Site" + ((h % 7) + 1);
  }

  function rowDummyPersonalType(r) {
    const h = Math.abs(
      hash32(
        "pt|" +
          String(r.businessKey || "") +
          "|" +
          String(r.kpiKey || "") +
          "|" +
          String(r.yearMonth || "")
      )
    );
    return h % 3 === 1 ? "Contractor" : "Employee";
  }

  function rowMatchesSiteFilter(r, f) {
    if (!f || !f.site || f.site === "all") return true;
    return rowDummySite(r) === f.site;
  }

  function rowMatchesPersonalTypeFilter(r, f) {
    if (!f || !f.personalType || f.personalType === "all") return true;
    return rowDummyPersonalType(r) === f.personalType;
  }

  function applyChartFilter(rows, f) {
    const keys =
      f.kpiKeys && f.kpiKeys.length
        ? f.kpiKeys.map(String)
        : [String(f.kpi)];
    if (!keys.length) return [];
    const range = new Set(effectiveChartMonths(f));
    return rows.filter((r) => {
      if (!keys.includes(String(r.kpiKey))) return false;
      if (!range.has(r.yearMonth)) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (!rowMatchesBusinessFilter(r, f)) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      if (!rowMatchesSiteFilter(r, f)) return false;
      if (!rowMatchesPersonalTypeFilter(r, f)) return false;
      return true;
    });
  }


  function getRowsForCategory(catKey) {
    if (catKey === LOCATION_VULN_CAT_KEY) {
      return DATA.factRows
        .filter(
          (r) =>
            r.categoryKey === LOCATION_VULN_SOURCE_CAT &&
            LOCATION_VULN_KPI_KEYS.has(Number(r.kpiKey))
        )
        .map((r) => ({ ...r, categoryKey: LOCATION_VULN_CAT_KEY }));
    }
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
    return {
      showKpi: catKey !== LOCATION_VULN_CAT_KEY,
      showState: catKey === LOCATION_VULN_CAT_KEY,
      showBusiness: true,
    };
  }

  function avg(nums) {
    if (!nums.length) return 0;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  }

  /** Incident Management: TRI first, then wireframe order (ΔRepeat, ΔFatal, Man-days, Vehicle). */
  const INCIDENT_KPI_ORDER = [
    21,
    1,
    2,
    3,
    4,
    5,
    7,
    8,
    INCIDENT_FIRE_KPI_KEY,
    INCIDENT_PROPERTY_DAMAGE_KPI_KEY,
    9,
    10,
    11,
    12,
    14,
    15,
    19,
    22,
    28,
    44,
  ];
  /** Hazard & Observation (leading): activity first + S-4/S-5 SRFA (from former Risk Control). */
  const HAZARD_KPI_ORDER = [38, 13, 6, 39, 40, 41, 42, 45, 46, 53, 21];
  /** Consequence Management + compliance KPIs moved from Risk Control. */
  const CONSEQUENCE_KPI_ORDER = [23, 24, 25, 26, 27, 51, 52, 54];

  function sortKpisForDisplay(catKey, kpisMeta) {
    const list = kpisMeta.slice();
    if (catKey === 1) {
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      return INCIDENT_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (catKey === HAZARD_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      const ordered = HAZARD_KPI_ORDER.map((id) => map.get(id)).filter(
        Boolean
      );
      const rest = list.filter(
        (k) => !HAZARD_KPI_ORDER.includes(Number(k.kpiKey))
      );
      return ordered.concat(rest);
    }
    if (catKey === ASSURANCE_CATEGORY_KEY) {
      const order = [501, 502, 503];
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      return order.map((id) => map.get(id)).filter(Boolean);
    }
    if (catKey === CONSEQUENCE_MANAGEMENT_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      return CONSEQUENCE_KPI_ORDER.map((id) => map.get(id))
        .filter(Boolean)
        .concat(
          list.filter((k) => !CONSEQUENCE_KPI_ORDER.includes(Number(k.kpiKey)))
        );
    }
    if (catKey === SPI_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      const ordered = SPI_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
      const rest = list.filter(
        (k) => !SPI_KPI_ORDER.includes(Number(k.kpiKey))
      );
      return ordered.concat(rest);
    }
    if (catKey === LOCATION_VULN_CAT_KEY) {
      const order = [13, 38, 39, 40, 45, 46, 53];
      const map = new Map(list.map((k) => [k.kpiKey, k]));
      return order.map((id) => map.get(id)).filter(Boolean);
    }
    const tri = list.find((k) => isTriKpiMeta(k));
    if (!tri) return list;
    const rest = list.filter((k) => k !== tri);
    return [tri, ...rest];
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

  /**
   * Preview data: for reference month vs prior month, reshape values per
   * category × KPI × business so Vs% is in ±5–±20% with mixed signs. Mutates
   * factRows in place so KPI tiles, charts, and BU matrix stay aligned.
   */
  function normalizePreviewFactRowsVsBand() {
    const data = DATA;
    if (!data || !Array.isArray(data.factRows) || !data.factRows.length) return;
    let ref = null;
    if (data.months && data.months.length) {
      ref = data.months[data.months.length - 1].yearMonth;
    }
    if (!ref && data.meta) ref = data.meta.lastDataMonth;
    if (!ref) return;
    const m1 = monthAdd(ref, -1);
    if (!m1) return;

    function aggRows(rows, ut) {
      if (!rows || !rows.length) return null;
      const nums = rows.map((r) => Number(r.value)).filter((v) => !Number.isNaN(v));
      if (!nums.length) return null;
      if (isAdditiveUnit(ut)) return nums.reduce((a, b) => a + b, 0);
      return avg(nums);
    }

    function applyNewAgg(rows, ut, newAgg) {
      const n = rows.length;
      if (!n || newAgg == null || Number.isNaN(Number(newAgg))) return;
      if (isAdditiveUnit(ut)) {
        const cur = rows.reduce((a, r) => a + Number(r.value), 0);
        if (Math.abs(cur) < 1e-9) {
          const each = Number(newAgg) / n;
          for (let i = 0; i < n; i++) rows[i].value = each;
          return;
        }
        const f = Number(newAgg) / cur;
        for (let i = 0; i < n; i++) rows[i].value = Number(rows[i].value) * f;
        return;
      }
      const curAvg = avg(rows.map((r) => Number(r.value)));
      const delta = Number(newAgg) - curAvg;
      for (let i = 0; i < n; i++) rows[i].value = Number(rows[i].value) + delta;
    }

    const index = new Map();
    for (let i = 0; i < data.factRows.length; i++) {
      const r = data.factRows[i];
      if (r.yearMonth !== ref && r.yearMonth !== m1) continue;
      const bn = String(r.businessName || "").trim();
      const key = r.categoryKey + "|" + r.kpiKey + "|" + bn + "|" + r.yearMonth;
      if (!index.has(key)) index.set(key, { rows: [] });
      index.get(key).rows.push(r);
    }

    const seriesKeys = new Set();
    for (const key of index.keys()) {
      const p = key.split("|");
      seriesKeys.add(p[0] + "|" + p[1] + "|" + p[2]);
    }

    for (const ser of seriesKeys) {
      const kM1 = ser + "|" + m1;
      const kRef = ser + "|" + ref;
      const slotM1 = index.get(kM1);
      const slotRef = index.get(kRef);
      if (!slotM1 || !slotM1.rows.length || !slotRef || !slotRef.rows.length)
        continue;

      const ut = slotM1.rows[0].unitType || slotRef.rows[0].unitType;
      const h = hash32(ser + "|vsband");
      const mag = 5 + (h % 16);
      const sign = (h >> 11) % 2 === 0 ? 1 : -1;
      const p = sign * mag;

      let baseAgg = aggRows(slotM1.rows, ut);
      if (baseAgg == null || Math.abs(baseAgg) < 1e-9) {
        const seed = 60 + (Math.abs(h) % 80);
        if (isAdditiveUnit(ut)) {
          const each = seed / slotM1.rows.length;
          for (let i = 0; i < slotM1.rows.length; i++) slotM1.rows[i].value = each;
        } else {
          for (let i = 0; i < slotM1.rows.length; i++) slotM1.rows[i].value = seed;
        }
        baseAgg = aggRows(slotM1.rows, ut);
      }
      if (baseAgg == null || Math.abs(baseAgg) < 1e-9) continue;

      const targetRef = baseAgg * (1 + p / 100);
      applyNewAgg(slotRef.rows, ut, targetRef);
    }
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

  /**
   * @param {boolean} [primaryOnly] If true, keep only the first selected KPI (for BU / vertical / radar windows).
   */
  function applyNonMonthFilters(rows, f, primaryOnly) {
    const keys = primaryOnly
      ? [String(f.kpi)]
      : f.kpiKeys && f.kpiKeys.length
        ? f.kpiKeys.map(String)
        : [String(f.kpi)];
    return rows.filter((r) => {
      if (!keys.includes(String(r.kpiKey))) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
      if (!rowMatchesBusinessFilter(r, f)) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      if (!rowMatchesSiteFilter(r, f)) return false;
      if (!rowMatchesPersonalTypeFilter(r, f)) return false;
      return true;
    });
  }

  /** For multi-KPI tiles: same geography / business / vertical filters, all KPIs. */
  function applyNonMonthFiltersAllKpis(rows, f) {
    return rows.filter((r) => {
      if (f.state !== "all" && r.state !== f.state) return false;
      if (!rowMatchesBusinessFilter(r, f)) return false;
      if (f.unitType !== "all" && r.unitType !== f.unitType) return false;
      if (!rowMatchesVariable(r, f.variable)) return false;
      if (!rowMatchesSiteFilter(r, f)) return false;
      if (!rowMatchesPersonalTypeFilter(r, f)) return false;
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

  /** KPI tile Vs line: same signed % and direction as buildKpiDetailMetrics (matches table when one BU is filtered). */
  function tileVsDisplay(item) {
    const raw = item.vsPct;
    if (raw == null || Number.isNaN(Number(raw))) {
      return { dir: "na", pct: "—" };
    }
    let d = item.vsDir;
    if (d === "neutral") d = "same";
    if (d !== "up" && d !== "down" && d !== "same") d = "na";
    return { dir: d, pct: formatSignedPct(Number(raw)) };
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

  /** Vs % for one BU × KPI (same windows as KPI tiles), using filtered pool. */
  function buKpiVsPct(basePool, buName, kpiMeta, f) {
    const kk = kpiMeta.kpiKey;
    const ut = kpiMeta.unitType;
    const add = isAdditiveUnit(ut);
    const ref = f.refMonth;
    const mode = f.vsMode || DEFAULT_VS_MODE;
    const m1 = monthAdd(ref, -1);
    const m2 = monthAdd(ref, -2);
    const m3 = monthAdd(ref, -3);
    const m4 = monthAdd(ref, -4);
    const m5 = monthAdd(ref, -5);
    const bu = factBusinessNameForPreview(buName);

    function aggMonth(ym) {
      if (!ym) return null;
      const rows = basePool.filter(
        (r) =>
          r.kpiKey === kk &&
          r.yearMonth === ym &&
          String(r.businessName || "").trim() === bu
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

    return {
      pct: pctChange(cur, baseVal),
      cur: cur,
      baseVal: baseVal,
    };
  }

  function isLowerIsBetterKpiName(name) {
    const n = String(name || "").toLowerCase();
    return (
      n.includes("trir") ||
      n.includes("ltifr") ||
      n.includes("lti") ||
      n.includes("fatal") ||
      n.includes("recordable incident") ||
      n.includes("incident rate") ||
      n.includes("lost time")
    );
  }

  function matrixCellGoodClass(pct, kpiName) {
    if (pct == null || Number.isNaN(Number(pct))) return "bu-cell__bar--na";
    const p = Number(pct);
    if (Math.abs(p) < 1e-9) return "bu-cell__bar--neutral";
    if (isLowerIsBetterKpiName(kpiName)) {
      return p < 0 ? "bu-cell__bar--good" : "bu-cell__bar--bad";
    }
    return p > 0 ? "bu-cell__bar--good" : "bu-cell__bar--bad";
  }

  /**
   * When Vs % cannot be computed (no rows), use a stable preview value so every matrix cell
   * shows a number + bar (Power BI: replace with BLANK() or alternate measure as needed).
   */
  function matrixFallbackPct(catKey, bu, kpiKey, f) {
    const h = hash32(
      String(catKey) +
        "|" +
        bu +
        "|" +
        String(kpiKey) +
        "|" +
        (f.refMonth || "") +
        "|" +
        (f.vsMode || "") +
        "|" +
        (f.state || "") +
        "|" +
        (f.business || "")
    );
    const x = (h % 20001) / 20001;
    return Math.round((x * 70 - 35) * 10) / 10;
  }

  function matrixEffectivePct(pct, catKey, bu, kpiKey, f) {
    if (pct != null && !Number.isNaN(Number(pct))) {
      return { value: Number(pct), imputed: false };
    }
    return {
      value: matrixFallbackPct(catKey, bu, kpiKey, f),
      imputed: true,
    };
  }

  /**
   * Diverging bar: width % within left or right half (0–100), from center axis.
   * Negative → left (red); positive → right (green). Scaled to max |%| in the matrix.
   */
  function matrixBarDivergingWidths(signedValue, maxAbs) {
    if (Math.abs(signedValue) < 1e-9) return { left: 0, right: 0 };
    if (maxAbs < 1e-12) return { left: 0, right: 0 };
    const t = Math.min(1, Math.abs(signedValue) / maxAbs);
    const w = t * 100;
    if (signedValue < 0) return { left: w, right: 0 };
    if (signedValue > 0) return { left: 0, right: w };
    return { left: 0, right: 0 };
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

  function cmpArrowWrapClass(dir) {
    if (dir === "up") {
      return "multi-kpi-tile__arrow-wrap multi-kpi-tile__arrow-wrap--up";
    }
    if (dir === "down") {
      return "multi-kpi-tile__arrow-wrap multi-kpi-tile__arrow-wrap--down";
    }
    if (dir === "same") {
      return "multi-kpi-tile__arrow-wrap multi-kpi-tile__arrow-wrap--same";
    }
    return "multi-kpi-tile__arrow-wrap multi-kpi-tile__arrow-wrap--na";
  }

  function arrowSvgHtml(dir) {
    const a =
      'class="multi-kpi-tile__arrow-svg" width="18" height="18" viewBox="0 0 24 24" focusable="false" aria-hidden="true"';
    if (dir === "up") {
      return (
        "<svg " +
        a +
        '><path fill="currentColor" d="M12 6L4 15h16L12 6z"/></svg>'
      );
    }
    if (dir === "down") {
      return (
        "<svg " +
        a +
        '><path fill="currentColor" d="M12 18l8-9H4l8 9z"/></svg>'
      );
    }
    if (dir === "same") {
      return (
        "<svg " +
        a +
        '><path fill="currentColor" d="M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8z"/></svg>'
      );
    }
    return (
      "<svg " +
      a +
      "><path fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" d=\"M6 12h12\"/></svg>"
    );
  }

  function shortBizChartLabel(name) {
    const s = String(name);
    if (s.length <= 10) return s;
    return s.slice(0, 9) + "\u2026";
  }

  /** Location compare chart: readable axis text (multiline for long names). */
  function locCompareBizLabel(name) {
    const s = String(name).trim();
    if (s.length <= 14) return s;
    const mid = Math.floor(s.length / 2);
    let cut = s.lastIndexOf(" ", mid + 5);
    if (cut < 4) cut = s.indexOf(" ", 10);
    if (cut > 0 && cut < s.length - 1) {
      return [s.slice(0, cut).trim(), s.slice(cut + 1).trim()];
    }
    return s.length <= 18 ? s : s.slice(0, 16) + "\u2026";
  }

  function shortKpiHeaderLabel(name) {
    const s = String(name);
    if (s.length <= 18) return s;
    return s.slice(0, 16) + "\u2026";
  }

  function appendKpiTileEl(
    grid,
    item,
    refMonth,
    vsMode,
    tileClass,
    selectedKpiKeys,
    presentationSeed,
    catKey
  ) {
    void refMonth;
    void presentationSeed;
    void selectedKpiKeys;
    const tile = document.createElement("div");
    tile.className = (tileClass || "multi-kpi-tile") + " multi-kpi-tile--action";
    tile.setAttribute("role", "button");
    tile.setAttribute("tabindex", "0");
    const vsShort = vsTagShort(vsMode || DEFAULT_VS_MODE);
    const pres = tileVsDisplay(item);
    const vDir = pres.dir;
    if (vDir === "up") tile.classList.add("multi-kpi-tile--trend-up");
    else if (vDir === "down") tile.classList.add("multi-kpi-tile--trend-down");
    const infoTitle = escapeAttr(vsFilterComparisonTooltip(vsMode));
    const vsLbl = vsOptionLabel(vsMode || DEFAULT_VS_MODE);
    tile.setAttribute(
      "aria-label",
      item.kpiName +
        ". Value " +
        formatValue(item.value, item.unitType) +
        ". " +
        vsLbl +
        " " +
        pres.pct +
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
      '<div class="multi-kpi-tile__cmp multi-kpi-tile__cmp--vs">' +
      '<span class="multi-kpi-tile__cmp-tag multi-kpi-tile__cmp-tag--short">' +
      escapeHtml(vsShort) +
      "</span>" +
      '<span class="' +
      cmpArrowWrapClass(vDir) +
      '">' +
      arrowSvgHtml(vDir) +
      "</span>" +
      '<span class="multi-kpi-tile__cmp-pct vs-pct-num">' +
      escapeHtml(pres.pct) +
      "</span>" +
      '<button type="button" class="multi-kpi-tile__info" title="' +
      infoTitle +
      '" aria-label="Vs comparison for current filter"><span class="multi-kpi-tile__info-icon" aria-hidden="true">i</span></button>' +
      "</div>";
    const infoBtn = tile.querySelector(".multi-kpi-tile__info");
    if (infoBtn) {
      infoBtn.addEventListener("click", function (e) {
        e.stopPropagation();
      });
    }
    function openCompareForKpi() {
      applyToolbarSingleKpiSelection(catKey, item.kpiKey);
      const tabc = document.getElementById("view-tab-compare");
      if (tabc) tabc.click();
    }
    tile.addEventListener("click", function (e) {
      if (e.target.closest(".multi-kpi-tile__info")) return;
      openCompareForKpi();
    });
    tile.addEventListener("keydown", function (e) {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openCompareForKpi();
      }
    });
    grid.appendChild(tile);
  }

  function renderMultiKpiCards(
    container,
    aggregatesList,
    refMonth,
    vsMode,
    selectedKpiKeys,
    presentationSeed,
    catKey
  ) {
    container.innerHTML = "";
    const sorted = aggregatesList.slice();
    const headItems = sorted.slice(0, 12);
    const tailItems = sorted.slice(12);

    const rowMain = document.createElement("div");
    rowMain.className = "multi-kpi-row";

    function appendMetricCard(startMetric, items) {
      if (!items.length) return;
      const card = document.createElement("div");
      card.className = "multi-kpi-card";
      const head = document.createElement("div");
      head.className = "multi-kpi-card__head";
      const endMetric = startMetric + items.length - 1;
      head.textContent =
        items.length === 1
          ? "Metric " + startMetric
          : "Metrics " + startMetric + "–" + endMetric;
      const grid = document.createElement("div");
      grid.className = "multi-kpi-card__grid";
      items.forEach((item) => {
        appendKpiTileEl(
          grid,
          item,
          refMonth,
          vsMode,
          "multi-kpi-tile",
          selectedKpiKeys,
          presentationSeed,
          catKey
        );
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
      rowMain.appendChild(card);
    }

    const chunks = chunkArray(headItems, 4);
    chunks.forEach((chunk, idx) => {
      appendMetricCard(idx * 4 + 1, chunk);
    });

    const tailA = tailItems.slice(0, 4);
    const tailB = tailItems.slice(4);
    if (tailA.length) {
      appendMetricCard(13, tailA);
    }
    if (tailB.length) {
      appendMetricCard(13 + tailA.length, tailB);
    }

    container.appendChild(rowMain);
  }

  /** Grouped bars — comparison (base) vs current KPI value per BU (all categories). */
  function buildBuComparisonChart(catKey, f) {
    const el = document.getElementById("chart-bu-compare");
    if (!el || typeof Chart === "undefined") return;
    const prev = Chart.getChart(el);
    if (prev) prev.destroy();
    const kMeta = getKpis(catKey).find(
      (x) => String(x.kpiKey) === String(f.kpi)
    );
    if (!kMeta) return;
    const basePool = applyNonMonthFiltersAllKpis(getRowsForCategory(catKey), f);
    const bus = PREVIEW_BUSINESS_NAMES.slice();
    const curData = [];
    const baseData = [];
    bus.forEach((bu) => {
      const x = buKpiVsPct(basePool, bu, kMeta, f);
      const c = x.cur;
      const b = x.baseVal;
      curData.push(c != null && !Number.isNaN(Number(c)) ? Number(c) : 0);
      baseData.push(b != null && !Number.isNaN(Number(b)) ? Number(b) : 0);
    });
    const vsShort = vsOptionLabel(f.vsMode || DEFAULT_VS_MODE);
    const unit = kMeta.unitType || "";
    const yTitle = unit === "PercentOrRate" ? "Value (%)" : "Value";
    new Chart(el, {
      type: "bar",
      data: {
        labels: bus.map(locCompareBizLabel),
        datasets: [
          {
            label: "Comparison (base) · " + vsShort,
            data: baseData,
            backgroundColor: "#8E278F",
            borderColor: "#6a1d6b",
            borderWidth: 1,
          },
          {
            label: "Current period",
            data: curData,
            backgroundColor: "#006DB6",
            borderWidth: 1,
            borderColor: "#004a80",
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        layout: {
          padding: { top: 10, right: 12, bottom: 64, left: 10 },
        },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              boxWidth: 12,
              font: {
                size: 9,
                family: FONT_UI,
              },
              color: "#231F20",
            },
          },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0].dataIndex;
                return bus[i] != null ? String(bus[i]) : "";
              },
            },
          },
        },
        scales: {
          x: {
            stacked: false,
            grid: { display: false },
            ticks: {
              autoSkip: false,
              maxTicksLimit: 48,
              font: {
                size: 10,
                weight: "500",
                family: FONT_UI,
              },
              maxRotation: 60,
              minRotation: 45,
              color: "#231F20",
              padding: 4,
            },
          },
          y: {
            beginAtZero: true,
            grace: "5%",
            ticks: { font: { size: 9 }, color: "#6D6E71" },
            grid: { color: "rgba(109, 110, 113, 0.18)" },
            title: {
              display: true,
              text: yTitle,
              font: { size: 10, family: FONT_UI },
              color: "#6D6E71",
            },
          },
        },
      },
    });
    const hint = document.getElementById("chart-bu-compare-hint");
    if (hint) hint.textContent = "(" + vsShort + " · 2 bars per BU)";
    const titleEl = document.getElementById("chart-bu-compare-title");
    if (titleEl)
      titleEl.textContent =
        shortKpiHeaderLabel(kMeta.kpiName) + " — comparison by business";
  }

  const HAZARD_CHART_FILL = [
    "rgba(0, 109, 182, 0.78)",
    "rgba(0, 177, 107, 0.78)",
    "rgba(142, 39, 143, 0.72)",
    "rgba(240, 76, 35, 0.75)",
    "rgba(59, 130, 246, 0.7)",
    "rgba(234, 179, 8, 0.8)",
    "rgba(99, 102, 241, 0.72)",
    "rgba(236, 72, 153, 0.7)",
  ];

  function hazardChartColors(n) {
    const out = [];
    for (let i = 0; i < n; i++) {
      out.push(HAZARD_CHART_FILL[i % HAZARD_CHART_FILL.length]);
    }
    return out;
  }

  /** One dataset per selected KPI; `kind` is "line" or "bar" (hazard trend). */
  function trendDatasetsForPreview(catKey, f, filteredTrend, lineLabels, kind) {
    const keys = effectiveKpiKeysForChartSeries(catKey, f);
    const kpisLine = getKpis(catKey);
    return keys.map((kk, idx) => {
      const kMeta = kpisLine.find((x) => String(x.kpiKey) === kk);
      const ut = kMeta ? kMeta.unitType : "";
      const data = lineLabels.map((ym) => {
        const slice = filteredTrend.filter(
          (r) => r.yearMonth === ym && String(r.kpiKey) === kk
        );
        if (!slice.length) return null;
        return avg(slice.map((r) => Number(r.value)));
      });
      const color = TREND_LINE_COLORS[idx % TREND_LINE_COLORS.length];
      const label = kMeta ? kpiDropdownLabel(kMeta) : kk;
      if (kind === "bar") {
        return {
          label: label,
          data: data,
          unitType: ut,
          backgroundColor: hexToRgba(color, 0.55),
          borderColor: color,
          borderWidth: 1,
          borderRadius: 4,
        };
      }
      return {
        label: label,
        data: data,
        unitType: ut,
        borderColor: color,
        backgroundColor: hexToRgba(color, 0.12),
        fill: true,
        tension: 0.2,
        borderWidth: 1.5,
        pointRadius: 2,
      };
    });
  }

  /** SPI first panel: hazard spotting & closure heatmap (week or month columns). */
  function renderSpiHazardHeatmap(f) {
    const host = document.getElementById("spi-hazard-heatmap-host");
    const foot = document.getElementById("spi-hazard-heatmap-foot");
    if (!host) return;
    const modeBtn = document.querySelector(
      ".spi-hm-period-btn.spi-hm-period-btn--active"
    );
    const mode =
      modeBtn && modeBtn.getAttribute("data-spi-hm-period") === "month"
        ? "month"
        : "week";
    const ref = (f.monthTo || f.refMonth) || getRefMonth();
    const bizTag = f.business === "all" ? "All" : String(f.business || "All");
    const monthShort = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];
    let colLabels;
    if (mode === "week") {
      colLabels = ["Week 1", "Week 2", "Week 3", "Week 4"];
    } else {
      colLabels = [];
      for (let i = -5; i <= 0; i++) {
        const ym = monthAdd(ref, i);
        if (!ym) continue;
        const p = ym.split("-");
        colLabels.push(
          (monthShort[Number(p[1]) - 1] || p[1]) + " " + String(p[0]).slice(2)
        );
      }
    }
    const colCount = colLabels.length;
    const rowTitles = ["Hazard spotting", "Hazard closure"];
    const matrix = [];
    for (let ri = 0; ri < rowTitles.length; ri++) {
      const row = [];
      for (let ci = 0; ci < colCount; ci++) {
        const seed = ref + "|" + bizTag + "|" + mode + "|" + ri + "|" + ci;
        const raw = -5 + (hash32(seed + "|v") % 200) / 10;
        let display;
        if (ci === colCount - 1) {
          const p = -15 + (hash32(seed + "|p") % 31);
          display = (p >= 0 ? "+ " : "− ") + Math.abs(p) + "%";
        } else {
          display =
            Math.abs(raw) < 1 ? String(raw.toFixed(4)) : String(raw.toFixed(2));
        }
        row.push({ display: display, raw: raw });
      }
      matrix.push(row);
    }

    function normForRow(ri, invert) {
      const row = matrix[ri];
      const end = row.length - 1;
      const vals = [];
      for (let i = 0; i < end; i++) vals.push(row[i].raw);
      const mn = Math.min.apply(null, vals);
      const mx = Math.max.apply(null, vals);
      return function (raw) {
        if (mx <= mn) return 0.5;
        let t = (raw - mn) / (mx - mn);
        if (invert) t = 1 - t;
        return t;
      };
    }
    const normSpot = normForRow(0, false);
    const normClose = normForRow(1, true);

    function classForCell(ri, ci, raw) {
      if (ci === colCount - 1) return "spi-hm-cell spi-hm-cell--pct";
      const t = ri === 0 ? normSpot(raw) : normClose(raw);
      if (ri === 0) {
        if (t < 0.34) return "spi-hm-cell spi-hm--spot-1";
        if (t < 0.67) return "spi-hm-cell spi-hm--spot-2";
        return "spi-hm-cell spi-hm--spot-3";
      }
      if (t < 0.34) return "spi-hm-cell spi-hm--cls-1";
      if (t < 0.67) return "spi-hm-cell spi-hm--cls-2";
      return "spi-hm-cell spi-hm--cls-3";
    }

    let html =
      '<table class="spi-hazard-heatmap-table" role="grid"><thead><tr><th class="spi-hm-corner" scope="col"></th>';
    for (let c = 0; c < colCount; c++) {
      html +=
        '<th scope="col" class="spi-hm-colhead">' +
        escapeHtml(colLabels[c]) +
        "</th>";
    }
    html += "</tr></thead><tbody>";
    for (let r = 0; r < rowTitles.length; r++) {
      html +=
        '<tr><th scope="row" class="spi-hm-rowhead">' +
        escapeHtml(rowTitles[r]) +
        "</th>";
      for (let c = 0; c < colCount; c++) {
        const cell = matrix[r][c];
        html +=
          '<td class="' +
          classForCell(r, c, cell.raw) +
          '">' +
          escapeHtml(cell.display) +
          "</td>";
      }
      html += "</tr>";
    }
    html += "</tbody></table>";
    host.innerHTML = html;

    if (foot) {
      const n = 8 + (hash32(ref + bizTag + "|spiHmFoot") % 15);
      const mom = -5 + (hash32(ref + bizTag + "|spiHmMom") % 21);
      const momStr =
        (mom >= 0 ? "▲" : "▼") + Math.abs(mom) + "% MoM";
      foot.innerHTML =
        "Total hazard-linked observations reached <strong>" +
        n +
        "</strong> (" +
        momStr +
        ") · <strong>Hazard closure</strong> indicators for the selected filters (preview sample).";
    }
  }

  /**
   * Hazard & Observation (leading): bar trend, horizontal BU bars, vertical doughnut (same 3-chart layout as other domains).
   */
  function buildHazardObservationCharts(
    f,
    poolAll,
    snapRows,
    winMonths,
    filteredTrend,
    utChart
  ) {
    const catKey = HAZARD_CATEGORY_KEY;
    const kMetaPrimary = getKpis(catKey).find(
      (x) => String(x.kpiKey) === String(f.kpi)
    );
    const lineSeriesName = kMetaPrimary
      ? kpiDropdownLabel(kMetaPrimary)
      : "Value";

    /** Primary panel: same week/month heatmap as Safety Performance Indices (Hazard category). */
    renderSpiHazardHeatmap(f);

    const byBizVals = {};
    snapRows.forEach((r) => {
      const b = r.businessName || "—";
      if (!byBizVals[b]) byBizVals[b] = [];
      byBizVals[b].push(Number(r.value));
    });
    function aggBizWindowForPreviewLabel(previewName) {
      const fact = factBusinessNameForPreview(previewName);
      const vals = byBizVals[fact] || byBizVals[previewName];
      if (!vals || !vals.length) return 0;
      if (utChart && isAdditiveUnit(utChart)) {
        return vals.reduce((a, x) => a + Number(x), 0);
      }
      return avg(vals);
    }
    function shortBu(name) {
      const s = String(name).trim();
      if (s.length <= 10) return s;
      return s.slice(0, 9) + "\u2026";
    }
    const buNames = PREVIEW_BUSINESS_NAMES.slice();
    const buShort = buNames.map(shortBu);
    const buVals = buNames.map((n) => aggBizWindowForPreviewLabel(n));
    const elBiz = document.getElementById("chart-biz");
    const emptyBiz = document.getElementById("chart-biz-empty");
    if (elBiz && typeof Chart !== "undefined") {
      if (!buVals.some((v) => v !== 0)) {
        elBiz.style.display = "none";
        if (emptyBiz) {
          emptyBiz.hidden = false;
          emptyBiz.textContent = "No BU data for current filters.";
        }
      } else {
        elBiz.style.display = "block";
        if (emptyBiz) emptyBiz.hidden = true;
        new Chart(elBiz, {
          type: "bar",
          data: {
            labels: buShort,
            datasets: [
              {
                label: lineSeriesName,
                data: buVals,
                backgroundColor: hazardChartColors(buVals.length),
                borderColor: "rgba(35, 31, 32, 0.28)",
                borderWidth: 1,
              },
            ],
          },
          options: {
            indexAxis: "y",
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 6, right: 14, bottom: 6, left: 4 } },
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  title(items) {
                    const i = items[0].dataIndex;
                    return buNames[i] != null ? String(buNames[i]) : "";
                  },
                  label(ctx) {
                    const v = Number(ctx.raw) || 0;
                    return " " + formatValue(v, utChart);
                  },
                },
              },
            },
            scales: {
              x: {
                beginAtZero: true,
                grace: "6%",
                ticks: { font: { size: 9, family: FONT_UI }, color: "#231F20" },
                grid: { color: "rgba(109, 110, 113, 0.18)" },
              },
              y: {
                ticks: { font: { size: 8, family: FONT_UI }, color: "#231F20" },
                grid: { display: false },
              },
            },
          },
        });
      }
    }

    const vertMap = {};
    snapRows.forEach((r) => {
      const vx = getRowCheckpoint(r);
      if (!vertMap[vx]) vertMap[vx] = [];
      vertMap[vx].push(Number(r.value));
    });
    const vertLabelsFull = Object.keys(vertMap).sort((a, b) =>
      a.localeCompare(b)
    );
    function shortVertLabel(name) {
      const s = String(name);
      return s.length > 12 ? s.slice(0, 11) + "\u2026" : s;
    }
    const vertLabels = vertLabelsFull.map(shortVertLabel);
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
    if (elVert && vertLabels.length && typeof Chart !== "undefined") {
      new Chart(elVert, {
        type: "doughnut",
        data: {
          labels: vertLabels,
          datasets: [
            {
              data: vertData,
              backgroundColor: hazardChartColors(vertData.length),
              borderColor: "#ffffff",
              borderWidth: 2,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: "56%",
          layout: { padding: { top: 4, right: 8, bottom: 4, left: 8 } },
          plugins: {
            legend: {
              display: true,
              position: "bottom",
              labels: {
                boxWidth: 10,
                padding: 6,
                font: { size: 8, family: FONT_UI },
              },
            },
            tooltip: {
              callbacks: {
                title(items) {
                  const i = items[0].dataIndex;
                  return vertLabelsFull[i] || "";
                },
                label(ctx) {
                  const v = Number(ctx.raw) || 0;
                  return " " + formatValue(v, utChart);
                },
              },
            },
          },
        },
      });
    }
  }

  /**
   * Semi-circular % speedometer (theme-aligned arc: Adani_Safety_MIS_Theme.json).
   */
  function drawPercentSpeedometerCanvas(canvasEl, percentRaw) {
    const p = Math.max(0, Math.min(100, Number(percentRaw) || 0));
    const dpr = window.devicePixelRatio || 1;
    const rect = canvasEl.getBoundingClientRect();
    let cssW = rect.width;
    let cssH = rect.height;
    if (cssW < 2 || cssH < 2) {
      const wrap = canvasEl.parentElement;
      if (wrap) {
        cssW = wrap.clientWidth || 280;
        cssH = wrap.clientHeight || 200;
      } else {
        cssW = 280;
        cssH = 200;
      }
    }
    canvasEl.width = Math.max(1, Math.round(cssW * dpr));
    canvasEl.height = Math.max(1, Math.round(cssH * dpr));
    const ctx = canvasEl.getContext("2d");
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);

    const padX = Math.max(10, cssW * 0.06);
    const padB = Math.max(14, cssH * 0.12);
    const cx = cssW / 2;
    const cy = cssH - padB;
    const R = Math.min(cssW - 2 * padX, (cssH - padB) * 1.05) * 0.48;
    const rFill = R * 0.88;
    const rTrack = R * 0.97;

    function pToTheta(pct) {
      return Math.PI + (Math.PI * pct) / 100;
    }

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, rFill, Math.PI, 2 * Math.PI, false);
    ctx.closePath();
    const grad = ctx.createRadialGradient(cx, cy, 0, cx, cy - R * 0.35, rFill);
    grad.addColorStop(0, "rgba(13, 148, 136, 0.22)");
    grad.addColorStop(0.5, "rgba(0, 163, 163, 0.14)");
    grad.addColorStop(1, "rgba(248, 250, 252, 0.92)");
    ctx.fillStyle = grad;
    ctx.fill();
    ctx.restore();

    const segColors = SPEEDOMETER_ARC_HEX;
    const gap = 0.04;
    for (let i = 0; i < 5; i++) {
      const p0 = i * 20 + gap * 10;
      const p1 = (i + 1) * 20 - gap * 10;
      const a0 = pToTheta(p0);
      const a1 = pToTheta(p1);
      ctx.beginPath();
      ctx.arc(cx, cy, rTrack, a0, a1, false);
      ctx.strokeStyle = segColors[i];
      ctx.lineWidth = Math.max(7, R * 0.09);
      ctx.lineCap = "round";
      ctx.stroke();
    }

    const theta = pToTheta(p);
    const needleLen = R * 0.78;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.lineTo(
      cx + needleLen * Math.cos(theta),
      cy + needleLen * Math.sin(theta)
    );
    ctx.strokeStyle = "#1F2937";
    ctx.lineWidth = Math.max(2, R * 0.028);
    ctx.lineCap = "round";
    ctx.stroke();

    ctx.beginPath();
    ctx.arc(cx, cy, Math.max(5, R * 0.065), 0, 2 * Math.PI);
    ctx.fillStyle = "#374151";
    ctx.fill();

    const labelR = needleLen + Math.max(12, R * 0.14);
    const tx = cx + labelR * Math.cos(theta);
    const ty = cy + labelR * Math.sin(theta);
    ctx.font = "600 " + Math.max(13, Math.round(R * 0.2)) + "px " + FONT_UI;
    ctx.fillStyle = "#111827";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(Math.round(p) + "%", tx, ty);
  }

  function renderPercentSpeedometerChart(canvasEl, percent) {
    canvasEl.setAttribute("data-speedometer-gauge", "1");
    canvasEl.setAttribute(
      "aria-label",
      "Speedometer: selected KPI rate for the current Vs window"
    );
    let lastPct = Number(percent) || 0;
    function redraw() {
      drawPercentSpeedometerCanvas(canvasEl, lastPct);
    }
    redraw();
    window.__adaniSpeedometerGaugeRedraw = redraw;
  }

  /**
   * Consequence Management / KPIs 24–25: CMP accountability — three 100% stacked
   * bars (preview proportions; match design reference).
   */
  function renderCmpAccountabilityBreakdownChart(canvasEl) {
    canvasEl.setAttribute(
      "aria-label",
      "CMP accountability breakdown: three stacked percentage bars"
    );
    function seg(label, color, triple) {
      return {
        label: label,
        data: triple,
        backgroundColor: color,
        borderWidth: 0,
        stack: "cmp",
      };
    }
    new Chart(canvasEl, {
      type: "bar",
      data: {
        labels: ["", "Action Category", "Job Band"],
        datasets: [
          seg("CMP Done", "#36D391", [72, null, null]),
          seg("Training", "#5DA5F9", [null, 32, null]),
          seg("PPE", "#A78BFA", [null, 22, null]),
          seg("Procedure", "#22D3EE", [null, 18, null]),
          seg("L1-L3", "#FBBF24", [null, null, 25]),
          seg("L4-L6", "#FB923C", [null, null, 30]),
          seg("L7+", "#FDE047", [null, null, 17]),
          {
            label: "Not Done",
            data: [28, 28, 28],
            backgroundColor: "#F87171",
            borderWidth: 0,
            stack: "cmp",
            borderRadius: { topLeft: 4, topRight: 4 },
            borderSkipped: false,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 8, right: 14, bottom: 2, left: 8 } },
        datasets: {
          bar: {
            barPercentage: 0.52,
            categoryPercentage: 0.72,
          },
        },
        plugins: {
          legend: {
            display: true,
            position: "bottom",
            labels: {
              usePointStyle: true,
              pointStyle: "circle",
              boxWidth: 7,
              padding: 10,
              font: { size: 8.5, family: FONT_UI },
            },
          },
          tooltip: {
            filter: function (item) {
              const v = item.raw;
              return v != null && !Number.isNaN(Number(v));
            },
            callbacks: {
              label: function (ctx) {
                const v = ctx.parsed.y;
                if (v == null || Number.isNaN(v)) return "";
                return " " + ctx.dataset.label + ": " + v + "%";
              },
            },
          },
        },
        scales: {
          x: {
            stacked: true,
            ticks: {
              font: { size: 9, family: FONT_UI },
              color: "#231F20",
              maxRotation: 0,
            },
            grid: { display: false },
            border: {
              display: true,
              color: "rgba(109, 110, 113, 0.45)",
            },
          },
          y: {
            stacked: true,
            min: 0,
            max: 100,
            ticks: {
              stepSize: 25,
              font: { size: 10, family: FONT_UI },
              color: "#231F20",
              padding: 6,
            },
            border: {
              display: true,
              color: "rgba(109, 110, 113, 0.45)",
            },
            grid: {
              color: "rgba(109, 110, 113, 0.35)",
              borderDash: [2, 4],
            },
            title: {
              display: true,
              text: "% of Total Incidents",
              font: { size: 10, family: FONT_UI },
              color: "#6D6E71",
              padding: { top: 0, bottom: 6, left: 0, right: 0 },
            },
          },
        },
      },
    });
  }

  /**
   * Trend + (SPI: quadrant + KPI row-mix) or (by business + by vertical). SPI secondary charts use all KPIs in category; trend line uses the KPI dropdown.
   */
  function buildCharts(catKey, f) {
    destroyCharts();

    const poolAll = getRowsForCategory(catKey);
    const filteredTrend = applyChartFilter(poolAll, f);
    const snapPool = applyNonMonthFilters(poolAll, f, true);
    const winMonths = new Set(
      bizUnitWindowMonths(f.vsMode || DEFAULT_VS_MODE, f.refMonth)
    );
    const snapRows = snapPool.filter((r) => winMonths.has(r.yearMonth));

    const monthSet = new Set();
    filteredTrend.forEach((r) => monthSet.add(r.yearMonth));
    const lineLabels = Array.from(monthSet).sort();

    const kpisLine = getKpis(catKey);
    const kMetaLine = kpisLine.find(
      (x) => String(x.kpiKey) === String(f.kpi)
    );
    const lineSeriesName = kMetaLine
      ? kpiDropdownLabel(kMetaLine)
      : "Value";
    const utChart = kpiUnitTypeForFilter(catKey, f);

    if (catKey === HAZARD_CATEGORY_KEY) {
      buildHazardObservationCharts(
        f,
        poolAll,
        snapRows,
        winMonths,
        filteredTrend,
        utChart
      );
      updateChartHints(f);
      return;
    }

    if (catKey === SPI_CATEGORY_KEY) {
      renderSpiHazardHeatmap(f);
      const spiPoolAllKpis = applyNonMonthFiltersAllKpis(poolAll, f);
      renderSpiRiskQuadrantChart(spiPoolAllKpis, f);
      renderSpiInsightsChart(spiPoolAllKpis, f);
      renderSpiKpiTrendLineChart(poolAll, f);
      updateChartHints(f);
      return;
    }

    const lineSets = trendDatasetsForPreview(
      catKey,
      f,
      filteredTrend,
      lineLabels,
      "line"
    );
    const elLine = document.getElementById("chart-line");
    const lineChartHost = elLine ? elLine.closest(".chart-box") : null;
    if (lineChartHost) {
      lineChartHost.classList.remove("chart-box--cmp-breakdown");
      lineChartHost.classList.remove("chart-box--speedometer-gauge");
      lineChartHost.classList.remove("chart-box--assurance-gauge");
    }

    const showCmpAccountability = shouldShowCmpAccountabilityChart(catKey, f);
    const showSpeedometerGauge = shouldShowPercentSpeedometerChart(catKey, f);
    const speedometerPct = showSpeedometerGauge
      ? Math.max(
          0,
          Math.min(100, speedometerPercentFromSnapRows(catKey, snapRows, f))
        )
      : null;

    if (
      showCmpAccountability &&
      elLine &&
      typeof Chart !== "undefined"
    ) {
      if (lineChartHost) lineChartHost.classList.add("chart-box--cmp-breakdown");
      renderCmpAccountabilityBreakdownChart(elLine);
    } else if (showSpeedometerGauge && elLine) {
      if (lineChartHost) {
        lineChartHost.classList.add("chart-box--speedometer-gauge");
      }
      renderPercentSpeedometerChart(elLine, speedometerPct);
    } else if (elLine && lineLabels.length && lineSets.length) {
      elLine.setAttribute(
        "aria-label",
        "Line chart: monthly average for the selected KPI in the trend window"
      );
      new Chart(elLine, {
        type: "line",
        data: {
          labels: lineLabels,
          datasets: lineSets,
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          layout: {
            padding: { top: 12, right: 16, bottom: 14, left: 22 },
          },
          plugins: {
            legend: {
              display: lineSets.length > 1,
              position: "bottom",
              labels: {
                boxWidth: 10,
                padding: 6,
                font: { size: 9, family: FONT_UI },
              },
            },
            tooltip: {
              callbacks: {
                label(ctx) {
                  const v =
                    ctx.parsed && ctx.parsed.y != null
                      ? ctx.parsed.y
                      : Number(ctx.raw) || 0;
                  const ut = ctx.dataset.unitType || utChart;
                  return " " + formatValue(v, ut);
                },
              },
            },
          },
          scales: {
            y: {
              beginAtZero: false,
              grace: "12%",
              ticks: {
                font: { size: 10 },
                color: "#231F20",
                padding: 6,
              },
              grid: { color: "rgba(109, 110, 113, 0.2)" },
            },
            x: {
              ticks: {
                font: { size: 10 },
                maxRotation: 45,
                color: "#231F20",
                padding: 4,
              },
              grid: { color: "rgba(109, 110, 113, 0.15)" },
            },
          },
        },
      });
    }

    if (catKey === LOCATION_VULN_CAT_KEY) {
      renderLocationVulnerabilityBuMap(snapRows, f);
      updateChartHints(f);
      return;
    }

    const byBizVals = {};
    snapRows.forEach((r) => {
      const b = r.businessName || "—";
      if (!byBizVals[b]) byBizVals[b] = [];
      byBizVals[b].push(Number(r.value));
    });
    function aggBizWindowForPreviewLabel(previewName) {
      const fact = factBusinessNameForPreview(previewName);
      const vals = byBizVals[fact] || byBizVals[previewName];
      if (!vals || !vals.length) return 0;
      if (utChart && isAdditiveUnit(utChart)) {
        return vals.reduce((a, x) => a + Number(x), 0);
      }
      return avg(vals);
    }
    const radarNames = PREVIEW_BUSINESS_NAMES.slice();
    const radarValues = radarNames.map((n) => aggBizWindowForPreviewLabel(n));
    renderBizBreakdown(radarNames, radarValues, lineSeriesName, utChart);

    /** By vertical (checkpoint mapping) — same roll-up window as By business. */
    const vertMap = {};
    snapRows.forEach((r) => {
      const vx = getRowCheckpoint(r);
      if (!vertMap[vx]) vertMap[vx] = [];
      vertMap[vx].push(Number(r.value));
    });
    const vertLabelsFull = Object.keys(vertMap).sort((a, b) =>
      a.localeCompare(b)
    );
    function shortVertLabel(name) {
      const s = String(name);
      return s.length > 14 ? s.slice(0, 13) + "…" : s;
    }
    const vertLabels = vertLabelsFull.map(shortVertLabel);
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
                font: { size: 10, family: FONT_UI },
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
    const lineTitleEl = document.getElementById("chart-line-title");
    const bizEl = document.getElementById("chart-biz-hint");
    const vertEl = document.getElementById("chart-vertical-hint");
    const n = chartMonthsForVsMode(mode, ref).length;
    const showCmpAcc =
      f &&
      shouldShowCmpAccountabilityChart(f.catKey, f);
    const showSpeedometerSpd =
      f && shouldShowPercentSpeedometerChart(f.catKey, f);
    const hmHintEl = document.getElementById("chart-hm-hint");
    if (lineTitleEl && f && f.catKey != null) {
      if (f.catKey === SPI_CATEGORY_KEY) {
        const list = kpiListForFilterDropdown(SPI_CATEGORY_KEY);
        const pk = String(
          f.kpiKeys && f.kpiKeys.length ? f.kpiKeys[0] : f.kpi || ""
        );
        const k = list.find((x) => String(x.kpiKey) === pk);
        lineTitleEl.textContent = k ? kpiDropdownLabel(k) : "SPI KPI trends";
      } else if (showCmpAcc) {
        lineTitleEl.textContent =
          "CMP ACCOUNTABILITY - BREAKDOWN ANALYSIS";
      } else {
        const list = kpiListForFilterDropdown(f.catKey);
        let keys =
          f.kpiKeys && f.kpiKeys.length
            ? f.kpiKeys.map(String)
            : f.kpi
              ? [String(f.kpi)]
              : [];
        if (!keys.length && f.kpi) keys = [String(f.kpi)];
        if (keys.length <= 1) {
          const k = list.find(
            (x) => String(x.kpiKey) === String(keys[0] || f.kpi)
          );
          lineTitleEl.textContent = k ? kpiDropdownLabel(k) : TRI_LABEL_FULL;
        } else {
          const first = list.find((x) => String(x.kpiKey) === keys[0]);
          const a = first ? kpiDropdownLabel(first) : keys[0];
          lineTitleEl.textContent =
            keys.length > 1 ? a + " +" + (keys.length - 1) + " more" : a;
        }
      }
    }
    const isHazard = f && f.catKey === HAZARD_CATEGORY_KEY;
    if (
      hmHintEl &&
      f &&
      (f.catKey === SPI_CATEGORY_KEY || f.catKey === HAZARD_CATEGORY_KEY)
    ) {
      const modeBtn = document.querySelector(
        ".spi-hm-period-btn.spi-hm-period-btn--active"
      );
      const hmMode =
        modeBtn && modeBtn.getAttribute("data-spi-hm-period") === "month"
          ? "month"
          : "week";
      hmHintEl.textContent =
        hmMode === "month" ? "(heatmap · month)" : "(heatmap · week)";
    }
    if (trendEl && f && f.catKey === SPI_CATEGORY_KEY) {
      trendEl.textContent =
        mode === "vs_last_year"
          ? "(lines · 12 mo)"
          : "(lines · " + n + " mo)";
    } else if (trendEl) {
      if (showSpeedometerSpd) {
        trendEl.textContent =
          mode === "vs_last_year"
            ? "(speedometer · 12 mo)"
            : "(speedometer · " + n + " mo)";
      } else if (showCmpAcc) {
        trendEl.textContent = "% of Fatal Incident";
      } else {
        trendEl.textContent =
          mode === "vs_last_year"
            ? "(lines · 12 mo)"
            : "(lines · " + n + " mo)";
      }
    }
    const bizHint = {
      vs_yesterday: "(radar · all BUs · latest mo)",
      vs_last_month: "(radar · all BUs · latest mo)",
      vs_last_week: "(radar · all BUs · 2 mo)",
      vs_last_quarter: "(radar · all BUs · 3 mo)",
      vs_last_year: "(radar · all BUs · YTD window)",
    };
    const bizHintHazard = {
      vs_yesterday: "(h-bars · all BUs · latest mo)",
      vs_last_month: "(h-bars · all BUs · latest mo)",
      vs_last_week: "(h-bars · all BUs · 2 mo)",
      vs_last_quarter: "(h-bars · all BUs · 3 mo)",
      vs_last_year: "(h-bars · all BUs · YTD window)",
    };
    const vertHint = {
      vs_yesterday: "(vertical · latest mo)",
      vs_last_month: "(vertical · latest mo)",
      vs_last_week: "(vertical · 2 mo)",
      vs_last_quarter: "(vertical · 3 mo)",
      vs_last_year: "(vertical · YTD window)",
    };
    const vertHintHazard = {
      vs_yesterday: "(doughnut · latest mo)",
      vs_last_month: "(doughnut · latest mo)",
      vs_last_week: "(doughnut · 2 mo)",
      vs_last_quarter: "(doughnut · 3 mo)",
      vs_last_year: "(doughnut · YTD window)",
    };
    const bh = isHazard
      ? bizHintHazard[mode] || "(h-bars)"
      : bizHint[mode] || "(share)";
    if (bizEl) bizEl.textContent = bh;
    if (vertEl) {
      vertEl.textContent = isHazard
        ? vertHintHazard[mode] || "(doughnut)"
        : vertHint[mode] || "(vertical)";
    }
    const spiBubbleHint = document.getElementById("chart-spi-bubble-hint");
    const spiMapHint = document.getElementById("chart-spi-map-hint");
    const spiRoll =
      mode === "vs_last_year"
        ? "(cross-plot · all SPI KPIs · YTD)"
        : "(cross-plot · all SPI KPIs · Vs window)";
    const spiMapRoll = "(100% stack · all SPI KPIs · 12 mo)";
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
    const kpisMeta = getKpis(catKey);
    const selectedSet = new Set(
      (f.kpiKeys && f.kpiKeys.length ? f.kpiKeys : [f.kpi]).map(String)
    );
    let colKpis = sortKpisForDisplay(catKey, kpisMeta).filter((k) =>
      selectedSet.has(String(k.kpiKey))
    );
    if (!colKpis.length) {
      const dk = defaultKpiKeyForCategory(
        catKey,
        kpiListForFilterDropdown(catKey)
      );
      colKpis = sortKpisForDisplay(catKey, kpisMeta).filter(
        (x) => String(x.kpiKey) === String(dk)
      );
    }
    colKpis = colKpis.slice(0, 12);
    const basePool = applyNonMonthFiltersAllKpis(getRowsForCategory(catKey), f);
    const buRows =
      f.business === "all"
        ? PREVIEW_BUSINESS_NAMES.slice()
        : [f.business].filter(Boolean);

    const thead = document.querySelector("#tbl-detail thead tr");
    if (thead) {
      if (!colKpis.length) {
        thead.innerHTML =
          '<th scope="col" class="bu-matrix__th-bu">BU</th><th scope="col" class="bu-matrix__th-empty">—</th>';
      } else {
        thead.innerHTML =
          '<th scope="col" class="bu-matrix__th-bu">BU</th>' +
          colKpis
            .map(
              (k) =>
                '<th scope="col" class="bu-matrix__th-kpi" title="' +
                escapeAttr(k.kpiName) +
                '"><span class="bu-matrix__kpi-name">' +
                escapeHtml(shortKpiHeaderLabel(k.kpiName)) +
                "</span></th>"
            )
            .join("");
      }
    }

    const tbody = document.getElementById("tbl-body");
    const summaryEl = document.getElementById("tbl-summary");
    const pager = document.querySelector("#view-panel-table .table-pager");

    if (pager) pager.hidden = true;

    if (!tbody) return;

    if (summaryEl) {
      summaryEl.textContent = "";
      summaryEl.removeAttribute("title");
    }

    if (!colKpis.length) {
      tbody.innerHTML =
        '<tr><td colspan="2" class="empty-msg">No KPI definitions for this category.</td></tr>';
      return;
    }

    if (!buRows.length) {
      tbody.innerHTML =
        '<tr><td colspan="' +
        (colKpis.length + 1) +
        '" class="empty-msg">No business row for current filters.</td></tr>';
      return;
    }

    const cellEff = buRows.map((bu) =>
      colKpis.map((k) => {
        const { pct } = buKpiVsPct(basePool, bu, k, f);
        return matrixEffectivePct(pct, catKey, bu, k.kpiKey, f);
      })
    );
    let maxAbs = 0;
    cellEff.forEach((row) => {
      row.forEach((c) => {
        const a = Math.abs(c.value);
        if (a > maxAbs) maxAbs = a;
      });
    });
    if (maxAbs < 1e-9) maxAbs = 1;

    tbody.innerHTML = buRows
      .map((bu, ri) => {
        const cells = colKpis
          .map((k, ki) => {
            const eff = cellEff[ri][ki];
            const pctStr = formatSignedPct(eff.value);
            const divW = matrixBarDivergingWidths(eff.value, maxAbs);
            return (
              '<td class="bu-matrix__cell">' +
              '<div class="bu-cell bu-cell--pbi">' +
              '<div class="bu-cell__databar bu-cell__databar--diverging" role="img" aria-label="' +
              escapeAttr(bu + ", " + k.kpiName + ": " + pctStr) +
              '">' +
              '<div class="bu-cell__track-split">' +
              '<div class="bu-cell__half bu-cell__half--neg">' +
              '<div class="bu-cell__fill bu-cell__fill--neg" style="width:' +
              divW.left +
              '%"></div></div>' +
              '<div class="bu-cell__centerline" aria-hidden="true"></div>' +
              '<div class="bu-cell__half bu-cell__half--pos">' +
              '<div class="bu-cell__fill bu-cell__fill--pos" style="width:' +
              divW.right +
              '%"></div></div></div>' +
              '<span class="bu-cell__value">' +
              escapeHtml(pctStr) +
              "</span></div></div></td>"
            );
          })
          .join("");
        return (
          '<tr class="' +
          (ri % 2 === 0 ? "bu-matrix__row--even" : "bu-matrix__row--odd") +
          '"><th scope="row" class="bu-matrix__bu">' +
          escapeHtml(bu) +
          "</th>" +
          cells +
          "</tr>"
        );
      })
      .join("");
  }

  function initCalendarDateInputs() {
    const mb = dataMonthBounds();
    const elFrom = document.getElementById("f-dt-from");
    const elTo = document.getElementById("f-dt-to");
    if (!elFrom || !elTo) return;
    const minD = mb.min + "-01";
    const maxD = lastDayOfMonthYm(mb.max);
    elFrom.min = minD;
    elFrom.max = maxD;
    elTo.min = minD;
    elTo.max = maxD;
    if (!elTo.value || elTo.value.length < 8) elTo.value = maxD;
    const endYm = elTo.value.slice(0, 7);
    let m = endYm;
    for (let i = 0; i < 11; i++) m = monthAdd(m, -1);
    if (!elFrom.value || elFrom.value.length < 8) {
      const startYm = m < mb.min ? mb.min : m;
      elFrom.value = startYm + "-01";
    }
  }

  function slugExportFilename(s) {
    return String(s || "export")
      .replace(/[^\w\-]+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 48);
  }

  function filterReportTextLines(catKey, f) {
    const cat = getCategory(catKey);
    const lines = [];
    lines.push("Adani Safety Performance Profile — applied filters (export)");
    lines.push("Generated: " + new Date().toISOString());
    lines.push("");
    lines.push("Category: " + (cat ? cat.categoryName : String(catKey)));
    lines.push("Business unit: " + (f.business === "all" ? "All" : f.business));
    lines.push("Site: " + (f.site === "all" ? "All" : f.site));
    lines.push(
      "Calendar: " +
        (f.monthFrom || "—") +
        " → " +
        (f.monthTo || f.refMonth || "—")
    );
    lines.push(
      "Personal type: " + (f.personalType === "all" ? "All" : f.personalType)
    );
    lines.push("Vs: " + vsOptionLabel(f.vsMode || DEFAULT_VS_MODE));
    if (f.state && f.state !== "all") lines.push("State: " + f.state);
    const vert =
      f.variable && f.variable.length
        ? f.variable.join(", ")
        : "All verticals";
    lines.push("Vertical: " + vert);
    const ids = f.kpiKeys && f.kpiKeys.length ? f.kpiKeys : [f.kpi];
    const kStr = ids
      .map((id) => {
        const m = getKpis(catKey).find((x) => String(x.kpiKey) === String(id));
        return m ? m.kpiName : String(id);
      })
      .join("; ");
    lines.push("KPI list: " + kStr);
    return lines;
  }

  function triggerDownloadBlob(blob, filename) {
    const u = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = u;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(u), 1500);
  }

  function downloadChartAsJpeg(canvasId, basename) {
    const el = document.getElementById(canvasId);
    if (!el) return;
    let url;
    if (typeof Chart !== "undefined") {
      const ch = Chart.getChart(el);
      if (ch) url = ch.toBase64Image("image/jpeg", 0.92);
    }
    if (!url && el.tagName === "CANVAS") {
      url = el.toDataURL("image/jpeg", 0.92);
    }
    if (!url) return;
    const a = document.createElement("a");
    a.href = url;
    a.download = slugExportFilename(basename) + ".jpg";
    a.click();
    announce("Downloaded chart as JPEG.");
  }

  function exportChartOrHostToJpeg(id, slug) {
    if (
      id === "chart-loc-bu-map" ||
      id === "chart-spi-map" ||
      id === "spi-hazard-heatmap-capture"
    ) {
      const host = document.getElementById(id);
      if (!host) return;
      captureElementToCanvas(host).then((canvas) => {
        if (!canvas) return;
        canvas.toBlob((blob) => {
          if (blob) {
            triggerDownloadBlob(blob, slugExportFilename(slug) + ".jpg");
          }
        }, "image/jpeg", 0.92);
        announce("Downloaded as JPEG.");
      });
      return;
    }
    downloadChartAsJpeg(id, slug);
  }

  function exportTableToCsv() {
    const tbl = document.getElementById("tbl-detail");
    if (!tbl) return;
    const rows = [...tbl.querySelectorAll("tr")].map((tr) =>
      [...tr.querySelectorAll("th,td")]
        .map((cell) => {
          let t = (cell.textContent || "").trim().replace(/\s+/g, " ");
          if (/[",\n]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
          return t;
        })
        .join(",")
    );
    const blob = new Blob([rows.join("\n")], {
      type: "text/csv;charset=utf-8",
    });
    triggerDownloadBlob(blob, slugExportFilename("bu-table") + ".csv");
    announce("Downloaded table (CSV).");
  }

  function exportTableToJpeg() {
    const panel = document.getElementById("view-panel-table");
    const scroll =
      panel && panel.querySelector(".table-scroll")
        ? panel.querySelector(".table-scroll")
        : document.querySelector(".table-scroll--bu-matrix");
    captureElementToCanvas(scroll || panel).then((canvas) => {
      if (!canvas) return;
      canvas.toBlob((blob) => {
        if (blob) {
          triggerDownloadBlob(blob, slugExportFilename("bu-table") + ".jpg");
        }
      }, "image/jpeg", 0.92);
      announce("Downloaded table (JPEG).");
    });
  }

  async function captureElementToCanvas(el) {
    if (!el) return null;
    if (typeof html2canvas !== "function") {
      announce("Page export needs html2canvas (check network).");
      return null;
    }
    try {
      return await html2canvas(el, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: "#ffffff",
      });
    } catch {
      announce("Capture failed.");
      return null;
    }
  }

  function wireExportDownloads(catKey) {
    const menu = document.getElementById("export-menu");
    if (!menu) return;
    const panel = menu.querySelector(".download-toolbar__panel");
    if (panel) {
      panel.addEventListener("click", (e) => e.stopPropagation());
    }
    function runPdfReport() {
      const f = readFilters(catKey);
      if (!f) return;
      const JSPDF = window.jspdf && window.jspdf.jsPDF;
      if (!JSPDF) {
        announce("PDF export needs jsPDF (check network).");
        return;
      }
      const pdf = new JSPDF({ unit: "pt", format: "a4" });
      const lines = filterReportTextLines(catKey, f);
      let y = 48;
      pdf.setFontSize(11);
      pdf.setTextColor(35, 31, 32);
      lines.forEach((ln) => {
        const parts = pdf.splitTextToSize(ln, 515);
        parts.forEach((p) => {
          if (y > 760) {
            pdf.addPage();
            y = 48;
          }
          pdf.text(p, 40, y);
          y += 14;
        });
      });
      pdf.save(slugExportFilename("filters-report") + ".pdf");
      announce("Downloaded filters report (PDF).");
    }
    async function runPagePdf() {
      const el = document.querySelector(".cat-view");
      const canvas = await captureElementToCanvas(el);
      if (!canvas) return;
      const JSPDF = window.jspdf && window.jspdf.jsPDF;
      if (!JSPDF) {
        announce("PDF export needs jsPDF.");
        return;
      }
      const imgData = canvas.toDataURL("image/jpeg", 0.9);
      const pdf = new JSPDF({
        orientation: "landscape",
        unit: "pt",
        format: "a4",
      });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const iw = canvas.width;
      const ih = canvas.height;
      const r = Math.min(pageW / iw, pageH / ih) * 0.96;
      const w = iw * r;
      const h = ih * r;
      pdf.addImage(imgData, "JPEG", (pageW - w) / 2, (pageH - h) / 2, w, h);
      pdf.save(slugExportFilename("dashboard-page") + ".pdf");
      announce("Downloaded page (PDF).");
    }
    async function runPageJpeg() {
      const el = document.querySelector(".cat-view");
      const canvas = await captureElementToCanvas(el);
      if (!canvas) return;
      canvas.toBlob((blob) => {
        if (blob) {
          triggerDownloadBlob(blob, slugExportFilename("dashboard-page") + ".jpg");
        }
      }, "image/jpeg", 0.92);
      announce("Downloaded page (JPEG).");
    }
    function bind(id, fn) {
      const b = document.getElementById(id);
      if (b) {
        b.addEventListener("click", () => {
          fn();
          try {
            menu.open = false;
          } catch {
            /* ignore */
          }
        });
      }
    }
    bind("dl-page-pdf", runPagePdf);
    bind("dl-page-jpeg", runPageJpeg);
    bind("dl-report-pdf", runPdfReport);
  }

  function wireAppRootExportClicks() {
    if (window.__adaniAppRootExportWired) return;
    window.__adaniAppRootExportWired = true;
    const root = document.getElementById("app-root");
    if (!root) return;
    root.addEventListener("click", function (e) {
      const cbtn = e.target.closest("[data-chart-export]");
      if (cbtn) {
        e.preventDefault();
        const id = cbtn.getAttribute("data-chart-export");
        const slug = cbtn.getAttribute("data-export-slug") || id;
        exportChartOrHostToJpeg(id, slug);
        return;
      }
      const tbtn = e.target.closest("[data-table-export]");
      if (tbtn) {
        e.preventDefault();
        const fmt = tbtn.getAttribute("data-table-export");
        if (fmt === "csv") exportTableToCsv();
        else if (fmt === "jpeg") exportTableToJpeg();
      }
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
    const cat = getCategory(catKey);
    const label = cat ? cat.categoryName + ". " : "";
    const buCount =
      f.business === "all" ? PREVIEW_BUSINESS_NAMES.length : f.business ? 1 : 0;
    const meta = getKpis(catKey);
    const sel = new Set(
      (f.kpiKeys && f.kpiKeys.length ? f.kpiKeys : [f.kpi]).map(String)
    );
    const kpiCols = sortKpisForDisplay(catKey, meta).filter((k) =>
      sel.has(String(k.kpiKey))
    ).length;
    announce(
      label.replace(/\.\s*$/, "") +
        (label ? " · " : "") +
        buCount +
        " · " +
        kpiCols
    );
  }

  function refreshCategoryView(catKey) {
    const kpisMeta = getKpis(catKey);
    const f = readFilters(catKey);
    if (!f) return;

    const aggList = buildKpiDetailMetrics(catKey, kpisMeta, f);

    const presentationSeed = [
      String(catKey),
      f.vsMode || "",
      f.refMonth || "",
      f.state || "",
      f.business || "",
      String(f.kpi || ""),
      JSON.stringify(f.kpiKeys || []),
      JSON.stringify(f.variable || {}),
    ].join("|");

    const multiWrap = document.getElementById("multi-kpi-wrap");
    if (multiWrap) {
      if (!aggList.length) {
        multiWrap.innerHTML =
          '<div class="empty-msg" style="padding:8px">No KPI data for this selection. Adjust Vs or geography filters.</div>';
      } else {
        renderMultiKpiCards(
          multiWrap,
          aggList,
          f.refMonth,
          f.vsMode,
          f.kpiKeys && f.kpiKeys.length ? f.kpiKeys : [f.kpi],
          presentationSeed,
          catKey
        );
      }
    }

    buildCharts(catKey, f);
    buildBuComparisonChart(catKey, f);
    renderTableBody(catKey);
    requestAnimationFrame(() => {
      requestAnimationFrame(() => resizeAllChartsIndex());
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
    tableState = { sortKey: "yearMonth", asc: false, page: 0 };
    destroyCharts();
    setShellLandingMode(false);

    const cat = getCategory(catKey);
    if (!cat) {
      renderCategories();
      return;
    }

    if (CATEGORY_DISABLED_NOT_IN_PREVIEW_KEYS.has(catKey)) {
      history.replaceState(null, "", "#categories");
      renderCategories();
      announce(
        "Leadership and Safety Governance is not included in this preview."
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

    const bizList = mergedBusinessList(
      distinctSorted(rowsForCat, (r) => r.businessName)
    );
    const bizOpts =
      '<option value="all">All business units</option>' +
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
    wrap.className = insightShell ? "cat-view cat-view--modern" : "cat-view";
    const siteOpts =
      '<option value="all">All sites</option>' +
      [1, 2, 3, 4, 5, 6, 7]
        .map(
          (n) =>
            '<option value="Site' +
            n +
            '">Site ' +
            n +
            "</option>"
        )
        .join("");
    const siteFieldHtml =
      '<div class="field"><label class="field-label" for="f-site">Site</label>' +
      '<select id="f-site">' +
      siteOpts +
      "</select></div>";
    const calendarFieldHtml =
      '<div class="field field--calendar-range field--calendar-compact">' +
      '<span class="field-label field-label--inline" id="f-cal-lbl">Calendar</span>' +
      '<div class="field-calendar-range" role="group" aria-labelledby="f-cal-lbl">' +
      '<label class="visually-hidden" for="f-dt-from">From date</label>' +
      '<input type="date" id="f-dt-from" class="toolbar-date" />' +
      '<span class="field-calendar-sep" aria-hidden="true">→</span>' +
      '<label class="visually-hidden" for="f-dt-to">To date</label>' +
      '<input type="date" id="f-dt-to" class="toolbar-date" />' +
      "</div></div>";
    const personalFieldHtml =
      '<div class="field"><label class="field-label" for="f-personal">Personal type</label>' +
      '<select id="f-personal">' +
      '<option value="all">All</option>' +
      '<option value="Employee">Employee</option>' +
      '<option value="Contractor">Contractor</option>' +
      "</select></div>";
    const vsFieldHtml =
      '<div class="field"><label class="field-label" for="f-vs">Vs</label>' +
      '<select id="f-vs">' +
      vsOpts +
      "</select></div>";
    const variableFieldHtml = variableFilterFieldHtml();
    const businessFieldHtml = cfg.showBusiness
      ? '<div class="field"><label class="field-label" for="f-biz">Business unit</label>' +
        '<select id="f-biz">' +
        bizOpts +
        "</select></div>"
      : "";
    const stateFieldHtml = cfg.showState
      ? '<div class="field"><label class="field-label" for="f-state">State</label>' +
        '<select id="f-state">' +
        stateOpts +
        "</select></div>"
      : "";
    const filterCoreInner =
      businessFieldHtml +
      siteFieldHtml +
      calendarFieldHtml +
      personalFieldHtml +
      vsFieldHtml +
      stateFieldHtml;
    const kpiSurfaceHtml =
      cfg.showKpi
        ? '<div class="cat-toolbar__kpi-surface" role="group" aria-labelledby="f-kpi-surface-lbl">' +
          '<span id="f-kpi-surface-lbl" class="cat-toolbar__kpi-surface__lbl">KPI list</span>' +
          '<div class="cat-toolbar__kpi-surface__control">' +
          kpiScopePanelHtml(catKey, kpisForUi) +
          "</div></div>"
        : "";
    const filtersAllScrollHtml =
      '<div class="cat-toolbar__filters-all-scroll">' +
      variableFieldHtml +
      filterCoreInner +
      kpiSurfaceHtml +
      "</div>";
    const lineChartTitleInner =
      '<span class="chart-analytics-title__label" id="chart-line-title">' +
      escapeHtml(TRI_LABEL_FULL) +
      '</span> <span id="chart-trend-hint" class="chart-box__hint">(lines · 12 mo)</span>';
    const lineChartBoxHtml =
      '<div class="chart-box">' +
      chartBlockTitleHtml(lineChartTitleInner, "chart-line", "trend") +
      '<div class="chart-canvas-wrap"><canvas id="chart-line" role="img" aria-label="Line chart: monthly average for the selected KPI in the trend window"></canvas></div></div>';
    const spiLineChartTitleInner =
      '<span class="chart-analytics-title__label" id="chart-line-title">' +
      escapeHtml("SPI KPI trends") +
      '</span> <span id="chart-trend-hint" class="chart-box__hint">(lines · 12 mo)</span>';
    const spiLineChartBoxHtml =
      '<div class="chart-box chart-box--spi-kpi-trend">' +
      chartBlockTitleHtml(spiLineChartTitleInner, "chart-line", "trend") +
      '<div class="chart-canvas-wrap chart-canvas-wrap--spi-trend"><canvas id="chart-line" role="img" aria-label="Line chart: monthly average for each SPI KPI in the trend window"></canvas></div></div>';
    /** Same SPI heatmap visual; used on Leading Hazard for Hazard Spotting (primary panel). */
    const hazardLeadingHeatmapBoxHtml =
      '<div class="chart-box chart-box--spi-hazard-hm chart-box--hazard-leading-hm">' +
      chartBlockTitleHtml(
        '<span class="spi-hm-chart-title" id="spi-hm-heading">Hazard Spotting</span> ' +
        '<span class="chart-box__hint" id="chart-hm-hint">(heatmap · week)</span>',
        "spi-hazard-heatmap-capture",
        "spi-hazard-heatmap"
      ) +
      '<div class="spi-hm-period-bar">' +
      '<span class="spi-hm-period-label">Period</span>' +
      '<div class="spi-hm-period-toggle" role="group" aria-label="Heat map period">' +
      '<button type="button" class="spi-hm-period-btn spi-hm-period-btn--active" data-spi-hm-period="week">Week</button>' +
      '<button type="button" class="spi-hm-period-btn" data-spi-hm-period="month">Month</button>' +
      "</div></div>" +
      '<div class="spi-hazard-heatmap-rule" aria-hidden="true"></div>' +
      '<div id="spi-hazard-heatmap-capture" class="spi-hazard-heatmap-capture">' +
      '<div id="spi-hazard-heatmap-host" class="spi-hazard-heatmap-host"></div>' +
      '<p class="spi-hazard-heatmap-foot" id="spi-hazard-heatmap-foot" aria-live="polite"></p>' +
      "</div></div>";
    const spiHazardHeatmapBoxHtml =
      '<div class="chart-box chart-box--spi-hazard-hm">' +
      chartBlockTitleHtml(
        '<span class="spi-hm-chart-title" id="spi-hm-heading">Hazard category</span> ' +
        '<span class="chart-box__hint" id="chart-hm-hint">(heatmap · week)</span>',
        "spi-hazard-heatmap-capture",
        "spi-hazard-heatmap"
      ) +
      '<div class="spi-hm-period-bar">' +
      '<span class="spi-hm-period-label">Period</span>' +
      '<div class="spi-hm-period-toggle" role="group" aria-label="Heat map period">' +
      '<button type="button" class="spi-hm-period-btn spi-hm-period-btn--active" data-spi-hm-period="week">Week</button>' +
      '<button type="button" class="spi-hm-period-btn" data-spi-hm-period="month">Month</button>' +
      "</div></div>" +
      '<div class="spi-hazard-heatmap-rule" aria-hidden="true"></div>' +
      '<div id="spi-hazard-heatmap-capture" class="spi-hazard-heatmap-capture">' +
      '<div id="spi-hazard-heatmap-host" class="spi-hazard-heatmap-host"></div>' +
      '<p class="spi-hazard-heatmap-foot" id="spi-hazard-heatmap-foot" aria-live="polite"></p>' +
      "</div></div>";
    const hazardChartsRowHtml =
      '<div class="cat-charts cat-charts--hazard" role="group" aria-label="Leading hazard and observation charts: hazard heat map, BU comparison, vertical mix">' +
      hazardLeadingHeatmapBoxHtml +
      '<div class="chart-box chart-box--biz chart-box--hazard-bu">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label">BU comparison</span> <span id="chart-biz-hint" class="chart-box__hint">(bars · all BUs)</span>',
        "chart-biz",
        "bu-comparison"
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--biz"><canvas id="chart-biz" role="img" aria-label="Horizontal bar chart: selected KPI by business unit"></canvas><p id="chart-biz-empty" class="chart-biz-empty" hidden></p></div></div>' +
      '<div class="chart-box chart-box--hazard-vert">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label">Vertical mix</span> <span id="chart-vertical-hint" class="chart-box__hint">(doughnut · Vs window)</span>',
        "chart-verticals",
        "vertical-mix"
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--hazard-doughnut"><canvas id="chart-verticals" role="img" aria-label="Doughnut chart: share of selected KPI by vertical"></canvas></div></div></div>';
    const chartsRowHtml =
      catKey === SPI_CATEGORY_KEY
        ? '<div class="cat-charts-wrap cat-charts-wrap--spi">' +
          '<div class="cat-charts cat-charts--spi" role="group" aria-label="Charts: hazard heat map, quadrant analysis, SPI KPI mix">' +
          spiHazardHeatmapBoxHtml +
          '<div class="chart-box chart-box--spi-bubble">' +
          chartBlockTitleHtml(
            '<span class="chart-analytics-title__label">Quadrant analysis</span> <span id="chart-spi-bubble-hint" class="chart-box__hint">(all SPI KPIs · cross-plot)</span>',
            "chart-spi-bubble",
            "quadrant"
          ) +
          '<div class="chart-canvas-wrap chart-canvas-wrap--spi-bubble"><canvas id="chart-spi-bubble" role="img" aria-label="Bubble quadrant chart: concern reporting rate versus fatality rate by business unit"></canvas></div></div>' +
          '<div class="chart-box chart-box--spi-map">' +
          chartBlockTitleHtml(
            '<span class="chart-analytics-title__label">SPI mix over time</span> <span id="chart-spi-map-hint" class="chart-box__hint">(all SPI KPIs · 100% · 12 mo)</span>',
            "chart-spi-insights",
            "spi-kpi-mix"
          ) +
          '<div class="chart-spi-map-host chart-spi-map-host--insights" id="chart-spi-map" role="region" aria-label="SPI one hundred percent stacked mix by month"></div></div></div>' +
          '<div class="cat-charts cat-charts--spi-trend" role="group" aria-label="SPI KPI trends: monthly line chart for all indices; click a series to focus a KPI">' +
          spiLineChartBoxHtml +
          "</div></div>"
        : catKey === HAZARD_CATEGORY_KEY
          ? hazardChartsRowHtml
        : catKey === LOCATION_VULN_CAT_KEY
          ? '<div class="cat-charts cat-charts--loc-map" role="group" aria-label="Location vulnerability: BU map">' +
            '<div class="chart-box chart-box--loc-bu-map">' +
            chartBlockTitleHtml(
              '<span class="chart-analytics-title__label">By business (India map)</span> <span id="chart-loc-bu-map-hint" class="chart-box__hint">(map · row counts · 24 BUs)</span>',
              "chart-loc-bu-map",
              "india-bu-map"
            ) +
            '<div class="chart-spi-map-host chart-spi-map-host--loc-bu" id="chart-loc-bu-map" role="presentation" aria-label="Map of India: circle size by filtered row count per business unit"></div></div></div>'
          : '<div class="cat-charts" role="group" aria-label="Charts for filtered data">' +
            lineChartBoxHtml +
            '<div class="chart-box chart-box--biz">' +
            chartBlockTitleHtml(
              '<span class="chart-analytics-title__label">By business</span> <span id="chart-biz-hint" class="chart-box__hint">(radar · all BUs)</span>',
              "chart-biz",
              "by-business-radar"
            ) +
            '<div class="chart-canvas-wrap chart-canvas-wrap--biz"><canvas id="chart-biz" role="img" aria-label="Radar chart: KPI value by business unit across all preview businesses"></canvas><p id="chart-biz-empty" class="chart-biz-empty" hidden></p></div></div>' +
            '<div class="chart-box">' +
            chartBlockTitleHtml(
              '<span class="chart-analytics-title__label">By vertical</span> <span id="chart-vertical-hint" class="chart-box__hint">(vertical · latest mo)</span>',
              "chart-verticals",
              "by-vertical"
            ) +
            '<div class="chart-canvas-wrap"><canvas id="chart-verticals" role="img" aria-label="Values by vertical for the selected KPI and Versus window"></canvas></div></div></div>';
    const compareTabHtml =
      '<button type="button" role="tab" id="view-tab-compare" class="view-tabs__btn" aria-selected="false" aria-controls="view-panel-compare" tabindex="-1" data-view="compare">Comparison view</button>';
    const comparePanelHtml =
      '<div id="view-panel-compare" class="view-panel view-panel--compare" role="tabpanel" aria-labelledby="view-tab-compare" hidden><div class="chart-box chart-box--bu-compare">' +
      '<h3 class="chart-analytics-title chart-analytics-title--toolbar">' +
      '<span class="chart-analytics-title__text">' +
      '<span class="chart-analytics-title__label" id="chart-bu-compare-title">KPI comparison by business</span> <span id="chart-bu-compare-hint" class="chart-box__hint">(grouped bars · all BUs)</span>' +
      "</span>" +
      chartDownloadButton("chart-bu-compare", "comparison-by-bu") +
      '</h3><div class="chart-canvas-wrap chart-canvas-wrap--bu-compare"><canvas id="chart-bu-compare" role="img" aria-label="Grouped bar chart: comparison period versus current KPI value for each business unit"></canvas></div></div></div>';
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
      '<div class="cat-toolbar__inner" role="group">' +
      '<div class="cat-toolbar__filters-scroll">' +
      filtersAllScrollHtml +
      "</div>" +
      '<div class="toolbar-actions">' +
      '<details class="download-toolbar" id="export-menu">' +
      '<summary class="toolbar-download-icon" title="Download page or filter report" aria-label="Download page or filter report">' +
      '<svg width="20" height="20" viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
      '<path fill="currentColor" d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>' +
      "</svg></summary>" +
      '<div class="download-toolbar__panel download-toolbar__panel--compact" role="menu" aria-label="Page export">' +
      '<button type="button" class="download-toolbar__btn" id="dl-page-pdf" role="menuitem">Page (PDF)</button>' +
      '<button type="button" class="download-toolbar__btn" id="dl-page-jpeg" role="menuitem">Page (JPEG)</button>' +
      '<button type="button" class="download-toolbar__btn" id="dl-report-pdf" role="menuitem">Filters report (PDF)</button>' +
      "</div></details>" +
      '<button type="button" class="btn btn--reset-compact" id="f-reset">Reset</button>' +
      "</div></div></fieldset>" +
      "</div>" +
      (catKey === LOCATION_VULN_CAT_KEY
        ? ""
        : '<fieldset class="kpi-summary-region">' +
          '<legend class="visually-hidden">KPI summary for current filters</legend>' +
          '<div class="multi-kpi-row" id="multi-kpi-wrap"></div>' +
          "</fieldset>") +
      '<div class="cat-main-view cat-main-view--charts" id="cat-main-view" data-view="charts">' +
      '<div class="view-tabs" role="tablist" aria-label="Chart, table, or comparison view">' +
      '<button type="button" role="tab" id="view-tab-charts" class="view-tabs__btn view-tabs__btn--active" aria-selected="true" aria-controls="view-panel-charts" data-view="charts">Chart view</button>' +
      '<button type="button" role="tab" id="view-tab-table" class="view-tabs__btn" aria-selected="false" aria-controls="view-panel-table" tabindex="-1" data-view="table">Table view</button>' +
      compareTabHtml +
      "</div>" +
      '<div id="view-panel-charts" class="view-panel view-panel--charts" role="tabpanel" aria-labelledby="view-tab-charts">' +
      chartsRowHtml +
      "</div>" +
      '<div id="view-panel-table" class="view-panel view-panel--table" role="tabpanel" aria-labelledby="view-tab-table" hidden>' +
      '<div class="table-zone">' +
      '<div class="table-zone__head">' +
      '<div class="table-zone__title">' +
      '<span class="table-zone__label">BU performance</span>' +
      '<span id="tbl-summary" class="table-zone__summary" aria-live="polite"></span>' +
      "</div>" +
      '<div class="table-zone__exports" role="group" aria-label="Export table">' +
      '<button type="button" class="chart-export-btn" data-table-export="csv" title="Download table (CSV)" aria-label="Download table as CSV">' +
      EXPORT_TABLE_CSV_SVG +
      "</button>" +
      '<button type="button" class="chart-export-btn" data-table-export="jpeg" title="Download table (JPEG)" aria-label="Download table as JPEG">' +
      EXPORT_DL_ICON_SVG +
      "</button></div>" +
      '<div class="table-pager" hidden>' +
      '<button type="button" id="tbl-prev" aria-label="Previous page">Prev</button>' +
      '<span id="tbl-pageinfo"></span>' +
      '<button type="button" id="tbl-next" aria-label="Next page">Next</button>' +
      "</div></div>" +
      '<div class="table-scroll table-scroll--bu-matrix" tabindex="0">' +
      '<table class="data-table bu-matrix" id="tbl-detail">' +
      "<thead><tr>" +
      '<th scope="col" class="bu-matrix__th-bu">BU</th>' +
      "</tr></thead>" +
      '<tbody id="tbl-body"></tbody></table>' +
      "</div></div></div>" +
      comparePanelHtml +
      "</div>" +
      '<p class="cat-context" id="cat-context">' +
      escapeHtml(cat.uxNote) +
      "</p>";

    if (insightShell) {
      wrap.innerHTML = wrap.innerHTML
        .replace('class="breadcrumb"', 'class="breadcrumb m2-breadcrumb"')
        .replace('class="cat-heading"', 'class="cat-heading m2-cat-heading"')
        .replace(
          '<div class="table-zone">',
          '<div class="table-zone m2-table-zone m2-evidence-zone">'
        )
        .replace('class="cat-context"', 'class="cat-context m2-cat-context"');
    }

    root.innerHTML = "";
    root.appendChild(wrap);

    document.getElementById("f-vs").value = DEFAULT_VS_MODE;
    if (cfg.showState) document.getElementById("f-state").value = "all";
    if (cfg.showBusiness) document.getElementById("f-biz").value = "all";
    const siteEl0 = document.getElementById("f-site");
    if (siteEl0) siteEl0.value = "all";
    const personalEl0 = document.getElementById("f-personal");
    if (personalEl0) personalEl0.value = "all";
    applyVariableFilterFromStorage();
    initCalendarDateInputs();

    function onFilterChange() {
      tableState.page = 0;
      refreshCategoryView(catKey);
    }

    [
      "f-vs",
      "f-state",
      "f-biz",
      "f-site",
      "f-personal",
      "f-dt-from",
      "f-dt-to",
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", onFilterChange);
    });
    wireVariableFilterControls(onFilterChange);
    wireToolbarScopeScrollPanels(wrap);

    if (catKey === SPI_CATEGORY_KEY || catKey === HAZARD_CATEGORY_KEY) {
      wrap.querySelectorAll(".spi-hm-period-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
          wrap.querySelectorAll(".spi-hm-period-btn").forEach((b) => {
            b.classList.remove("spi-hm-period-btn--active");
          });
          btn.classList.add("spi-hm-period-btn--active");
          refreshCategoryView(catKey);
        });
      });
    }

    if (cfg.showKpi) {
      updateKpiScopeCount();
      document
        .querySelectorAll('#f-kpi-panel input[name="f-kpi-cb"]')
        .forEach((cb) => {
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
            .forEach((c) => {
              c.checked = true;
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
            .forEach((c) => {
              c.checked = false;
            });
          saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
          updateKpiScopeCount();
          onFilterChange();
        });
      }
    }

    document.getElementById("f-reset").addEventListener("click", () => {
      document.getElementById("f-vs").value = DEFAULT_VS_MODE;
      if (cfg.showKpi) {
        try {
          localStorage.removeItem(LS_KPI_PREFIX + catKey);
        } catch {
          /* ignore */
        }
        const def = loadKpiSelection(catKey, kpisForUi);
        document
          .querySelectorAll('#f-kpi-panel input[name="f-kpi-cb"]')
          .forEach((cb) => {
            cb.checked = def.includes(String(cb.value));
          });
        saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
        updateKpiScopeCount();
      }
      if (cfg.showState) document.getElementById("f-state").value = "all";
      if (cfg.showBusiness) document.getElementById("f-biz").value = "all";
      const siteEl = document.getElementById("f-site");
      if (siteEl) siteEl.value = "all";
      const personalEl = document.getElementById("f-personal");
      if (personalEl) personalEl.value = "all";
      initCalendarDateInputs();
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
    wireCatMainViewIndex();
    wireExportDownloads(catKey);
    refreshCategoryView(catKey);
    const h = document.getElementById("cat-heading");
    if (h) h.focus();
    updateHeaderNavState();
  }

  function renderLanding() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#landing");
    setShellLandingMode(true);

    const box = document.createElement("div");
    box.className =
      "landing landing--bg-only" +
      (insightShell ? " landing--modern" : "");
    box.setAttribute("role", "region");
    box.setAttribute("aria-label", "Adani Safety Performance Profile home");
    box.innerHTML =
      '<div class="landing__cta">' +
      '<button type="button" class="landing__start" id="btn-landing-start">' +
      '<span class="landing__start-text">Start Now</span>' +
      '<svg class="landing__start-arrow" width="16" height="16" viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
      '<path fill="currentColor" d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8-8-8z"/>' +
      "</svg>" +
      "</button></div>";

    root.innerHTML = "";
    root.appendChild(box);
    setHeaderLastUpdatedToday();
    const startBtn = document.getElementById("btn-landing-start");
    if (startBtn) {
      try {
        startBtn.focus({ preventScroll: true });
      } catch {
        startBtn.focus();
      }
      startBtn.addEventListener("click", () => {
        history.replaceState(null, "", "#categories");
        renderCategories();
      });
    }
    announce(
      insightShell
        ? "Insights home. Choose Start Now to browse safety domains."
        : "Adani Safety Performance Profile home."
    );
    updateHeaderNavState();
  }

  /** Categories list styled for Insights (`m2-cat-directory` rows). */
  function renderCategoriesInsightsDirectory() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#categories");
    setShellLandingMode(false);

    const box = document.createElement("div");
    box.className =
      "home-body home-body--modern m2-cat-page m2-cat-page--directory";
    box.innerHTML =
      '<div class="m2-cat-dir-head">' +
      journeyStepsHtml(2) +
      '<div class="m2-cat-dir-intro">' +
      '<h2 id="home-h">Safety domains</h2>' +
      '<p class="m2-cat-dir-lede">Search by <strong>category or KPI name</strong>. All domains below open the same charts, table, comparison, and exports as the main dashboard preview—presented in the Insights layout.</p>' +
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
        const interactive = !CATEGORY_DISABLED_NOT_IN_PREVIEW_KEYS.has(
          cat.categoryKey
        );
        const el = interactive
          ? document.createElement("button")
          : document.createElement("div");
        if (interactive) el.type = "button";
        const order = String(cat.sortOrder || cat.categoryKey).padStart(2, "0");
        el.className =
          "m2-cat-row m2-cat-row--k" +
          cat.categoryKey +
          (interactive ? " m2-cat-row--open" : " m2-cat-row--disabled");
        el.setAttribute("role", "listitem");
        el.setAttribute(
          "aria-label",
          interactive
            ? cat.categoryName + ", " + cat.kpiCount + " KPIs, open dashboard"
            : cat.categoryName +
                ", " +
                cat.kpiCount +
                " KPIs, not in preview"
        );
        if (!interactive) {
          el.setAttribute("aria-disabled", "true");
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
          '<span class="m2-cat-row__chip m2-cat-row__chip--' +
          (interactive ? "live" : "muted") +
          '">' +
          (interactive ? "Available" : "Not in Preview") +
          "</span>" +
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
          (interactive ? "Open" : "—") +
          "</span>" +
          "</span>";
        if (interactive) {
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

  function renderCategories() {
    if (insightShell) {
      renderCategoriesInsightsDirectory();
      return;
    }
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(null, "", "#categories");
    setShellLandingMode(false);

    const box = document.createElement("div");
    box.className = "home-body home-body--launchpad";
    box.innerHTML =
      journeyStepsHtml(2) +
      '<div class="home-intro home-intro--launchpad">' +
      '<h2 id="home-h">Categories</h2>' +
      '<p class="home-lede">All <strong>safety domains</strong> below are available in this preview—including <strong>Incident Management</strong>, <strong>Hazard &amp; Observation (Leading)</strong>, <strong>Assurance</strong>, <strong>Safety Performance Indices</strong>, and <strong>Location Vulnerability</strong>. Use <strong>Home</strong> in the header to return to the start screen.</p>' +
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
        const active = !CATEGORY_DISABLED_NOT_IN_PREVIEW_KEYS.has(
          cat.categoryKey
        );
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
                " KPIs. Not in preview."
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
          (active ? "Available" : "Not in Preview") +
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
          (active ? "Explore" : "—") +
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

  wireAppRootExportClicks();

  normalizePreviewFactRowsVsBand();

  if (insightShell && meta.lastUpdateISO) {
    setHeaderTimestamp(meta.lastUpdateISO);
  } else {
    setHeaderLastUpdatedToday();
  }

  const m0 = location.hash.match(/cat=(\d+)/);
  if (m0) renderCategory(parseInt(m0[1], 10));
  else if (location.hash === "#categories") renderCategories();
  else renderLanding();
})();
