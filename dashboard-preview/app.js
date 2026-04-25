/** Adani Safety Performance Dashboard — HTML preview (fixed canvas; open via local server). */
(function () {
  "use strict";

  const root = document.getElementById("app-root");
  const liveRegion = document.getElementById("sr-live");
  const DATA = window.__DASHBOARD_DATA__;
  /** Matches styles.css --font (Segoe UI family + UI sans fallbacks). */
  const FONT_UI =
    '"Segoe UI", "Segoe UI Variable", "Segoe UI Historic", system-ui, sans-serif';
  /** Softer than pure black — axes, legends, ticks (pairs with light surfaces in styles.css). */
  const CHART_INK = "#334155";

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

  /** Fixed preview stamp — same across sessions (aligns with static help pages). */
  function setHeaderLastUpdatedToday() {
    const hu = document.getElementById("header-updated");
    if (!hu) return;
    hu.textContent = "Tuesday, April 14, 2026";
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
        '<div class="boot-error" role="alert" style="padding:12px 14px;font-size:13px;color:#0f172a;max-width:52rem">' +
        '<h1 class="boot-error__title" style="margin:0 0 10px;font-size:1.05rem;font-weight:700;color:#8e278f">' +
        "Adani Safety Performance Dashboard</h1>" +
        "<p style=\"margin:0 0 10px;line-height:1.45\"><strong>Data not loaded.</strong> " +
        "The site header above should still be visible; this panel explains what went wrong.</p>" +
        "<ul style=\"margin:0;padding-left:1.25rem;line-height:1.5\">" +
        "<li>Place <code>embedded-data.js</code> in the same folder as <code>index.html</code> (dashboard-preview).</li>" +
        "<li>Open via a local HTTP server (e.g. <code>npx serve</code> from that folder) so the browser is not blocking large scripts on <code>file://</code>.</li>" +
        "<li>Regenerate data: run <code>dashboard-preview\\Refresh-PreviewData.ps1</code> (requires PowerBI_Assets CSVs).</li>" +
        "</ul></div>";
    }
    return;
  }

  if (!DATA.factRows) DATA.factRows = [];
  if (!DATA.months) DATA.months = [];
  if (!DATA.states) DATA.states = [];
  if (!DATA.categories) DATA.categories = [];
  if (!DATA.kpiDetailByCategory) {
    DATA.kpiDetailByCategory = [];
  } else if (!Array.isArray(DATA.kpiDetailByCategory)) {
    const raw = DATA.kpiDetailByCategory;
    DATA.kpiDetailByCategory =
      raw && typeof raw === "object" && Object.keys(raw).length
        ? Object.values(raw)
        : [];
  }
  if (!DATA.monthlyByCategory) DATA.monthlyByCategory = [];
  if (!DATA.businessBreakdown) DATA.businessBreakdown = [];

  (function patchTrainingKpiDisplayNames() {
    function fixName(o) {
      if (!o || o.kpiName !== "Training Count") return;
      o.kpiName = "Total Training Conducted";
    }
    (DATA.kpiDetailByCategory || []).forEach((block) => {
      (block.kpis || []).forEach(fixName);
    });
    (DATA.factRows || []).forEach(fixName);
  })();

  /**
   * Category Selection page: display order, names, and card descriptions (workbook "Categories" sheet).
   * sortOrder controls grid order; categoryKey stays stable for routes and data joins.
   * ensureCategoryCatalogueBaseline merges any missing keys from this catalogue into embedded data.
   */
  const CATEGORY_SELECTION_CATALOG_META = [
    {
      categoryKey: 1,
      sortOrder: 1,
      categoryName: "Incident Management",
      uxNote:
        "Tracks occurrence, type, trends, and closure of all safety incidents across all verticals",
    },
    {
      categoryKey: 3,
      sortOrder: 2,
      categoryName: "Safety Performance Indices (SPI)",
      uxNote:
        "Measures normalized safety performance using frequency, severity, and rate-based metrics",
    },
    {
      categoryKey: 4,
      sortOrder: 3,
      categoryName: "Consequence Management",
      uxNote:
        "Tracks disciplinary actions and compliance taken in response to safety incidents",
    },
    {
      categoryKey: 5,
      sortOrder: 4,
      categoryName: "Assurance and Compliance",
      uxNote:
        "Evaluates implementation of safety standards, learnings, and compliance across operations",
    },
    {
      categoryKey: 2,
      sortOrder: 5,
      categoryName: "Hazard & Observation Management",
      uxNote:
        "Monitors proactive risk identification, near misses, unsafe acts, and closure effectiveness",
    },
    {
      categoryKey: 7,
      sortOrder: 6,
      categoryName: "Risk Control Programs",
      uxNote:
        "Monitors effectiveness and closure of structured risk mitigation initiatives",
    },
    {
      categoryKey: 8,
      sortOrder: 7,
      categoryName: "Leadership & Safety Governance",
      uxNote: "",
    },
    {
      categoryKey: 6,
      sortOrder: 8,
      categoryName: "Training & Competency Development",
      uxNote:
        "Evaluates workforce safety capability through training coverage, intensity, and skill development",
    },
    {
      categoryKey: 9,
      sortOrder: 9,
      categoryName: "Digital & Technology Intervention",
      uxNote:
        "Measures adoption and utilization of safety systems, digital tools & technological solutions",
    },
    {
      categoryKey: 10,
      sortOrder: 10,
      categoryName: "Vulnerable Location",
      uxNote:
        "Highlights geographic and site-level exposure patterns to as Resilient & Vulnerable Site prioritize interventions",
    },
  ];

  function normalizeCategorySelectionCatalog() {
    if (!DATA || !Array.isArray(DATA.categories)) return;
    const byKey = new Map(
      CATEGORY_SELECTION_CATALOG_META.map((m) => [Number(m.categoryKey), m])
    );
    DATA.categories.forEach((cat) => {
      const m = byKey.get(Number(cat.categoryKey));
      if (!m) return;
      cat.sortOrder = m.sortOrder;
      cat.categoryName = m.categoryName;
      cat.uxNote = m.uxNote;
    });
    DATA.categories.sort(
      (a, b) => (a.sortOrder || 0) - (b.sortOrder || 0)
    );
  }

  function ensureCategoryCatalogueBaseline() {
    if (!DATA || !Array.isArray(DATA.categories)) return;
    const BASE = CATEGORY_SELECTION_CATALOG_META.map((m) => ({
      categoryKey: m.categoryKey,
      categoryName: m.categoryName,
      sortOrder: m.sortOrder,
      uxNote: m.uxNote,
      kpiCount: 0,
      latestMonthIndex: 0,
    }));
    const present = new Set(DATA.categories.map((c) => c.categoryKey));
    for (let i = 0; i < BASE.length; i++) {
      const b = BASE[i];
      if (!present.has(b.categoryKey)) {
        DATA.categories.push({ ...b });
        present.add(b.categoryKey);
      }
    }
    DATA.categories.sort(
      (a, b) => (a.sortOrder || 0) - (b.sortOrder || 0)
    );
    if (Array.isArray(DATA.kpiDetailByCategory)) {
      DATA.categories.forEach((cat) => {
        const det = DATA.kpiDetailByCategory.find(
          (x) => x.categoryKey === cat.categoryKey
        );
        if (det && Array.isArray(det.kpis) && det.kpis.length) {
          cat.kpiCount = det.kpis.length;
        }
      });
    }
  }

  /** Resolve KPI display name for tooltips and map popups (preview catalogue + fact rows). */
  function kpiNameForKey(kpiKey) {
    const kk = String(kpiKey);
    const blocks = DATA.kpiDetailByCategory || [];
    for (let i = 0; i < blocks.length; i++) {
      const kpis = blocks[i].kpis;
      if (!kpis) continue;
      const hit = kpis.find((x) => String(x.kpiKey) === kk);
      if (hit && hit.kpiName) return hit.kpiName;
    }
    const fr = DATA.factRows;
    if (fr && fr.length) {
      const sample = fr.find((r) => String(r.kpiKey) === kk);
      if (sample && sample.kpiName) return sample.kpiName;
    }
    return "KPI " + kk;
  }

  /** Insights shell (`insights.html`): modern chrome + directory home; same logic as dashboard. */
  const insightShell =
    typeof document !== "undefined" &&
    document.querySelector('.shell--modern[data-layout="insights"]') != null;

  /**
   * Main `index.html` category grid sequence (stable categoryKey). Insights layout unchanged.
   */
  const CLASSIC_INDEX_CATEGORY_GRID_ORDER = [
    10, 1, 3, 2, 7, 4, 6, 5, 8, 9,
  ];

  function classicIndexCategoryGridRank(categoryKey) {
    const k = Number(categoryKey);
    const i = CLASSIC_INDEX_CATEGORY_GRID_ORDER.indexOf(k);
    return i === -1 ? CLASSIC_INDEX_CATEGORY_GRID_ORDER.length + k : i;
  }

  /** Labels for main preview only; `categoryKey` and data joins stay unchanged. */
  function classicIndexCategoryDisplayName(categoryKey, dataName) {
    if (insightShell) return dataName || "";
    const k = Number(categoryKey);
    if (k === 1) return "Incident Data";
    if (k === 10) return "Vulnerable & Resilient Locations";
    return dataName || "";
  }

  if (typeof Chart !== "undefined") {
    Chart.defaults.font.family = FONT_UI;
    Chart.defaults.font.size = 10;
    /** Default ink for all Chart.js text (legends, ticks, titles) */
    Chart.defaults.color = CHART_INK;
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

  /** Synthetic site buckets (aligned with Site filter checkboxes). */
  const PREVIEW_SITE_LABELS = [
    "Site1",
    "Site2",
    "Site3",
    "Site4",
    "Site5",
    "Site6",
    "Site7",
  ];

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
    if (!f || f.business === "all") return true;
    const rb = String(r.businessName || "").trim();
    if (Array.isArray(f.business)) {
      if (!f.business.length) return false;
      return f.business.some((bu) => {
        const sel = String(bu || "").trim();
        if (rb === sel) return true;
        return factBusinessNameForPreview(sel) === rb;
      });
    }
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
  /** Incident-only additions (Dim keys 18 / 19 are used elsewhere — 56 / 57 / 58 here). */
  const INCIDENT_FIRE_KPI_KEY = 56;
  const INCIDENT_PROPERTY_DAMAGE_KPI_KEY = 57;
  const INCIDENT_TOTAL_RECORDABLE_INJURIES_KPI_KEY = 58;

  /** Align packaged JSON meta with header product name (avoid editing embedded-data.js). */
  function previewApplyDataMetaTitles() {
    if (!DATA || !DATA.meta) return;
    DATA.meta.dashboardTitle = "Adani Safety Performance Dashboard";
    DATA.meta.subtitle =
      "Safety performance indicators — Interactive Preview";
  }

  /**
   * Incident Management: add Fire, Property Damage, and Total Recordable Injuries (preview Dim keys).
   * Seeded fact rows; `ensureCategoryCatalogueBaseline` syncs category `kpiCount` from `detail.kpis`.
   */
  function previewAddIncidentFirePropertyKpis() {
    if (!DATA || !Array.isArray(DATA.factRows)) return;
    const INCIDENT_CAT = 1;
    const detail = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === INCIDENT_CAT
    );
    if (!detail || !Array.isArray(detail.kpis)) return;
    const have = new Set(detail.kpis.map((k) => Number(k.kpiKey)));
    if (
      have.has(INCIDENT_FIRE_KPI_KEY) &&
      have.has(INCIDENT_PROPERTY_DAMAGE_KPI_KEY) &&
      have.has(INCIDENT_TOTAL_RECORDABLE_INJURIES_KPI_KEY)
    ) {
      return;
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

    const fireMeta = {
      kpiKey: INCIDENT_FIRE_KPI_KEY,
      kpiName: "Fire incidents Count",
      unitType: "Count",
      latestValue: 0,
    };
    const propMeta = {
      kpiKey: INCIDENT_PROPERTY_DAMAGE_KPI_KEY,
      kpiName: "Property Damage Count",
      unitType: "Count",
      latestValue: 0,
    };
    const triMeta = {
      kpiKey: INCIDENT_TOTAL_RECORDABLE_INJURIES_KPI_KEY,
      kpiName: "Total Recordable Injuries",
      unitType: "Count",
      latestValue: 0,
    };

    const toAdd = [];
    if (!have.has(INCIDENT_FIRE_KPI_KEY)) toAdd.push(fireMeta);
    if (!have.has(INCIDENT_PROPERTY_DAMAGE_KPI_KEY)) toAdd.push(propMeta);
    if (!have.has(INCIDENT_TOTAL_RECORDABLE_INJURIES_KPI_KEY))
      toAdd.push(triMeta);
    if (!toAdd.length) return;

    detail.kpis.push(...toAdd);
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
        for (let ki = 0; ki < toAdd.length; ki++) {
          const kpi = toAdd[ki];
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

  /** Incident Data (category 1): approved KPI titles + sync `kpiDetailByCategory` and fact rows. */
  function previewApplyIncidentKpiDisplayNames() {
    if (!DATA) return;
    const INC_CAT = 1;
    const LABEL = {
      1: "Process safety Incidents",
      2: "Process Safety Near Misses",
      3: "Leak/spill count",
      4: "Injury Illness incidents",
      5: "Repeat incident cases count",
      7: "Dangerous occurrence count",
      8: "Fatality Count",
      9: "LTI Count",
      10: "MTC Count",
      11: "RWC Count",
      12: "FAC Count",
      14: "Delta of LTI count",
      15: "Delta of Repeat incident count",
      19: "Delta of Fatal incident",
      22: "Total Man day lost",
      28: "Total Vehicle incidents count",
      44: "Investigations Closure %",
      56: "Fire incidents Count",
      57: "Property Damage Count",
      58: "Total Recordable Injuries",
    };
    const blocks = Array.isArray(DATA.kpiDetailByCategory)
      ? DATA.kpiDetailByCategory
      : [];
    const detail = blocks.find((x) => Number(x.categoryKey) === INC_CAT);
    if (detail && Array.isArray(detail.kpis)) {
      detail.kpis.forEach((k) => {
        const kk = Number(k.kpiKey);
        if (LABEL[kk]) k.kpiName = LABEL[kk];
      });
    }
    if (Array.isArray(DATA.factRows)) {
      for (let i = 0; i < DATA.factRows.length; i++) {
        const r = DATA.factRows[i];
        if (Number(r.categoryKey) !== INC_CAT) continue;
        const kk = Number(r.kpiKey);
        if (LABEL[kk]) r.kpiName = LABEL[kk];
      }
    }
  }

  /**
   * (Not invoked) Previously merged Risk Control (7) into Hazard / Consequence, which removed
   * category 7 from the Categories list. Kept for reference; the preview shows all domains.
   */
  function previewStripRiskControlPrograms() {
    const RISK_CONTROL_CATEGORY_KEY = 7;
    const HAZARD_CAT = 2;
    const CONSEQUENCE_CAT = 4;
    const KPIS_TO_HAZARD = new Set([41, 42]);
    /** Exclude kpiKey 54 — used by Hazard (SI planned vs actual), not Risk Control merge. */
    const KPIS_TO_CONSEQUENCE = new Set([51, 52]);
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
      const kk = Number(k);
      DATA.categories = (DATA.categories || []).filter(
        (c) => Number(c.categoryKey) !== kk
      );
      DATA.factRows = (DATA.factRows || []).filter(
        (r) => Number(r.categoryKey) !== kk
      );
      if (DATA.monthlyByCategory) {
        DATA.monthlyByCategory = DATA.monthlyByCategory.filter(
          (x) => Number(x.categoryKey) !== kk
        );
      }
      if (DATA.kpiDetailByCategory) {
        DATA.kpiDetailByCategory = DATA.kpiDetailByCategory.filter(
          (x) => Number(x.categoryKey) !== kk
        );
      }
      if (DATA.businessBreakdown) {
        DATA.businessBreakdown = DATA.businessBreakdown.filter(
          (x) => Number(x.categoryKey) !== kk
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
        kpiKey: 503,
        kpiName: "Incident Key Learning Implementation %",
        unitType: "PercentOrRate",
        latestValue: 81,
      },
      {
        kpiKey: 502,
        kpiName:
          "% Standard Implementation against the fatal four top critical Safety Standards",
        unitType: "PercentOrRate",
        latestValue: 72,
      },
      {
        kpiKey: 501,
        kpiName: "FRC compliance Rate",
        unitType: "PercentOrRate",
        latestValue: 88,
      },
    ];
    DATA.categories.push({
      categoryKey: ASSURANCE_CATEGORY_KEY,
      categoryName: "Assurance and Compliance",
      sortOrder: 5,
      uxNote:
        "Evaluates implementation of safety standards, learnings, and compliance across operations",
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
      series: rollupSeries(ASSURANCE_CATEGORY_KEY, 503),
    });
  }

  /** Hazard category: SI planned vs actual (preview-seeded from SI rounds rows). */
  function previewAddHazardSiPlannedActualKpi() {
    if (!DATA || !Array.isArray(DATA.factRows)) return;
    const haz = 2;
    const siKpi = 54;
    const block = DATA.kpiDetailByCategory.find((x) => x.categoryKey === haz);
    if (!block || !Array.isArray(block.kpis)) return;
    const siMeta = {
      kpiKey: siKpi,
      kpiName: "SI (Planned vs Actual)",
      unitType: "PercentOrRate",
      latestValue: 84.2,
    };
    if (!block.kpis.some((k) => Number(k.kpiKey) === siKpi)) {
      block.kpis.push(siMeta);
      const cat = DATA.categories.find((c) => c.categoryKey === haz);
      if (cat) cat.kpiCount = block.kpis.length;
    }
    if (
      DATA.factRows.some(
        (r) => r.categoryKey === haz && Number(r.kpiKey) === siKpi
      )
    ) {
      return;
    }
    const template = DATA.factRows.filter(
      (r) => r.categoryKey === haz && Number(r.kpiKey) === 53
    );
    if (!template.length) return;
    const newRows = template.map((r) => {
      const v = Number(r.value);
      const n = Number.isFinite(v)
        ? Math.min(100, Math.max(0, Math.round(v * 2.35 + 22)))
        : 0;
      return {
        ...r,
        kpiKey: siKpi,
        kpiName: siMeta.kpiName,
        value: n,
      };
    });
    DATA.factRows.push(...newRows);
    const catSync = DATA.categories.find((c) => c.categoryKey === haz);
    if (catSync) catSync.kpiCount = block.kpis.length;
  }

  /**
   * Risk Control (7): remove KPIs that mirror Assurance and Compliance (cat 5) preview measures
   * — Incident Key Learning %, FRC Compliance %, Critical Standards % (Dim keys 51 / 52 / 54).
   * Hazard still uses kpiKey 54 as "SI (Planned vs Actual)" under category 2 only.
   */
  function previewStripRiskControlAssuranceDuplicateKpis() {
    if (!DATA) return;
    const RISK = 7;
    const drop = new Set([51, 52, 54]);
    const detail = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === RISK
    );
    if (detail && Array.isArray(detail.kpis)) {
      detail.kpis = detail.kpis.filter((k) => !drop.has(Number(k.kpiKey)));
    }
    if (Array.isArray(DATA.factRows)) {
      DATA.factRows = DATA.factRows.filter(
        (r) =>
          !(Number(r.categoryKey) === RISK && drop.has(Number(r.kpiKey)))
      );
    }
    const mbc = DATA.monthlyByCategory.find(
      (x) => Number(x.categoryKey) === RISK
    );
    if (mbc && Array.isArray(mbc.series) && Array.isArray(DATA.months)) {
      const monthKeys = DATA.months.map((m) => m.yearMonth).filter(Boolean);
      mbc.series = monthKeys.map((ym) => {
        const nums = DATA.factRows
          .filter(
            (r) =>
              String(r.yearMonth) === String(ym) &&
              Number(r.categoryKey) === RISK
          )
          .map((r) => Number(r.value))
          .filter((v) => Number.isFinite(v));
        const v = nums.length
          ? nums.reduce((a, b) => a + b, 0) / nums.length
          : 0;
        return { yearMonth: ym, value: Math.round(v * 100) / 100 };
      });
    }
    const bb = DATA.businessBreakdown.find(
      (x) => Number(x.categoryKey) === RISK
    );
    if (bb && Array.isArray(bb.bars)) {
      const lastYm =
        DATA.months && DATA.months.length
          ? DATA.months[DATA.months.length - 1].yearMonth
          : null;
      if (lastYm) {
        for (let i = 0; i < bb.bars.length; i++) {
          const b = bb.bars[i];
          const name = String(b.business || "").trim();
          if (!name) continue;
          const sum = DATA.factRows
            .filter(
              (r) =>
                Number(r.categoryKey) === RISK &&
                String(r.yearMonth) === String(lastYm) &&
                String(r.businessName || "").trim() === name
            )
            .reduce((acc, r) => acc + Number(r.value || 0), 0);
          b.value = Math.round(sum * 100) / 100;
        }
      }
    }
  }

  /**
   * Risk Control Programs (7): exactly six KPIs in stakeholder order — Total Actions Closure %,
   * Safety Inspection Observation closure %, Total SRFA Closure %, S-4 / S-5 SRFA closure %, Near Miss FR.
   * Clones facts from Hazard (2) and SPI (3) before Hazard KPI restriction drops shared keys.
   */
  function previewSeedRiskControlProgramsKpis() {
    if (!DATA || !Array.isArray(DATA.factRows)) return;
    const RISK = 7;
    const HAZ = 2;
    const SPI = 3;
    const order = [46, 45, 40, 41, 42, 20];
    const displayNames = {
      46: "Total Actions Closure %",
      45: "Safety Inspection Observation closure %",
      40: "Total SRFA Closure %",
      41: "S-4 SRFA Closure %",
      42: "S-5 SRFA Closure %",
      20: "Near Miss FR",
    };

    function metaFrom(catKey, kpiKey) {
      const b = DATA.kpiDetailByCategory.find(
        (x) => Number(x.categoryKey) === catKey
      );
      if (!b || !Array.isArray(b.kpis)) return null;
      return (
        b.kpis.find((k) => Number(k.kpiKey) === Number(kpiKey)) || null
      );
    }

    function nameFor(id) {
      return displayNames[id] || "KPI " + id;
    }

    const templates = {};
    for (let i = 0; i < order.length; i++) {
      const id = order[i];
      const preferCat =
        id === 41 || id === 42 ? RISK : id === 20 ? SPI : HAZ;
      let rows = DATA.factRows.filter(
        (r) =>
          Number(r.categoryKey) === preferCat && Number(r.kpiKey) === id
      );
      if (!rows.length && (id === 41 || id === 42)) {
        rows = DATA.factRows.filter(
          (r) => Number(r.categoryKey) === HAZ && Number(r.kpiKey) === id
        );
      }
      templates[id] = rows;
    }

    const kpisOut = order.map((id) => {
      const m =
        metaFrom(HAZ, id) || metaFrom(RISK, id) || metaFrom(SPI, id);
      const sampleVal =
        templates[id] && templates[id].length
          ? Number(templates[id][templates[id].length - 1].value)
          : 0;
      const base = m || {
        kpiKey: id,
        kpiName: nameFor(id),
        unitType: "PercentOrRate",
        latestValue: sampleVal,
      };
      return {
        kpiKey: base.kpiKey,
        kpiName: nameFor(id),
        unitType: base.unitType || "PercentOrRate",
        latestValue:
          base.latestValue != null && base.latestValue !== ""
            ? base.latestValue
            : sampleVal,
      };
    });

    let riskDetail = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === RISK
    );
    if (!riskDetail) {
      riskDetail = { categoryKey: RISK, kpis: [] };
      DATA.kpiDetailByCategory.push(riskDetail);
    }
    riskDetail.kpis = kpisOut;

    DATA.factRows = DATA.factRows.filter(
      (r) => Number(r.categoryKey) !== RISK
    );

    for (let i = 0; i < order.length; i++) {
      const id = order[i];
      const rows = templates[id] || [];
      for (let j = 0; j < rows.length; j++) {
        const r = rows[j];
        DATA.factRows.push({
          ...r,
          categoryKey: RISK,
          kpiKey: id,
          kpiName: nameFor(id),
        });
      }
    }

    const catRow = DATA.categories.find((c) => Number(c.categoryKey) === RISK);
    if (catRow) catRow.kpiCount = order.length;

    const mbc = DATA.monthlyByCategory.find(
      (x) => Number(x.categoryKey) === RISK
    );
    if (mbc && Array.isArray(DATA.months)) {
      const monthKeys = DATA.months.map((m) => m.yearMonth).filter(Boolean);
      mbc.series = monthKeys.map((ym) => {
        const nums = DATA.factRows
          .filter(
            (r) =>
              String(r.yearMonth) === String(ym) &&
              Number(r.categoryKey) === RISK
          )
          .map((r) => Number(r.value))
          .filter((v) => Number.isFinite(v));
        const v = nums.length
          ? nums.reduce((a, b) => a + b, 0) / nums.length
          : 0;
        return { yearMonth: ym, value: Math.round(v * 100) / 100 };
      });
    }

    const bb = DATA.businessBreakdown.find(
      (x) => Number(x.categoryKey) === RISK
    );
    if (bb && Array.isArray(bb.bars) && DATA.months && DATA.months.length) {
      const lastYm = DATA.months[DATA.months.length - 1].yearMonth;
      for (let bi = 0; bi < bb.bars.length; bi++) {
        const bar = bb.bars[bi];
        const bn = String(bar.business || "").trim();
        if (!bn) continue;
        const sum = DATA.factRows
          .filter(
            (r) =>
              Number(r.categoryKey) === RISK &&
              String(r.yearMonth) === String(lastYm) &&
              String(r.businessName || "").trim() === bn
          )
          .reduce((acc, r) => acc + Number(r.value || 0), 0);
        bar.value = Math.round(sum * 100) / 100;
      }
    }
  }

  /**
   * Vulnerable Location (10): catalogue row + KPI list cloned from Hazard map measures (no extra facts).
   */
  function previewSeedVulnerableLocationCategory() {
    if (!DATA || !Array.isArray(DATA.factRows)) return;
    const LOC = 10;
    const HAZ = 2;
    const LV_ORDER = [13, 38, 39, 40, 45, 46, 53];
    DATA.factRows = DATA.factRows.filter(
      (r) =>
        !(
          Number(r.categoryKey) === LOC &&
          [601, 602, 603].includes(Number(r.kpiKey))
        )
    );
    DATA.kpiDetailByCategory = DATA.kpiDetailByCategory.filter(
      (x) => Number(x.categoryKey) !== LOC
    );
    const haz = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === HAZ
    );
    const locKpis = [];
    if (haz && Array.isArray(haz.kpis)) {
      const map = new Map(haz.kpis.map((k) => [Number(k.kpiKey), k]));
      for (let i = 0; i < LV_ORDER.length; i++) {
        const k = map.get(LV_ORDER[i]);
        if (k) {
          locKpis.push({
            kpiKey: k.kpiKey,
            kpiName: k.kpiName,
            unitType: k.unitType || "Count",
            latestValue: k.latestValue,
          });
        }
      }
    }
    if (!locKpis.length) {
      locKpis.push(
        {
          kpiKey: 13,
          kpiName: "Near Miss Count",
          unitType: "Count",
          latestValue: 0,
        },
        {
          kpiKey: 38,
          kpiName: "Hazard Spotting Rate",
          unitType: "Count",
          latestValue: 0,
        }
      );
    }
    DATA.kpiDetailByCategory.push({
      categoryKey: LOC,
      kpis: locKpis,
    });
    DATA.monthlyByCategory = (DATA.monthlyByCategory || []).filter(
      (x) => Number(x.categoryKey) !== LOC
    );
    DATA.businessBreakdown = (DATA.businessBreakdown || []).filter(
      (x) => Number(x.categoryKey) !== LOC
    );
    const meta = CATEGORY_SELECTION_CATALOG_META.find(
      (m) => Number(m.categoryKey) === LOC
    );
    let catRow = DATA.categories.find((c) => Number(c.categoryKey) === LOC);
    if (!catRow && meta) {
      catRow = {
        categoryKey: LOC,
        categoryName: meta.categoryName,
        sortOrder: meta.sortOrder,
        uxNote: meta.uxNote,
        kpiCount: locKpis.length,
        latestMonthIndex: 1,
      };
      DATA.categories.push(catRow);
    } else     if (catRow) {
      catRow.kpiCount = locKpis.length;
      catRow.latestMonthIndex = 1;
    }
  }

  /**
   * Hazard & Observation Management: stakeholder KPIs in fixed order
   * (Observations → Closure % → SI planned vs actual → unsafe acts/hr in SI → Near miss → Life saved).
   * Runs after Vulnerable Location seed (which clones the fuller hazard catalogue first).
   */
  function previewRestrictHazardObservationKpis() {
    if (!DATA || !Array.isArray(DATA.kpiDetailByCategory)) return;
    const haz = 2;
    const allowOrder = [38, 39, 54, 53, 13, 6];
    const allowSet = new Set(allowOrder);
    const displayNames = {
      38: "Hazard Spotting Observations",
      39: "Hazard Spotting Closure %",
      54: "SI (Planned vs Actual)",
      53: "Unsafe Acts Identified per hour in SI Round",
      13: "Near Miss Count",
      6: "Life Saved Cases Count",
    };
    const block = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === haz
    );
    if (!block || !Array.isArray(block.kpis)) return;
    const byKey = new Map(block.kpis.map((k) => [Number(k.kpiKey), k]));
    block.kpis = allowOrder
      .map((id) => {
        const row = byKey.get(id);
        if (!row) return null;
        const label = displayNames[id];
        if (label) row.kpiName = label;
        return row;
      })
      .filter(Boolean);
    const catRow = DATA.categories.find((c) => Number(c.categoryKey) === haz);
    if (catRow) catRow.kpiCount = block.kpis.length;

    if (Array.isArray(DATA.factRows)) {
      DATA.factRows = DATA.factRows.filter((r) => {
        if (Number(r.categoryKey) !== haz) return true;
        return allowSet.has(Number(r.kpiKey));
      });
      DATA.factRows.forEach((r) => {
        if (Number(r.categoryKey) !== haz) return;
        const lab = displayNames[Number(r.kpiKey)];
        if (lab) r.kpiName = lab;
      });
    }
  }

  /**
   * Consequence Management (4): exactly five KPIs in stakeholder order — Total CMP action Taken,
   * Fatality vs CMP, Job band, Category, LTI vs CMP. Drops any merged Risk Control duplicates (e.g. 51/52).
   */
  function previewRestrictConsequenceManagementKpis() {
    if (!DATA || !Array.isArray(DATA.kpiDetailByCategory)) return;
    const cons = 4;
    const allowOrder = [25, 23, 27, 26, 24];
    const allowSet = new Set(allowOrder);
    const displayNames = {
      25: "Total CMP action Taken",
      23: "Fatality vs CMP action Compliance",
      27: "CMP action Job band wise",
      26: "CMP action category wise",
      24: "LTI vs CMP action Compliance",
    };
    const block = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === cons
    );
    if (!block || !Array.isArray(block.kpis)) return;
    const byKey = new Map(block.kpis.map((k) => [Number(k.kpiKey), k]));
    block.kpis = allowOrder
      .map((id) => {
        const row = byKey.get(id);
        if (!row) return null;
        const label = displayNames[id];
        if (label) row.kpiName = label;
        return row;
      })
      .filter(Boolean);
    const catRow = DATA.categories.find((c) => Number(c.categoryKey) === cons);
    if (catRow) catRow.kpiCount = block.kpis.length;

    if (Array.isArray(DATA.factRows)) {
      DATA.factRows = DATA.factRows.filter((r) => {
        if (Number(r.categoryKey) !== cons) return true;
        return allowSet.has(Number(r.kpiKey));
      });
      DATA.factRows.forEach((r) => {
        if (Number(r.categoryKey) !== cons) return;
        const lab = displayNames[Number(r.kpiKey)];
        if (lab) r.kpiName = lab;
      });
    }

    const mbc = DATA.monthlyByCategory.find(
      (x) => Number(x.categoryKey) === cons
    );
    if (mbc && Array.isArray(DATA.months)) {
      const monthKeys = DATA.months.map((m) => m.yearMonth).filter(Boolean);
      mbc.series = monthKeys.map((ym) => {
        const nums = DATA.factRows
          .filter(
            (r) =>
              String(r.yearMonth) === String(ym) &&
              Number(r.categoryKey) === cons
          )
          .map((r) => Number(r.value))
          .filter((v) => Number.isFinite(v));
        const v = nums.length
          ? nums.reduce((a, b) => a + b, 0) / nums.length
          : 0;
        return { yearMonth: ym, value: Math.round(v * 100) / 100 };
      });
    }

    const bb = DATA.businessBreakdown.find(
      (x) => Number(x.categoryKey) === cons
    );
    if (bb && Array.isArray(bb.bars) && DATA.months && DATA.months.length) {
      const lastYm = DATA.months[DATA.months.length - 1].yearMonth;
      for (let bi = 0; bi < bb.bars.length; bi++) {
        const bar = bb.bars[bi];
        const bn = String(bar.business || "").trim();
        if (!bn) continue;
        const sum = DATA.factRows
          .filter(
            (r) =>
              Number(r.categoryKey) === cons &&
              String(r.yearMonth) === String(lastYm) &&
              String(r.businessName || "").trim() === bn
          )
          .reduce((acc, r) => acc + Number(r.value || 0), 0);
        bar.value = Math.round(sum * 100) / 100;
      }
    }
  }

  /**
   * Training & Competency Development (6): five KPIs in stakeholder order with catalogue labels.
   */
  function previewRestrictTrainingCompetencyDevelopmentKpis() {
    if (!DATA || !Array.isArray(DATA.kpiDetailByCategory)) return;
    const tr = 6;
    const allowOrder = [34, 35, 37, 36, 49];
    const allowSet = new Set(allowOrder);
    const displayNames = {
      34: "Total Training Conducted",
      35: "Training Intensity Rate",
      37:
        "Executive training Index (Functional Managers, Managing Managers, Managing Others & Managing Self )",
      36: "Frontline Worker Training Index",
      49: "SAKSHAM Implementation %",
    };
    const block = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === tr
    );
    if (!block || !Array.isArray(block.kpis)) return;
    const byKey = new Map(block.kpis.map((k) => [Number(k.kpiKey), k]));
    block.kpis = allowOrder
      .map((id) => {
        const row = byKey.get(id);
        if (!row) return null;
        const label = displayNames[id];
        if (label) row.kpiName = label;
        return row;
      })
      .filter(Boolean);
    const catRow = DATA.categories.find((c) => Number(c.categoryKey) === tr);
    if (catRow) catRow.kpiCount = block.kpis.length;

    if (Array.isArray(DATA.factRows)) {
      DATA.factRows = DATA.factRows.filter((r) => {
        if (Number(r.categoryKey) !== tr) return true;
        return allowSet.has(Number(r.kpiKey));
      });
      DATA.factRows.forEach((r) => {
        if (Number(r.categoryKey) !== tr) return;
        const lab = displayNames[Number(r.kpiKey)];
        if (lab) r.kpiName = lab;
      });
    }

    const mbc = DATA.monthlyByCategory.find(
      (x) => Number(x.categoryKey) === tr
    );
    if (mbc && Array.isArray(DATA.months)) {
      const monthKeys = DATA.months.map((m) => m.yearMonth).filter(Boolean);
      mbc.series = monthKeys.map((ym) => {
        const nums = DATA.factRows
          .filter(
            (r) =>
              String(r.yearMonth) === String(ym) &&
              Number(r.categoryKey) === tr
          )
          .map((r) => Number(r.value))
          .filter((v) => Number.isFinite(v));
        const v = nums.length
          ? nums.reduce((a, b) => a + b, 0) / nums.length
          : 0;
        return { yearMonth: ym, value: Math.round(v * 100) / 100 };
      });
    }

    const bb = DATA.businessBreakdown.find(
      (x) => Number(x.categoryKey) === tr
    );
    if (bb && Array.isArray(bb.bars) && DATA.months && DATA.months.length) {
      const lastYm = DATA.months[DATA.months.length - 1].yearMonth;
      for (let bi = 0; bi < bb.bars.length; bi++) {
        const bar = bb.bars[bi];
        const bn = String(bar.business || "").trim();
        if (!bn) continue;
        const sum = DATA.factRows
          .filter(
            (r) =>
              Number(r.categoryKey) === tr &&
              String(r.yearMonth) === String(lastYm) &&
              String(r.businessName || "").trim() === bn
          )
          .reduce((acc, r) => acc + Number(r.value || 0), 0);
        bar.value = Math.round(sum * 100) / 100;
      }
    }
  }

  /**
   * Leadership & Safety Governance (8): four KPIs in stakeholder order; keeps multi-KPI / standard charts.
   */
  function previewRestrictLeadershipSafetyGovernanceKpis() {
    if (!DATA || !Array.isArray(DATA.kpiDetailByCategory)) return;
    const ld = 8;
    const allowOrder = [47, 48, 43, 55];
    const allowSet = new Set(allowOrder);
    const displayNames = {
      47: "Business Safety Council Meeting",
      48: "6 Taskforce governance meeting",
      43: "Leadership Safety Walkthroughs",
      55: "Leadership Safety Observations",
    };
    const block = DATA.kpiDetailByCategory.find(
      (x) => Number(x.categoryKey) === ld
    );
    if (!block || !Array.isArray(block.kpis)) return;
    const byKey = new Map(block.kpis.map((k) => [Number(k.kpiKey), k]));
    block.kpis = allowOrder
      .map((id) => {
        const row = byKey.get(id);
        if (!row) return null;
        const label = displayNames[id];
        if (label) row.kpiName = label;
        return row;
      })
      .filter(Boolean);
    const catRow = DATA.categories.find((c) => Number(c.categoryKey) === ld);
    if (catRow) catRow.kpiCount = block.kpis.length;

    if (Array.isArray(DATA.factRows)) {
      DATA.factRows = DATA.factRows.filter((r) => {
        if (Number(r.categoryKey) !== ld) return true;
        return allowSet.has(Number(r.kpiKey));
      });
      DATA.factRows.forEach((r) => {
        if (Number(r.categoryKey) !== ld) return;
        const lab = displayNames[Number(r.kpiKey)];
        if (lab) r.kpiName = lab;
      });
    }

    const mbc = DATA.monthlyByCategory.find(
      (x) => Number(x.categoryKey) === ld
    );
    if (mbc && Array.isArray(DATA.months)) {
      const monthKeys = DATA.months.map((m) => m.yearMonth).filter(Boolean);
      mbc.series = monthKeys.map((ym) => {
        const nums = DATA.factRows
          .filter(
            (r) =>
              String(r.yearMonth) === String(ym) &&
              Number(r.categoryKey) === ld
          )
          .map((r) => Number(r.value))
          .filter((v) => Number.isFinite(v));
        const v = nums.length
          ? nums.reduce((a, b) => a + b, 0) / nums.length
          : 0;
        return { yearMonth: ym, value: Math.round(v * 100) / 100 };
      });
    }

    const bb = DATA.businessBreakdown.find(
      (x) => Number(x.categoryKey) === ld
    );
    if (bb && Array.isArray(bb.bars) && DATA.months && DATA.months.length) {
      const lastYm = DATA.months[DATA.months.length - 1].yearMonth;
      for (let bi = 0; bi < bb.bars.length; bi++) {
        const bar = bb.bars[bi];
        const bn = String(bar.business || "").trim();
        if (!bn) continue;
        const sum = DATA.factRows
          .filter(
            (r) =>
              Number(r.categoryKey) === ld &&
              String(r.yearMonth) === String(lastYm) &&
              String(r.businessName || "").trim() === bn
          )
          .reduce((acc, r) => acc + Number(r.value || 0), 0);
        bar.value = Math.round(sum * 100) / 100;
      }
    }
  }

  previewApplyDataMetaTitles();
  previewAddIncidentFirePropertyKpis();
  previewApplyIncidentKpiDisplayNames();
  previewCategoryDataBootstrap();
  previewAddHazardSiPlannedActualKpi();
  previewStripRiskControlAssuranceDuplicateKpis();
  previewRestrictConsequenceManagementKpis();
  previewSeedRiskControlProgramsKpis();
  previewSeedVulnerableLocationCategory();
  previewRestrictHazardObservationKpis();
  previewRestrictTrainingCompetencyDevelopmentKpis();
  previewRestrictLeadershipSafetyGovernanceKpis();
  ensureCategoryCatalogueBaseline();
  normalizeCategorySelectionCatalog();

  /** Detail table rows per page (keep in sync with styles.css --detail-table-body-rows) */
  const PAGE_SIZE = 5;
  let tableState = {
    sortKey: "yearMonth",
    asc: false,
    page: 0,
  };

  let currentCategoryKey = null;
  let catSearchAnnounceTimer = null;
  /** Retries `buildBuComparisonChart` when Chart.js is still loading (defer). */
  let vlBuCompareChartAttempts = 0;

  const INCIDENT_MANAGEMENT_CATEGORY_KEY = 1;
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
  const LEADERSHIP_CATEGORY_KEY = 8;
  /** Risk Control Programs — preview roster matches stakeholder workbook order (six KPIs). */
  const RISK_CONTROL_PROGRAMS_CATEGORY_KEY = 7;
  /** Main dashboard (`index.html`) only — mandatory meeting types for compliance matrix. */
  const LEADERSHIP_MEETING_TYPE_LABELS = [
    "Executive safety board",
    "BU safety council",
    "Cross-functional review",
    "Contractor safety forum",
    "Incident learning review",
    "Training & competency forum",
    "Audit action review",
  ];
  /** Month-view drill-down: department pillars (wireframe). */
  const LEADERSHIP_DEPT_LABELS = [
    "BSC",
    "SRP",
    "Contractor",
    "Training",
    "Logistic",
    "Incident",
    "Audit",
  ];
  const VULNERABLE_LOCATION_CATEGORY_KEY = 10;
  /** Hazard KPIs aggregated on the Vulnerable Location map (see KPI reference). */
  const VULNERABLE_LOCATION_KPI_ORDER = [13, 38, 39, 40, 45, 46, 53];
  const VULNERABLE_LOCATION_KPI_KEYS = new Set(VULNERABLE_LOCATION_KPI_ORDER);
  /** Categories that appear on the list but cannot be opened (none in current preview). */
  const CATEGORY_DISABLED_NOT_IN_PREVIEW_KEYS = new Set();
  /**
   * Leading vs lagging — used on the Categories list (#categories) chips only.
   */
  const CATEGORY_NATURE_BY_KEY = {
    1: {
      kind: "lagging",
      label: "Lagging",
      blurb:
        "Tracks occurrence, type, trends, and closure of all safety incidents across all verticals.",
    },
    2: {
      kind: "leading",
      label: "Leading",
      blurb:
        "Monitors proactive risk identification, near misses, unsafe acts, and closure effectiveness.",
    },
    3: {
      kind: "lagging",
      label: "Lagging (rate-based)",
      blurb:
        "Measures normalized safety performance using frequency, severity, and rate-based metrics.",
    },
    4: {
      kind: "lagging",
      label: "Lagging",
      blurb:
        "Tracks disciplinary actions and compliance taken in response to safety incidents.",
    },
    5: {
      kind: "lagging",
      label: "Lagging",
      blurb:
        "Evaluates implementation of safety standards, learnings, and compliance across operations.",
    },
    6: {
      kind: "leading",
      label: "Leading",
      blurb:
        "Evaluates workforce safety capability through training coverage, intensity, and skill development.",
    },
    7: {
      kind: "leading",
      label: "Leading / preventive",
      blurb:
        "Monitors effectiveness and closure of structured risk mitigation initiatives.",
    },
    8: {
      kind: "strategic",
      label: "Strategic Overview",
      blurb: "",
    },
    9: {
      kind: "leading",
      label: "Leading",
      blurb:
        "Measures adoption and utilization of safety systems, digital tools & technological solutions.",
    },
    10: {
      kind: "strategic",
      label: "Strategic Overview",
      blurb:
        "Highlights geographic and site-level exposure patterns to as Resilient & Vulnerable Site prioritize interventions.",
    },
  };

  /** Subtitle under main title on category drill (#cat=…): category name + (Lagging)/(Leading)/(Strategic Overview). */
  function categoryNatureBracketSuffix(catKey) {
    const n = CATEGORY_NATURE_BY_KEY[Number(catKey)];
    if (!n) return "";
    if (n.kind === "lagging") return " (Lagging)";
    if (n.kind === "leading") return " (Leading)";
    if (n.kind === "strategic")
      return " (" + (n.label || "Strategic Overview") + ")";
    return "";
  }

  function categoryListChipLeadingLagging(catKey) {
    const n = CATEGORY_NATURE_BY_KEY[catKey];
    if (!n) return { label: "—", chipKind: "leading" };
    if (n.kind === "lagging")
      return { label: "Lagging", chipKind: "lagging" };
    if (n.kind === "strategic")
      return { label: n.label || "Strategic Overview", chipKind: "leading" };
    return { label: "Leading", chipKind: "leading" };
  }
  /** Adani approved palette — #00B16B #006DB6 #8E278F #F04C23 #E6E7E8 (charts: KPI dashboard only). */
  const ADANI_GREEN = "#00B16B";
  const ADANI_BLUE = "#006DB6";
  const ADANI_PURPLE = "#8E278F";
  const ADANI_ORANGE = "#F04C23";
  const ADANI_GREY = "#E6E7E8";
  /** Percent speedometer arc: concern → OK using brand hues. */
  const SPEEDOMETER_ARC_HEX = [
    ADANI_ORANGE,
    "#FF8A50",
    "#F5C400",
    ADANI_GREEN,
    ADANI_BLUE,
  ];
  /** Incident Key Learning (503): dual Group vs business gauge when selected in Assurance. */
  const INCIDENT_KEY_LEARNING_SPEEDOMETER_KPI_KEY = "503";
  const EVENT_LEVEL_LABELS = [
    "0 Near Miss",
    "1 Minor",
    "2 Moderate",
    "3 Serious",
    "4 Major",
    "5 Catastrophic",
    "No Level",
  ];
  /** Safety Performance Indices (category 3): Fatality Rate → LTIFR → LTISR → TRIR → Vehicle Incident Rate → Near Miss Rate. */
  const SPI_KPI_ORDER = [16, 17, 18, 21, 29, 20];
  /** Heat map row order (matches reference: Fatality Rate → LTIFR → LTISR → Near Miss Rate → TRIR → Vehicle Incident Rate). */
  const SPI_HEATMAP_KPI_ORDER = [16, 17, 18, 20, 21, 29];
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

  function formatBusinessSiteRowLabel(bu) {
    return bu;
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
        return vsOptionLabel(id);
    }
  }

  function periodComparisonTooltip(f) {
    return (
      comparisonVsPeriodCaption(f) +
      ". Value compares totals or averages across the Current Period vs the Comparison Period, by KPI unit rules."
    );
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

  function escapeHtmlMultiline(s) {
    if (s == null) return "";
    return String(s)
      .split("\n")
      .map((line) => escapeHtml(line))
      .join("<br>");
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

  /**
   * Chart title row: label + hint, optional “i” info (hover/focus shows parameters),
   * JPEG export. infoText: plain lines separated by \n (escaped).
   */
  function chartBlockTitleHtml(innerSpansHtml, dataId, slug, infoText) {
    const sid = escapeAttr(String(slug || dataId || "chart"));
    const insightOnly = infoText ? chartHelpInsightOnly(infoText) : "";
    const tip =
      insightOnly &&
      '<span class="chart-info-wrap">' +
        '<button type="button" class="chart-info-btn" aria-describedby="ti-' +
        sid +
        '" aria-label="What this chart shows">i</button>' +
        '<span class="chart-info-tooltip" id="ti-' +
        sid +
        '" role="tooltip">' +
        escapeHtmlMultiline(insightOnly) +
        "</span></span>";
    return (
      '<h3 class="chart-analytics-title chart-analytics-title--toolbar">' +
      '<span class="chart-analytics-title__text">' +
      innerSpansHtml +
      "</span>" +
      (tip || "") +
      chartDownloadButton(dataId, slug) +
      "</h3>"
    );
  }

  const CHART_HELP = {
    trend:
      "Insight: Month-on-month trend for the active KPI.\nUses: Selected KPI(s), Current Period, Comparison Period, Business, State, Site, Personnel Type, Verticals.\nExport: JPEG via camera icon.",
    spiTrend:
      "Insight: How each Safety Performance Index KPI moves over the last months.\nUses: All SPI KPIs in parallel, same global filters and calendar end month as the page.\nExport: JPEG.",
    hazardHeat:
      "Insight: Leading hazard activity — density by week or month from preview heat map seed.\nUses: Hazard period toggle, Current Period, Business / Vertical filters.\nExport: JPEG.",
    hazardBiz:
      "Insight: Compare the selected KPI across business units (horizontal bars).\nUses: KPI scope checkboxes, Business filter, Current / Comparison periods.\nExport: JPEG.",
    hazardVert:
      "Insight: Share of the selected KPI across Field / O&M / Office / Projects for the Current Period.\nUses: Vertical scope, KPI selection, filters.\nExport: JPEG.",
    spiHeat:
      "Insight: Six-month grid of SPI indices — each cell is the monthly average for your filters; shading runs light→dark within that KPI row (compare months for the same measure).\nRows use three color families: green (Fatality Rate / LTIFR / LTISR / vehicle), red (Near Miss Rate), orange (TRIR).\nUses: All SPI KPIs, same filters as the rest of the page.\nExport: JPEG via camera icon.",
    spiMix:
      "Insight: 100% stacked trend of SPI KPI mix by month — see which indices dominate over time.\nUses: All SPI KPIs, filters.\nExport: JPEG.",
    bizRadar:
      "Insight: Radar comparing the same KPI across every preview business — shape shows balance.\nUses: Primary KPI from scope, Current Period months.\nExport: JPEG.",
    vertical:
      "Insight: Line across verticals for the active KPI (latest month slice).\nUses: KPI selection, Vertical scope, Business / State filters.\nExport: JPEG.",
    compareBu:
      "Insight: Grouped bars — Comparison Period vs Current Period for the active KPI, every BU.\nUses: Current and Comparison period ranges, KPI selection, filters.\nExport: JPEG.",
    compareBuVl:
      "Insight: Grouped bars by business unit — same KPI aggregated for the current period, split between sites in resilient states (top 5) and vulnerable states (bottom 5) using the same map ranking.\nUses: Current period window, KPI selection, business / site / vertical filters.\nExport: JPEG.",
  };

  /** Chart "i" tooltip: keep Insight block only (drops Uses / Export lines). */
  function chartHelpInsightOnly(raw) {
    if (raw == null || raw === "") return "";
    const s = String(raw);
    const lower = s.toLowerCase();
    const start = lower.indexOf("insight:");
    if (start === -1) {
      return s
        .split("\n")
        .filter(
          (line) =>
            !/^\s*uses:\s*/i.test(line) && !/^\s*export:\s*/i.test(line)
        )
        .join("\n")
        .trim();
    }
    let end = s.length;
    const u = lower.indexOf("\nuses:", start);
    const e = lower.indexOf("\nexport:", start);
    if (u !== -1) end = Math.min(end, u);
    if (e !== -1) end = Math.min(end, e);
    return s.slice(start, end).trim();
  }

  /** Trend chart subtitle under title: active filters only (no series window). */
  function chartTrendHintFiltersOnly(f) {
    if (!f) return "";
    const parts = [];
    parts.push(
      "Current " + formatMonthRangeShort(currentPeriodMonthList(f))
    );
    parts.push(
      "Comparison " + formatMonthRangeShort(comparisonPeriodMonthList(f))
    );
    if (f.business !== "all") {
      if (Array.isArray(f.business)) {
        if (!f.business.length) parts.push("Business: none");
        else parts.push(f.business.join(", "));
      } else if (f.business) {
        parts.push(String(f.business));
      }
    }
    if (f.state && f.state !== "all") parts.push(String(f.state));
    if (f.site !== "all") {
      if (Array.isArray(f.site)) {
        if (!f.site.length) parts.push("Site: none");
        else parts.push("Site: " + f.site.join(", "));
      } else if (f.site) {
        parts.push("Site: " + String(f.site));
      }
    }
    if (f.personalType && f.personalType !== "all") {
      parts.push("Personnel type: " + String(f.personalType));
    }
    if (f.variable && f.variable.length) {
      parts.push(
        f.variable.length +
          " vertical" +
          (f.variable.length === 1 ? "" : "s")
      );
    }
    return parts.join(" · ");
  }

  function trendInsightTooltipForChartLine(f) {
    if (!f) return chartHelpInsightOnly(CHART_HELP.trend);
    if (insightShell && shouldShowPercentSpeedometerChart(f.catKey, f)) {
      return "Insight: % completion vs group average (preview gauge).";
    }
    return chartHelpInsightOnly(CHART_HELP.trend);
  }

  function refreshTrendInsightTooltips(f) {
    const tiSpi = document.getElementById("ti-trend-spi");
    if (tiSpi) {
      tiSpi.textContent = chartHelpInsightOnly(CHART_HELP.spiTrend);
    }
    const tiLine = document.getElementById("ti-trend-line");
    if (tiLine && f) {
      tiLine.textContent = trendInsightTooltipForChartLine(f);
    }
  }

  const LS_VARIABLE_FILTER = "mis-variable-checkpoints";
  const LS_KPI_PREFIX = insightShell
    ? "insights-kpi-keys-"
    : "preview-kpi-keys-";
  const LS_CAT_MAIN_VIEW = insightShell
    ? "insights_cat_main_view"
    : "adani_cat_main_view";
  const LS_CAT_MAIN_VIEW_VL = insightShell
    ? "insights_vl_main_view"
    : "adani_vl_main_view";

  const TREND_LINE_COLORS = [
    ADANI_BLUE,
    ADANI_GREEN,
    ADANI_PURPLE,
    ADANI_ORANGE,
    "#1E9FD6",
    "#26C9A0",
    "#C77DCC",
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
    if (catKey === INCIDENT_MANAGEMENT_CATEGORY_KEY && kpisMeta.length) {
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

  /** Business unit — same interaction pattern as KPI list (details + All/None + checkboxes). */
  function businessFilterFieldHtml(bizList) {
    const list = bizList || [];
    const boxes = list
      .map((b, i) => {
        const id = "f-biz-cb-" + i;
        return (
          '<label class="kpi-cb" for="' +
          id +
          '">' +
          '<input type="checkbox" class="f-biz-cb" id="' +
          id +
          '" name="f-biz-cb" value="' +
          escapeAttr(b) +
          '"/>' +
          '<span class="kpi-cb__text">' +
          escapeHtml(b) +
          "</span></label>"
        );
      })
      .join("");
    return (
      '<div class="field field--toolbar-scope field--biz-site-scope">' +
      '<span class="field-label" id="f-biz-field-lbl">Business</span>' +
      '<details class="kpi-scope kpi-scope--toolbar" id="f-biz-details">' +
      '<summary class="kpi-scope__summary" aria-labelledby="f-biz-field-lbl" title="Business">' +
      '<span class="kpi-scope__summary-text">' +
      '<span class="kpi-scope__title">Business</span>' +
      '<span class="kpi-scope__count" id="f-biz-hint"></span>' +
      "</span></summary>" +
      '<div class="kpi-panel" id="f-biz-panel" role="group" aria-labelledby="f-biz-field-lbl">' +
      '<div class="kpi-panel__bar">' +
      '<button type="button" class="btn kpi-panel__btn" id="f-biz-btn-all">All</button>' +
      '<button type="button" class="btn kpi-panel__btn" id="f-biz-btn-none">None</button>' +
      "</div>" +
      '<div class="kpi-panel__list">' +
      boxes +
      "</div></div></details></div>"
    );
  }

  /** Site — same pattern as business / KPI list. */
  function siteFilterFieldHtml() {
    const n = [1, 2, 3, 4, 5, 6, 7];
    const boxes = n
      .map((num) => {
        const id = "f-site-cb-" + num;
        const val = "Site" + num;
        return (
          '<label class="kpi-cb" for="' +
          id +
          '">' +
          '<input type="checkbox" class="f-site-cb" id="' +
          id +
          '" name="f-site-cb" value="' +
          val +
          '"/>' +
          '<span class="kpi-cb__text">Site ' +
          num +
          "</span></label>"
        );
      })
      .join("");
    return (
      '<div class="field field--toolbar-scope field--biz-site-scope field--filter-compact">' +
      '<span class="field-label" id="f-site-field-lbl">Site</span>' +
      '<details class="kpi-scope kpi-scope--toolbar" id="f-site-details">' +
      '<summary class="kpi-scope__summary" aria-labelledby="f-site-field-lbl" title="Site">' +
      '<span class="kpi-scope__summary-text">' +
      '<span class="kpi-scope__title">Site</span>' +
      '<span class="kpi-scope__count" id="f-site-hint"></span>' +
      "</span></summary>" +
      '<div class="kpi-panel" id="f-site-panel" role="group" aria-labelledby="f-site-field-lbl">' +
      '<div class="kpi-panel__bar">' +
      '<button type="button" class="btn kpi-panel__btn" id="f-site-btn-all">All</button>' +
      '<button type="button" class="btn kpi-panel__btn" id="f-site-btn-none">None</button>' +
      "</div>" +
      '<div class="kpi-panel__list kpi-panel__list--site">' +
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
      '<span class="field-label" id="f-var-lbl">Verticals</span>' +
      '<details class="var-scope var-scope--toolbar" id="f-var-details">' +
      '<summary class="var-scope__summary" aria-labelledby="f-var-lbl" title="Verticals">' +
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

  /** Badge text for KPI unit type (Dim_KPI.Unit_Type → Count, % / Rate, Days, Hours). */
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
    const ck = Number(catKey);
    return DATA.categories.find((c) => Number(c.categoryKey) === ck);
  }

  function getKpis(catKey) {
    const ck = Number(catKey);
    const blocks = Array.isArray(DATA.kpiDetailByCategory)
      ? DATA.kpiDetailByCategory
      : [];
    let k = blocks.find((x) => Number(x.categoryKey) === ck);
    if (!k) {
      k = blocks.find((x) => String(x.categoryKey) === String(catKey));
    }
    let list =
      k && Array.isArray(k.kpis) ? k.kpis.slice() : [];
    if (!list.length && Array.isArray(DATA.factRows)) {
      const byKey = new Map();
      for (let i = 0; i < DATA.factRows.length; i++) {
        const r = DATA.factRows[i];
        if (r == null) continue;
        const rc = r.categoryKey;
        if (Number(rc) !== ck && String(rc) !== String(ck)) continue;
        const kk = Number(r.kpiKey);
        if (Number.isNaN(kk)) continue;
        if (byKey.has(kk)) continue;
        byKey.set(kk, {
          kpiKey: r.kpiKey,
          kpiName: r.kpiName || "KPI " + kk,
          unitType: r.unitType || "Count",
          latestValue: null,
        });
      }
      list = Array.from(byKey.values());
    }
    if (ck === SPI_CATEGORY_KEY) return list;
    const filtered = list.filter((kpi) => !isTriKpiMeta(kpi));
    return filtered.length ? filtered : list;
  }

  /** Preferred default KPI in dropdown (TRI / TRIR); fallback: first KPI in display order. */
  const TRI_LABEL_FULL = "TRIR";

  function defaultKpiKeyForCategory(catKey, kpisMetaForUi) {
    const cat = Number(catKey);
    const merged =
      kpisMetaForUi && kpisMetaForUi.length
        ? kpisMetaForUi
        : kpiListForFilterDropdown(catKey);
    if (cat === HAZARD_CATEGORY_KEY) {
      const sorted = sortKpisForDisplay(catKey, merged);
      return sorted.length ? String(sorted[0].kpiKey) : "38";
    }
    if (cat === SPI_CATEGORY_KEY) {
      const tri = merged.find((k) => isTriKpiMeta(k));
      if (tri) return String(tri.kpiKey);
    }
    const sorted = sortKpisForDisplay(catKey, merged);
    return sorted.length ? String(sorted[0].kpiKey) : "1";
  }

  function kpiDropdownLabel(k) {
    if (String(k.kpiKey) === "21" || isTriKpiMeta(k)) return TRI_LABEL_FULL;
    return k.kpiName;
  }

  function isTriKpiMeta(k) {
    /** Dim 58 = Incident Data “Total Recordable Injuries” — not SPI TRIR; must stay in cat 1 lists. */
    if (String(k.kpiKey) === "58") return false;
    return (
      String(k.kpiKey) === "21" ||
      /Total Recordable|TRIR|\bTRI\b|Recordable Incident Rate/i.test(
        k.kpiName || ""
      )
    );
  }

  /** KPI dropdown: TRIR / TRI only under Safety Performance Indices (category 3). */
  function kpiListForFilterDropdown(catKey) {
    const cat = Number(catKey);
    let base = getKpis(catKey);
    if (cat === SPI_CATEGORY_KEY && !base.some(isTriKpiMeta)) {
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
    if (!insightShell) {
      const shown = classicIndexCategoryDisplayName(
        cat.categoryKey,
        cat.categoryName
      );
      if ((shown || "").toLowerCase().includes(qLower)) return true;
    }
    if ((cat.uxNote || "").toLowerCase().includes(qLower)) return true;
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

  function destroyCharts() {
    [
      "chart-line",
      "chart-verticals",
      "chart-biz",
      "chart-spi-bubble",
      "chart-spi-insights",
      "chart-bu-compare",
      "chart-hazard-alt",
      "chart-spi-abs-gauge",
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.destroy();
      }
    });
    destroySpiLeafletMap();
    if (window.__adaniVlMap) {
      try {
        window.__adaniVlMap.remove();
      } catch {
        /* ignore */
      }
      window.__adaniVlMap = null;
    }
    const vlMapHost = document.getElementById("vl-leaflet-map");
    if (vlMapHost) vlMapHost.innerHTML = "";
    const spiHmHost = document.getElementById("spi-hazard-heatmap-host");
    if (spiHmHost) spiHmHost.innerHTML = "";
    ["chart-hazard-alt", "chart-spi-abs-gauge"].forEach(function(id) {
      const el = document.getElementById(id);
      if (el) {
        const c2 = el.getContext("2d");
        if (c2) c2.clearRect(0, 0, el.width, el.height);
      }
    });
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
    if (typeof window.__adaniSpeedometerGaugeRedraw === "function") {
      try {
        window.__adaniSpeedometerGaugeRedraw();
      } catch {
        /* ignore */
      }
    }
  }

  const SPI_HM_PAL = {
    green: {
      light: [220, 255, 232],
      dark: [13, 68, 48],
      textDark: [15, 40, 30],
      textLight: [255, 255, 255],
    },
    red: {
      light: [255, 230, 240],
      dark: [183, 28, 28],
      textDark: [80, 24, 32],
      textLight: [255, 255, 255],
    },
    orange: {
      light: [255, 243, 224],
      dark: [230, 81, 0],
      textDark: [100, 48, 0],
      textLight: [255, 255, 255],
    },
  };

  function spiHeatmapPalette(kpiKey) {
    const k = Number(kpiKey);
    if (k === 20) return SPI_HM_PAL.red;
    if (k === 21) return SPI_HM_PAL.orange;
    return SPI_HM_PAL.green;
  }

  function lerpRgbSpiHm(a, b, t) {
    return [
      Math.round(a[0] + (b[0] - a[0]) * t),
      Math.round(a[1] + (b[1] - a[1]) * t),
      Math.round(a[2] + (b[2] - a[2]) * t),
    ];
  }

  function styleForSpiHeatmapCell(tNorm, pal) {
    const rgb = lerpRgbSpiHm(pal.light, pal.dark, tNorm);
    const lum =
      (0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]) / 255;
    const textRgb = lum > 0.55 ? pal.textDark : pal.textLight;
    return (
      "background:rgb(" +
      rgb.join(",") +
      ");color:rgb(" +
      textRgb.join(",") +
      ")"
    );
  }

  function formatSpiHeatmapValue(v) {
    if (v == null || Number.isNaN(Number(v))) return "—";
    const n = Number(v);
    const a = Math.abs(n);
    if (a > 0 && a < 1) return n.toFixed(4);
    return n.toFixed(2);
  }

  /**
   * SPI heat map: KPI rows × last six calendar months; per-row color scale (greens / Near Miss reds / Incident oranges).
   */
  function renderSpiPerfHeatmapChart(snapPool, f) {
    const el = document.getElementById("chart-spi-bubble");
    if (!el) return;

    const colMonths = calendarClippedMonthKeys(f, 6);
    if (!colMonths.length) {
      el.innerHTML =
        '<p class="spi-perf-hm-fallback">No months in range for this heat map.</p>';
      el.setAttribute("aria-label", "SPI heat map: no data in calendar range.");
      return;
    }

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
    const colLabels = colMonths.map((ym) => {
      const p = ym.split("-");
      const mon = monthShort[Number(p[1]) - 1] || p[1];
      const yy = String(p[0]).slice(2);
      return mon + "\u2019" + yy;
    });

    const kpisMeta = getKpis(SPI_CATEGORY_KEY);
    const byKey = new Map(
      kpisMeta.map((k) => [Number(k.kpiKey), k])
    );

    const matrix = [];
    for (let ri = 0; ri < SPI_HEATMAP_KPI_ORDER.length; ri++) {
      const kpid = SPI_HEATMAP_KPI_ORDER[ri];
      const row = [];
      for (let ci = 0; ci < colMonths.length; ci++) {
        const ym = colMonths[ci];
        const sub = snapPool.filter(
          (r) =>
            String(r.kpiKey) === String(kpid) &&
            String(r.yearMonth) === String(ym)
        );
        const vals = sub
          .map((r) => Number(r.value))
          .filter((v) => !Number.isNaN(v));
        row.push(vals.length ? avg(vals) : null);
      }
      matrix.push(row);
    }

    function normFnForRow(ri) {
      const row = matrix[ri];
      const vals = row.filter(
        (v) => v != null && !Number.isNaN(Number(v))
      );
      if (!vals.length) {
        return function () {
          return 0.5;
        };
      }
      const mn = Math.min.apply(null, vals);
      const mx = Math.max.apply(null, vals);
      if (mx <= mn) {
        return function () {
          return 0.5;
        };
      }
      return function (v) {
        if (v == null || Number.isNaN(Number(v))) return 0;
        return (Number(v) - mn) / (mx - mn);
      };
    }

    let html =
      '<table class="spi-perf-heatmap-table" role="grid"><thead><tr><th class="spi-perf-hm-corner" scope="col"></th>';
    for (let c = 0; c < colLabels.length; c++) {
      html +=
        '<th scope="col" class="spi-perf-hm-colhead">' +
        escapeHtml(colLabels[c]) +
        "</th>";
    }
    html += "</tr></thead><tbody>";
    for (let r = 0; r < SPI_HEATMAP_KPI_ORDER.length; r++) {
      const kpid = SPI_HEATMAP_KPI_ORDER[r];
      const meta = byKey.get(kpid);
      const rowLabel = meta
        ? kpiDropdownLabel(meta)
        : "KPI " + kpid;
      const norm = normFnForRow(r);
      const pal = spiHeatmapPalette(kpid);
      html +=
        '<tr><th scope="row" class="spi-perf-hm-rowhead">' +
        escapeHtml(rowLabel) +
        "</th>";
      for (let c = 0; c < colMonths.length; c++) {
        const v = matrix[r][c];
        if (v == null || Number.isNaN(Number(v))) {
          html += '<td class="spi-perf-hm-cell spi-perf-hm-cell--na">—</td>';
        } else {
          const t = norm(v);
          html +=
            '<td class="spi-perf-hm-cell" style="' +
            escapeAttr(styleForSpiHeatmapCell(t, pal)) +
            '">' +
            escapeHtml(formatSpiHeatmapValue(v)) +
            "</td>";
        }
      }
      html += "</tr>";
    }
    html += "</tbody></table>";
    el.innerHTML = html;
    el.setAttribute(
      "aria-label",
      "Safety performance indices heat map: SPI KPI rows by month; color intensity reflects relative value within each row for the current filters."
    );
  }

  const SPI_INSIGHT_PALETTE = [
    { fill: "rgba(0, 109, 182, 0.42)", stroke: ADANI_BLUE },
    { fill: "rgba(0, 177, 107, 0.42)", stroke: ADANI_GREEN },
    { fill: "rgba(142, 39, 143, 0.38)", stroke: ADANI_PURPLE },
    { fill: "rgba(240, 76, 35, 0.35)", stroke: ADANI_ORANGE },
    { fill: "rgba(230, 231, 235, 0.85)", stroke: "#9ca3af" },
    { fill: "rgba(0, 109, 182, 0.28)", stroke: "#2E8FD0" },
    { fill: "rgba(0, 177, 107, 0.3)", stroke: "#26C98A" },
    { fill: "rgba(142, 39, 143, 0.28)", stroke: "#A855BC" },
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
          subtitle: {
            display: true,
            text: spiMixCompositionCaption(f),
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 4 },
          },
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
              color: CHART_INK,
            },
            grid: { color: "rgba(109, 110, 113, 0.1)" },
            title: {
              display: true,
              text: "Month",
              font: { size: 10 },
              color: CHART_INK,
            },
          },
          y: {
            stacked: true,
            min: 0,
            max: 100,
            ticks: {
              font: { size: 9 },
              color: CHART_INK,
              callback(v) {
                return v + "%";
              },
            },
            title: {
              display: true,
              text: "% of rows by KPI (same filters)",
              font: { size: 10 },
              color: CHART_INK,
            },
            grid: { color: "rgba(109, 110, 113, 0.085)" },
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

    const selectedKeys = effectiveKpiKeysForChartSeries(catKey, f);
    const selectedSet = new Set(selectedKeys.map(String));
    const kpisMetaFiltered = kpisMeta.filter((k) =>
      selectedSet.has(String(k.kpiKey))
    );
    if (!kpisMetaFiltered.length) return;

    const allKeys = new Set(kpisMetaFiltered.map((k) => String(k.kpiKey)));
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

    const datasets = kpisMetaFiltered.map((kMeta, idx) => {
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
      selectedKeys.length > 1
        ? "Line chart: monthly average for selected Safety Performance Index KPIs; click a series or legend entry to focus that KPI."
        : "Line chart: monthly average for each Safety Performance Index KPI; click a series or legend entry to focus that KPI."
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
          subtitle: {
            display: true,
            text: chartTrendHintFiltersOnly(f),
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 4 },
          },
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
                const name = ctx.dataset.label
                  ? String(ctx.dataset.label)
                  : "";
                return (
                  " " +
                  (name ? name + ": " : "") +
                  formatValue(v, ut)
                );
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
              color: CHART_INK,
              padding: 4,
            },
            grid: { color: "rgba(109, 110, 113, 0.1)" },
            title: {
              display: true,
              text: "KPI value",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
            },
          },
          x: {
            ticks: {
              font: { size: 9 },
              maxRotation: 45,
              color: CHART_INK,
              padding: 2,
            },
            grid: { color: "rgba(109, 110, 113, 0.08)" },
            title: {
              display: true,
              text: "Month",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
            },
          },
        },
      },
    });
  }

  function wireCatMainViewIndex() {
    const main = document.getElementById("cat-main-view");
    if (!main) return;
    if (document.getElementById("view-tab-vl-map")) {
      wireCatMainViewVulnerable();
      return;
    }
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
        updateKpiToolbarButtonsVisibility("compare");
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
      updateKpiToolbarButtonsVisibility(charts ? "charts" : "table");
      if (charts && enforceChartViewSingleKpiSelection() && currentCategoryKey != null) {
        refreshCategoryView(currentCategoryKey);
      }
      if (!charts && currentCategoryKey != null) {
        renderTableBody(currentCategoryKey);
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

    /** Delegated tab switch — pointerup + click (dedupe only duplicate events for same tab) */
    let viewTabLastMs = 0;
    let viewTabLastMode = "";
    function onViewTabActivate(ev) {
      const btn = ev.target.closest("button.view-tabs__btn[data-view]");
      if (!btn || !main.contains(btn)) return;
      const mode = btn.getAttribute("data-view");
      if (mode !== "charts" && mode !== "table" && mode !== "compare") return;
      const t = Date.now();
      if (
        mode === viewTabLastMode &&
        t - viewTabLastMs < 350
      ) {
        return;
      }
      viewTabLastMs = t;
      viewTabLastMode = mode;
      apply(mode);
    }
    main.addEventListener("pointerup", onViewTabActivate);
    main.addEventListener("click", onViewTabActivate);
  }

  /** Map layer rows: all India states; state toolbar filter does not clip map totals. */
  function vulnerableLocationMapRows(f) {
    const fAll = Object.assign({}, f, { state: "all" });
    const base = getRowsForCategory(VULNERABLE_LOCATION_CATEGORY_KEY);
    return applyNonMonthFiltersAllKpis(base, fAll).filter((r) =>
      VULNERABLE_LOCATION_KPI_KEYS.has(Number(r.kpiKey))
    );
  }

  /** State/UT rollups + top/bottom five (resilient / vulnerable) for map and charts. */
  function vulnerableLocationStateRankingFromSnap(snap) {
    const byState = {};
    for (let i = 0; i < snap.length; i++) {
      const r = snap[i];
      const st = String(r.state || "").trim();
      if (!st) continue;
      if (!byState[st]) byState[st] = { sum: 0, kpiCount: 0, seen: {} };
      const kk = Number(r.kpiKey);
      if (!Number.isNaN(kk)) {
        if (!byState[st].seen[kk]) {
          byState[st].seen[kk] = 1;
          byState[st].kpiCount += 1;
        }
      }
      const v = Number(r.value);
      if (Number.isFinite(v)) byState[st].sum += v;
    }
    let maxKpi = 1;
    let maxSum = 1;
    INDIA_STATES_UT.forEach((nm) => {
      const a = byState[nm];
      if (!a) return;
      if (a.kpiCount > maxKpi) maxKpi = a.kpiCount;
      if (a.sum > maxSum) maxSum = a.sum;
    });
    const sums = [];
    INDIA_STATES_UT.forEach((nm) => {
      const a = byState[nm] || { sum: 0, kpiCount: 0, seen: {} };
      sums.push({ state: nm, sum: a.sum || 0 });
    });
    sums.sort((a, b) => b.sum - a.sum);
    const top5 = sums.slice(0, 5).map((x) => x.state);
    const bottom5 = sums.slice(-5).map((x) => x.state);
    return { byState, top5, bottom5, maxKpi, maxSum };
  }

  function vulnerableLocationStateRankingFromFilters(f) {
    const winM = new Set(effectiveBizWindowMonths(f));
    const snap = vulnerableLocationMapRows(f).filter((r) => winM.has(r.yearMonth));
    return vulnerableLocationStateRankingFromSnap(snap);
  }

  /** Preview-only man-hour tiles (deterministic from filters). */
  function vulnerableLocationDummyManHours(f) {
    const seed =
      String(f.currentFrom || "") +
      "|" +
      String(f.currentTo || "") +
      "|" +
      String(f.comparisonFrom || "") +
      "|" +
      String(f.comparisonTo || "") +
      "|" +
      (Array.isArray(f.business)
        ? f.business.join("+")
        : String(f.business || "")) +
      "|" +
      (Array.isArray(f.site) ? f.site.join("+") : String(f.site || "")) +
      "|" +
      JSON.stringify(f.variable || {});
    const emp = 1280400 + (Math.abs(hash32(seed + "|emp")) % 498000);
    const con = 915200 + (Math.abs(hash32(seed + "|con")) % 362000);
    const total = emp + con;
    const safeFrac = 0.912 + (Math.abs(hash32(seed + "|safe")) % 65) / 1000;
    const safe = Math.round(total * safeFrac);
    return { emp: emp, con: con, total: total, safe: safe };
  }

  function populateVulnerableLocationManHourTiles(f) {
    const d = vulnerableLocationDummyManHours(f);
    const fmt = (n) =>
      Number(n).toLocaleString("en-IN", { maximumFractionDigits: 0 });
    const setText = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.textContent = v;
    };
    setText("vl-kpi-emp", fmt(d.emp));
    setText("vl-kpi-con", fmt(d.con));
    setText("vl-kpi-total", fmt(d.total));
    setText("vl-kpi-safe", fmt(d.safe));
  }

  function renderVulnerableLocationMap(f) {
    const host = document.getElementById("vl-leaflet-map");
    if (!host) return;
    if (window.__adaniVlMap) {
      try {
        window.__adaniVlMap.remove();
      } catch {
        /* ignore */
      }
      window.__adaniVlMap = null;
    }
    host.innerHTML = "";
    if (typeof L === "undefined") {
      host.innerHTML =
        '<p class="vl-map-fallback" role="status">Map preview needs Leaflet (check network) and a local HTTP server.</p>';
      return;
    }
    const inner = document.createElement("div");
    inner.className = "vl-leaflet-map-inner";
    inner.id = "vl-leaflet-map-inner";
    host.appendChild(inner);

    const rank = vulnerableLocationStateRankingFromFilters(f);
    const byState = rank.byState;
    const maxKpi = rank.maxKpi;
    const maxSum = rank.maxSum;
    const top5 = rank.top5;
    const bottom5 = rank.bottom5;

    const map = L.map(inner, {
      zoomControl: true,
      attributionControl: true,
    }).setView([22.5, 78], 5);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution:
        '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
    }).addTo(map);
    const sel = f.state && f.state !== "all" ? String(f.state).trim() : "";

    function heatColor(t) {
      const u = Math.max(0, Math.min(1, t));
      const r = Math.round(230 + (0 - 230) * u);
      const g = Math.round(240 + (177 - 240) * u);
      const b = Math.round(250 + (107 - 250) * u);
      return "rgb(" + r + "," + g + "," + b + ")";
    }

    INDIA_STATES_UT.forEach((nm) => {
      const ll = STATE_CENTROID_BY_STATE[nm];
      if (!ll) return;
      const agg = byState[nm] || { sum: 0, kpiCount: 0, seen: {} };
      const tK = agg.kpiCount / maxKpi;
      const tS = agg.sum / maxSum;
      const t = 0.55 * tK + 0.45 * tS;
      const fill = heatColor(t);
      const radius = 6 + agg.kpiCount * 2.2 + Math.min(14, Math.sqrt(agg.sum + 1));
      // determine resilient/vulnerable membership
      const isTop = top5.indexOf(nm) !== -1;
      const isBottom = bottom5.indexOf(nm) !== -1;
      // read user toggles (default true)
      const showRes = (() => {
        const el = document.getElementById("vl-filter-resilient");
        return el ? !!el.checked : true;
      })();
      const showVul = (() => {
        const el = document.getElementById("vl-filter-vulnerable");
        return el ? !!el.checked : true;
      })();
      // decide visibility and color
      let strokeCol = "rgba(35,31,32,0.35)";
      let fillCol = fill;
      if (isTop) {
        strokeCol = "#16a34a";
        fillCol = "#a7f3d0";
      } else if (isBottom) {
        strokeCol = "#dc2626";
        fillCol = "#fca5a5";
      } else {
        strokeCol = "rgba(35,31,32,0.25)";
        fillCol = "rgba(200,200,200,0.18)";
      }
      // decide whether to show (always show, but muted if not selected)
      const showSelectedOnly = !(showRes && showVul);
      const isSelected =
        (isTop && showRes) || (isBottom && showVul) || (!isTop && !isBottom && (showRes && showVul));
      const opacity = isSelected ? 0.9 : 0.28;
      const weight = isSelected ? 2 : 1;
      const mk = L.circleMarker(ll, {
        radius: radius,
        stroke: true,
        color: sel && sel === nm ? ADANI_PURPLE : strokeCol,
        weight: sel && sel === nm ? 3 : weight,
        fillColor: fillCol,
        fillOpacity: opacity,
      }).addTo(map);
      mk.bindPopup(
        "<strong>" +
          escapeHtml(nm) +
          "</strong><br/>" +
          "KPIs with data: " +
          agg.kpiCount +
          " / " +
          VULNERABLE_LOCATION_KPI_ORDER.length +
          "<br/>" +
          "Σ value (window): " +
          formatValue(agg.sum, "Count")
      );
    });
    map.fitBounds(
      L.latLngBounds(
        L.latLng(INDIA_MAP_BOUNDS_SW[0], INDIA_MAP_BOUNDS_SW[1]),
        L.latLng(INDIA_MAP_BOUNDS_NE[0], INDIA_MAP_BOUNDS_NE[1])
      ),
      { padding: [10, 14] }
    );
    window.__adaniVlMap = map;
  }

  function renderVulnerableLocationTrend(f) {
    const el1 = document.getElementById("chart-vl-trend");
    const el2 = document.getElementById("chart-vl-dist");
    const tTitle = document.getElementById("chart-vl-trend-title");
    const dTitle = document.getElementById("chart-vl-dist-title");
    if (!el1 || !el2) return;
    try {
      const c1 = Chart.getChart(el1);
      if (c1) c1.destroy();
    } catch {}
    try {
      const c2 = Chart.getChart(el2);
      if (c2) c2.destroy();
    } catch {}

    const pk =
      f.kpiKeys && f.kpiKeys.length
        ? String(f.kpiKeys[0])
        : String(f.kpi);
    const kMeta = getKpis(VULNERABLE_LOCATION_CATEGORY_KEY).find(
      (x) => String(x.kpiKey) === pk
    );
    const kLabel = kMeta ? kpiDropdownLabel(kMeta) : "KPI " + pk;
    if (tTitle) {
      tTitle.textContent =
        "Trend — " + kLabel + " (resilient vs vulnerable locations)";
    }
    if (dTitle) {
      dTitle.textContent =
        "Distribution — " + kLabel + " (latest month in current period)";
    }

    const rank = vulnerableLocationStateRankingFromFilters(f);
    const topSet = new Set(rank.top5);
    const botSet = new Set(rank.bottom5);
    const basePool = applyNonMonthFiltersAllKpis(
      getRowsForCategory(VULNERABLE_LOCATION_CATEGORY_KEY),
      f
    );
    const months = currentPeriodMonthList(f);

    function monthAgg(stateSet, ym) {
      if (!kMeta || !ym) return null;
      const kk = kMeta.kpiKey;
      const ut = kMeta.unitType || "";
      const rows = basePool.filter(
        (r) =>
          String(r.kpiKey) === String(kk) &&
          String(r.yearMonth) === String(ym) &&
          stateSet.has(String(r.state || "").trim())
      );
      if (!rows.length) return null;
      if (isAdditiveUnit(ut)) {
        return rows.reduce((s, r) => s + Number(r.value || 0), 0);
      }
      return avg(rows.map((r) => r.value));
    }

    const labels = months.length
      ? months.map((ym) => formatMonthYear(ym))
      : ["—"];
    const resData = months.length
      ? months.map((ym) => monthAgg(topSet, ym) ?? 0)
      : [0];
    const vulData = months.length
      ? months.map((ym) => monthAgg(botSet, ym) ?? 0)
      : [0];
    const lastM = months.length ? months[months.length - 1] : null;
    const distLabels = ["Resilient (top 5 states)", "Vulnerable (bottom 5 states)"];
    const distData = lastM
      ? [
          monthAgg(topSet, lastM) ?? 0,
          monthAgg(botSet, lastM) ?? 0,
        ]
      : [0, 0];

    const vlLegendBottom = {
      display: true,
      position: "bottom",
      align: "center",
      labels: {
        boxWidth: 10,
        padding: 6,
        font: { size: 9, family: FONT_UI },
        color: CHART_INK,
      },
    };
    const vlChartPlugins = {
      legend: vlLegendBottom,
      subtitle: {
        display: true,
        text:
          "Legend: resilient = top 5 states/UTs by combined KPI values in the map window; vulnerable = bottom 5 (same rule as map markers).",
        color: CHART_INK,
        font: { size: 8.5, family: FONT_UI },
        padding: { bottom: 2 },
      },
    };

    if (typeof Chart !== "undefined") {
      new Chart(el1, {
        type: "line",
        data: {
          labels: labels,
          datasets: [
            {
              label: "Resilient locations (top 5 states)",
              data: resData,
              borderColor: "#34d399",
              backgroundColor: "rgba(52, 211, 153, 0.12)",
              tension: 0.2,
              fill: true,
              borderWidth: 2,
              pointRadius: 3,
            },
            {
              label: "Vulnerable locations (bottom 5 states)",
              data: vulData,
              borderColor: "#fb7185",
              backgroundColor: "rgba(251, 113, 133, 0.12)",
              tension: 0.2,
              fill: true,
              borderWidth: 2,
              pointRadius: 3,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          plugins: vlChartPlugins,
        },
      });
      new Chart(el2, {
        type: "bar",
        data: {
          labels: distLabels,
          datasets: [
            {
              label: kLabel + " · latest month",
              data: distData,
              backgroundColor: ["rgba(52, 211, 153, 0.55)", "rgba(251, 113, 133, 0.55)"],
              borderColor: ["#34d399", "#fb7185"],
              borderWidth: 1,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: vlChartPlugins,
        },
      });
    } else {
      const p1 = el1.parentElement;
      const p2 = el2.parentElement;
      if (p1) p1.innerHTML = "<div class='chart-placeholder'>Trend chart (requires Chart.js)</div>";
      if (p2) p2.innerHTML = "<div class='chart-placeholder'>Distribution chart (requires Chart.js)</div>";
    }
  }

  function refreshVulnerableLocationView(catKey) {
    if (Number(catKey) !== VULNERABLE_LOCATION_CATEGORY_KEY) return;
    const f = readFilters(catKey);
    if (!f) return;
    populateVulnerableLocationManHourTiles(f);
    const main = document.getElementById("cat-main-view");
    const mode = main ? main.dataset.view || "trend" : "trend";
    // render multi KPI cards for this category (match Incident UI)
    const kpisMeta = getKpis(catKey);
    const multiWrap = document.getElementById("multi-kpi-wrap");
    if (multiWrap) {
      const aggList = buildKpiDetailMetrics(catKey, kpisMeta, f);
      if (!aggList.length) {
        multiWrap.innerHTML =
          '<div class="empty-msg" style="padding:8px">No KPI data for this selection. Adjust periods or geography filters.</div>';
      } else {
        renderMultiKpiCards(
          multiWrap,
          aggList,
          f,
          f.kpiKeys && f.kpiKeys.length ? f.kpiKeys : [f.kpi],
          [
            String(catKey),
            f.currentFrom || "",
            f.currentTo || "",
            f.comparisonFrom || "",
            f.comparisonTo || "",
            f.refMonth || "",
            f.state || "",
            (Array.isArray(f.business)
              ? JSON.stringify(f.business.slice().sort())
              : String(f.business || "")),
            String(f.kpi || ""),
            JSON.stringify(f.kpiKeys || []),
            JSON.stringify(f.variable || {}),
          ].join("|"),
          catKey
        );
      }
    }

    if (mode === "map") {
      renderVulnerableLocationMap(f);
    } else if (mode === "trend") {
      renderVulnerableLocationTrend(f);
    } else if (mode === "table") {
      renderTableBody(catKey);
    } else {
      buildBuComparisonChart(catKey, f);
    }
    const cat = getCategory(catKey);
    const ctx = document.getElementById("cat-context");
    if (ctx && cat) {
      ctx.textContent = cat.uxNote || "";
    }
  }

  function wireCatMainViewVulnerable() {
    const main = document.getElementById("cat-main-view");
    if (!main) return;
    const tTrend = document.getElementById("view-tab-vl-trend");
    const tMap = document.getElementById("view-tab-vl-map");
    const tTable = document.getElementById("view-tab-vl-table");
    const tCmp = document.getElementById("view-tab-vl-compare");
    const pTrend = document.getElementById("view-panel-vl-trend");
    const pMap = document.getElementById("view-panel-vl-map");
    const pTable = document.getElementById("view-panel-vl-table");
    const pCmp = document.getElementById("view-panel-vl-compare");
    if (!tTrend || !pTrend || !tMap || !pMap || !tTable || !pTable || !tCmp || !pCmp) {
      return;
    }

    function apply(mode) {
      const isTrend = mode === "trend";
      const isMap = mode === "map";
      const isTable = mode === "table";
      const isCmp = mode === "compare";
      main.dataset.view = mode;
      main.classList.remove(
        "cat-main-view--charts",
        "cat-main-view--table",
        "cat-main-view--compare",
        "cat-main-view--vl-map",
        "cat-main-view--vl-table",
        "cat-main-view--vl-compare"
      );
      if (isMap) main.classList.add("cat-main-view--vl-map");
      else if (isTable) main.classList.add("cat-main-view--vl-table");
      else if (isCmp) main.classList.add("cat-main-view--vl-compare");
      else if (isTrend) main.classList.add("cat-main-view--charts");
      [tTrend, tMap, tTable, tCmp].forEach((btn) => {
        if (!btn) return;
        const m = btn.getAttribute("data-view");
        const on = m === mode;
        btn.setAttribute("aria-selected", on ? "true" : "false");
        btn.classList.toggle("view-tabs__btn--active", on);
        btn.tabIndex = on ? 0 : -1;
      });
      if (pTrend) pTrend.hidden = !isTrend;
      if (pMap) pMap.hidden = !isMap;
      if (pTable) pTable.hidden = !isTable;
      if (pCmp) pCmp.hidden = !isCmp;
      try {
        localStorage.setItem(LS_CAT_MAIN_VIEW_VL, mode);
      } catch {
        /* ignore */
      }
      if (currentCategoryKey != null) {
        refreshVulnerableLocationView(currentCategoryKey);
      }
      requestAnimationFrame(() => {
        requestAnimationFrame(() => resizeAllChartsIndex());
      });
    }

    /** Same pattern as `wireCatMainViewIndex`: delegate on #cat-main-view so tabs always work
     * even if initial `apply()` throws during chart/table render. Dedupe pointerup+click per tab. */
    let vlTabLastMs = 0;
    let vlTabLastMode = "";
    function onVlViewTabActivate(ev) {
      const btn = ev.target.closest("button.view-tabs__btn[data-view]");
      if (!btn || !main.contains(btn)) return;
      const mode = btn.getAttribute("data-view");
      if (mode !== "map" && mode !== "table" && mode !== "compare" && mode !== "trend") return;
      const t = Date.now();
      if (mode === vlTabLastMode && t - vlTabLastMs < 350) return;
      vlTabLastMs = t;
      vlTabLastMode = mode;
      apply(mode);
    }
    main.addEventListener("pointerup", onVlViewTabActivate);
    main.addEventListener("click", onVlViewTabActivate);

    let initial = "trend";
    try {
      const s = localStorage.getItem(LS_CAT_MAIN_VIEW_VL);
      if (s === "map" || s === "table" || s === "compare" || s === "trend") initial = s;
    } catch {
      /* ignore */
    }
    try {
      apply(initial);
    } catch {
      /* Initial chart/table render can throw; tab handlers are already wired. */
    }
  }

  /** Distinct hues for radar spokes (one color per business unit). */
  function radarSpokeColors(count) {
    const n = Math.max(1, Math.min(Number(count) || 1, 48));
    const out = [];
    for (let i = 0; i < n; i++) {
      const h = Math.round((360 * i) / n + 14) % 360;
      out.push({
        stroke: "hsl(" + h + " 62% 46%)",
        fill: "hsla(" + h + " 58% 52% / 0.9)",
        fillSoft: "hsla(" + h + " 55% 55% / 0.22)",
      });
    }
    return out;
  }

  /**
   * By business: radar chart — one spoke per preview BU (24); value from filtered rows or 0.
   */
  function renderBizBreakdown(fullNames, values, datasetLabel, unitType, f) {
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

    const periodSub =
      f && typeof f === "object" && f.refMonth
        ? bizWindowRowsCaption(f)
        : "";

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
            angleLines: { color: "rgba(109, 110, 113, 0.15)" },
            grid: { color: "rgba(109, 110, 113, 0.09)" },
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
              color: CHART_INK,
            },
            title: {
              display: true,
              text:
                unitType === "PercentOrRate"
                  ? "KPI value (%) — radial scale"
                  : "KPI value — radial scale",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
            },
          },
        },
        plugins: {
          subtitle: {
            display: !!periodSub,
            text: periodSub,
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 6 },
          },
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
              color: CHART_INK,
            },
          },
          tooltip: {
            enabled: true,
            backgroundColor: "rgba(35, 31, 32, 0.96)",
            titleColor: "#ffffff",
            bodyColor: "#ffffff",
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

  /**
   * Trend & Distribution layout rules: single-site, single-BU (site radar), multi-KPI.
   */
  function chartUiMode(catKey, f) {
    const keys = effectiveKpiKeysForChartSeries(catKey, f);
    const multiKpi = keys.length > 1;
    const singleSite =
      f && Array.isArray(f.site) && f.site.length === 1;
    const singleBu =
      f && Array.isArray(f.business) && f.business.length === 1;
    const radarVisible =
      !singleSite && !multiKpi &&
      Number(catKey) !== CONSEQUENCE_MANAGEMENT_CATEGORY_KEY;
    const verticalVisible =
      !singleSite && (multiKpi || !singleBu);
    return {
      keys,
      multiKpi,
      singleSite,
      singleBu,
      radarVisible,
      verticalVisible,
      siteRadarInsteadOfBu: Boolean(singleBu && radarVisible),
    };
  }

  /** Trend / SPI / Hazard grids: hide boxes + set column count for reflow. */
  function applyChartLayoutVisibility(catKey, f) {
    const mode = chartUiMode(catKey, f);
    const trend = document.getElementById("chart-box-trend");
    const biz = document.getElementById("chart-box-biz");
    const vert = document.getElementById("chart-box-vertical");
    const hm = document.getElementById("chart-box-hazard-hm");
    const gridStd = document.querySelector(".cat-charts.cat-charts--standard");
    const gridConseq = document.querySelector(".cat-charts.cat-charts--consequence");
    const gridSpi = document.querySelector(".cat-charts.cat-charts--spi");
    const gridHz = document.querySelector(".cat-charts.cat-charts--hazard");

    function resetGridClasses(g) {
      if (!g) return;
      g.classList.remove(
        "cat-charts--cols-1",
        "cat-charts--cols-2",
        "cat-charts--cols-3"
      );
    }
    [gridStd, gridConseq, gridSpi, gridHz].forEach(resetGridClasses);

    function show(el, on) {
      if (!el) return;
      el.hidden = !on;
      el.classList.toggle("chart-box--hidden", !on);
    }

    if (catKey === HAZARD_CATEGORY_KEY) {
      show(hm, true);
      show(biz, mode.radarVisible);
      show(vert, mode.verticalVisible);
      const n = 1 + (mode.radarVisible ? 1 : 0) + (mode.verticalVisible ? 1 : 0);
      if (gridHz) gridHz.classList.add("cat-charts--cols-" + n);
      return;
    }
    if (catKey === SPI_CATEGORY_KEY) {
      show(trend, true);
      show(vert, true);
      if (gridSpi) gridSpi.classList.add("cat-charts--cols-2");
      return;
    }
    if (Number(catKey) === CONSEQUENCE_MANAGEMENT_CATEGORY_KEY) {
      show(trend, true);
      show(vert, mode.verticalVisible);
      const n = 1 + (mode.verticalVisible ? 1 : 0);
      if (gridConseq) gridConseq.classList.add("cat-charts--cols-" + n);
      return;
    }
    show(trend, true);
    show(biz, mode.radarVisible);
    show(vert, mode.verticalVisible);
    const n = 1 + (mode.radarVisible ? 1 : 0) + (mode.verticalVisible ? 1 : 0);
    if (gridStd) gridStd.classList.add("cat-charts--cols-" + n);
  }

  function updateComparisonTabVisibility(catKey, f) {
    const tab = document.getElementById("view-tab-compare");
    if (!tab) return;
    const singleSite =
      f && Array.isArray(f.site) && f.site.length === 1;
    tab.style.display = singleSite ? "none" : "";
    if (singleSite) {
      const main = document.getElementById("cat-main-view");
      if (main && main.getAttribute("data-view") === "compare") {
        const chartsBtn = document.getElementById("view-tab-charts");
        if (chartsBtn) chartsBtn.click();
      }
    }
  }

  function updateTableZoneLabel(f) {
    const el = document.getElementById("tbl-zone-label");
    if (!el || !f) return;
    if (Number(currentCategoryKey) === VULNERABLE_LOCATION_CATEGORY_KEY) {
      el.textContent =
        "Business, site & KPIs (current period window)";
      return;
    }
    const singleSite =
      Array.isArray(f.site) && f.site.length === 1;
    const singleBu =
      Array.isArray(f.business) && f.business.length === 1;
    if (singleSite) {
      el.textContent = "Site performance";
    } else if (singleBu) {
      el.textContent = "Site performance";
    } else {
      el.textContent = "BU performance";
    }
  }

  function kpiScopeTitleWithPlusN(catKey, f) {
    const list = kpiListForFilterDropdown(catKey);
    const keyStrs = effectiveKpiKeysForChartSeries(catKey, f);
    if (!keyStrs.length) return TRI_LABEL_FULL;
    const first = list.find((x) => String(x.kpiKey) === String(keyStrs[0]));
    if (keyStrs.length <= 1) {
      return first ? kpiDropdownLabel(first) : String(keyStrs[0]);
    }
    const primaryName = first ? kpiDropdownLabel(first) : String(keyStrs[0]);
    const n = keyStrs.length - 1;
    return primaryName + " + " + n;
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
    if (!keys.length || keys.length > 1) return null;
    const primary = keys[0];
    const meta = getKpis(catKey).find((x) => String(x.kpiKey) === String(primary));
    if (!meta || meta.unitType !== "PercentOrRate") return null;
    return primary;
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

  /** Group roll-up for KPI 503 (all BUs): same Current Period / geo / vertical filters, business ignored. */
  function incidentKeyLearningGroupPercent(poolAll, f) {
    const snapPool = applyNonMonthFiltersExceptBusiness(poolAll, f, true);
    const winMonths = new Set(effectiveBizWindowMonths(f));
    const rows = snapPool.filter((r) => winMonths.has(r.yearMonth));
    const nums = rows
      .filter((r) => String(r.kpiKey) === INCIDENT_KEY_LEARNING_SPEEDOMETER_KPI_KEY)
      .map((r) => Number(r.value))
      .filter(Number.isFinite);
    if (!nums.length) return 0;
    return Math.max(0, Math.min(100, avg(nums)));
  }

  function readBusinessSelectionFromDom() {
    const cbs = document.querySelectorAll("input.f-biz-cb");
    if (!cbs.length) return "all";
    const checked = Array.from(cbs).filter((cb) => cb.checked);
    if (!checked.length) return [];
    if (checked.length === cbs.length) return "all";
    return checked.map((cb) => cb.value);
  }

  function readSiteSelectionFromDom() {
    const cbs = document.querySelectorAll("input.f-site-cb");
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

  function updateSiteFilterSummary() {
    const cbs = document.querySelectorAll("input.f-site-cb");
    const el = document.getElementById("f-site-hint");
    if (!el) return;
    const total = cbs.length;
    if (!total) {
      el.textContent = "";
      return;
    }
    const n = document.querySelectorAll("input.f-site-cb:checked").length;
    el.textContent = n + "/" + total;
  }

  function initBusinessSiteCheckboxFilters() {
    document.querySelectorAll("input.f-biz-cb").forEach((cb) => {
      cb.checked = true;
    });
    document.querySelectorAll("input.f-site-cb").forEach((cb) => {
      cb.checked = true;
    });
    updateBizFilterSummary();
    updateSiteFilterSummary();
  }

  /** Checkbox lists (KPI-style) for Business unit + Site. */
  function wireBusinessAndSiteFilterControls(onChange) {
    function wireOne(kind) {
      const isBiz = kind === "biz";
      const panelId = isBiz ? "f-biz-panel" : "f-site-panel";
      const detailsId = isBiz ? "f-biz-details" : "f-site-details";
      const allId = isBiz ? "f-biz-btn-all" : "f-site-btn-all";
      const noneId = isBiz ? "f-biz-btn-none" : "f-site-btn-none";
      const cbSel = isBiz ? "input.f-biz-cb" : "input.f-site-cb";
      const panel = document.getElementById(panelId);
      const allBtn = document.getElementById(allId);
      const noneBtn = document.getElementById(noneId);
      if (!panel) return;
      function listCbs() {
        return panel.querySelectorAll(cbSel);
      }
      function refresh() {
        if (isBiz) updateBizFilterSummary();
        else updateSiteFilterSummary();
        if (onChange) onChange();
      }
      function closeIfAllSelected() {
        const cbs = listCbs();
        const tot = cbs.length;
        const n = panel.querySelectorAll(cbSel + ":checked").length;
        if (tot && n === tot) {
          const det = document.getElementById(detailsId);
          if (det) det.open = false;
        }
      }
      if (allBtn) {
        allBtn.addEventListener("click", () => {
          listCbs().forEach((cb) => {
            cb.checked = true;
          });
          refresh();
          const det = document.getElementById(detailsId);
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
          if (isBiz) updateBizFilterSummary();
          else updateSiteFilterSummary();
          closeIfAllSelected();
          if (onChange) onChange();
        });
      });
      if (isBiz) updateBizFilterSummary();
      else updateSiteFilterSummary();
    }
    wireOne("biz");
    wireOne("site");
  }

  function readFilters(catKey) {
    const elSt = document.getElementById("f-state");
    const elPersonal = document.getElementById("f-personal");
    const elCurFrom = document.getElementById("f-cur-from");
    const elCurTo = document.getElementById("f-cur-to");
    const elCmpFrom = document.getElementById("f-cmp-from");
    const elCmpTo = document.getElementById("f-cmp-to");
    const kpisMeta = kpiListForFilterDropdown(catKey);
    const rawKeys = document.getElementById("f-kpi-panel")
      ? readSelectedKpiKeysFromDom()
      : [];
    let kpiKeys = orderKpiKeysForCharts(catKey, rawKeys);
    const dk = defaultKpiKeyForCategory(catKey, kpisMeta);
    const kpi = kpiKeys.length ? kpiKeys[0] : dk;
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
    if (!comparisonFrom) comparisonFrom = defs.cmpFrom || monthAdd(currentFrom, -12);

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

    function normalizeToolbarMulti(sel) {
      if (Array.isArray(sel) && sel.length === 0) return "all";
      return sel;
    }

    return {
      catKey,
      kpi: kpi,
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
      business: normalizeToolbarMulti(readBusinessSelectionFromDom()),
      site: normalizeToolbarMulti(readSiteSelectionFromDom()),
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

  /** Inclusive list of YYYY-MM between bounds (fact data is month-granularity). */
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

  function defaultToolbarPeriodRanges() {
    const mb = dataMonthBounds();
    const curTo = mb.max;
    const curFrom = monthAdd(curTo, -11);
    const cmpTo = monthAdd(curTo, -12);
    const cmpFrom = monthAdd(curFrom, -12);
    return { curFrom, curTo, cmpFrom, cmpTo };
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

  /** Trend and line charts: months in the selected Current Period. */
  function effectiveChartMonths(f) {
    const months = currentPeriodMonthList(f);
    if (months.length) return months;
    const ref = (f && (f.monthTo || f.refMonth)) || getRefMonth();
    return ref ? [ref] : [];
  }

  /** SPI helpers: up to `count` months from Current Period (end-aligned). */
  function calendarClippedMonthKeys(f, count) {
    const full = currentPeriodMonthList(f);
    if (!full.length) return [];
    if (full.length <= count) return full;
    return full.slice(-count);
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
    const d = rowDummySite(r);
    if (Array.isArray(f.site)) {
      if (!f.site.length) return false;
      return f.site.indexOf(d) !== -1;
    }
    return d === f.site;
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


  /**
   * Preview-only: scale VL fact values by state index so top/bottom five states
   * (by combined KPI sum) separate clearly — trend, comparison, and map show green vs red.
   */
  function vlFactRowPreviewContrast(r) {
    const st = String(r.state || "").trim();
    const idx = INDIA_STATES_UT.indexOf(st);
    if (idx === -1) return r;
    const v = Number(r.value);
    if (!Number.isFinite(v)) return r;
    const n = INDIA_STATES_UT.length;
    const t = n > 1 ? idx / (n - 1) : 0;
    const mult = 2.35 - t * 1.75;
    const nv = Math.max(0.0001, v * mult);
    return Object.assign({}, r, { value: nv });
  }

  function getRowsForCategory(catKey) {
    const ck = Number(catKey);
    if (ck === VULNERABLE_LOCATION_CATEGORY_KEY) {
      if (!Array.isArray(DATA.factRows)) return [];
      return DATA.factRows
        .filter((r) => VULNERABLE_LOCATION_KPI_KEYS.has(Number(r.kpiKey)))
        .map((r) =>
          vlFactRowPreviewContrast(
            Object.assign({}, r, { categoryKey: VULNERABLE_LOCATION_CATEGORY_KEY })
          )
        );
    }
    return DATA.factRows.filter((r) => Number(r.categoryKey) === ck);
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
    const ck = Number(catKey);
    if (ck === VULNERABLE_LOCATION_CATEGORY_KEY) {
      return {
        showKpi: false,
        showState: true,
        showBusiness: false,
      };
    }
    return {
      showKpi: true,
      showState: false,
      showBusiness: true,
    };
  }

  function categoryShowsKpiToolbar(catKey) {
    return getFilterConfig(catKey).showKpi;
  }

  /** Trend & Distribution view: one KPI checkbox; Table / Comparison allow multi-select. */
  function updateKpiToolbarButtonsVisibility(mode) {
    if (currentCategoryKey == null || !categoryShowsKpiToolbar(currentCategoryKey)) {
      return;
    }
    const all = document.getElementById("f-kpi-all");
    const none = document.getElementById("f-kpi-none");
    if (!all || !none) return;
    const showMulti = mode !== "charts";
    all.hidden = !showMulti;
    none.hidden = !showMulti;
  }

  function enforceChartViewSingleKpiSelection() {
    if (currentCategoryKey == null || !categoryShowsKpiToolbar(currentCategoryKey)) {
      return false;
    }
    const main = document.getElementById("cat-main-view");
    if (!main || main.dataset.view !== "charts") return false;
    const panel = document.getElementById("f-kpi-panel");
    if (!panel) return false;
    const boxes = panel.querySelectorAll('input[name="f-kpi-cb"]');
    if (!boxes.length) return false;
    const catKey = currentCategoryKey;
    const checked = Array.from(boxes).filter((b) => b.checked);
    if (checked.length <= 1) return false;

    const ordered = orderKpiKeysForCharts(
      catKey,
      checked.map((x) => x.value)
    );
    const keepKey = ordered[0] ? String(ordered[0]) : String(checked[0].value);
    boxes.forEach((cb) => {
      cb.checked = String(cb.value) === keepKey;
    });
    saveKpiSelection(catKey, readSelectedKpiKeysFromDom());
    return true;
  }

  function avg(nums) {
    if (!nums.length) return 0;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  }

  /**
   * Incident Data: Fatality → FAC → Injury Illness incidents (4) → Total Recordable Injuries (58) → …
   * (preview-seeded 56 / 57 / 58).
   */
  const INCIDENT_KPI_ORDER = [
    8,
    9,
    10,
    11,
    12,
    4,
    INCIDENT_TOTAL_RECORDABLE_INJURIES_KPI_KEY,
    3,
    28,
    5,
    7,
    INCIDENT_FIRE_KPI_KEY,
    22,
    1,
    2,
    44,
    14,
    15,
    19,
    INCIDENT_PROPERTY_DAMAGE_KPI_KEY,
    21,
  ];
  /**
   * Hazard & Observation (leading): stakeholder sequence for the main preview —
   * Observations → Spotting closure % → SI (planned vs actual) → unsafe acts/hr in SI → Near miss → Life saved.
   */
  const HAZARD_KPI_ORDER = [38, 39, 54, 53, 13, 6];
  /**
   * Consequence Management (4): stakeholder order — Total CMP action Taken → Fatality vs CMP →
   * Job band → Category → LTI vs CMP (five KPIs; no merged extras).
   */
  const CONSEQUENCE_KPI_ORDER = [25, 23, 27, 26, 24];
  /**
   * Training & Competency Development (6): Total Training Conducted → Intensity → Executive index →
   * Frontline Worker index → SAKSHAM (five KPIs).
   */
  const TRAINING_KPI_ORDER = [34, 35, 37, 36, 49];
  /**
   * Leadership & Safety Governance (8): council → taskforce governance → walkthroughs → observations (four KPIs).
   */
  const LEADERSHIP_KPI_ORDER = [47, 48, 43, 55];
  /**
   * Risk Control Programs (7): Total Actions → SI observation closure → Total SRFA →
   * S-4 / S-5 SRFA → Near Miss FR (preview-seeded; Dim keys 46 / 45 / 40 / 41 / 42 / 20).
   */
  const RISK_CONTROL_KPI_ORDER = [46, 45, 40, 41, 42, 20];
  /**
   * Assurance and Compliance (5): Incident Key Learning → Standards implementation → FRC rate.
   */
  const ASSURANCE_KPI_ORDER = [503, 502, 501];

  function sortKpisForDisplay(catKey, kpisMeta) {
    const cat = Number(catKey);
    const list = kpisMeta.slice();
    if (cat === 1) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      const ordered = INCIDENT_KPI_ORDER.map((id) => map.get(id)).filter(
        Boolean
      );
      const used = new Set(ordered.map((k) => Number(k.kpiKey)));
      const rest = list.filter((k) => !used.has(Number(k.kpiKey)));
      return ordered.concat(rest);
    }
    if (cat === HAZARD_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      return HAZARD_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (cat === ASSURANCE_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      return ASSURANCE_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (cat === CONSEQUENCE_MANAGEMENT_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      return CONSEQUENCE_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (cat === SPI_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      const ordered = SPI_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
      const rest = list.filter(
        (k) => !SPI_KPI_ORDER.includes(Number(k.kpiKey))
      );
      return ordered.concat(rest);
    }
    if (cat === TRAINING_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      return TRAINING_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (cat === LEADERSHIP_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      return LEADERSHIP_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (cat === RISK_CONTROL_PROGRAMS_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      return RISK_CONTROL_KPI_ORDER.map((id) => map.get(id)).filter(Boolean);
    }
    if (cat === VULNERABLE_LOCATION_CATEGORY_KEY) {
      const map = new Map(list.map((k) => [Number(k.kpiKey), k]));
      return VULNERABLE_LOCATION_KPI_ORDER.map((id) => map.get(id)).filter(
        Boolean
      );
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

  /** Snapshot charts (BU / vertical / radar): aggregate over Current Period months. */
  function effectiveBizWindowMonths(f) {
    const months = currentPeriodMonthList(f);
    if (months.length) return months;
    const ref = (f && f.refMonth) || getRefMonth();
    return ref ? [ref] : [];
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

  function calendarSliceCaption(_f) {
    void _f;
    return "";
  }

  /** Human-readable current vs comparison periods (toolbar date ranges). */
  function comparisonVsPeriodCaption(f) {
    const cur = formatMonthRangeShort(currentPeriodMonthList(f));
    const cmp = formatMonthRangeShort(comparisonPeriodMonthList(f));
    return cur + " (current) vs " + cmp + " (comparison)";
  }

  /** Trend line / SPI multi-line: months from `effectiveChartMonths` (Current Period). */
  function trendSeriesPeriodCaption(f) {
    const months = effectiveChartMonths(f);
    const n = months.length;
    const range = formatMonthRangeShort(months);
    return (
      "Series: " + range + (n !== 1 ? " (" + n + " mo)" : "")
    );
  }

  /** By business / radar / vertical / scatter: months rolled into each snapshot row. */
  function bizWindowRowsCaption(f) {
    const months = effectiveBizWindowMonths(f);
    const n = months.length;
    const range = formatMonthRangeShort(months);
    return (
      "Aggregated: " +
      range +
      (n !== 1 ? " (" + n + " mo)" : "")
    );
  }

  /** SPI 100% stacked mix: up to 12 months ending at calendar “to”. */
  function spiMixCompositionCaption(f) {
    const months = calendarClippedMonthKeys(f, 12);
    const n = months.length;
    const range = formatMonthRangeShort(months);
    return (
      "Months: " + range + (n !== 1 ? " (" + n + " mo)" : "")
    );
  }

  /** Short subtitle for SPI heat map (no duplicate calendar lines). */
  function spiHeatmapChartHint(f) {
    const months = calendarClippedMonthKeys(f, 6);
    if (!months.length) {
      return "Adjust the Current Period range to include months.";
    }
    const range = formatMonthRangeShort(months);
    return (
      "Monthly averages · " +
      range +
      " · darker = higher within that KPI row"
    );
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

  /**
   * Same as `applyNonMonthFilters` for the primary KPI, but rolls up across all businesses
   * (for Group-level metrics vs a single BU slice).
   */
  function applyNonMonthFiltersExceptBusiness(rows, f, primaryOnly) {
    const keys = primaryOnly
      ? [String(f.kpi)]
      : f.kpiKeys && f.kpiKeys.length
        ? f.kpiKeys.map(String)
        : [String(f.kpi)];
    return rows.filter((r) => {
      if (!keys.includes(String(r.kpiKey))) return false;
      if (f.state !== "all" && r.state !== f.state) return false;
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

  /** BU matrix table: CSS classes for pastel +/− styling (styles.css .bu-cell__value--*). */
  function matrixVsValueClass(pct) {
    if (pct == null || Number.isNaN(Number(pct))) {
      return "bu-cell__value bu-cell__value--na";
    }
    const p = Number(pct);
    if (Math.abs(p) < 1e-9) {
      return "bu-cell__value bu-cell__value--neutral";
    }
    return p > 0
      ? "bu-cell__value bu-cell__value--pos"
      : "bu-cell__value bu-cell__value--neg";
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

  /** Roll up one KPI over a list of year-months (additive sum or average by unit rules). */
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
    const meta = Array.isArray(kpisMeta) ? kpisMeta : [];
    const sorted = sortKpisForDisplay(catKey, meta);
    const base = applyNonMonthFiltersAllKpis(getRowsForCategory(catKey), f);
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

  /** Vs % for one BU × KPI — Current Period vs Comparison Period (toolbar ranges). */
  function buKpiVsPct(basePool, buName, kpiMeta, f) {
    const kk = kpiMeta.kpiKey;
    const ut = kpiMeta.unitType;
    const bu = factBusinessNameForPreview(buName);
    const curM = currentPeriodMonthList(f);
    const cmpM = comparisonPeriodMonthList(f);

    function aggregateOverMonths(months) {
      const add = isAdditiveUnit(ut);
      function aggMonth(ym) {
        if (!ym) return null;
        const rows = basePool.filter(
          (r) =>
            String(r.kpiKey) === String(kk) &&
            String(r.yearMonth) === String(ym) &&
            String(r.businessName || "").trim() === bu
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

    const cur = aggregateOverMonths(curM);
    const baseVal = aggregateOverMonths(cmpM);

    return {
      pct: pctChange(cur, baseVal),
      cur: cur,
      baseVal: baseVal,
    };
  }

  /**
   * One BU × KPI — aggregate **current period** months only, rows whose state is in `stateSet`.
   * Used for Strategic Overview comparison: resilient vs vulnerable states per BU.
   */
  function buKpiCurrentPeriodForStateSet(basePool, buName, kpiMeta, f, stateSet) {
    const kk = kpiMeta.kpiKey;
    const ut = kpiMeta.unitType;
    const bu = factBusinessNameForPreview(buName);
    const curM = currentPeriodMonthList(f);
    const stSet =
      stateSet instanceof Set
        ? stateSet
        : new Set(
            Array.from(stateSet || []).map((s) => String(s || "").trim()).filter(Boolean)
          );

    function aggregateOverMonths(months) {
      const add = isAdditiveUnit(ut);
      function aggMonth(ym) {
        if (!ym) return null;
        const rows = basePool.filter(
          (r) =>
            String(r.kpiKey) === String(kk) &&
            String(r.yearMonth) === String(ym) &&
            String(r.businessName || "").trim() === bu &&
            stSet.has(String(r.state || "").trim())
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

    const v = aggregateOverMonths(curM);
    return v != null && !Number.isNaN(Number(v)) ? Number(v) : null;
  }

  /** Vs % for one Site × KPI (dummy site label) within filtered business slice. */
  function siteKpiVsPct(basePool, siteLabel, kpiMeta, f) {
    const kk = kpiMeta.kpiKey;
    const ut = kpiMeta.unitType;
    const curM = currentPeriodMonthList(f);
    const cmpM = comparisonPeriodMonthList(f);
    const site = String(siteLabel || "").trim();

    function aggregateOverMonths(months) {
      const add = isAdditiveUnit(ut);
      function aggMonth(ym) {
        if (!ym) return null;
        const rows = basePool.filter(
          (r) =>
            String(r.kpiKey) === String(kk) &&
            String(r.yearMonth) === String(ym) &&
            rowDummySite(r) === site
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

    const cur = aggregateOverMonths(curM);
    const baseVal = aggregateOverMonths(cmpM);

    return {
      pct: pctChange(cur, baseVal),
      cur: cur,
      baseVal: baseVal,
    };
  }

  /** Group roll-up across all preview BUs — same periods as `buKpiVsPct`. */
  function groupKpiVsPct(basePool, kpiMeta, f) {
    const kk = kpiMeta.kpiKey;
    const ut = kpiMeta.unitType;
    const add = isAdditiveUnit(ut);
    const bus = PREVIEW_BUSINESS_NAMES.slice();
    const curM = currentPeriodMonthList(f);
    const cmpM = comparisonPeriodMonthList(f);

    function aggMonth(ym) {
      if (!ym) return null;
      const buVals = [];
      bus.forEach((previewBu) => {
        const b = factBusinessNameForPreview(previewBu);
        const rows = basePool.filter(
          (r) =>
            r.kpiKey === kk &&
            r.yearMonth === ym &&
            String(r.businessName || "").trim() === b
        );
        if (!rows.length) return;
        let v;
        if (add) {
          v = rows.reduce((a, r) => a + Number(r.value), 0);
        } else {
          v = avg(rows.map((r) => r.value));
        }
        buVals.push(v);
      });
      if (!buVals.length) return null;
      if (add) {
        return buVals.reduce((a, b) => a + b, 0);
      }
      return avg(buVals);
    }

    function aggregateOverMonths(months) {
      const vals = months
        .map(aggMonth)
        .filter((v) => v != null && !Number.isNaN(v));
      if (!vals.length) return null;
      if (add) return vals.reduce((a, b) => a + b, 0);
      return avg(vals);
    }

    const cur = aggregateOverMonths(curM);
    const baseVal = aggregateOverMonths(cmpM);

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
        (f.currentFrom || "") +
        "|" +
        (f.comparisonTo || "") +
        "|" +
        (f.state || "") +
        "|" +
        (Array.isArray(f.business)
          ? f.business
              .slice()
              .sort()
              .join("\u001f")
          : String(f.business || ""))
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
    f,
    tileClass,
    selectedKpiKeys,
    presentationSeed,
    catKey
  ) {
    void presentationSeed;
    void selectedKpiKeys;
    const tile = document.createElement("div");
    tile.className = (tileClass || "multi-kpi-tile") + " multi-kpi-tile--action";
    tile.setAttribute("role", "button");
    tile.setAttribute("tabindex", "0");
    const vsShort = vsTagShort(item.vsMode || "vs_period");
    const pres = tileVsDisplay(item);
    const vDir = pres.dir;
    if (vDir === "up") tile.classList.add("multi-kpi-tile--trend-up");
    else if (vDir === "down") tile.classList.add("multi-kpi-tile--trend-down");
    const infoTitle = escapeAttr(periodComparisonTooltip(f));
    const vsLbl = "Current vs comparison period";
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
      '<div class="multi-kpi-tile__head">' +
      '<div class="multi-kpi-tile__name">' +
      escapeHtml(item.kpiName) +
      "</div>" +
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
      '" aria-label="Period comparison for current filter"><span class="multi-kpi-tile__info-icon" aria-hidden="true">i</span></button>' +
      "</div>";
    const infoBtn = tile.querySelector(".multi-kpi-tile__info");
    if (infoBtn) {
      infoBtn.addEventListener("click", function (e) {
        e.stopPropagation();
      });
    }
    function openCompareForKpi() {
      applyToolbarSingleKpiSelection(catKey, item.kpiKey);
      if (Number(catKey) === LEADERSHIP_CATEGORY_KEY) {
        return;
      }
      const tCharts = document.getElementById("view-tab-charts");
      if (tCharts) tCharts.click();
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
    f,
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

    function appendMetricChunk(items) {
      if (!items.length) return;
      const card = document.createElement("div");
      card.className = "multi-kpi-card";
      const grid = document.createElement("div");
      grid.className = "multi-kpi-card__grid";
      items.forEach((item) => {
        appendKpiTileEl(
          grid,
          item,
          f,
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
      card.appendChild(grid);
      rowMain.appendChild(card);
    }

    const chunks = chunkArray(headItems, 4);
    chunks.forEach((chunk) => {
      appendMetricChunk(chunk);
    });

    const tailA = tailItems.slice(0, 4);
    const tailB = tailItems.slice(4);
    if (tailA.length) {
      appendMetricChunk(tailA);
    }
    if (tailB.length) {
      appendMetricChunk(tailB);
    }

    container.appendChild(rowMain);
  }

  /** Strategic Overview: grouped bars per BU — resilient (top 5 states) vs vulnerable (bottom 5) for current period. */
  function buildVulnerableLocationResilientVsVulnerableBuComparison(
    catKey,
    f,
    el,
    hint,
    titleEl,
    mode
  ) {
    if (mode.singleSite) {
      if (hint) {
        hint.textContent =
          "Single site selected · resilient vs vulnerable comparison by business unit is hidden · " +
          comparisonVsPeriodCaption(f);
      }
      if (titleEl) {
        titleEl.textContent = "Resilient vs vulnerable sites — by business unit";
      }
      return;
    }
    const pk =
      f.kpiKeys && f.kpiKeys.length
        ? String(f.kpiKeys[0])
        : String(f.kpi);
    let kMeta = getKpis(catKey).find((x) => String(x.kpiKey) === pk);
    if (!kMeta) {
      const ordered = sortKpisForDisplay(catKey, getKpis(catKey));
      kMeta = ordered[0];
    }
    if (!kMeta) return;

    const rank = vulnerableLocationStateRankingFromFilters(f);
    const resSet = new Set(
      (rank.top5 || []).map((s) => String(s || "").trim()).filter(Boolean)
    );
    const vulSet = new Set(
      (rank.bottom5 || []).map((s) => String(s || "").trim()).filter(Boolean)
    );
    const basePool = applyNonMonthFiltersAllKpis(getRowsForCategory(catKey), f);
    const bus = PREVIEW_BUSINESS_NAMES.slice();
    const resData = [];
    const vulData = [];
    bus.forEach((bu) => {
      const rv = buKpiCurrentPeriodForStateSet(basePool, bu, kMeta, f, resSet);
      const vv = buKpiCurrentPeriodForStateSet(basePool, bu, kMeta, f, vulSet);
      resData.push(rv != null ? rv : 0);
      vulData.push(vv != null ? vv : 0);
    });
    const labelList = bus.map(locCompareBizLabel);
    const locListForTooltip = bus.slice();
    const unit = kMeta.unitType || "";
    const yTitle = unit === "PercentOrRate" ? "Value (%)" : "Value";
    const curRange = formatMonthRangeShort(currentPeriodMonthList(f));
    const subLines =
      "Current period " +
      curRange +
      " · each pair of bars is one BU — green = KPI total in resilient states (top 5), coral = in vulnerable states (bottom 5), same ranking as the map.";

    new Chart(el, {
      type: "bar",
      data: {
        labels: labelList,
        datasets: [
          {
            label: "Resilient sites (top 5 states/UTs)",
            data: resData,
            backgroundColor: "rgba(52, 211, 153, 0.62)",
            borderColor: "#34d399",
            borderWidth: 1,
          },
          {
            label: "Vulnerable sites (bottom 5 states/UTs)",
            data: vulData,
            backgroundColor: "rgba(251, 113, 133, 0.62)",
            borderColor: "#fb7185",
            borderWidth: 1,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        layout: {
          padding: { top: 4, right: 6, bottom: 28, left: 6 },
        },
        plugins: {
          subtitle: {
            display: true,
            text: subLines,
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 4 },
          },
          legend: {
            display: true,
            position: "bottom",
            align: "center",
            labels: {
              boxWidth: 10,
              padding: 4,
              font: {
                size: 9,
                family: FONT_UI,
              },
              color: CHART_INK,
            },
          },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0].dataIndex;
                return locListForTooltip[i] != null
                  ? String(locListForTooltip[i])
                  : "";
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
              color: CHART_INK,
              padding: 4,
            },
            title: {
              display: true,
              text: "Business unit",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
              padding: { top: 2, bottom: 0 },
            },
          },
          y: {
            beginAtZero: true,
            grace: "5%",
            ticks: { font: { size: 9 }, color: CHART_INK },
            grid: { color: "rgba(109, 110, 113, 0.09)" },
            title: {
              display: true,
              text: yTitle,
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
              padding: { top: 0, bottom: 4 },
            },
          },
        },
      },
    });
    if (hint) {
      hint.textContent =
        "(Current period · resilient vs vulnerable states per BU · map ranking) · " +
        curRange;
    }
    if (titleEl) {
      titleEl.textContent =
        shortKpiHeaderLabel(kMeta.kpiName) +
        " — resilient vs vulnerable sites by business unit";
    }
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        const ch = Chart.getChart(el);
        if (ch) {
          try {
            ch.resize();
          } catch {
            /* ignore */
          }
        }
      });
    });
  }

  /** Grouped bars — comparison (base) vs current KPI value per BU or per site (single-BU slice). */
  function buildBuComparisonChart(catKey, f) {
    const el = document.getElementById("chart-bu-compare");
    if (!el) return;
    if (typeof Chart === "undefined") {
      if (
        Number(catKey) === VULNERABLE_LOCATION_CATEGORY_KEY &&
        vlBuCompareChartAttempts < 80
      ) {
        vlBuCompareChartAttempts++;
        window.setTimeout(function () {
          if (Number(currentCategoryKey) === VULNERABLE_LOCATION_CATEGORY_KEY) {
            buildBuComparisonChart(catKey, f);
          }
        }, 50);
      }
      return;
    }
    vlBuCompareChartAttempts = 0;
    const prev = Chart.getChart(el);
    if (prev) prev.destroy();
    const hint = document.getElementById("chart-bu-compare-hint");
    const titleEl = document.getElementById("chart-bu-compare-title");
    const mode = chartUiMode(catKey, f);
    if (Number(catKey) === VULNERABLE_LOCATION_CATEGORY_KEY) {
      buildVulnerableLocationResilientVsVulnerableBuComparison(
        catKey,
        f,
        el,
        hint,
        titleEl,
        mode
      );
      return;
    }
    if (mode.singleSite) {
      if (hint) {
        hint.textContent =
          "Single site selected · comparison across locations is hidden · " +
          comparisonVsPeriodCaption(f);
      }
      if (titleEl) titleEl.textContent = "Location comparison";
      return;
    }
    const pk =
      f.kpiKeys && f.kpiKeys.length
        ? String(f.kpiKeys[0])
        : String(f.kpi);
    let kMeta = getKpis(catKey).find((x) => String(x.kpiKey) === pk);
    if (!kMeta) {
      const ordered = sortKpisForDisplay(catKey, getKpis(catKey));
      kMeta = ordered[0];
    }
    if (!kMeta) return;
    const basePool = applyNonMonthFiltersAllKpis(getRowsForCategory(catKey), f);
    const isAssurance = catKey === ASSURANCE_CATEGORY_KEY;
    const bus = PREVIEW_BUSINESS_NAMES.slice();
    const sites = PREVIEW_SITE_LABELS.slice();
    const curData = [];
    const baseData = [];
    let labelList = [];
    let xAxisTitle = "Business unit";
    let locListForTooltip = bus;

    if (mode.siteRadarInsteadOfBu) {
      xAxisTitle = "Site";
      locListForTooltip = sites;
      sites.forEach((site) => {
        const x = siteKpiVsPct(basePool, site, kMeta, f);
        const c = x.cur;
        const b = x.baseVal;
        curData.push(c != null && !Number.isNaN(Number(c)) ? Number(c) : 0);
        baseData.push(b != null && !Number.isNaN(Number(b)) ? Number(b) : 0);
      });
      labelList = sites.map(locCompareBizLabel);
    } else {
      bus.forEach((bu) => {
        const x = buKpiVsPct(basePool, bu, kMeta, f);
        const c = x.cur;
        const b = x.baseVal;
        curData.push(c != null && !Number.isNaN(Number(c)) ? Number(c) : 0);
        baseData.push(b != null && !Number.isNaN(Number(b)) ? Number(b) : 0);
      });
      if (isAssurance) {
        const g = groupKpiVsPct(basePool, kMeta, f);
        const gc =
          g.cur != null && !Number.isNaN(Number(g.cur)) ? Number(g.cur) : 0;
        const gb =
          g.baseVal != null && !Number.isNaN(Number(g.baseVal))
            ? Number(g.baseVal)
            : 0;
        curData.unshift(gc);
        baseData.unshift(gb);
      }
      labelList = isAssurance
        ? ["Group"].concat(bus.map(locCompareBizLabel))
        : bus.map(locCompareBizLabel);
    }
    const cmpRange = formatMonthRangeShort(comparisonPeriodMonthList(f));
    const unit = kMeta.unitType || "";
    const yTitle = unit === "PercentOrRate" ? "Value (%)" : "Value";
    const hasGroup = isAssurance && !mode.siteRadarInsteadOfBu;
    new Chart(el, {
      type: "bar",
      data: {
        labels: labelList,
        datasets: [
          {
            label: "Comparison (base) · " + cmpRange,
            data: baseData,
            backgroundColor: "rgba(196, 181, 253, 0.75)",
            borderColor: "rgba(139, 92, 246, 0.85)",
            borderWidth: 1,
          },
          {
            label: "Current period",
            data: curData,
            backgroundColor: "rgba(125, 211, 252, 0.82)",
            borderWidth: 1,
            borderColor: "rgba(14, 165, 233, 0.9)",
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        layout: {
          padding: { top: 4, right: 6, bottom: 28, left: 6 },
        },
        plugins: {
          subtitle: {
            display: true,
            text: comparisonVsPeriodCaption(f),
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 4 },
          },
          legend: {
            display: true,
            position: "bottom",
            align: "center",
            labels: {
              boxWidth: 10,
              padding: 4,
              font: {
                size: 9,
                family: FONT_UI,
              },
              color: CHART_INK,
            },
          },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0].dataIndex;
                if (hasGroup && i === 0) return "Group";
                const bi = hasGroup ? i - 1 : i;
                return locListForTooltip[bi] != null
                  ? String(locListForTooltip[bi])
                  : "";
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
              color: CHART_INK,
              padding: 4,
            },
            title: {
              display: true,
              text: mode.siteRadarInsteadOfBu
                ? "Site"
                : isAssurance
                  ? "Group / business unit"
                  : "Business unit",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
              padding: { top: 2, bottom: 0 },
            },
          },
          y: {
            beginAtZero: true,
            grace: "5%",
            ticks: { font: { size: 9 }, color: CHART_INK },
            grid: { color: "rgba(109, 110, 113, 0.09)" },
            title: {
              display: true,
              text: yTitle,
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
              padding: { top: 0, bottom: 4 },
            },
          },
        },
      },
    });
    if (hint) {
      if (mode.siteRadarInsteadOfBu) {
        hint.textContent =
          "(" +
          vsShort +
          " · 2 bars per column · sites in selected BU) · " +
          comparisonVsPeriodCaption(f);
      } else {
        hint.textContent =
          "(" +
          vsShort +
          " · 2 bars per column" +
          (isAssurance ? " · Group + all BUs" : " · all BUs") +
          ") · " +
          comparisonVsPeriodCaption(f);
      }
    }
    if (titleEl) {
      if (mode.siteRadarInsteadOfBu) {
        titleEl.textContent =
          shortKpiHeaderLabel(kMeta.kpiName) + " — comparison by site";
      } else {
        titleEl.textContent =
          shortKpiHeaderLabel(kMeta.kpiName) +
          (isAssurance
            ? " — Group vs business comparison"
            : " — comparison by business");
      }
    }
  }

  const HAZARD_CHART_FILL = [
    "rgba(0, 109, 182, 0.65)",
    "rgba(0, 177, 107, 0.62)",
    "rgba(142, 39, 143, 0.58)",
    "rgba(240, 76, 35, 0.55)",
    "rgba(0, 109, 182, 0.45)",
    "rgba(0, 177, 107, 0.48)",
    "rgba(142, 39, 143, 0.48)",
    "rgba(240, 76, 35, 0.45)",
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

  /** Leading Hazard (and shared markup): hazard spotting & closure heatmap (week or month columns). */
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
    const bizTag =
      f.business === "all"
        ? "All"
        : Array.isArray(f.business) && f.business.length
          ? f.business.join("+")
          : String(f.business || "All");
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
    const isHazardSpotting = f && String(f.kpi) === "38";
    const isUnsafeActsPerHour = f && String(f.kpi) === "53";
    /** Absolute counts + only green/red vs previous period (Week or Month). */
    const useHazardAbsDeltaColors = isHazardSpotting || isUnsafeActsPerHour;
    const rowTitles = isUnsafeActsPerHour
      ? ["Unsafe Acts Identified per hour in SI Round"]
      : isHazardSpotting
        ? ["Unsafe Acts", "Unsafe Condition"]
        : ["Hazard Spotting Observations", "Hazard Spotting Closure %"];
    const matrix = [];
    for (let ri = 0; ri < rowTitles.length; ri++) {
      const row = [];
      for (let ci = 0; ci < colCount; ci++) {
        const seed = ref + "|" + bizTag + "|" + mode + "|" + ri + "|" + ci;
        let raw, display;
        if (isHazardSpotting || isUnsafeActsPerHour) {
          if (mode === "week") {
            raw = 40 + (hash32(seed + "|v") % 61);
          } else {
            raw = 200 + (hash32(seed + "|v") % 201);
          }
          display = String(raw);
        } else {
          raw = -5 + (hash32(seed + "|v") % 200) / 10;
          if (ci === colCount - 1) {
            const p = -15 + (hash32(seed + "|p") % 31);
            display = (p >= 0 ? "+ " : "− ") + Math.abs(p) + "%";
          } else {
            display =
              Math.abs(raw) < 1
                ? String(raw.toFixed(4))
                : String(raw.toFixed(2));
          }
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
    const normClose =
      matrix.length > 1 ? normForRow(1, true) : normSpot;

    function prevRawForDelta(ri, ci, raw) {
      const cur = Number(raw);
      if (ci > 0) return Number(matrix[ri][ci - 1].raw);
      const h = hash32(
        ref + "|" + bizTag + "|" + mode + "|r" + ri + "|prv0"
      );
      return Math.max(0, cur - 18 + (h % 36));
    }

    function classForCell(ri, ci, raw) {
      if (useHazardAbsDeltaColors) {
        const cur = Number(raw);
        const prev = prevRawForDelta(ri, ci, raw);
        return cur >= prev
          ? "spi-hm-cell spi-hm--ua-ph-up"
          : "spi-hm-cell spi-hm--ua-ph-down";
      }
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
    if (!isUnsafeActsPerHour) {
      if (isHazardSpotting) {
        function totalRawAt(ci) {
          return (
            Number(matrix[0][ci].raw) + Number(matrix[1][ci].raw)
          );
        }
        function prevTotalRawForDelta(ci) {
          if (ci > 0) return totalRawAt(ci - 1);
          const h = hash32(
            ref + "|" + bizTag + "|" + mode + "|tot|prv0"
          );
          const cur = totalRawAt(0);
          return Math.max(0, cur - 24 + (h % 48));
        }
        html +=
          '<tr><th scope="row" class="spi-hm-rowhead">' +
          "Total (Acts + Condition)" +
          "</th>";
        for (let c = 0; c < colCount; c++) {
          const tr = totalRawAt(c);
          const disp = String(Math.round(tr));
          const prev = prevTotalRawForDelta(c);
          const cls =
            tr >= prev
              ? "spi-hm-cell spi-hm--ua-ph-up"
              : "spi-hm-cell spi-hm--ua-ph-down";
          html +=
            '<td class="' + cls + '">' + escapeHtml(disp) + "</td>";
        }
        html += "</tr>";
      } else {
        const totalNums = [];
        for (let tc = 0; tc < colCount - 1; tc++) {
          totalNums.push(matrix[0][tc].raw + matrix[1][tc].raw);
        }
        const tMn = totalNums.length ? Math.min.apply(null, totalNums) : 0;
        const tMx = totalNums.length ? Math.max.apply(null, totalNums) : 1;
        function normTotal(raw) {
          if (tMx <= tMn) return 0.5;
          return (raw - tMn) / (tMx - tMn);
        }
        function classForTotalCell(ci, totalRaw) {
          if (ci === colCount - 1) {
            return "spi-hm-cell spi-hm-cell--pct spi-hm-cell--pct-total";
          }
          const t = normTotal(totalRaw);
          if (t < 0.34) return "spi-hm-cell spi-hm--total-1";
          if (t < 0.67) return "spi-hm-cell spi-hm--total-2";
          return "spi-hm-cell spi-hm--total-3";
        }
        html +=
          '<tr><th scope="row" class="spi-hm-rowhead">' +
          "Total" +
          "</th>";
        for (let c = 0; c < colCount; c++) {
          if (c === colCount - 1) {
            html +=
              '<td class="' +
              classForTotalCell(c, 0) +
              '">—</td>';
          } else {
            const tr = matrix[0][c].raw + matrix[1][c].raw;
            const disp =
              Math.abs(tr) < 1
                ? String(tr.toFixed(4))
                : String(tr.toFixed(2));
            html +=
              '<td class="' +
              classForTotalCell(c, tr) +
              '">' +
              escapeHtml(disp) +
              "</td>";
          }
        }
        html += "</tr>";
      }
    }
    html += "</tbody></table>";
    host.innerHTML = html;

    if (foot) {
      const n = 8 + (hash32(ref + bizTag + "|spiHmFoot") % 15);
      const mom = -5 + (hash32(ref + bizTag + "|spiHmMom") % 21);
      const momStr =
        (mom >= 0 ? "▲" : "▼") + Math.abs(mom) + "% MoM";
      const periodWord = mode === "week" ? "week" : "month";
      foot.innerHTML = useHazardAbsDeltaColors
        ? "<strong>Absolute counts only</strong> (no percentages; preview sample). " +
          (isHazardSpotting
            ? "Rows: <strong>Unsafe Acts</strong>, <strong>Unsafe Condition</strong>, <strong>Total (Acts + Condition)</strong>. "
            : "<strong>Unsafe Acts Identified per hour in SI Round</strong>. ") +
          "<strong>Pastel green</strong> = same or higher than the previous " +
          periodWord +
          " (first column vs a seeded baseline for color). <strong>Pastel red</strong> = lower. Chart uses only these two colours."
        : "Total hazard-linked observations reached <strong>" +
          n +
          "</strong> (" +
          momStr +
          ") · <strong>Hazard Spotting Closure %</strong> indicators for the selected filters (preview sample).";
    }
  }

  /**
   * Hazard & Observation (leading): bar trend, horizontal BU bars, vertical doughnut (same 3-chart layout as other domains).
   */
  function buildHazardObservationCharts(
    f,
    poolAll,
    snapRows,
    snapRowsMulti,
    filteredTrend,
    utChart
  ) {
    const catKey = HAZARD_CATEGORY_KEY;
    const mode = chartUiMode(catKey, f);
    const kMetaPrimary = getKpis(catKey).find(
      (x) => String(x.kpiKey) === String(f.kpi)
    );
    const lineSeriesName = kMetaPrimary
      ? kpiDropdownLabel(kMetaPrimary)
      : "Value";

    const bizTitleLbl = document.getElementById("chart-biz-title-label");
    if (bizTitleLbl) {
      bizTitleLbl.textContent = mode.siteRadarInsteadOfBu
        ? "By site"
        : "By business";
    }

    /** Primary trend panel: heatmap for spotting KPIs, speedometer for % KPIs, line for count KPIs. */
    const HAZARD_SPEEDOMETER_KPI_KEYS = new Set(["39", "40", "45", "46", "54"]);
    const HAZARD_LINE_TREND_KPI_KEYS = new Set(["6", "13"]);
    const primaryKpiStr = String(f.kpi || "38");
    const hmHost = document.getElementById("spi-hazard-heatmap-host");
    const hmFoot = document.getElementById("spi-hazard-heatmap-foot");
    const hmPeriodBar = document.getElementById("spi-hm-period-bar");
    const hmRule = document.getElementById("spi-hazard-heatmap-rule");
    const altCanvas = document.getElementById("chart-hazard-alt");

    function showHeatmapElements(on) {
      if (hmHost) hmHost.style.display = on ? "" : "none";
      if (hmFoot) hmFoot.style.display = on ? "" : "none";
      if (hmPeriodBar) hmPeriodBar.style.display = on ? "" : "none";
      if (hmRule) hmRule.style.display = on ? "" : "none";
      if (altCanvas) altCanvas.hidden = on;
    }

    /** Multiple KPIs: multi-line trend on primary canvas (heatmap hidden). */
    if (mode.multiKpi) {
      showHeatmapElements(false);
      if (altCanvas && typeof Chart !== "undefined") {
        const monthSetM = new Set();
        filteredTrend.forEach((r) => monthSetM.add(r.yearMonth));
        const lineLabelsM = Array.from(monthSetM).sort();
        const lineSetsM = trendDatasetsForPreview(
          catKey,
          f,
          filteredTrend,
          lineLabelsM,
          "line"
        );
        if (lineLabelsM.length && lineSetsM.length) {
          altCanvas.setAttribute(
            "aria-label",
            "Line chart: monthly trend for each selected Hazard KPI"
          );
          new Chart(altCanvas, {
            type: "line",
            data: {
              labels: lineLabelsM,
              datasets: lineSetsM,
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              layout: {
                padding: { top: 10, right: 14, bottom: 10, left: 18 },
              },
              plugins: {
                legend: {
                  display: lineSetsM.length > 1,
                  position: "bottom",
                  labels: {
                    boxWidth: 10,
                    padding: 6,
                    font: { size: 8, family: FONT_UI },
                  },
                },
                subtitle: {
                  display: true,
                  text: bizWindowRowsCaption(f),
                  color: CHART_INK,
                  font: { size: 8.5, family: FONT_UI },
                  padding: { bottom: 4 },
                },
                tooltip: {
                  callbacks: {
                    label(ctx) {
                      const v =
                        ctx.parsed && ctx.parsed.y != null
                          ? ctx.parsed.y
                          : Number(ctx.raw) || 0;
                      const ut = ctx.dataset.unitType || "";
                      const name = ctx.dataset.label
                        ? String(ctx.dataset.label)
                        : "";
                      return (
                        " " +
                        (name ? name + ": " : "") +
                        formatValue(v, ut)
                      );
                    },
                  },
                },
              },
              scales: {
                y: {
                  beginAtZero: false,
                  grace: "12%",
                  ticks: {
                    font: { size: 9, family: FONT_UI },
                    color: CHART_INK,
                  },
                  grid: { color: "rgba(109,110,113,0.09)" },
                },
                x: {
                  ticks: {
                    font: { size: 8, family: FONT_UI },
                    color: CHART_INK,
                  },
                  grid: { color: "rgba(109,110,113,0.09)" },
                },
              },
            },
          });
        }
      }
    } else if (HAZARD_SPEEDOMETER_KPI_KEYS.has(primaryKpiStr)) {
      showHeatmapElements(false);
      if (altCanvas) {
        const pct = speedometerPercentFromSnapRows(catKey, snapRows, f);
        renderPercentSpeedometerChart(
          altCanvas,
          Math.max(0, Math.min(100, pct)),
          f
        );
      }
    } else if (HAZARD_LINE_TREND_KPI_KEYS.has(primaryKpiStr)) {
      showHeatmapElements(false);
      if (altCanvas && filteredTrend.length) {
        const monthSet2 = new Set();
        filteredTrend.forEach((r) => monthSet2.add(r.yearMonth));
        const lineLabels2 = Array.from(monthSet2).sort();
        const trendData2 = lineLabels2.map((ym) => {
          const vals = filteredTrend
            .filter(
              (r) =>
                String(r.kpiKey) === primaryKpiStr && r.yearMonth === ym
            )
            .map((r) => Number(r.value))
            .filter(Number.isFinite);
          return vals.length ? avg(vals) : null;
        });
        altCanvas.setAttribute(
          "aria-label",
          "Line chart: " + lineSeriesName + " trend"
        );
        new Chart(altCanvas, {
          type: "line",
          data: {
            labels: lineLabels2,
            datasets: [
              {
                label: lineSeriesName,
                data: trendData2,
                borderColor: ADANI_GREEN,
                backgroundColor: "rgba(0,177,107,0.10)",
                borderWidth: 2,
                pointRadius: 3,
                fill: true,
                tension: 0.35,
              },
            ],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
              padding: { top: 10, right: 14, bottom: 10, left: 18 },
            },
            plugins: {
              legend: { display: false },
              subtitle: {
                display: true,
                text: bizWindowRowsCaption(f),
                color: CHART_INK,
                font: { size: 8.5, family: FONT_UI },
                padding: { bottom: 4 },
              },
              tooltip: {
                callbacks: {
                  label(ctx) {
                    return " " + formatValue(Number(ctx.raw) || 0, utChart);
                  },
                },
              },
            },
            scales: {
              x: {
                ticks: { font: { size: 8, family: FONT_UI }, color: CHART_INK },
                grid: { color: "rgba(109,110,113,0.09)" },
              },
              y: {
                beginAtZero: true,
                ticks: { font: { size: 9, family: FONT_UI }, color: CHART_INK },
                grid: { color: "rgba(109,110,113,0.09)" },
              },
            },
          },
        });
      }
    } else {
      showHeatmapElements(true);
      renderSpiHazardHeatmap(f);
    }

    const hazardRollupCaption = bizWindowRowsCaption(f);

    const byBizVals = {};
    const bySiteVals = {};
    snapRows.forEach((r) => {
      const b = r.businessName || "—";
      if (!byBizVals[b]) byBizVals[b] = [];
      byBizVals[b].push(Number(r.value));
      const st = rowDummySite(r);
      if (!bySiteVals[st]) bySiteVals[st] = [];
      bySiteVals[st].push(Number(r.value));
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
    function aggSiteWindow(siteLabel) {
      const vals = bySiteVals[siteLabel];
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
    const siteNames = PREVIEW_SITE_LABELS.slice();
    const siteShort = siteNames.map(shortBu);
    const siteVals = siteNames.map((n) => aggSiteWindow(n));
    const axisNames = mode.siteRadarInsteadOfBu ? siteNames : buNames;
    const axisShort = mode.siteRadarInsteadOfBu ? siteShort : buShort;
    const axisVals = mode.siteRadarInsteadOfBu ? siteVals : buVals;
    const yAxisTitle = mode.siteRadarInsteadOfBu ? "Site" : "Business unit";
    const elBiz = document.getElementById("chart-biz");
    const emptyBiz = document.getElementById("chart-biz-empty");
    if (mode.radarVisible && elBiz && typeof Chart !== "undefined") {
      if (!axisVals.some((v) => v !== 0)) {
        elBiz.style.display = "none";
        if (emptyBiz) {
          emptyBiz.hidden = false;
          emptyBiz.textContent = mode.siteRadarInsteadOfBu
            ? "No site data for current filters."
            : "No BU data for current filters.";
        }
      } else {
        elBiz.style.display = "block";
        if (emptyBiz) emptyBiz.hidden = true;
        new Chart(elBiz, {
          type: "bar",
          data: {
            labels: axisShort,
            datasets: [
              {
                label: lineSeriesName,
                data: axisVals,
                backgroundColor: hazardChartColors(axisVals.length),
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
              subtitle: {
                display: true,
                text: hazardRollupCaption,
                color: CHART_INK,
                font: { size: 8.5, family: FONT_UI },
                padding: { bottom: 4 },
              },
              legend: { display: false },
              tooltip: {
                callbacks: {
                  title(items) {
                    const i = items[0].dataIndex;
                    return axisNames[i] != null ? String(axisNames[i]) : "";
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
                ticks: { font: { size: 9, family: FONT_UI }, color: CHART_INK },
                grid: { color: "rgba(109, 110, 113, 0.09)" },
                title: {
                  display: true,
                  text:
                    utChart === "PercentOrRate"
                      ? "KPI value (%)"
                      : "KPI value",
                  font: { size: 10, family: FONT_UI },
                  color: CHART_INK,
                },
              },
              y: {
                ticks: { font: { size: 8, family: FONT_UI }, color: CHART_INK },
                grid: { display: false },
                title: {
                  display: true,
                  text: yAxisTitle,
                  font: { size: 10, family: FONT_UI },
                  color: CHART_INK,
                },
              },
            },
          },
        });
      }
    }

    if (mode.verticalVisible) {
      if (mode.multiKpi) {
        renderVerticalCheckpointLineChartMulti(
          catKey,
          f,
          snapRowsMulti
        );
      } else {
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
                subtitle: {
                  display: true,
                  text: hazardRollupCaption,
                  color: CHART_INK,
                  font: { size: 8.5, family: FONT_UI },
                  padding: { bottom: 4 },
                },
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
    ctx.fillStyle = "#eef1f4";
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
    ctx.fillStyle = "#334155";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(Math.round(p) + "%", tx, ty);
  }

  /**
   * Dual semi-circular gauge (reference: outer scale + inner scale, Group vs Business on % SPI).
   * Outer arc: traffic-light thirds; inner arc: orange (0–75%) + grey (75–100%); light needle = Group,
   * dark blue tick = Business; purple tick on inner = Business.
   */
  function drawDualGaugeSpeedometerCanvas(canvasEl, groupPct, bizPct) {
    const g = Math.max(0, Math.min(100, Number(groupPct) || 0));
    const b = Math.max(0, Math.min(100, Number(bizPct) || 0));
    const dpr = window.devicePixelRatio || 1;
    const rect = canvasEl.getBoundingClientRect();
    let cssW = rect.width;
    let cssH = rect.height;
    if (cssW < 2 || cssH < 2) {
      const wrap = canvasEl.parentElement;
      if (wrap) {
        cssW = wrap.clientWidth || 320;
        cssH = wrap.clientHeight || 220;
      } else {
        cssW = 320;
        cssH = 220;
      }
    }
    canvasEl.width = Math.max(1, Math.round(cssW * dpr));
    canvasEl.height = Math.max(1, Math.round(cssH * dpr));
    const ctx = canvasEl.getContext("2d");
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);

    const padX = Math.max(10, cssW * 0.06);
    const padB = Math.max(30, cssH * 0.2);
    const cx = cssW / 2;
    const cy = cssH - padB;
    const R = Math.min(cssW - 2 * padX, (cssH - padB) * 1.05) * 0.48;

    function pToTheta(pct) {
      return Math.PI + (Math.PI * pct) / 100;
    }

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, R * 0.88, Math.PI, 2 * Math.PI, false);
    ctx.closePath();
    ctx.fillStyle = "#eef1f4";
    ctx.fill();
    ctx.restore();

    const rInnerMid = R * 0.58;
    const lwIn = Math.max(6, R * 0.1);
    ctx.lineWidth = lwIn;
    ctx.lineCap = "round";
    ctx.beginPath();
    ctx.arc(cx, cy, rInnerMid, pToTheta(0), pToTheta(75), false);
    ctx.strokeStyle = "#FFAB40";
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(cx, cy, rInnerMid, pToTheta(75), pToTheta(100), false);
    ctx.strokeStyle = "#ECEFF1";
    ctx.stroke();

    const rOuterMid = R * 0.92;
    const lwOut = Math.max(8, R * 0.1);
    ctx.lineWidth = lwOut;
    const t1 = 100 / 3;
    const t2 = (2 * 100) / 3;
    const outerCols = ["#E57373", "#FFD54F", "#66BB6A"];
    const segBounds = [0, t1, t2, 100];
    for (let i = 0; i < 3; i++) {
      ctx.beginPath();
      ctx.arc(
        cx,
        cy,
        rOuterMid,
        pToTheta(segBounds[i]),
        pToTheta(segBounds[i + 1]),
        false
      );
      ctx.strokeStyle = outerCols[i];
      ctx.stroke();
    }

    const tb = pToTheta(b);
    const tg = pToTheta(g);
    ctx.lineWidth = Math.max(3, R * 0.04);
    ctx.strokeStyle = ADANI_PURPLE;
    ctx.beginPath();
    ctx.moveTo(
      cx + Math.cos(tb) * rInnerMid * 0.82,
      cy + Math.sin(tb) * rInnerMid * 0.82
    );
    ctx.lineTo(
      cx + Math.cos(tb) * rInnerMid * 1.12,
      cy + Math.sin(tb) * rInnerMid * 1.12
    );
    ctx.stroke();

    const needleLen = rOuterMid * 0.88;
    ctx.lineWidth = Math.max(2, R * 0.025);
    ctx.strokeStyle = "#4FC3F7";
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.lineTo(
      cx + Math.cos(tg) * needleLen,
      cy + Math.sin(tg) * needleLen
    );
    ctx.stroke();

    const tickR0 = rInnerMid * 1.05;
    const tickR1 = rOuterMid * 1.02;
    ctx.lineWidth = Math.max(4, R * 0.035);
    ctx.strokeStyle = "#0D47A1";
    ctx.beginPath();
    ctx.moveTo(cx + Math.cos(tb) * tickR0, cy + Math.sin(tb) * tickR0);
    ctx.lineTo(cx + Math.cos(tb) * tickR1, cy + Math.sin(tb) * tickR1);
    ctx.stroke();

    ctx.beginPath();
    ctx.arc(cx, cy, Math.max(5, R * 0.065), 0, 2 * Math.PI);
    ctx.fillStyle = "#1565C0";
    ctx.fill();

    ctx.font = "600 " + Math.max(10, Math.round(R * 0.14)) + "px " + FONT_UI;
    ctx.fillStyle = "#334155";
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.fillText(
      "Group " + Math.round(g) + "% · Business " + Math.round(b) + "%",
      cx,
      cy + R * 0.1
    );

    ctx.font = "500 " + Math.max(8, Math.round(R * 0.095)) + "px " + FONT_UI;
    ctx.fillStyle = "#64748b";
    ctx.textAlign = "center";
    ctx.fillText("0% — 100% (same scale for both rings)", cx, cy + R * 0.26);
  }

  function renderDualGaugeSpeedometerChart(canvasEl, groupPct, bizPct, f) {
    canvasEl.setAttribute("data-speedometer-gauge", "1");
    canvasEl.classList.add("chart-line--dual-gauge");
    let lastG = Number(groupPct) || 0;
    let lastB = Number(bizPct) || 0;
    function redraw() {
      drawDualGaugeSpeedometerCanvas(canvasEl, lastG, lastB);
    }
    redraw();
    window.__adaniSpeedometerGaugeRedraw = redraw;
    const cap =
      "Dual gauge: Group " +
      Math.round(lastG) +
      "%, Business " +
      Math.round(lastB) +
      "% · " +
      (f ? bizWindowRowsCaption(f) : "Current period");
    canvasEl.setAttribute("aria-label", cap);
  }

  function renderPercentSpeedometerChart(canvasEl, percent, f) {
    canvasEl.classList.remove("chart-line--dual-gauge");
    canvasEl.setAttribute("data-speedometer-gauge", "1");
    const pk =
      f && f.catKey != null ? speedometerPrimaryKpiKey(f.catKey, f) : null;
    const km =
      pk && f
        ? getKpis(f.catKey).find((x) => String(x.kpiKey) === String(pk))
        : null;
    const kpiPart = km ? kpiDropdownLabel(km) + " · " : "";
    const cap =
      f && f.refMonth
        ? "Speedometer: " + kpiPart + bizWindowRowsCaption(f)
        : "Speedometer: " +
          (km ? kpiDropdownLabel(km) + " · " : "") +
          (f ? bizWindowRowsCaption(f) : "Current period");
    canvasEl.setAttribute("aria-label", cap);
    let lastPct = Number(percent) || 0;
    function redraw() {
      drawPercentSpeedometerCanvas(canvasEl, lastPct);
    }
    redraw();
    window.__adaniSpeedometerGaugeRedraw = redraw;
  }

  /**
   * SPI absolute-value gauge thresholds.
   * Each entry: zones [{lo, hi, color}] (low → high), max scale, reverse (higher = better).
   */
  function getSpiKpiAbsThresholds(kpiKey) {
    const G = "#00B16B", Y = "#F5C400", R = "#E53935";
    const k = Number(kpiKey);
    switch (k) {
      case 16:
        return { zones: [{lo:0,hi:0.05,color:G},{lo:0.05,hi:0.10,color:Y},{lo:0.10,hi:Infinity,color:R}], max: 0.15, unit: "Rate", reverse: false };
      case 17:
        return { zones: [{lo:0,hi:0.10,color:G},{lo:0.10,hi:0.20,color:Y},{lo:0.20,hi:Infinity,color:R}], max: 0.30, unit: "Rate", reverse: false };
      case 18:
        return { zones: [{lo:0,hi:0.20,color:G},{lo:0.20,hi:0.40,color:Y},{lo:0.40,hi:Infinity,color:R}], max: 0.60, unit: "Rate", reverse: false };
      case 21:
        return { zones: [{lo:0,hi:0.50,color:G},{lo:0.50,hi:1.00,color:Y},{lo:1.00,hi:Infinity,color:R}], max: 1.50, unit: "Rate", reverse: false };
      case 29:
        return { zones: [{lo:0,hi:3,color:G},{lo:3,hi:5,color:Y},{lo:5,hi:Infinity,color:R}], max: 7, unit: "Rate", reverse: false };
      case 20:
        return { zones: [{lo:0,hi:3,color:R},{lo:3,hi:4,color:Y},{lo:4,hi:Infinity,color:G}], max: 6, unit: "Rate", reverse: true };
      default:
        return null;
    }
  }

  /**
   * Draw a semi-circular absolute-value threshold gauge.
   * Zones are painted as colored arc segments (Green/Yellow/Red),
   * with tick marks and labels at zone boundaries.
   */
  function drawSpiAbsThresholdGaugeCanvas(canvasEl, value, thresholds) {
    const cfg = thresholds;
    const maxVal = cfg.max;
    const dpr = window.devicePixelRatio || 1;
    const rect = canvasEl.getBoundingClientRect();
    let cssW = rect.width || (canvasEl.parentElement ? canvasEl.parentElement.clientWidth : 0) || 300;
    let cssH = rect.height || (canvasEl.parentElement ? canvasEl.parentElement.clientHeight : 0) || 220;
    if (cssW < 2) cssW = 300;
    if (cssH < 2) cssH = 220;
    canvasEl.width = Math.max(1, Math.round(cssW * dpr));
    canvasEl.height = Math.max(1, Math.round(cssH * dpr));
    const ctx = canvasEl.getContext("2d");
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);

    const padX = Math.max(28, cssW * 0.10);
    const padB = Math.max(36, cssH * 0.18);
    const cx = cssW / 2;
    const cy = cssH - padB;
    const R = Math.min(cssW / 2 - padX, (cssH - padB) * 1.05) * 0.92;
    const rTrack = R * 0.95;
    const trackW = Math.max(8, R * 0.11);

    function valToTheta(v) {
      const t = Math.max(0, Math.min(1, v / maxVal));
      return Math.PI + Math.PI * t;
    }

    // background arc
    ctx.save();
    ctx.beginPath();
    ctx.arc(cx, cy, rTrack, Math.PI, 2 * Math.PI, false);
    ctx.strokeStyle = "#e5e7eb";
    ctx.lineWidth = trackW + 2;
    ctx.lineCap = "butt";
    ctx.stroke();
    ctx.restore();

    // colored zone segments
    for (let i = 0; i < cfg.zones.length; i++) {
      const z = cfg.zones[i];
      const a0 = valToTheta(Math.max(0, z.lo));
      const a1 = valToTheta(Math.min(maxVal, z.hi === Infinity ? maxVal : z.hi));
      ctx.beginPath();
      ctx.arc(cx, cy, rTrack, a0, a1, false);
      ctx.strokeStyle = z.color;
      ctx.lineWidth = trackW;
      ctx.lineCap = "butt";
      ctx.stroke();
    }

    // zone boundary tick marks + labels
    const boundaries = [];
    for (let i = 0; i < cfg.zones.length - 1; i++) {
      const v = cfg.zones[i].hi;
      if (v !== Infinity && v > 0) boundaries.push(v);
    }
    boundaries.push(maxVal);
    ctx.font = "600 " + Math.max(9, Math.round(R * 0.115)) + "px " + FONT_UI;
    ctx.fillStyle = "#374151";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    boundaries.forEach(function(v) {
      const theta = valToTheta(v);
      const tx0 = cx + (rTrack + trackW * 0.85) * Math.cos(theta);
      const ty0 = cy + (rTrack + trackW * 0.85) * Math.sin(theta);
      const tx1 = cx + (rTrack - trackW * 0.85) * Math.cos(theta);
      const ty1 = cy + (rTrack - trackW * 0.85) * Math.sin(theta);
      ctx.save();
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = Math.max(2, R * 0.022);
      ctx.lineCap = "round";
      ctx.beginPath();
      ctx.moveTo(tx0, ty0);
      ctx.lineTo(tx1, ty1);
      ctx.stroke();
      ctx.restore();
      if (v !== maxVal) {
        const lr = rTrack + trackW * 1.8;
        const lx = cx + lr * Math.cos(theta);
        const ly = cy + lr * Math.sin(theta);
        const labelStr = v < 1 ? v.toFixed(2) : String(v);
        ctx.fillText(labelStr, lx, ly);
      }
    });

    // 0 label
    {
      const t0 = Math.PI;
      const lr = rTrack + trackW * 1.8;
      ctx.font = "600 " + Math.max(9, Math.round(R * 0.115)) + "px " + FONT_UI;
      ctx.fillStyle = "#374151";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText("0", cx + lr * Math.cos(t0), cy + lr * Math.sin(t0));
    }

    // needle
    const clampedVal = Math.max(0, Math.min(maxVal, Number(value) || 0));
    const needleTheta = valToTheta(clampedVal);
    const needleLen = R * 0.78;
    ctx.save();
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.lineTo(cx + needleLen * Math.cos(needleTheta), cy + needleLen * Math.sin(needleTheta));
    ctx.strokeStyle = "#1F2937";
    ctx.lineWidth = Math.max(2, R * 0.028);
    ctx.lineCap = "round";
    ctx.stroke();
    ctx.restore();

    // needle pivot
    ctx.beginPath();
    ctx.arc(cx, cy, Math.max(5, R * 0.065), 0, 2 * Math.PI);
    ctx.fillStyle = "#374151";
    ctx.fill();

    // value label near needle tip
    const valLabelR = needleLen + Math.max(14, R * 0.16);
    const vlx = cx + valLabelR * Math.cos(needleTheta);
    const vly = cy + valLabelR * Math.sin(needleTheta);
    const dispStr = clampedVal < 1 ? clampedVal.toFixed(4) : clampedVal.toFixed(2);
    ctx.font = "700 " + Math.max(12, Math.round(R * 0.20)) + "px " + FONT_UI;
    ctx.fillStyle = "#0f172a";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(dispStr, vlx, vly);

    // status label at bottom
    let zone = cfg.zones[cfg.zones.length - 1];
    for (let i = 0; i < cfg.zones.length; i++) {
      const z = cfg.zones[i];
      if (clampedVal >= z.lo && (z.hi === Infinity || clampedVal < z.hi)) {
        zone = z;
        break;
      }
    }
    const statusLabel =
      zone.color === "#00B16B" ? "Within Target" :
      zone.color === "#F5C400" ? "Caution" : "Above Limit";
    ctx.font = "500 " + Math.max(10, Math.round(R * 0.135)) + "px " + FONT_UI;
    ctx.fillStyle = zone.color;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(statusLabel, cx, cy - R * 0.12);
  }

  function renderSpiAbsThresholdGauge(canvasEl, value, kpiKey, f) {
    const cfg = getSpiKpiAbsThresholds(kpiKey);
    if (!cfg) return;
    const kpis = getKpis(SPI_CATEGORY_KEY);
    const meta = kpis.find((x) => String(x.kpiKey) === String(kpiKey));
    const kpiLabel = meta ? kpiDropdownLabel(meta) : ("KPI " + kpiKey);
    canvasEl.setAttribute(
      "aria-label",
      "Threshold gauge: " + kpiLabel + " · current value " + (Number(value) || 0).toFixed(4)
    );
    drawSpiAbsThresholdGaugeCanvas(canvasEl, value, cfg);
  }

  /**
   * Consequence Management / KPIs 24–25: CMP accountability — three 100% stacked
   * bars (preview proportions; match design reference).
   */
  function renderCmpAccountabilityBreakdownChart(canvasEl, f) {
    const cmpR = f ? formatMonthRangeShort(comparisonPeriodMonthList(f)) : "—";
    const periodLine =
      "Design preview · Ref. " +
      formatMonthYear((f && f.refMonth) || "") +
      " · Comparison " +
      cmpR;
    canvasEl.setAttribute(
      "aria-label",
      "CMP accountability breakdown: three stacked percentage bars. " + periodLine
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
          seg("CMP Done", ADANI_GREEN, [72, null, null]),
          seg("Training", "#2E9FE6", [null, 32, null]),
          seg("PPE", "#A855BC", [null, 22, null]),
          seg("Procedure", "#4FC3F7", [null, 18, null]),
          seg("L1-L3", "#FFD54F", [null, null, 25]),
          seg("L4-L6", ADANI_ORANGE, [null, null, 30]),
          seg("L7+", "#FFEE58", [null, null, 17]),
          {
            label: "Not Done",
            data: [28, 28, 28],
            backgroundColor: "#FF8A80",
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
          subtitle: {
            display: true,
            text: periodLine,
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 2 },
          },
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
              color: CHART_INK,
              maxRotation: 0,
            },
            grid: { display: false },
            border: {
              display: true,
              color: "rgba(109, 110, 113, 0.28)",
            },
            title: {
              display: true,
              text: "CMP breakdown dimension",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
            },
          },
          y: {
            stacked: true,
            min: 0,
            max: 100,
            ticks: {
              stepSize: 25,
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
              padding: 6,
            },
            border: {
              display: true,
              color: "rgba(109, 110, 113, 0.28)",
            },
            grid: {
              color: "rgba(109, 110, 113, 0.2)",
              borderDash: [2, 4],
            },
            title: {
              display: true,
              text: "% of Total Incidents",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
              padding: { top: 0, bottom: 6, left: 0, right: 0 },
            },
          },
        },
      },
    });
  }

  /**
   * By vertical (checkpoint line) — same roll-up window as By business (radar).
   */
  function renderVerticalCheckpointLineChart(f, snapRows, utChart) {
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
    if (
      !elVert ||
      !vertLabels.length ||
      typeof Chart === "undefined"
    ) {
      return;
    }
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
          subtitle: {
            display: true,
            text: bizWindowRowsCaption(f),
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 4 },
          },
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
              color: CHART_INK,
              padding: 6,
            },
            grid: { color: "rgba(109, 110, 113, 0.1)" },
            title: {
              display: true,
              text: "Value",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
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
              color: CHART_INK,
              padding: 4,
            },
            grid: { color: "rgba(109, 110, 113, 0.065)" },
            title: {
              display: true,
              text: "Vertical",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
            },
          },
        },
      },
    });
  }

  /**
   * By vertical — one line per selected KPI (Trend & Distribution when multiple KPIs selected).
   */
  function renderVerticalCheckpointLineChartMulti(catKey, f, snapRows) {
    const keys = effectiveKpiKeysForChartSeries(catKey, f);
    const kpisLine = getKpis(catKey);
    const vertSet = new Set();
    snapRows.forEach((r) => vertSet.add(getRowCheckpoint(r)));
    const vertLabelsFull = Array.from(vertSet).sort((a, b) =>
      String(a).localeCompare(String(b))
    );
    function shortVertLabel(name) {
      const s = String(name);
      return s.length > 14 ? s.slice(0, 13) + "…" : s;
    }
    const vertLabels = vertLabelsFull.map(shortVertLabel);

    const datasets = keys.map((kk, idx) => {
      const kMeta = kpisLine.find((x) => String(x.kpiKey) === String(kk));
      const ut = kMeta ? kMeta.unitType : "";
      const vertData = vertLabelsFull.map((label) => {
        const slice = snapRows.filter(
          (r) =>
            String(r.kpiKey) === String(kk) && getRowCheckpoint(r) === label
        );
        const vals = slice.map((r) => Number(r.value));
        if (ut && isAdditiveUnit(ut)) {
          return vals.reduce(
            (a, x) => a + (Number.isFinite(x) ? x : 0),
            0
          );
        }
        const nums = vals.filter((x) => Number.isFinite(x));
        return nums.length ? avg(nums) : 0;
      });
      const color = TREND_LINE_COLORS[idx % TREND_LINE_COLORS.length];
      return {
        label: kMeta ? kpiDropdownLabel(kMeta) : kk,
        data: vertData,
        unitType: ut,
        borderColor: color,
        backgroundColor: hexToRgba(color, 0.08),
        fill: false,
        tension: 0.35,
        borderWidth: 2,
        pointRadius: 3,
        pointBackgroundColor: color,
      };
    });

    const elVert = document.getElementById("chart-verticals");
    if (!elVert || !vertLabels.length || typeof Chart === "undefined") {
      return;
    }
    new Chart(elVert, {
      type: "line",
      data: {
        labels: vertLabels,
        datasets: datasets,
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
          padding: { top: 12, right: 20, bottom: 16, left: 28 },
        },
        plugins: {
          subtitle: {
            display: true,
            text: bizWindowRowsCaption(f),
            color: CHART_INK,
            font: { size: 8.5, family: FONT_UI },
            padding: { bottom: 4 },
          },
          legend: {
            display: datasets.length > 1,
            position: "bottom",
            labels: {
              boxWidth: 10,
              padding: 6,
              font: { size: 9, family: FONT_UI },
            },
          },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0].dataIndex;
                return vertLabelsFull[i] || "";
              },
              label(ctx) {
                const v =
                  ctx.parsed && ctx.parsed.y != null
                    ? ctx.parsed.y
                    : Number(ctx.raw) || 0;
                const ut = ctx.dataset.unitType || "";
                return (
                  " " +
                  (ctx.dataset.label || "") +
                  ": " +
                  formatValue(v, ut)
                );
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
              color: CHART_INK,
              padding: 6,
            },
            grid: { color: "rgba(109, 110, 113, 0.1)" },
            title: {
              display: true,
              text: "Value",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
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
              color: CHART_INK,
              padding: 4,
            },
            grid: { color: "rgba(109, 110, 113, 0.065)" },
            title: {
              display: true,
              text: "Vertical",
              font: { size: 10, family: FONT_UI },
              color: CHART_INK,
            },
          },
        },
      },
    });
  }

  /**
   * Trend + (SPI: by vertical only) or (by business + by vertical). SPI uses multi-KPI line + vertical for focused KPI.
   */
  function buildCharts(catKey, f) {
    destroyCharts();

    const poolAll = getRowsForCategory(catKey);
    const filteredTrend = applyChartFilter(poolAll, f);
    const mode = chartUiMode(catKey, f);
    applyChartLayoutVisibility(catKey, f);

    const snapPoolSingle = applyNonMonthFilters(poolAll, f, true);
    const snapPoolMulti = applyNonMonthFilters(poolAll, f, false);
    const winMonths = new Set(effectiveBizWindowMonths(f));
    const snapRows = snapPoolSingle.filter((r) => winMonths.has(r.yearMonth));
    const snapRowsMulti = snapPoolMulti.filter((r) => winMonths.has(r.yearMonth));

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
        snapRowsMulti,
        filteredTrend,
        utChart
      );
      updateChartHints(f);
      return;
    }

    if (catKey === SPI_CATEGORY_KEY) {
      const snapPoolSpi = applyNonMonthFilters(poolAll, f, true);
      const winMonthsSpi = new Set(effectiveBizWindowMonths(f));
      const snapRowsSpi = snapPoolSpi.filter((r) =>
        winMonthsSpi.has(r.yearMonth)
      );
      const elLineSpi = document.getElementById("chart-line");
      const lineChartHostSpi = elLineSpi
        ? elLineSpi.closest(".chart-box")
        : null;
      if (lineChartHostSpi) {
        lineChartHostSpi.classList.remove(
          "chart-box--cmp-breakdown",
          "chart-box--assurance-gauge",
          "chart-box--speedometer-gauge"
        );
      }
      // SPI: always show trend line on left
      renderSpiKpiTrendLineChart(poolAll, f);

      // SPI right panel: abs threshold gauge for single KPI, by-vertical for multi
      const elGauge = document.getElementById("chart-spi-abs-gauge");
      const elVert = document.getElementById("chart-verticals");
      const spiKeys = effectiveKpiKeysForChartSeries(catKey, f);
      const spiSingleKey = spiKeys.length === 1 ? spiKeys[0] : null;
      const spiThresholds = spiSingleKey ? getSpiKpiAbsThresholds(spiSingleKey) : null;
      const vertTitleEl = document.getElementById("chart-vertical-title-label");
      const vertHintEl = document.getElementById("chart-vertical-hint");

      if (spiThresholds && elGauge) {
        // Show gauge, hide vertical chart
        elGauge.hidden = false;
        if (elVert) elVert.hidden = true;
        if (vertTitleEl) {
          const spiMeta = getKpis(catKey).find(
            (x) => String(x.kpiKey) === String(spiSingleKey)
          );
          vertTitleEl.textContent = "Gauge" + (spiMeta ? ": " + kpiDropdownLabel(spiMeta) : "");
        }
        if (vertHintEl) vertHintEl.textContent = "Threshold bands · current value";
        const absVal = (function() {
          const rows = snapRowsSpi.filter(
            (r) => String(r.kpiKey) === String(spiSingleKey)
          );
          const nums = rows.map((r) => Number(r.value)).filter(Number.isFinite);
          return nums.length ? avg(nums) : 0;
        }());
        renderSpiAbsThresholdGauge(elGauge, absVal, spiSingleKey, f);
      } else {
        // Show vertical chart, hide gauge
        if (elGauge) elGauge.hidden = true;
        if (elVert) elVert.hidden = false;
        if (vertTitleEl) vertTitleEl.textContent = "By vertical";
        if (vertHintEl) vertHintEl.textContent = "Verticals · Current Period";
        if (mode.multiKpi) {
          renderVerticalCheckpointLineChartMulti(catKey, f, snapRowsMulti);
        } else {
          renderVerticalCheckpointLineChart(f, snapRowsSpi, utChart);
        }
      }
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
      renderCmpAccountabilityBreakdownChart(elLine, f);
    } else if (showSpeedometerGauge && elLine) {
      if (lineChartHost) {
        lineChartHost.classList.add("chart-box--speedometer-gauge");
      }
      const pkSpd = speedometerPrimaryKpiKey(catKey, f);
      if (
        pkSpd === INCIDENT_KEY_LEARNING_SPEEDOMETER_KPI_KEY &&
        Number(catKey) === ASSURANCE_CATEGORY_KEY
      ) {
        renderDualGaugeSpeedometerChart(
          elLine,
          incidentKeyLearningGroupPercent(poolAll, f),
          Math.max(0, Math.min(100, speedometerPct)),
          f
        );
      } else {
        renderPercentSpeedometerChart(elLine, speedometerPct, f);
      }
    } else if (elLine && lineLabels.length && lineSets.length) {
      elLine.setAttribute(
        "aria-label",
        lineSets.length > 1
          ? "Line chart: monthly average for each selected KPI in the trend window"
          : "Line chart: monthly average for the selected KPI in the trend window"
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
            subtitle: {
              display: true,
              text: trendSeriesPeriodCaption(f),
              color: CHART_INK,
              font: { size: 8.5, family: FONT_UI },
              padding: { bottom: 4 },
            },
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
                  const name = ctx.dataset.label
                    ? String(ctx.dataset.label)
                    : "";
                  return (
                    " " +
                    (name ? name + ": " : "") +
                    formatValue(v, ut)
                  );
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
                color: CHART_INK,
                padding: 6,
              },
              grid: { color: "rgba(109, 110, 113, 0.1)" },
              title: {
                display: true,
                text: "KPI value",
                font: { size: 10, family: FONT_UI },
                color: CHART_INK,
              },
            },
            x: {
              ticks: {
                font: { size: 10 },
                maxRotation: 45,
                color: CHART_INK,
                padding: 4,
              },
              grid: { color: "rgba(109, 110, 113, 0.08)" },
              title: {
                display: true,
                text: "Month",
                font: { size: 10, family: FONT_UI },
                color: CHART_INK,
              },
            },
          },
        },
      });
    }

    const bizTitleLbl = document.getElementById("chart-biz-title-label");
    if (bizTitleLbl) {
      bizTitleLbl.textContent = mode.siteRadarInsteadOfBu
        ? "By site"
        : "By business";
    }

    if (mode.radarVisible) {
      if (mode.siteRadarInsteadOfBu) {
        const bySiteVals = {};
        snapRows.forEach((r) => {
          const s = rowDummySite(r);
          if (!bySiteVals[s]) bySiteVals[s] = [];
          bySiteVals[s].push(Number(r.value));
        });
        function aggSiteWindow(siteLabel) {
          const vals = bySiteVals[siteLabel];
          if (!vals || !vals.length) return 0;
          if (utChart && isAdditiveUnit(utChart)) {
            return vals.reduce((a, x) => a + Number(x), 0);
          }
          return avg(vals);
        }
        const siteNames = PREVIEW_SITE_LABELS.slice();
        const radarValues = siteNames.map((n) => aggSiteWindow(n));
        const radarDisplayLabels = siteNames.map((n) =>
          formatBusinessSiteRowLabel(n)
        );
        renderBizBreakdown(
          radarDisplayLabels,
          radarValues,
          lineSeriesName,
          utChart,
          f
        );
      } else {
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
        const radarValues = radarNames.map((n) =>
          aggBizWindowForPreviewLabel(n)
        );
        const radarDisplayLabels = radarNames.map((n) =>
          formatBusinessSiteRowLabel(n)
        );
        renderBizBreakdown(
          radarDisplayLabels,
          radarValues,
          lineSeriesName,
          utChart,
          f
        );
      }
    }

    if (mode.verticalVisible) {
      if (mode.multiKpi) {
        renderVerticalCheckpointLineChartMulti(
          catKey,
          f,
          snapRowsMulti
        );
      } else {
        renderVerticalCheckpointLineChart(f, snapRows, utChart);
      }
    }

    updateChartHints(f);
  }

  function updateChartHints(f) {
    const trendEl = document.getElementById("chart-trend-hint");
    const lineTitleEl = document.getElementById("chart-line-title");
    const bizEl = document.getElementById("chart-biz-hint");
    const vertEl = document.getElementById("chart-vertical-hint");
    const showCmpAcc =
      f &&
      shouldShowCmpAccountabilityChart(f.catKey, f);
    const showSpeedometerSpd =
      f && shouldShowPercentSpeedometerChart(f.catKey, f);
    const showDualGaugeIk =
      f &&
      showSpeedometerSpd &&
      speedometerPrimaryKpiKey(f.catKey, f) ===
        INCIDENT_KEY_LEARNING_SPEEDOMETER_KPI_KEY;
    const hmHintEl = document.getElementById("chart-hm-hint");
    const kpiScopeTitleEl = document.getElementById("kpi-charts-scope-title");
    if (lineTitleEl && f && f.catKey != null) {
      if (showCmpAcc) {
        lineTitleEl.textContent =
          "CMP ACCOUNTABILITY - BREAKDOWN ANALYSIS";
      } else if (showSpeedometerSpd || showDualGaugeIk) {
        lineTitleEl.textContent = kpiScopeTitleWithPlusN(f.catKey, f);
      } else {
        lineTitleEl.textContent = kpiScopeTitleWithPlusN(f.catKey, f);
      }
    }
    if (kpiScopeTitleEl && f && categoryShowsKpiToolbar(f.catKey)) {
      kpiScopeTitleEl.textContent = kpiScopeTitleWithPlusN(f.catKey, f);
    }
    const isHazard = f && f.catKey === HAZARD_CATEGORY_KEY;
    if (hmHintEl && f && f.catKey === HAZARD_CATEGORY_KEY) {
      const modeBtn = document.querySelector(
        ".spi-hm-period-btn.spi-hm-period-btn--active"
      );
      const hmMode =
        modeBtn && modeBtn.getAttribute("data-spi-hm-period") === "month"
          ? "month"
          : "week";
      hmHintEl.textContent =
        (hmMode === "month" ? "(month" : "(week") +
        " · ref " +
        formatMonthYear(f.refMonth) +
        ")";
    }
    if (trendEl && f && f.catKey === SPI_CATEGORY_KEY) {
      trendEl.textContent = showSpeedometerSpd
        ? (showDualGaugeIk
            ? "Group vs business gauge · " + chartTrendHintFiltersOnly(f)
            : "Speedometer · " + chartTrendHintFiltersOnly(f))
        : "SPI KPI lines · " + chartTrendHintFiltersOnly(f);
    } else if (trendEl && f) {
      if (showDualGaugeIk) {
        trendEl.textContent =
          "Group vs business gauge · " + chartTrendHintFiltersOnly(f);
      } else if (showSpeedometerSpd) {
        trendEl.textContent =
          "Speedometer · " + chartTrendHintFiltersOnly(f);
      } else if (showCmpAcc) {
        trendEl.textContent =
          "% of Fatal Incident · " + chartTrendHintFiltersOnly(f);
      } else {
        trendEl.textContent = chartTrendHintFiltersOnly(f);
      }
    }
    if (bizEl && f) {
      const cm = chartUiMode(f.catKey, f);
      if (isHazard) {
        bizEl.textContent =
          (cm.siteRadarInsteadOfBu
            ? "(h-bars · by site) · "
            : "(h-bars · all BUs) · ") + bizWindowRowsCaption(f);
      } else {
        bizEl.textContent =
          (cm.siteRadarInsteadOfBu ? "By site · " : "") +
          bizWindowRowsCaption(f);
      }
    }
    if (vertEl && f) {
      const cm = chartUiMode(f.catKey, f);
      const prefix = isHazard
        ? cm.multiKpi
          ? "(multi-line · vertical) · "
          : "(doughnut · vertical) · "
        : cm.multiKpi
          ? "Multi-KPI by vertical · "
          : "By vertical · ";
      vertEl.textContent = prefix + bizWindowRowsCaption(f);
    }
    refreshTrendInsightTooltips(f);
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

  /** Table value column: value at the row month (see rowCompareForMode for compare logic). */
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

  function renderVulnerableLocationTableBody(catKey, f) {
    updateTableZoneLabel(f);
    const kpisMeta = sortKpisForDisplay(catKey, getKpis(catKey));
    const pool = applyNonMonthFiltersAllKpis(
      getRowsForCategory(VULNERABLE_LOCATION_CATEGORY_KEY),
      f
    ).filter((r) => VULNERABLE_LOCATION_KPI_KEYS.has(Number(r.kpiKey)));
    const win = new Set(effectiveBizWindowMonths(f));
    const snap = pool.filter((r) => win.has(r.yearMonth));
    const thead = document.querySelector("#tbl-detail thead tr");
    const tbody = document.getElementById("tbl-body");
    const sm = document.getElementById("tbl-summary");
    if (sm) sm.textContent = "";
    if (!thead || !tbody) return;

    const rowKeys = [];
    const keySet = new Set();
    for (let i = 0; i < snap.length; i++) {
      const r = snap[i];
      const bu = String(r.businessName || "").trim() || "—";
      const site = rowDummySite(r);
      const key = bu + "\0" + site;
      if (!keySet.has(key)) {
        keySet.add(key);
        rowKeys.push({ bu: bu, site: site });
      }
    }
    rowKeys.sort((a, b) => {
      const c = a.bu.localeCompare(b.bu);
      if (c !== 0) return c;
      return a.site.localeCompare(b.site);
    });

    thead.innerHTML =
      '<th scope="col" class="bu-matrix__th-bu">Business</th>' +
      '<th scope="col" class="bu-matrix__th-bu">Site</th>' +
      kpisMeta
        .map(
          (k) =>
            '<th scope="col" class="col-num">' +
            escapeHtml(shortKpiHeaderLabel(k.kpiName)) +
            "</th>"
        )
        .join("");
    let html = "";
    for (let ri = 0; ri < rowKeys.length; ri++) {
      const row = rowKeys[ri];
      html +=
        '<tr><th scope="row">' +
        escapeHtml(row.bu) +
        '</th><td class="bu-matrix__cell-site">' +
        escapeHtml(row.site) +
        "</td>";
      for (let ki = 0; ki < kpisMeta.length; ki++) {
        const k = kpisMeta[ki];
        const slice = snap.filter(
          (r) =>
            String(r.businessName || "").trim() === row.bu &&
            rowDummySite(r) === row.site &&
            String(r.kpiKey) === String(k.kpiKey)
        );
        const vals = slice
          .map((r) => Number(r.value))
          .filter((x) => Number.isFinite(x));
        let cell = "—";
        if (vals.length) {
          const ut = k.unitType || "";
          const agg = isAdditiveUnit(ut)
            ? vals.reduce((a, b) => a + b, 0)
            : avg(vals);
          cell = formatValue(agg, ut);
        }
        html += '<td class="col-num">' + escapeHtml(cell) + "</td>";
      }
      html += "</tr>";
    }
    tbody.innerHTML = html;
  }

  function renderTableBody(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    if (Number(catKey) === VULNERABLE_LOCATION_CATEGORY_KEY) {
      renderVulnerableLocationTableBody(catKey, f);
      return;
    }
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
    const singleSite =
      Array.isArray(f.site) && f.site.length === 1;
    const singleBu =
      Array.isArray(f.business) && f.business.length === 1;
    let axisRows =
      f.business === "all"
        ? PREVIEW_BUSINESS_NAMES.slice()
        : Array.isArray(f.business) && f.business.length
          ? f.business.slice()
          : [f.business].filter(Boolean);
    let useSiteAxis = false;
    if (singleSite) {
      const s = String(f.site[0] || "").trim();
      axisRows = s ? [s] : ["Site"];
      useSiteAxis = true;
    } else if (singleBu) {
      axisRows = PREVIEW_SITE_LABELS.slice();
      useSiteAxis = true;
    }
    updateTableZoneLabel(f);

    const tblDetail = document.getElementById("tbl-detail");
    const thead = tblDetail ? tblDetail.querySelector("thead tr") : null;
    const buColLabel = useSiteAxis ? "Site" : "BU";
    if (thead) {
      if (!colKpis.length) {
        thead.innerHTML =
          '<th scope="col" class="bu-matrix__th-bu">' +
          escapeHtml(buColLabel) +
          '</th><th scope="col" class="bu-matrix__th-empty">—</th>';
      } else {
        thead.innerHTML =
          '<th scope="col" class="bu-matrix__th-bu">' +
          escapeHtml(buColLabel) +
          "</th>" +
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

    let tbody = tblDetail
      ? tblDetail.querySelector("tbody")
      : document.getElementById("tbl-body");
    if (tblDetail && !tbody) {
      tbody = document.createElement("tbody");
      tbody.id = "tbl-body";
      tblDetail.appendChild(tbody);
    }
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

    if (!axisRows.length) {
      tbody.innerHTML =
        '<tr><td colspan="' +
        (colKpis.length + 1) +
        '" class="empty-msg">No row for current filters.</td></tr>';
      return;
    }

    const cellEff = axisRows.map((loc) =>
      colKpis.map((k) => {
        const { pct } = useSiteAxis
          ? siteKpiVsPct(basePool, loc, k, f)
          : buKpiVsPct(basePool, loc, k, f);
        return matrixEffectivePct(pct, catKey, loc, k.kpiKey, f);
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

    tbody.innerHTML = axisRows
      .map((bu, ri) => {
        const rowLabel = formatBusinessSiteRowLabel(bu);
        const cells = colKpis
          .map((k, ki) => {
            const eff = cellEff[ri][ki];
            const pctStr = formatSignedPct(eff.value);
            const divW = matrixBarDivergingWidths(eff.value, maxAbs);
            return (
              '<td class="bu-matrix__cell">' +
              '<div class="bu-cell bu-cell--pbi">' +
              '<div class="bu-cell__databar bu-cell__databar--diverging" role="img" aria-label="' +
              escapeAttr(rowLabel + ", " + k.kpiName + ": " + pctStr) +
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
              '<span class="' +
              matrixVsValueClass(eff.value) +
              '">' +
              escapeHtml(pctStr) +
              "</span></div></div></td>"
            );
          })
          .join("");
        return (
          '<tr class="' +
          (ri % 2 === 0 ? "bu-matrix__row--even" : "bu-matrix__row--odd") +
          '"><th scope="row" class="bu-matrix__bu">' +
          escapeHtml(rowLabel) +
          "</th>" +
          cells +
          "</tr>"
        );
      })
      .join("");
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

  function slugExportFilename(s) {
    return String(s || "export")
      .replace(/[^\w\-]+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 48);
  }

  function filterReportTextLines(catKey, f) {
    const cat = getCategory(catKey);
    const lines = [];
    lines.push("Adani Safety Performance Dashboard — applied filters (export)");
    lines.push("Generated: " + new Date().toISOString());
    lines.push("");
    lines.push(
      "Category: " +
        (cat
          ? classicIndexCategoryDisplayName(catKey, cat.categoryName)
          : String(catKey))
    );
    lines.push(
      "Business unit: " +
        (f.business === "all"
          ? "All"
          : Array.isArray(f.business)
            ? f.business.length
              ? f.business.join(", ")
              : "None"
            : f.business)
    );
    lines.push(
      "Site: " +
        (f.site === "all"
          ? "All"
          : Array.isArray(f.site)
            ? f.site.length
              ? f.site.join(", ")
              : "None"
            : f.site)
    );
    lines.push(
      "Current Period: " +
        (f.currentFrom || f.monthFrom || "—") +
        " → " +
        (f.currentTo || f.monthTo || f.refMonth || "—")
    );
    lines.push(
      "Comparison Period: " +
        (f.comparisonFrom || "—") +
        " → " +
        (f.comparisonTo || "—")
    );
    lines.push(
      "Personnel type: " +
        (f.personalType === "all" ? "All" : f.personalType)
    );
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
    if (id === "chart-spi-bubble") {
      const host = document.getElementById("chart-spi-bubble");
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
    if (
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

  /**
   * Leadership Governance page (`index.html`): multi-KPI summary row (same card pattern as Training).
   */
  function refreshLeadershipGovernanceKpiStrip(catKey) {
    if (Number(catKey) !== LEADERSHIP_CATEGORY_KEY) return;
    const multiWrap = document.getElementById("multi-kpi-wrap");
    if (!multiWrap) return;
    const kpisMeta = getKpis(catKey);
    const f = readFilters(catKey);
    if (!f) return;
    const aggList = buildKpiDetailMetrics(catKey, kpisMeta, f);
    const presentationSeed = [
      String(catKey),
      f.currentFrom || "",
      f.currentTo || "",
      f.comparisonFrom || "",
      f.comparisonTo || "",
      f.refMonth || "",
      f.state || "",
      (Array.isArray(f.business)
        ? JSON.stringify(f.business.slice().sort())
        : String(f.business || "")),
      String(f.kpi || ""),
      JSON.stringify(f.kpiKeys || []),
      JSON.stringify(f.variable || {}),
    ].join("|");
    if (!aggList.length) {
      multiWrap.innerHTML =
        '<div class="empty-msg" style="padding:8px">No KPI data for this selection.</div>';
    } else {
      renderMultiKpiCards(
        multiWrap,
        aggList,
        f,
        f.kpiKeys && f.kpiKeys.length ? f.kpiKeys : [f.kpi],
        presentationSeed,
        catKey
      );
    }
  }

  function announceFilterSummary(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    const cat = getCategory(catKey);
    if (!insightShell && Number(catKey) === LEADERSHIP_CATEGORY_KEY) {
      announce(
        (cat
          ? classicIndexCategoryDisplayName(catKey, cat.categoryName)
          : "Leadership") + ". Governance review matrix · slicers and drill-down."
      );
      return;
    }
    if (Number(catKey) === VULNERABLE_LOCATION_CATEGORY_KEY) {
      announce(
        (cat
          ? classicIndexCategoryDisplayName(catKey, cat.categoryName)
          : classicIndexCategoryDisplayName(
              VULNERABLE_LOCATION_CATEGORY_KEY,
              "Vulnerable Location"
            )) +
          ". State filter: " +
          (f.state && f.state !== "all" ? f.state : "All states") +
          ". Mapped KPIs: " +
          VULNERABLE_LOCATION_KPI_ORDER.length +
          "."
      );
      return;
    }
    const label = cat
      ? classicIndexCategoryDisplayName(catKey, cat.categoryName) + ". "
      : "";
    const buCount =
      f.business === "all"
        ? PREVIEW_BUSINESS_NAMES.length
        : Array.isArray(f.business)
          ? f.business.length
          : f.business
            ? 1
            : 0;
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

    if (!insightShell && Number(catKey) === LEADERSHIP_CATEGORY_KEY) {
      if (typeof window.__adaniLeadershipMeetingsRerender === "function") {
        window.__adaniLeadershipMeetingsRerender();
      }
      refreshLeadershipGovernanceKpiStrip(catKey);
      announceFilterSummary(catKey);
      return;
    }

    if (Number(catKey) === VULNERABLE_LOCATION_CATEGORY_KEY) {
      refreshVulnerableLocationView(catKey);
      announceFilterSummary(catKey);
      return;
    }

    const multiWrap = document.getElementById("multi-kpi-wrap");
    if (multiWrap) {
      const aggList = buildKpiDetailMetrics(catKey, kpisMeta, f);
      const presentationSeed = [
        String(catKey),
        f.currentFrom || "",
        f.currentTo || "",
        f.comparisonFrom || "",
        f.comparisonTo || "",
        f.refMonth || "",
        f.state || "",
        (Array.isArray(f.business)
          ? JSON.stringify(f.business.slice().sort())
          : String(f.business || "")),
        String(f.kpi || ""),
        JSON.stringify(f.kpiKeys || []),
        JSON.stringify(f.variable || {}),
      ].join("|");
      if (!aggList.length) {
        multiWrap.innerHTML =
          '<div class="empty-msg" style="padding:8px">No KPI data for this selection. Adjust periods or geography filters.</div>';
      } else {
        renderMultiKpiCards(
          multiWrap,
          aggList,
          f,
          f.kpiKeys && f.kpiKeys.length ? f.kpiKeys : [f.kpi],
          presentationSeed,
          catKey
        );
      }
    }

    /** Table must render even if chart code throws (otherwise tbody stays empty). */
    renderTableBody(catKey);
    buildCharts(catKey, f);
    buildBuComparisonChart(catKey, f);
    updateComparisonTabVisibility(catKey, f);
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

  function setCategoryHeaderSubtitle(name) {
    const el = document.getElementById("header-category-subtitle");
    if (!el) return;
    if (name) {
      el.textContent = name;
      el.hidden = false;
    } else {
      el.textContent = "";
      el.hidden = true;
    }
  }

  /**
   * Main preview (`index.html` only): Leadership & Safety Governance —
   * meeting compliance matrix wireframe (no KPI strip / chart tabs).
   */
  function mountLeadershipMeetingCompliancePage(catKey, cat) {
    const refM = getRefMonth();
    const months = [];
    for (let d = -5; d <= 0; d++) {
      const ym = monthAdd(refM, d);
      if (ym) months.push(ym);
    }
    const bus = PREVIEW_BUSINESS_NAMES.slice();

    function lcConductedCount(bu, ym) {
      const h = hash32(String(bu) + "|" + String(ym));
      return h % 8;
    }

    function lcMeetingFlags(bu, ym) {
      const c = lcConductedCount(bu, ym);
      const order = [0, 1, 2, 3, 4, 5, 6].sort(
        (a, b) =>
          hash32(String(bu) + "|" + String(ym) + "|o|" + a) -
          hash32(String(bu) + "|" + String(ym) + "|o|" + b)
      );
      const f = new Array(7).fill(false);
      for (let k = 0; k < c; k++) f[order[k]] = true;
      return f;
    }

    function lcDeptAttended(bu, ym, deptIdx) {
      return (
        (hash32(String(bu) + "|" + String(ym) + "|dept|" + deptIdx) & 1) ===
        0
      );
    }

    function lcRagFromCount(c) {
      if (c >= 7) return "green";
      if (c >= 5) return "yellow";
      return "red";
    }

    const ragKeys = ["green", "yellow", "red"];
    const ragShown = new Set(ragKeys);
    const buShown = new Set(bus);
    const monthShown = new Set(months);
    const selBu = new Set();
    const selMo = new Set();
    let matrixAxisMode = "bu";

    const wrap = document.createElement("div");
    wrap.className = "cat-view cat-view--leadership-meetings";
    const lcKpiTilesHtml =
      '<fieldset class="kpi-summary-region">' +
      '<legend class="visually-hidden">KPI summary for current filters</legend>' +
      '<div class="multi-kpi-row" id="multi-kpi-wrap"></div>' +
      "</fieldset>";
    wrap.innerHTML =
      '<div class="lc-page">' +
      lcKpiTilesHtml +
      '<div class="cat-top-bar cat-top-bar--filters-only">' +
      '<fieldset id="cat-heading" tabindex="-1" class="cat-toolbar cat-toolbar--compact" aria-label="Governance review filters">' +
      '<legend class="visually-hidden">Governance review filters</legend>' +
      '<div class="cat-toolbar__inner" role="group">' +
      '<div class="cat-toolbar__filters-scroll">' +
      '<div class="cat-toolbar__filters-all-scroll">' +
      '<div class="field field--variable field--var-inline">' +
      '<span class="field-label" id="lc-rag-lbl">RAG</span>' +
      '<details class="var-scope var-scope--toolbar" id="lc-rag-details">' +
      '<summary class="var-scope__summary" aria-labelledby="lc-rag-lbl" title="RAG">' +
      '<span class="var-scope__summary-text">' +
      '<span class="var-scope__hint" id="lc-rag-hint">All RAG bands</span></span>' +
      '<span class="var-scope__chev" aria-hidden="true"></span></summary>' +
      '<div class="var-scope__panel" role="group" aria-labelledby="lc-rag-lbl">' +
      '<div class="var-scope__menu">' +
      '<label class="field-variable-check field-variable-check--row field-variable-check--all">' +
      '<input type="checkbox" id="lc-rag-all" checked />' +
      '<span class="field-variable-check__text">All RAG bands</span></label>' +
      '<div class="var-scope__divider" aria-hidden="true"></div>' +
      '<div class="var-scope__options" id="lc-slicer-rag-options"></div>' +
      "</div></div></details></div>" +
      '<div class="field field--variable field--var-inline">' +
      '<span class="field-label" id="lc-bu-lbl">BU</span>' +
      '<details class="var-scope var-scope--toolbar" id="lc-bu-details">' +
      '<summary class="var-scope__summary" aria-labelledby="lc-bu-lbl" title="Business unit">' +
      '<span class="var-scope__summary-text">' +
      '<span class="var-scope__hint" id="lc-bu-hint">All BUs</span></span>' +
      '<span class="var-scope__chev" aria-hidden="true"></span></summary>' +
      '<div class="var-scope__panel" role="group" aria-labelledby="lc-bu-lbl">' +
      '<div class="var-scope__menu">' +
      '<label class="field-variable-check field-variable-check--row field-variable-check--all">' +
      '<input type="checkbox" id="lc-bu-all" checked />' +
      '<span class="field-variable-check__text">All BUs</span></label>' +
      '<div class="var-scope__divider" aria-hidden="true"></div>' +
      '<div class="var-scope__options lc-meeting-scope--bu" id="lc-slicer-bu-options"></div>' +
      "</div></div></details></div>" +
      '<div class="field field--variable field--var-inline">' +
      '<span class="field-label" id="lc-mo-lbl">Month</span>' +
      '<details class="var-scope var-scope--toolbar" id="lc-mo-details">' +
      '<summary class="var-scope__summary" aria-labelledby="lc-mo-lbl" title="Month">' +
      '<span class="var-scope__summary-text">' +
      '<span class="var-scope__hint" id="lc-mo-hint">All months</span></span>' +
      '<span class="var-scope__chev" aria-hidden="true"></span></summary>' +
      '<div class="var-scope__panel" role="group" aria-labelledby="lc-mo-lbl">' +
      '<div class="var-scope__menu">' +
      '<label class="field-variable-check field-variable-check--row field-variable-check--all">' +
      '<input type="checkbox" id="lc-mo-all" checked />' +
      '<span class="field-variable-check__text">All months</span></label>' +
      '<div class="var-scope__divider" aria-hidden="true"></div>' +
      '<div class="var-scope__options" id="lc-slicer-mo-options"></div>' +
      "</div></div></details></div>" +
      "</div></div>" +
      '<div class="toolbar-actions">' +
      '<button type="button" class="btn btn--reset-compact" id="lc-reset">Reset</button>' +
      "</div></div></fieldset></div>" +
      '<header class="lc-header">' +
      '<h2 class="lc-title">Governance Review &amp; Assurance Tracker</h2>' +
      '<div class="lc-matrix-mode" role="group" aria-label="Matrix rows">' +
      '<span class="lc-matrix-mode__lbl">View</span>' +
      '<div class="lc-matrix-mode__toggle">' +
      '<button type="button" class="lc-mode-btn lc-mode-btn--active" id="lc-mode-bu" aria-pressed="true">BU</button>' +
      '<button type="button" class="lc-mode-btn" id="lc-mode-mo" aria-pressed="false">Month</button>' +
      "</div></div>" +
      '<p class="lc-subtitle" id="lc-matrix-subtitle"></p>' +
      "</header>" +
      '<div class="lc-primary">' +
      '<div class="lc-legend" aria-label="RAG legend">' +
      '<span class="lc-legend__item"><span class="lc-legend__sw lc-legend__sw--g"></span> Green = Full compliance (7/7)</span>' +
      '<span class="lc-legend__item"><span class="lc-legend__sw lc-legend__sw--y"></span> Amber = Partial (5–6)</span>' +
      '<span class="lc-legend__item"><span class="lc-legend__sw lc-legend__sw--r"></span> Red = Low (≤4)</span>' +
      '' +
      "</div>" +
      '<div class="lc-matrix-scroll" tabindex="0">' +
      '<table class="lc-matrix lc-matrix--bu-axis" id="lc-matrix" aria-label="Governance review and assurance matrix">' +
      "<thead><tr>" +
      '<th scope="col" class="lc-matrix__corner">BU <span class="lc-matrix__hint">/ Month</span></th>' +
      "</tr></thead>" +
      '<tbody id="lc-matrix-body"></tbody></table></div></div>' +
      '<section class="lc-drill" aria-labelledby="lc-drill-h">' +
      '<h3 class="lc-drill__title" id="lc-drill-h">Drill-down detail</h3>' +
      '<p class="lc-drill__hint" id="lc-drill-hint"></p>' +
      '<div class="lc-drill__body" id="lc-drill-body"></div>' +
      "</section>" +
      '<p class="cat-context lc-footnote" id="cat-context">' +
      escapeHtml(cat.uxNote || "") +
      "</p></div>";

    function buildSlicerNodes() {
      const ragEl = wrap.querySelector("#lc-slicer-rag-options");
      const buEl = wrap.querySelector("#lc-slicer-bu-options");
      const moEl = wrap.querySelector("#lc-slicer-mo-options");
      ragEl.innerHTML = ragKeys
        .map((k) => {
          const lab =
            k === "green" ? "Green" : k === "yellow" ? "Amber" : "Red";
          return (
            '<label class="field-variable-check field-variable-check--row">' +
            '<input type="checkbox" class="lc-rag-cb" value="' +
            escapeAttr(k) +
            '" />' +
            '<span class="field-variable-check__text">' +
            escapeHtml(lab) +
            "</span></label>"
          );
        })
        .join("");
      buEl.innerHTML = bus
        .map(
          (b) =>
            '<label class="field-variable-check field-variable-check--row">' +
            '<input type="checkbox" class="lc-bu-cb" value="' +
            escapeAttr(b) +
            '" />' +
            '<span class="field-variable-check__text">' +
            escapeHtml(b) +
            "</span></label>"
        )
        .join("");
      moEl.innerHTML = months
        .map((ym) => {
          return (
            '<label class="field-variable-check field-variable-check--row">' +
            '<input type="checkbox" class="lc-mo-cb" value="' +
            escapeAttr(ym) +
            '" />' +
            '<span class="field-variable-check__text">' +
            escapeHtml(spiChartMonthTick(ym)) +
            "</span></label>"
          );
        })
        .join("");
    }

    function updateLcRagHint() {
      const hint = wrap.querySelector("#lc-rag-hint");
      const all = wrap.querySelector("#lc-rag-all");
      const cbs = wrap.querySelectorAll(".lc-rag-cb");
      if (!hint || !all) return;
      if (all.checked) {
        hint.textContent = "All RAG bands";
        return;
      }
      const n = [...cbs].filter((cb) => cb.checked).length;
      hint.textContent = n ? n + " selected" : "All RAG bands";
    }

    function updateLcBuHint() {
      const hint = wrap.querySelector("#lc-bu-hint");
      const all = wrap.querySelector("#lc-bu-all");
      const cbs = wrap.querySelectorAll(".lc-bu-cb");
      if (!hint || !all) return;
      if (all.checked) {
        hint.textContent = "All BUs";
        return;
      }
      const n = [...cbs].filter((cb) => cb.checked).length;
      hint.textContent = n ? n + " selected" : "All BUs";
    }

    function updateLcMoHint() {
      const hint = wrap.querySelector("#lc-mo-hint");
      const all = wrap.querySelector("#lc-mo-all");
      const cbs = wrap.querySelectorAll(".lc-mo-cb");
      if (!hint || !all) return;
      if (all.checked) {
        hint.textContent = "All months";
        return;
      }
      const n = [...cbs].filter((cb) => cb.checked).length;
      hint.textContent = n ? n + " selected" : "All months";
    }

    function updateLcFilterHints() {
      updateLcRagHint();
      updateLcBuHint();
      updateLcMoHint();
    }

    function readRagShownFromDom() {
      const all = wrap.querySelector("#lc-rag-all");
      ragShown.clear();
      if (all && all.checked) {
        ragKeys.forEach((k) => ragShown.add(k));
        return;
      }
      wrap.querySelectorAll(".lc-rag-cb").forEach((inp) => {
        if (inp.checked) ragShown.add(inp.value);
      });
      if (!ragShown.size) ragKeys.forEach((k) => ragShown.add(k));
    }
    function readBuShownFromDom() {
      const all = wrap.querySelector("#lc-bu-all");
      buShown.clear();
      if (all && all.checked) {
        bus.forEach((b) => buShown.add(b));
        return;
      }
      wrap.querySelectorAll(".lc-bu-cb").forEach((inp) => {
        if (inp.checked) buShown.add(inp.value);
      });
      if (!buShown.size) bus.forEach((b) => buShown.add(b));
    }
    function readMonthShownFromDom() {
      const all = wrap.querySelector("#lc-mo-all");
      monthShown.clear();
      if (all && all.checked) {
        months.forEach((m) => monthShown.add(m));
        return;
      }
      wrap.querySelectorAll(".lc-mo-cb").forEach((inp) => {
        if (inp.checked) monthShown.add(inp.value);
      });
      if (!monthShown.size) months.forEach((m) => monthShown.add(m));
    }

    function visibleMonths() {
      return months.filter((m) => monthShown.has(m));
    }

    function visibleBus() {
      return bus.filter((b) => buShown.has(b));
    }

    function rowPassesRagFilter(bu) {
      const cols = visibleMonths();
      if (!cols.length) return false;
      for (let j = 0; j < cols.length; j++) {
        const ym = cols[j];
        const rag = lcRagFromCount(lcConductedCount(bu, ym));
        if (ragShown.has(rag)) return true;
      }
      return false;
    }

    function rowPassesRagFilterMonth(ym) {
      const busB = visibleBus();
      for (let i = 0; i < busB.length; i++) {
        const rag = lcRagFromCount(lcConductedCount(busB[i], ym));
        if (ragShown.has(rag)) return true;
      }
      return false;
    }

    function paintMatrix() {
      readRagShownFromDom();
      readBuShownFromDom();
      readMonthShownFromDom();
      const cols = visibleMonths();
      const busV = visibleBus();
      const theadRow = wrap.querySelector("#lc-matrix thead tr");
      const tb = wrap.querySelector("#lc-matrix-body");
      const mat = wrap.querySelector("#lc-matrix");

      function lcCellInnerHtml(c, rag, filtered) {
        const fade = filtered ? " lc-cell-stack--muted" : "";
        const pct = Math.round((Number(c) / 7) * 100);
        const showPctLabel = matrixAxisMode !== "month";
        return (
          '<span class="lc-cell-stack' +
          fade +
          '">' +
          // progress bar
          '<span class="lc-progress" role="progressbar" aria-valuemin="0" aria-valuemax="7" aria-valuenow="' +
          String(c) +
          '">' +
          '<span class="lc-progress__fill lc-progress__fill--' +
          rag +
          '" style="width:' +
          String(pct) +
          '%"></span>' +
          "</span>" +
          // count line
          '<span class="lc-cell-countline"><span class="lc-cell-count">' +
          String(c) +
          '</span><span class="lc-cell-suffix">/7</span></span>' +
          // percent label (omit in Month view)
          (showPctLabel
            ? '<span class="lc-progress__label">' + String(pct) + "%</span>"
            : "") +
          "</span>"
        );
      }

      if (matrixAxisMode === "bu") {
        mat.setAttribute(
          "aria-label",
          "Governance review: business units by month"
        );
        const rows = busV.filter((b) => rowPassesRagFilter(b));
        theadRow.innerHTML =
          '<th scope="col" class="lc-matrix__corner">BU <span class="lc-matrix__hint">/ Month</span></th>' +
          cols
            .map(
              (ym) =>
                '<th scope="col" class="lc-matrix__mo' +
                (selMo.has(ym) ? " lc-matrix__mo--sel" : "") +
                '" data-lc-colhead="' +
                escapeAttr(ym) +
                '" title="Select month for drill-down">' +
                escapeHtml(spiChartMonthTick(ym)) +
                "</th>"
            )
            .join("");
        tb.innerHTML = rows
          .map((bu) => {
            const rowSel = selBu.has(bu) ? " lc-matrix__row--sel" : "";
            let tr =
              '<tr class="lc-matrix__row' +
              rowSel +
              '"><th scope="row" class="lc-matrix__bu" data-lc-rowhead="' +
              escapeAttr(bu) +
              '" title="Select BU for drill-down">' +
              escapeHtml(bu) +
              "</th>";
            for (let j = 0; j < cols.length; j++) {
              const ym = cols[j];
              const c = lcConductedCount(bu, ym);
              const rag = lcRagFromCount(c);
              const on = ragShown.has(rag);
              const cls =
                "lc-cell lc-cell--" +
                rag +
                (on ? "" : " lc-cell--filtered") +
                (selMo.has(ym) ? " lc-cell--col-sel" : "") +
                (selBu.has(bu) ? " lc-cell--row-sel" : "");
              tr +=
                '<td class="' +
                cls +
                '" data-lc-bu="' +
                escapeAttr(bu) +
                '" data-lc-ym="' +
                escapeAttr(ym) +
                '" title="' +
                escapeAttr(String(c) + "/7 meetings conducted") +
                '">' +
                lcCellInnerHtml(c, rag, !on) +
                "</td>";
            }
            tr += "</tr>";
            return tr;
          })
          .join("");
        return;
      }

      mat.setAttribute(
        "aria-label",
        "Governance review: months by business unit"
      );
      const rowMonths = cols.filter((ym) => rowPassesRagFilterMonth(ym));
      theadRow.innerHTML =
        '<th scope="col" class="lc-matrix__corner">Month <span class="lc-matrix__hint">/ BU</span></th>' +
        busV
          .map(
            (bu) =>
              '<th scope="col" class="lc-matrix__bu-col' +
              (selBu.has(bu) ? " lc-matrix__bu-col--sel" : "") +
              '" data-lc-bu-colhead="' +
              escapeAttr(bu) +
              '" title="Select BU for drill-down">' +
              escapeHtml(bu) +
              "</th>"
          )
          .join("");
      tb.innerHTML = rowMonths
        .map((ym) => {
          const rowSel = selMo.has(ym) ? " lc-matrix__row--sel" : "";
          let tr =
            '<tr class="lc-matrix__row' +
            rowSel +
            '"><th scope="row" class="lc-matrix__bu lc-matrix__mo-row" data-lc-mo-rowhead="' +
            escapeAttr(ym) +
            '" title="Select month for drill-down">' +
            escapeHtml(spiChartMonthTick(ym)) +
            "</th>";
          for (let i = 0; i < busV.length; i++) {
            const bu = busV[i];
            const c = lcConductedCount(bu, ym);
            const rag = lcRagFromCount(c);
            const on = ragShown.has(rag);
            const cls =
              "lc-cell lc-cell--" +
              rag +
              (on ? "" : " lc-cell--filtered") +
              (selBu.has(bu) ? " lc-cell--col-sel" : "") +
              (selMo.has(ym) ? " lc-cell--row-sel" : "");
            tr +=
              '<td class="' +
              cls +
              '" data-lc-bu="' +
              escapeAttr(bu) +
              '" data-lc-ym="' +
              escapeAttr(ym) +
              '" title="' +
              escapeAttr(String(c) + "/7 meetings conducted") +
              '">' +
              lcCellInnerHtml(c, rag, !on) +
              "</td>";
          }
          tr += "</tr>";
          return tr;
        })
        .join("");
    }

    function renderScenarioA(bu) {
      const cols = visibleMonths();
      let h =
        '<div class="lc-dd-block"><h4 class="lc-dd-block__title">BU detail — ' +
        escapeHtml(bu) +
        "</h4>" +
        '<div class="lc-dd-scroll"><table class="lc-dd-table">' +
        "<thead><tr><th scope=\"col\">Meeting type</th>";
      for (let j = 0; j < cols.length; j++) {
        h +=
          '<th scope="col">' + escapeHtml(spiChartMonthTick(cols[j])) + "</th>";
      }
      h += "</tr></thead><tbody>";
      const flagsByYm = {};
      for (let j = 0; j < cols.length; j++) {
        flagsByYm[cols[j]] = lcMeetingFlags(bu, cols[j]);
      }
      for (let mi = 0; mi < LEADERSHIP_MEETING_TYPE_LABELS.length; mi++) {
        h +=
          "<tr><th scope=\"row\">" +
          escapeHtml(LEADERSHIP_MEETING_TYPE_LABELS[mi]) +
          "</th>";
        for (let j = 0; j < cols.length; j++) {
          const ym = cols[j];
          const ok = flagsByYm[ym][mi];
          h +=
            "<td>" +
            (ok
              ? '<span class="lc-icon lc-icon--ok" title="Conducted">✔</span>'
              : '<span class="lc-icon lc-icon--no" title="Not conducted">✖</span>') +
            "</td>";
        }
        h += "</tr>";
      }
      h += "</tbody></table></div></div>";
      return h;
    }

    function renderScenarioB(ym) {
      const busB = visibleBus();
      let h =
        '<div class="lc-dd-block"><h4 class="lc-dd-block__title">Month — ' +
        escapeHtml(spiChartMonthTick(ym)) +
        "</h4>" +
        '<div class="lc-dd-scroll"><table class="lc-dd-table">' +
        "<thead><tr><th scope=\"col\">BU</th>";
      for (let mi = 0; mi < LEADERSHIP_MEETING_TYPE_LABELS.length; mi++) {
        h +=
          '<th scope="col" class="lc-dd-table__mt">' +
          escapeHtml(LEADERSHIP_MEETING_TYPE_LABELS[mi]) +
          "</th>";
      }
      h += "</tr></thead><tbody>";
      for (let i = 0; i < busB.length; i++) {
        const bu = busB[i];
        const fl = lcMeetingFlags(bu, ym);
        h += "<tr><th scope=\"row\">" + escapeHtml(bu) + "</th>";
        for (let mi = 0; mi < 7; mi++) {
          h +=
            "<td>" +
            (fl[mi]
              ? '<span class="lc-icon lc-icon--ok" title="Conducted">✔</span>'
              : '<span class="lc-icon lc-icon--no" title="Not conducted">✖</span>') +
            "</td>";
        }
        h += "</tr>";
      }
      h += "</tbody></table></div></div>";
      return h;
    }

    function renderScenarioMonthDeptGrid(ym) {
      const busB = visibleBus();
      let h =
        '<div class="lc-dd-block"><h4 class="lc-dd-block__title">' +
        escapeHtml(spiChartMonthTick(ym)) +
        " — departments × BU</h4>" +
        '<div class="lc-dd-scroll"><table class="lc-dd-table"><thead><tr><th scope="col">Department</th>';
      for (let i = 0; i < busB.length; i++) {
        h +=
          '<th scope="col" class="lc-dd-table__mt">' +
          escapeHtml(busB[i]) +
          "</th>";
      }
      h += "</tr></thead><tbody>";
      for (let di = 0; di < LEADERSHIP_DEPT_LABELS.length; di++) {
        h +=
          "<tr><th scope=\"row\">" +
          escapeHtml(LEADERSHIP_DEPT_LABELS[di]) +
          "</th>";
        for (let i = 0; i < busB.length; i++) {
          const ok = lcDeptAttended(busB[i], ym, di);
          h +=
            "<td>" +
            (ok
              ? '<span class="lc-icon lc-icon--ok" title="Done">✔</span>'
              : '<span class="lc-icon lc-icon--no" title="Not done">✖</span>') +
            "</td>";
        }
        h += "</tr>";
      }
      h += "</tbody></table></div></div>";
      return h;
    }

    function renderScenarioBuDeptAcrossMonths(bu) {
      const cols = visibleMonths();
      let h =
        '<div class="lc-dd-block"><h4 class="lc-dd-block__title">BU — ' +
        escapeHtml(bu) +
        " (departments × month)</h4>" +
        '<div class="lc-dd-scroll"><table class="lc-dd-table"><thead><tr><th scope="col">Department</th>';
      for (let j = 0; j < cols.length; j++) {
        h +=
          '<th scope="col">' + escapeHtml(spiChartMonthTick(cols[j])) + "</th>";
      }
      h += "</tr></thead><tbody>";
      for (let di = 0; di < LEADERSHIP_DEPT_LABELS.length; di++) {
        h +=
          "<tr><th scope=\"row\">" +
          escapeHtml(LEADERSHIP_DEPT_LABELS[di]) +
          "</th>";
        for (let j = 0; j < cols.length; j++) {
          const ok = lcDeptAttended(bu, cols[j], di);
          h +=
            "<td>" +
            (ok
              ? '<span class="lc-icon lc-icon--ok" title="Done">✔</span>'
              : '<span class="lc-icon lc-icon--no" title="Not done">✖</span>') +
            "</td>";
        }
        h += "</tr>";
      }
      h += "</tbody></table></div></div>";
      return h;
    }

    function paintDrill() {
      const hint = wrap.querySelector("#lc-drill-hint");
      const body = wrap.querySelector("#lc-drill-body");
      const buList = Array.from(selBu).filter((b) => buShown.has(b));
      const moList = Array.from(selMo).filter((m) => monthShown.has(m));
      if (!buList.length && !moList.length) {
        hint.textContent =
          matrixAxisMode === "month"
            ? "Month view: select a month row (or cell) for departments × BU, or a BU column for departments × months. Shift-click a cell toggles BU and month."
            : "BU view: select a month column or BU row for drill-down. Shift-click a cell toggles BU and month.";
        body.innerHTML =
          '<p class="lc-drill__empty">No drill-down selection yet.</p>';
        return;
      }
      if (matrixAxisMode === "month") {
        if (moList.length) {
          hint.textContent =
            "Departments (BSC, SRP, Contractor, Training, Logistic, Incident, Audit) × BU for each selected month.";
          const parts = [];
          moList
            .sort()
            .forEach((ym) => parts.push(renderScenarioMonthDeptGrid(ym)));
          body.innerHTML = parts.join("");
          return;
        }
        hint.textContent =
          "Departments × months for each selected BU.";
        const parts = [];
        buList.forEach((b) => parts.push(renderScenarioBuDeptAcrossMonths(b)));
        body.innerHTML = parts.join("");
        return;
      }
      if (moList.length) {
        hint.textContent =
          "Month drill-down: BUs × mandatory meeting types. BU-only drill-down is hidden while a month is selected.";
        const parts = [];
        moList.sort().forEach((ym) => parts.push(renderScenarioB(ym)));
        body.innerHTML = parts.join("");
        return;
      }
      hint.textContent =
        "BU drill-down: meeting types × visible months.";
      const parts = [];
      buList.forEach((b) => parts.push(renderScenarioA(b)));
      body.innerHTML = parts.join("");
    }

    function paintAll() {
      paintMatrix();
      paintDrill();
    }

    function setMatrixAxisMode(mode) {
      if (mode !== "bu" && mode !== "month") return;
      if (matrixAxisMode === mode) return;
      matrixAxisMode = mode;
      selBu.clear();
      selMo.clear();
      const btnBu = wrap.querySelector("#lc-mode-bu");
      const btnMo = wrap.querySelector("#lc-mode-mo");
      const sub = wrap.querySelector("#lc-matrix-subtitle");
      const mat = wrap.querySelector("#lc-matrix");
      if (btnBu && btnMo) {
        btnBu.classList.toggle("lc-mode-btn--active", mode === "bu");
        btnMo.classList.toggle("lc-mode-btn--active", mode === "month");
        btnBu.setAttribute("aria-pressed", mode === "bu" ? "true" : "false");
        btnMo.setAttribute("aria-pressed", mode === "month" ? "true" : "false");
      }
      if (sub) {
        sub.textContent = "";
      }
      if (mat) {
        mat.classList.toggle("lc-matrix--bu-axis", mode === "bu");
        mat.classList.toggle("lc-matrix--month-axis", mode === "month");
      }
      paintAll();
      announce(
        mode === "bu" ? "Matrix view: BU rows." : "Matrix view: Month rows."
      );
    }

    function onMatrixClick(e) {
      const cell = e.target.closest("td[data-lc-bu]");
      if (cell) {
        const b = cell.getAttribute("data-lc-bu");
        const ym = cell.getAttribute("data-lc-ym");
        if (e.shiftKey) {
          if (selBu.has(b)) selBu.delete(b);
          else selBu.add(b);
          if (selMo.has(ym)) selMo.delete(ym);
          else selMo.add(ym);
        } else if (matrixAxisMode === "bu") {
          if (selBu.has(b)) selBu.delete(b);
          else selBu.add(b);
        } else {
          if (selMo.has(ym)) selMo.delete(ym);
          else selMo.add(ym);
        }
        paintAll();
        announce(
          "Matrix: " +
            selBu.size +
            " BU(s) and " +
            selMo.size +
            " month(s) in drill-down."
        );
        return;
      }
      const moRh = e.target.closest("th[data-lc-mo-rowhead]");
      if (moRh) {
        const ym = moRh.getAttribute("data-lc-mo-rowhead");
        if (selMo.has(ym)) selMo.delete(ym);
        else selMo.add(ym);
        paintAll();
        announce("Drill-down: " + selMo.size + " month(s) selected.");
        return;
      }
      const buCol = e.target.closest("th[data-lc-bu-colhead]");
      if (buCol) {
        const b = buCol.getAttribute("data-lc-bu-colhead");
        if (selBu.has(b)) selBu.delete(b);
        else selBu.add(b);
        paintAll();
        announce("Drill-down: " + selBu.size + " BU(s) selected.");
        return;
      }
      const rh = e.target.closest("th[data-lc-rowhead]");
      if (rh) {
        const b = rh.getAttribute("data-lc-rowhead");
        if (selBu.has(b)) selBu.delete(b);
        else selBu.add(b);
        paintAll();
        announce("Drill-down: " + selBu.size + " BU(s) selected.");
        return;
      }
      const ch = e.target.closest("th[data-lc-colhead]");
      if (ch) {
        const ym = ch.getAttribute("data-lc-colhead");
        if (selMo.has(ym)) selMo.delete(ym);
        else selMo.add(ym);
        paintAll();
        announce("Drill-down: " + selMo.size + " month(s) selected.");
      }
    }

    function onLcFilterChange() {
      updateLcFilterHints();
      Array.from(selBu).forEach((b) => {
        if (!buShown.has(b)) selBu.delete(b);
      });
      Array.from(selMo).forEach((m) => {
        if (!monthShown.has(m)) selMo.delete(m);
      });
      paintAll();
      announce("Matrix filters updated.");
    }

    function wireLcScopeGroup(allId, cbSel, detailsId) {
      const allEl = wrap.querySelector(allId);
      if (!allEl) return;
      function cbs() {
        return wrap.querySelectorAll(cbSel);
      }
      function subChange() {
        const any = [...cbs()].some((cb) => cb.checked);
        if (any) allEl.checked = false;
        else allEl.checked = true;
        updateLcFilterHints();
        onLcFilterChange();
      }
      function allChange() {
        if (allEl.checked) {
          cbs().forEach((cb) => {
            cb.checked = false;
          });
          const det = wrap.querySelector(detailsId);
          if (det) det.open = false;
        }
        updateLcFilterHints();
        onLcFilterChange();
      }
      allEl.addEventListener("change", allChange);
      cbs().forEach((cb) => cb.addEventListener("change", subChange));
    }

    root.innerHTML = "";
    root.appendChild(wrap);
    buildSlicerNodes();
    wireToolbarScopeScrollPanels(wrap);

    wireLcScopeGroup("#lc-rag-all", ".lc-rag-cb", "#lc-rag-details");
    wireLcScopeGroup("#lc-bu-all", ".lc-bu-cb", "#lc-bu-details");
    wireLcScopeGroup("#lc-mo-all", ".lc-mo-cb", "#lc-mo-details");
    updateLcFilterHints();

    wrap.querySelector("#lc-mode-bu").addEventListener("click", () => {
      setMatrixAxisMode("bu");
    });
    wrap.querySelector("#lc-mode-mo").addEventListener("click", () => {
      setMatrixAxisMode("month");
    });

    wrap.querySelector("#lc-matrix").addEventListener("click", onMatrixClick);

    wrap.querySelector("#lc-reset").addEventListener("click", () => {
      selBu.clear();
      selMo.clear();
      matrixAxisMode = "bu";
      const btnBu = wrap.querySelector("#lc-mode-bu");
      const btnMo = wrap.querySelector("#lc-mode-mo");
      const subEl = wrap.querySelector("#lc-matrix-subtitle");
      const matEl = wrap.querySelector("#lc-matrix");
      if (btnBu && btnMo) {
        btnBu.classList.add("lc-mode-btn--active");
        btnMo.classList.remove("lc-mode-btn--active");
        btnBu.setAttribute("aria-pressed", "true");
        btnMo.setAttribute("aria-pressed", "false");
      }
      if (subEl) {
        subEl.textContent = "";
      }
      if (matEl) {
        matEl.classList.add("lc-matrix--bu-axis");
        matEl.classList.remove("lc-matrix--month-axis");
      }
      const ragAll = wrap.querySelector("#lc-rag-all");
      const buAll = wrap.querySelector("#lc-bu-all");
      const moAll = wrap.querySelector("#lc-mo-all");
      if (ragAll) ragAll.checked = true;
      if (buAll) buAll.checked = true;
      if (moAll) moAll.checked = true;
      wrap.querySelectorAll(".lc-rag-cb").forEach((cb) => {
        cb.checked = false;
      });
      wrap.querySelectorAll(".lc-bu-cb").forEach((cb) => {
        cb.checked = false;
      });
      wrap.querySelectorAll(".lc-mo-cb").forEach((cb) => {
        cb.checked = false;
      });
      wrap.querySelectorAll("details.var-scope--toolbar").forEach((d) => {
        d.open = false;
      });
      updateLcFilterHints();
      paintAll();
      announce("Governance review filters reset.");
    });

    window.__adaniLeadershipMeetingsRerender = paintAll;
    paintAll();
    refreshLeadershipGovernanceKpiStrip(catKey);
    announceFilterSummary(catKey);
    const h = document.getElementById("cat-heading");
    if (h) h.focus();
    updateHeaderNavState();
  }

  function mountVulnerableLocationCategoryPage(
    catKey,
    cat,
    _cfg,
    periodRangesFieldHtml,
    bizList
  ) {
    void _cfg;
    const businessFieldHtml = businessFilterFieldHtml(bizList || []);
    const variableFieldHtml = variableFilterFieldHtml();
    const siteFieldHtml = siteFilterFieldHtml();
    const filtersScroll =
      businessFieldHtml +
      variableFieldHtml +
      siteFieldHtml +
      periodRangesFieldHtml;
    const comparePanelHtml =
      '<div id="view-panel-vl-compare" class="view-panel view-panel--vl-compare" role="tabpanel" aria-labelledby="view-tab-vl-compare" hidden><div class="chart-box chart-box--bu-compare">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-bu-compare-title">Resilient vs vulnerable sites by business unit</span> <span id="chart-bu-compare-hint" class="chart-box__hint">Current period · top 5 vs bottom 5 states (map rule) · all BUs</span>',
        "chart-bu-compare",
        "comparison-by-bu-vl",
        CHART_HELP.compareBuVl
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--bu-compare"><canvas id="chart-bu-compare" role="img" aria-label="Resilient versus vulnerable sites by business unit, current period"></canvas></div></div></div>';
    const wrap = document.createElement("div");
    wrap.className = insightShell
      ? "cat-view cat-view--modern cat-view--vulnerable-location"
      : "cat-view cat-view--vulnerable-location";
    wrap.innerHTML =
      '<div class="cat-top-bar cat-top-bar--filters-only">' +
      '<fieldset id="cat-heading" tabindex="-1" class="cat-toolbar cat-toolbar--compact cat-toolbar--vl" aria-label="Filters: business, verticals, sites, and periods">' +
      '<legend class="visually-hidden">Business, verticals, sites, and period filters</legend>' +
      '<div class="cat-toolbar__inner" role="group">' +
      '<div class="cat-toolbar__filters-scroll">' +
      '<div class="cat-toolbar__filters-all-scroll">' +
      filtersScroll +
      "</div></div>" +
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
      "</div></div></fieldset></div>" +
      // KPI surface (four KPI tiles)
      '<div class="cat-main-kpis" id="cat-main-kpis">' +
      '<div class="kpi-row vl-kpi-row">' +
      '<div class="vl-kpi"><div class="vl-kpi__label">Total man hours (employees)</div><div class="vl-kpi__value" id="vl-kpi-emp"></div></div>' +
      '<div class="vl-kpi"><div class="vl-kpi__label">Total man hours (contractors)</div><div class="vl-kpi__value" id="vl-kpi-con"></div></div>' +
      '<div class="vl-kpi"><div class="vl-kpi__label">Total man hours (emp + con)</div><div class="vl-kpi__value" id="vl-kpi-total"></div></div>' +
      '<div class="vl-kpi"><div class="vl-kpi__label">Total safe man hours</div><div class="vl-kpi__value" id="vl-kpi-safe"></div></div>' +
      "</div></div>" +
      '<div class="cat-main-view cat-main-view--vl" id="cat-main-view" data-view="trend">' +
      '<div class="view-tabs" role="tablist" aria-label="Trend and Distribution, table, comparison, or map">' +
      '<button type="button" role="tab" id="view-tab-vl-trend" class="view-tabs__btn view-tabs__btn--active" aria-selected="true" aria-controls="view-panel-vl-trend" data-view="trend">Trend &amp; Distribution View</button>' +
      '<button type="button" role="tab" id="view-tab-vl-table" class="view-tabs__btn" aria-selected="false" aria-controls="view-panel-vl-table" tabindex="-1" data-view="table">Table view</button>' +
      '<button type="button" role="tab" id="view-tab-vl-compare" class="view-tabs__btn" aria-selected="false" aria-controls="view-panel-vl-compare" tabindex="-1" data-view="compare">Comparison view</button>' +
      '<button type="button" role="tab" id="view-tab-vl-map" class="view-tabs__btn" aria-selected="false" aria-controls="view-panel-vl-map" tabindex="-1" data-view="map">Map view</button>' +
      "</div>" +
      // Trend & Distribution panel (two charts)
      '<div id="view-panel-vl-trend" class="view-panel view-panel--vl-trend" role="tabpanel" aria-labelledby="view-tab-vl-trend">' +
      '<div class="chart-row">' +
      '<div class="chart-box chart-box--vl-trend">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-vl-trend-title">Trend — selected KPI</span>',
        "chart-vl-trend",
        "vl-trend",
        CHART_HELP.trend
      ) +
      '<div class="chart-canvas-wrap"><canvas id="chart-vl-trend" role="img" aria-label="Trend chart"></canvas></div></div>' +
      '<div class="chart-box chart-box--vl-dist">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-vl-dist-title">Distribution — selected KPI</span>',
        "chart-vl-dist",
        "vl-dist",
        CHART_HELP.dist
      ) +
      '<div class="chart-canvas-wrap"><canvas id="chart-vl-dist" role="img" aria-label="Distribution chart"></canvas></div></div>' +
      "</div></div>" +
      '<div id="view-panel-vl-map" class="view-panel view-panel--vl-map" role="tabpanel" aria-labelledby="view-tab-vl-map" hidden>' +
      '<div class="vl-map-wrap">' +
      // map controls: resilient/vulnerable toggles and top5 lists
      '<div class="vl-map-controls">' +
      '<label class="vl-map-controls__lbl">' +
      '<input type="checkbox" id="vl-filter-resilient" checked />' +
      '<span class="vl-map-cb-swatch vl-map-cb-swatch--resilient" aria-hidden="true" title="Map marker colour: resilient"></span>' +
      "<span>Resilient sites</span></label>" +
      '<label class="vl-map-controls__lbl">' +
      '<input type="checkbox" id="vl-filter-vulnerable" checked />' +
      '<span class="vl-map-cb-swatch vl-map-cb-swatch--vulnerable" aria-hidden="true" title="Map marker colour: vulnerable"></span>' +
      "<span>Vulnerable sites</span></label>" +
      "</div>" +
      '' +
      '<div id="vl-leaflet-map" class="vl-map-host" role="application" aria-label="India map: vulnerable & resilient locations"></div>' +
      '<div class="vl-map-legend-block" id="vl-map-legend" aria-label="Map legend for vulnerable and resilient locations">' +
      '<p class="vl-map-legend-block__title">Legend — vulnerable &amp; resilient locations</p>' +
      '<ul class="vl-map-legend-block__list">' +
      '<li><span class="vl-map-legend__sw vl-map-legend__sw--res" aria-hidden="true"></span><strong>Resilient locations</strong> — green markers: top five states/UTs by combined KPI total in the current window (preview ranking).</li>' +
      '<li><span class="vl-map-legend__sw vl-map-legend__sw--vul" aria-hidden="true"></span><strong>Vulnerable locations</strong> — red/coral markers: bottom five states/UTs by the same ranking.</li>' +
      '<li><span class="vl-map-legend__sw vl-map-legend__sw--oth" aria-hidden="true"></span><strong>Other</strong> — neutral markers; use the checkboxes to emphasise only resilient or vulnerable sites.</li>' +
      "</ul></div>" +
      "</div></div>" +
      '<div id="view-panel-vl-table" class="view-panel view-panel--vl-table" role="tabpanel" aria-labelledby="view-tab-vl-table" hidden>' +
      '<div class="table-zone">' +
      '<div class="table-zone__head">' +
      '<div class="table-zone__title">' +
      '<span class="table-zone__label" id="tbl-zone-label">Business &amp; site KPIs</span>' +
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
      '<th scope="col" class="bu-matrix__th-bu">Business</th>' +
      '<th scope="col" class="bu-matrix__th-bu">Site</th>' +
      "</tr></thead>" +
      '<tbody id="tbl-body"></tbody></table>' +
      "</div></div></div>" +
      comparePanelHtml +
      "</div>" +
      '<p class="cat-context" id="cat-context">' +
      escapeHtml(cat.uxNote || "") +
      "</p>";

    root.innerHTML = "";
    root.appendChild(wrap);

    initBusinessSiteCheckboxFilters();
    applyVariableFilterFromStorage();
    initPeriodRangeInputs();

    function onFilterChange() {
      tableState.page = 0;
      refreshCategoryView(catKey);
    }
    ["f-cur-from", "f-cur-to", "f-cmp-from", "f-cmp-to"].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", onFilterChange);
    });
    wireBusinessAndSiteFilterControls(onFilterChange);
    wireVariableFilterControls(onFilterChange);
    wireToolbarScopeScrollPanels(wrap);

    document.getElementById("f-reset").addEventListener("click", () => {
      initBusinessSiteCheckboxFilters();
      ["f-cur-from", "f-cur-to", "f-cmp-from", "f-cmp-to"].forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.value = "";
      });
      initPeriodRangeInputs();
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
    // wire map filter toggles (resilient / vulnerable)
    const rChk = document.getElementById("vl-filter-resilient");
    const vChk = document.getElementById("vl-filter-vulnerable");
    if (rChk) rChk.addEventListener("change", () => refreshVulnerableLocationView(catKey));
    if (vChk) vChk.addEventListener("change", () => refreshVulnerableLocationView(catKey));
  }

  function renderCategory(catKey) {
    currentCategoryKey = catKey;
    tableState = { sortKey: "yearMonth", asc: false, page: 0 };
    destroyCharts();
    window.__adaniLeadershipMeetingsRerender = null;
    setShellLandingMode(false);

    const cat = getCategory(catKey);
    if (!cat) {
      renderCategories();
      return;
    }
    setCategoryHeaderSubtitle(
      classicIndexCategoryDisplayName(catKey, cat.categoryName || "") +
        categoryNatureBracketSuffix(catKey)
    );

    if (CATEGORY_DISABLED_NOT_IN_PREVIEW_KEYS.has(catKey)) {
      history.replaceState(null, "", "#categories");
      renderCategories();
      announce("This category is not included in this preview.");
      return;
    }

    const kpisMeta = getKpis(catKey);
    const kpisForUi = kpiListForFilterDropdown(catKey);
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

    const bizList = mergedBusinessList(
      distinctSorted(rowsForCat, (r) => r.businessName)
    );
    const wrap = document.createElement("div");
    wrap.className = insightShell ? "cat-view cat-view--modern" : "cat-view";
    const siteFieldHtml = siteFilterFieldHtml();
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

    if (Number(catKey) === VULNERABLE_LOCATION_CATEGORY_KEY) {
      const bizListVl = mergedBusinessList(
        distinctSorted(rowsForCat, (r) => r.businessName)
      );
      mountVulnerableLocationCategoryPage(
        catKey,
        cat,
        cfg,
        periodRangesFieldHtml,
        bizListVl
      );
      return;
    }

    if (!insightShell && Number(catKey) === LEADERSHIP_CATEGORY_KEY) {
      mountLeadershipMeetingCompliancePage(catKey, cat);
      return;
    }

    const personalFieldHtml =
      '<div class="field field--filter-compact field--personal-toolbar"><label class="field-label" for="f-personal">Personnel Type</label>' +
      '<select id="f-personal">' +
      '<option value="all">All</option>' +
      '<option value="Employee">Employee</option>' +
      '<option value="Contractor">Contractor</option>' +
      "</select></div>";
    const variableFieldHtml = variableFilterFieldHtml();
    const businessFieldHtml = cfg.showBusiness
      ? businessFilterFieldHtml(bizList)
      : "";
    const stateFieldHtml = cfg.showState
      ? '<div class="field"><label class="field-label" for="f-state">State</label>' +
        '<select id="f-state">' +
        stateOpts +
        "</select></div>"
      : "";
    const filterCoreInner =
      businessFieldHtml +
      variableFieldHtml +
      siteFieldHtml +
      personalFieldHtml +
      periodRangesFieldHtml +
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
      filterCoreInner +
      kpiSurfaceHtml +
      "</div>";
    const kpiChartsScopeWrap = (innerRowHtml) =>
      '<div class="kpi-charts-scope" id="kpi-charts-scope">' +
      '<div class="kpi-charts-scope__head">' +
      '<span class="kpi-charts-scope__title" id="kpi-charts-scope-title"></span>' +
      "</div>" +
      innerRowHtml +
      "</div>";
    const trendLineChartTitleInner =
      '<span class="chart-analytics-title__label" id="chart-line-title">Trend</span> <span id="chart-trend-hint" class="chart-box__hint">Monthly values · Current Period</span>';
    const trendLineChartBoxHtml =
      '<div class="chart-box chart-box--trend" id="chart-box-trend">' +
      chartBlockTitleHtml(
        trendLineChartTitleInner,
        "chart-line",
        "trend-line",
        CHART_HELP.trend
      ) +
      '<div class="chart-canvas-wrap"><canvas id="chart-line" role="img" aria-label="Line chart: monthly values for the selected KPI"></canvas></div></div>';
    const standardThreeChartsInner =
      '<div class="cat-charts cat-charts--standard" role="group" aria-label="Charts for filtered data">' +
      trendLineChartBoxHtml +
      '<div class="chart-box chart-box--biz" id="chart-box-biz">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-biz-title-label">By business</span> <span id="chart-biz-hint" class="chart-box__hint">Same KPI · every BU · latest window</span>',
        "chart-biz",
        "by-business-radar",
        CHART_HELP.bizRadar
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--biz"><canvas id="chart-biz" role="img" aria-label="Radar chart: KPI value by business unit across all preview businesses"></canvas><p id="chart-biz-empty" class="chart-biz-empty" hidden></p></div></div>' +
      '<div class="chart-box chart-box--vertical" id="chart-box-vertical">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-vertical-title-label">By vertical</span> <span id="chart-vertical-hint" class="chart-box__hint">Verticals vs KPI · latest month</span>',
        "chart-verticals",
        "by-vertical",
        CHART_HELP.vertical
      ) +
      '<div class="chart-canvas-wrap"><canvas id="chart-verticals" role="img" aria-label="Values by vertical for the selected KPI and Current Period"></canvas></div></div></div>';
    const spiLineChartTitleInner =
      '<span class="chart-analytics-title__label" id="chart-line-title">Trend</span> <span id="chart-trend-hint" class="chart-box__hint">SPI KPI lines · Current Period</span>';
    const spiLineChartBoxHtml =
      '<div class="chart-box chart-box--spi-kpi-trend chart-box--trend" id="chart-box-trend">' +
      chartBlockTitleHtml(
        spiLineChartTitleInner,
        "chart-line",
        "trend-spi",
        CHART_HELP.spiTrend
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--spi-trend"><canvas id="chart-line" role="img" aria-label="Line chart: SPI KPI trends"></canvas></div></div>';
    /** Same SPI heatmap visual; used on Leading Hazard for Hazard Spotting (primary panel). */
    const hazardLeadingHeatmapBoxHtml =
      '<div class="chart-box chart-box--spi-hazard-hm chart-box--hazard-leading-hm chart-box--trend" id="chart-box-hazard-hm">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="spi-hm-heading">Trend</span> ' +
        '<span class="chart-box__hint" id="chart-hm-hint">Week or month density</span>',
        "spi-hazard-heatmap-capture",
        "spi-hazard-heatmap",
        CHART_HELP.hazardHeat
      ) +
      '<div class="spi-hm-period-bar" id="spi-hm-period-bar">' +
      '<span class="spi-hm-period-label">Period</span>' +
      '<div class="spi-hm-period-toggle" role="group" aria-label="Trend period">' +
      '<button type="button" class="spi-hm-period-btn spi-hm-period-btn--active" data-spi-hm-period="week">Week</button>' +
      '<button type="button" class="spi-hm-period-btn" data-spi-hm-period="month">Month</button>' +
      "</div></div>" +
      '<div class="spi-hazard-heatmap-rule" id="spi-hazard-heatmap-rule" aria-hidden="true"></div>' +
      '<div id="spi-hazard-heatmap-capture" class="spi-hazard-heatmap-capture">' +
      '<div id="spi-hazard-heatmap-host" class="spi-hazard-heatmap-host"></div>' +
      '<canvas id="chart-hazard-alt" class="chart-hazard-alt" role="img" aria-label="Trend or gauge for selected Hazard KPI" hidden></canvas>' +
      '<p class="spi-hazard-heatmap-foot" id="spi-hazard-heatmap-foot" aria-live="polite"></p>' +
      "</div></div>";
    const hazardChartsRowHtml =
      '<div class="cat-charts cat-charts--hazard" role="group" aria-label="Leading hazard and observation charts: hazard heat map, BU comparison, vertical mix">' +
      hazardLeadingHeatmapBoxHtml +
      '<div class="chart-box chart-box--biz chart-box--hazard-bu" id="chart-box-biz">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-biz-title-label">By business</span> <span id="chart-biz-hint" class="chart-box__hint">Horizontal bars · selected KPI · all BUs</span>',
        "chart-biz",
        "bu-comparison",
        CHART_HELP.hazardBiz
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--biz"><canvas id="chart-biz" role="img" aria-label="Horizontal bar chart: selected KPI by business unit"></canvas><p id="chart-biz-empty" class="chart-biz-empty" hidden></p></div></div>' +
      '<div class="chart-box chart-box--hazard-vert chart-box--vertical" id="chart-box-vertical">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-vertical-title-label">By vertical</span> <span id="chart-vertical-hint" class="chart-box__hint">Doughnut · share in Current Period</span>',
        "chart-verticals",
        "vertical-mix",
        CHART_HELP.hazardVert
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--hazard-doughnut"><canvas id="chart-verticals" role="img" aria-label="Doughnut chart: share of selected KPI by vertical"></canvas></div></div></div>';
    const spiTwoChartsInner =
      '<div class="cat-charts cat-charts--spi" role="group" aria-label="Charts: SPI KPI trends and gauge or by vertical">' +
      spiLineChartBoxHtml +
      '<div class="chart-box chart-box--vertical chart-box--spi-right" id="chart-box-vertical">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-vertical-title-label">By vertical</span> <span id="chart-vertical-hint" class="chart-box__hint">Verticals · Current Period</span>',
        "chart-verticals",
        "spi-by-vertical",
        CHART_HELP.vertical
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--spi-right">' +
      '<canvas id="chart-spi-abs-gauge" class="chart-spi-abs-gauge" role="img" aria-label="SPI absolute threshold gauge" hidden></canvas>' +
      '<canvas id="chart-verticals" role="img" aria-label="Values by vertical for the selected KPI and Current Period"></canvas>' +
      '</div></div></div>';
    const consequenceTwoChartsInner =
      '<div class="cat-charts cat-charts--consequence" role="group" aria-label="Charts for Consequence Management filtered data">' +
      trendLineChartBoxHtml +
      '<div class="chart-box chart-box--vertical" id="chart-box-vertical">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-vertical-title-label">By vertical</span> <span id="chart-vertical-hint" class="chart-box__hint">Verticals vs KPI · latest month</span>',
        "chart-verticals",
        "by-vertical",
        CHART_HELP.vertical
      ) +
      '<div class="chart-canvas-wrap"><canvas id="chart-verticals" role="img" aria-label="Values by vertical for the selected KPI and Current Period"></canvas></div></div></div>';
    const chartsRowHtml =
      catKey === SPI_CATEGORY_KEY
        ? kpiChartsScopeWrap(spiTwoChartsInner)
        : catKey === HAZARD_CATEGORY_KEY
          ? kpiChartsScopeWrap(hazardChartsRowHtml)
          : Number(catKey) === CONSEQUENCE_MANAGEMENT_CATEGORY_KEY
            ? kpiChartsScopeWrap(consequenceTwoChartsInner)
            : kpiChartsScopeWrap(standardThreeChartsInner);
    const compareTabHtml =
      '<button type="button" role="tab" id="view-tab-compare" class="view-tabs__btn" aria-selected="false" aria-controls="view-panel-compare" tabindex="-1" data-view="compare">Comparison view</button>';
    const comparePanelHtml =
      '<div id="view-panel-compare" class="view-panel view-panel--compare" role="tabpanel" aria-labelledby="view-tab-compare" hidden><div class="chart-box chart-box--bu-compare">' +
      chartBlockTitleHtml(
        '<span class="chart-analytics-title__label" id="chart-bu-compare-title">BU comparison — current vs base</span> <span id="chart-bu-compare-hint" class="chart-box__hint">Grouped bars · same KPI · all BUs</span>',
        "chart-bu-compare",
        "comparison-by-bu",
        CHART_HELP.compareBu
      ) +
      '<div class="chart-canvas-wrap chart-canvas-wrap--bu-compare"><canvas id="chart-bu-compare" role="img" aria-label="Grouped bar chart: base vs current period by business unit"></canvas></div></div></div>';
    wrap.innerHTML =
      '<div class="cat-top-bar cat-top-bar--filters-only">' +
      '<fieldset id="cat-heading" tabindex="-1" class="cat-toolbar cat-toolbar--compact" aria-label="Refine results">' +
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
      '<fieldset class="kpi-summary-region">' +
      '<legend class="visually-hidden">KPI summary for current filters</legend>' +
      '<div class="multi-kpi-row" id="multi-kpi-wrap"></div>' +
      "</fieldset>" +
      '<div class="cat-main-view cat-main-view--charts" id="cat-main-view" data-view="charts">' +
      '<div class="view-tabs" role="tablist" aria-label="Trend & Distribution, table, or comparison view">' +
      '<button type="button" role="tab" id="view-tab-charts" class="view-tabs__btn view-tabs__btn--active" aria-selected="true" aria-controls="view-panel-charts" data-view="charts">Trend & Distribution View</button>' +
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
      '<span class="table-zone__label" id="tbl-zone-label">BU performance</span>' +
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
        .replace(
          '<div class="table-zone">',
          '<div class="table-zone m2-table-zone m2-evidence-zone">'
        )
        .replace('class="cat-context"', 'class="cat-context m2-cat-context"');
    }

    root.innerHTML = "";
    root.appendChild(wrap);

    if (cfg.showState) document.getElementById("f-state").value = "all";
    initBusinessSiteCheckboxFilters();
    const personalEl0 = document.getElementById("f-personal");
    if (personalEl0) personalEl0.value = "all";
    applyVariableFilterFromStorage();
    initPeriodRangeInputs();

    function onFilterChange() {
      tableState.page = 0;
      refreshCategoryView(catKey);
    }

    [
      "f-state",
      "f-personal",
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

    if (catKey === HAZARD_CATEGORY_KEY) {
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
            if (cfg.showKpi) {
              const main = document.getElementById("cat-main-view");
              if (main && main.dataset.view === "charts") {
                const boxes = document.querySelectorAll(
                  '#f-kpi-panel input[name="f-kpi-cb"]'
                );
                const checkedAfter = Array.from(boxes).filter((b) => b.checked);
                if (checkedAfter.length > 1 && cb.checked) {
                  boxes.forEach((b) => {
                    if (b !== cb) b.checked = false;
                  });
                }
                const after = Array.from(boxes).filter((b) => b.checked);
                if (after.length === 0 && boxes.length) {
                  boxes[0].checked = true;
                }
              }
            }
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
      initBusinessSiteCheckboxFilters();
      const personalEl = document.getElementById("f-personal");
      if (personalEl) personalEl.value = "all";
      ["f-cur-from", "f-cur-to", "f-cmp-from", "f-cmp-to"].forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.value = "";
      });
      initPeriodRangeInputs();
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
    setCategoryHeaderSubtitle("");
    history.replaceState(null, "", "#landing");
    setShellLandingMode(true);

    const box = document.createElement("div");
    box.className =
      "landing landing--bg-only" +
      (insightShell ? " landing--modern" : "");
    box.setAttribute("role", "region");
    box.setAttribute("aria-label", "Adani Safety Performance Dashboard home");
    box.innerHTML =
      '<div class="landing__cta">' +
      '<button type="button" class="landing__start" id="btn-landing-start">' +
      '<span class="landing__start-text">Explore Safety KPIs</span>' +
      '<svg class="landing__start-arrow" width="16" height="16" viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
      '<path fill="currentColor" d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8-8-8z"/>' +
      "</svg>" +
      "</button></div>" +
      '<div class="landing__doc-row" role="navigation" aria-label="Help and KPI reference">' +
      '<a class="landing__doc-btn" href="guide.html">How to use this dashboard</a>' +
      '<a class="landing__doc-btn landing__doc-btn--secondary" href="kpi-reference.html">KPI Catalogue</a>' +
      "</div>";

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
        ? "Insights home. Choose Explore Safety KPIs to browse dashboard categories."
        : "Adani Safety Performance Dashboard home."
    );
    updateHeaderNavState();
  }

  /** Categories list styled for Insights (`m2-cat-directory` rows). */
  function renderCategoriesInsightsDirectory() {
    currentCategoryKey = null;
    destroyCharts();
    setCategoryHeaderSubtitle("");
    history.replaceState(null, "", "#categories");
    setShellLandingMode(false);

    const box = document.createElement("div");
    box.className =
      "home-body home-body--modern m2-cat-page m2-cat-page--directory";
    box.innerHTML =
      '<div class="m2-cat-dir-head">' +
      '<div class="m2-cat-dir-intro">' +
      '<h2 id="home-h" class="categories-page__title">Dashboard Categories</h2>' +
      '<p class="m2-cat-dir-lede">Choose a category to analyze related safety KPIs and performance. All categories below open the same charts, table, comparison, and exports as the main dashboard preview—presented in the Insights layout.</p>' +
      '<div class="home-tools m2-cat-dir-search" role="search">' +
      '<label class="home-search-label" for="cat-q">Search by category or KPI</label>' +
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
      const cats = DATA.categories
        .filter((c) => categoryMatchesSearchQuery(c, q))
        .sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0));
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
        const ll = categoryListChipLeadingLagging(cat.categoryKey);
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
            ? cat.categoryName +
                ", " +
                cat.kpiCount +
                " KPIs, " +
                ll.label +
                ", open dashboard"
            : cat.categoryName +
                ", " +
                cat.kpiCount +
                " KPIs, " +
                ll.label +
                ", not in preview"
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
          ll.chipKind +
          '">' +
          ll.label +
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
    setCategoryHeaderSubtitle("");
    history.replaceState(null, "", "#categories");
    setShellLandingMode(false);

    const box = document.createElement("div");
    box.className = "home-body home-body--launchpad";
    box.innerHTML =
      '<div class="home-intro home-intro--launchpad">' +
      '<h2 id="home-h" class="categories-page__title">Dashboard Categories</h2>' +
      '<p class="home-lede">Choose a category to analyze related safety KPIs and performance</p>' +
      '<div class="home-tools" role="search">' +
      '<label class="home-search-label" for="cat-q">Search by category or KPI</label>' +
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
      const cats = DATA.categories
        .filter((c) => categoryMatchesSearchQuery(c, q))
        .sort(
          (a, b) =>
            classicIndexCategoryGridRank(a.categoryKey) -
            classicIndexCategoryGridRank(b.categoryKey)
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
        const displayName = classicIndexCategoryDisplayName(
          cat.categoryKey,
          cat.categoryName
        );
        const ll = categoryListChipLeadingLagging(cat.categoryKey);
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
            ? displayName +
                ", " +
                cat.kpiCount +
                " KPIs, " +
                ll.label
            : displayName +
                ", " +
                cat.kpiCount +
                " KPIs, " +
                ll.label +
                ". Not in preview."
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
          '<span class="category-card__lp-badge category-card__lp-badge--' +
          ll.chipKind +
          '">' +
          ll.label +
          "</span></div>" +
          '<div class="category-card__lp-body">' +
          '<span class="category-card__name">' +
          escapeHtml(displayName) +
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
