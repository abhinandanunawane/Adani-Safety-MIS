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

  const meta = DATA.meta || {};

  const PAGE_SIZE = 5;
  let tableState = {
    sortKey: "yearMonth",
    asc: false,
    page: 0,
  };

  let currentCategoryKey = null;

  function announce(msg) {
    if (liveRegion) liveRegion.textContent = msg;
  }

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
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
    ["chart-line", "chart-bar", "chart-doughnut"].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.destroy();
      }
    });
  }

  const gradientColors = [
    "#0a7c86",
    "#1565c0",
    "#5b21b6",
    "#be185d",
    "#c62828",
    "#00838f",
    "#283593",
    "#6a1b9a",
  ];

  function readFilters(catKey) {
    const elKpi = document.getElementById("f-kpi");
    const elFrom = document.getElementById("f-from");
    const elTo = document.getElementById("f-to");
    const elSt = document.getElementById("f-state");
    if (!elKpi || !elFrom || !elTo || !elSt) return null;
    const elBiz = document.getElementById("f-biz");
    const elUnit = document.getElementById("f-unit");
    let mf = elFrom.value;
    let mt = elTo.value;
    if (mf > mt) {
      const t = mf;
      mf = mt;
      mt = t;
    }
    return {
      catKey,
      kpi: elKpi.value,
      monthFrom: mf,
      monthTo: mt,
      state: elSt.value,
      business: elBiz ? elBiz.value : "all",
      unitType: elUnit ? elUnit.value : "all",
    };
  }

  function applyRowFilter(rows, f) {
    return rows.filter((r) => {
      if (f.kpi !== "all" && String(r.kpiKey) !== f.kpi) return false;
      if (r.yearMonth < f.monthFrom || r.yearMonth > f.monthTo) return false;
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
    const sorted = sortKpisForDisplay(catKey, kpisMeta);
    const base = applyNonMonthFilters(getRowsForCategory(catKey), f);
    const refMonth = f.monthTo;
    const lmMonth = monthAdd(refMonth, -1);
    const lyMonth = monthAdd(refMonth, -12);

    return sorted.map((k) => {
      const kk = k.kpiKey;
      function avgFor(ym) {
        const rows = base.filter(
          (r) => r.kpiKey === kk && r.yearMonth === ym
        );
        if (!rows.length) return null;
        return avg(rows.map((r) => r.value));
      }
      const cur = avgFor(refMonth);
      const lm = avgFor(lmMonth);
      const ly = avgFor(lyMonth);
      return {
        kpiKey: kk,
        kpiName: k.kpiName,
        unitType: k.unitType,
        value: cur,
        lmPct: pctChange(cur, lm),
        lyPct: pctChange(cur, ly),
        vsLmDir: vsDir(cur, lm),
        vsLyDir: vsDir(cur, ly),
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

  function renderMultiKpiCards(container, aggregatesList) {
    container.innerHTML = "";
    const chunks = chunkArray(aggregatesList, 4);
    chunks.forEach((chunk, idx) => {
      const card = document.createElement("div");
      card.className = "multi-kpi-card";
      const head = document.createElement("div");
      head.className = "multi-kpi-card__head";
      const start = idx * 4 + 1;
      const end = idx * 4 + chunk.length;
      head.textContent = "Metrics " + start + "–" + end;
      const grid = document.createElement("div");
      grid.className = "multi-kpi-card__grid";
      chunk.forEach((item) => {
        const tile = document.createElement("div");
        tile.className = "multi-kpi-tile";
        const lmPct = formatSignedPct(item.lmPct);
        const lyPct = formatSignedPct(item.lyPct);
        tile.setAttribute(
          "aria-label",
          item.kpiName +
            ". Current " +
            formatValue(item.value, item.unitType) +
            ". Versus last month " +
            lmPct +
            ". Versus last year " +
            lyPct +
            "."
        );
        tile.title = item.kpiName;
        tile.innerHTML =
          '<div class="multi-kpi-tile__name">' +
          escapeHtml(item.kpiName) +
          "</div>" +
          '<div class="multi-kpi-tile__value-row">' +
          '<span class="multi-kpi-tile__val">' +
          formatValue(item.value, item.unitType) +
          "</span>" +
          '<span class="' +
          cmpArrowClass(item.vsLmDir) +
          '" aria-hidden="true">' +
          arrowGlyph(item.vsLmDir) +
          "</span>" +
          "</div>" +
          '<div class="multi-kpi-tile__unit">' +
          escapeHtml(item.unitType) +
          "</div>" +
          '<div class="multi-kpi-tile__cmp">' +
          '<div class="multi-kpi-tile__cmp-line">' +
          '<span class="multi-kpi-tile__cmp-tag">LM</span>' +
          '<span class="' +
          cmpArrowClass(item.vsLmDir) +
          '">' +
          arrowGlyph(item.vsLmDir) +
          "</span>" +
          '<span class="multi-kpi-tile__cmp-pct">' +
          escapeHtml(lmPct) +
          "</span>" +
          "</div>" +
          '<div class="multi-kpi-tile__cmp-line">' +
          '<span class="multi-kpi-tile__cmp-tag">LY</span>' +
          '<span class="' +
          cmpArrowClass(item.vsLyDir) +
          '">' +
          arrowGlyph(item.vsLyDir) +
          "</span>" +
          '<span class="multi-kpi-tile__cmp-pct">' +
          escapeHtml(lyPct) +
          "</span>" +
          "</div>" +
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

  function buildCharts(filteredRows) {
    destroyCharts();

    const byMonth = {};
    filteredRows.forEach((r) => {
      if (!byMonth[r.yearMonth]) byMonth[r.yearMonth] = [];
      byMonth[r.yearMonth].push(r.value);
    });
    const monthKeys = Object.keys(byMonth).sort();
    const lineLabels = monthKeys;
    const lineData = monthKeys.map((m) => avg(byMonth[m]));

    const byBiz = {};
    filteredRows.forEach((r) => {
      byBiz[r.businessName] = (byBiz[r.businessName] || 0) + r.value;
    });
    const bizLabels = Object.keys(byBiz);
    const bizData = bizLabels.map((b) => byBiz[b]);

    const unitCount = {};
    filteredRows.forEach((r) => {
      unitCount[r.unitType] = (unitCount[r.unitType] || 0) + 1;
    });
    const unitLabels = Object.keys(unitCount);
    const unitData = unitLabels.map((u) => unitCount[u]);

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
              borderColor: "#0a7c86",
              backgroundColor: "rgba(10, 124, 134, 0.1)",
              fill: true,
              tension: 0.2,
              borderWidth: 2,
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

    const elBar = document.getElementById("chart-bar");
    if (elBar && bizLabels.length) {
      new Chart(elBar, {
        type: "bar",
        data: {
          labels: bizLabels,
          datasets: [
            {
              data: bizData,
              backgroundColor: bizLabels.map(
                (_, i) => gradientColors[i % gradientColors.length]
              ),
            },
          ],
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            x: { beginAtZero: true, ticks: { font: { size: 9 } } },
            y: { ticks: { font: { size: 8 } } },
          },
        },
      });
    }

    const elD = document.getElementById("chart-doughnut");
    if (elD && unitLabels.length) {
      new Chart(elD, {
        type: "doughnut",
        data: {
          labels: unitLabels,
          datasets: [
            {
              data: unitData,
              backgroundColor: gradientColors,
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

  function renderTableBody(catKey) {
    const f = readFilters(catKey);
    if (!f) return;
    let rows = applyRowFilter(getRowsForCategory(catKey), f);
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
          "No rows match. Widen the month range, set State to All, or reset filters.";
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
        '<tr><td colspan="7" class="empty-msg">No rows for current filters.</td></tr>';
    } else {
      tbody.innerHTML = pageRows
        .map(
          (r) =>
            "<tr title=\"" +
            escapeHtml(r.kpiName + " — " + r.businessName + " — " + r.yearMonth) +
            "\">" +
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
            '<td class="col-num">' +
            formatValue(r.value, r.unitType) +
            "</td>" +
            '<td class="col-num">' +
            (r.target == null || r.target === ""
              ? "—"
              : formatValue(r.target, r.unitType)) +
            "</td>" +
            "</tr>"
        )
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
        " rows match filters. Sort with column headers; Prev and Next move pages."
    );
  }

  function refreshCategoryView(catKey) {
    const kpisMeta = getKpis(catKey);
    const f = readFilters(catKey);
    if (!f) return;

    const filtered = applyRowFilter(getRowsForCategory(catKey), f);
    const aggList = buildKpiDetailMetrics(catKey, kpisMeta, f);

    const multiWrap = document.getElementById("multi-kpi-wrap");
    if (multiWrap) {
      if (!aggList.length) {
        multiWrap.innerHTML =
          '<div class="empty-msg" style="padding:8px">No KPI data for this selection. Widen the month range or clear filters.</div>';
      } else {
        renderMultiKpiCards(multiWrap, aggList);
      }
    }

    buildCharts(filtered);
    renderTableBody(catKey);
    announceFilterSummary(catKey);
  }

  function renderCategory(catKey) {
    currentCategoryKey = catKey;
    tableState = { sortKey: "yearMonth", asc: false, page: 0 };
    destroyCharts();

    const cat = getCategory(catKey);
    if (!cat) {
      renderHome();
      return;
    }

    const kpisMeta = getKpis(catKey);
    const rowsForCat = getRowsForCategory(catKey);
    const cfg = getFilterConfig(catKey);
    const months = DATA.months.length
      ? DATA.months
      : [{ yearMonth: meta.lastDataMonth, dateKey: 0 }];
    const firstM = months[0].yearMonth;
    const lastM = months[months.length - 1].yearMonth;

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

    const monthOpts = months
      .map(
        (m) =>
          '<option value="' +
          escapeHtml(m.yearMonth) +
          '">' +
          escapeHtml(m.yearMonth) +
          "</option>"
      )
      .join("");

    const stateOpts =
      '<option value="all">All states</option>' +
      (DATA.states || [])
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
      '<div class="field"><label class="field-label" for="f-from">From</label>' +
      '<select id="f-from">' +
      monthOpts +
      "</select></div>" +
      '<div class="field"><label class="field-label" for="f-to">To</label>' +
      '<select id="f-to">' +
      monthOpts +
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
      '<div class="cat-top-bar__title">' +
      '<nav class="breadcrumb" aria-label="Breadcrumb">' +
      '<ol><li><a href="#" id="bc-home">Home</a></li><li aria-current="page">' +
      '<h2 class="cat-heading" id="cat-heading" tabindex="-1">' +
      escapeHtml(cat.categoryName) +
      "</h2></li></ol></nav>" +
      "</div>" +
      '<fieldset class="cat-toolbar cat-toolbar--compact" aria-label="Refine results">' +
      '<legend class="visually-hidden">Refine results</legend>' +
      '<div class="cat-toolbar__inner" role="group" aria-describedby="filter-hint">' +
      '<p id="filter-hint" class="visually-hidden">Narrow by KPI, month range, or state. Results update charts and the table below.</p>' +
      coreFields +
      '<div class="toolbar-actions">' +
      '<button type="button" class="btn" id="f-reset">Reset</button>' +
      "</div></div></fieldset>" +
      "</div>" +
      '<p class="cat-context" id="cat-context">' +
      escapeHtml(cat.uxNote) +
      "</p>" +
      '<fieldset class="kpi-summary-region">' +
      '<legend class="visually-hidden">KPI summary for current filters</legend>' +
      '<div class="multi-kpi-row" id="multi-kpi-wrap"></div>' +
      "</fieldset>" +
      '<div class="cat-charts" role="group" aria-label="Charts for filtered data">' +
      '<div class="chart-box"><h3>Trend</h3><div class="chart-canvas-wrap"><canvas id="chart-line" role="img" aria-label="Line chart: average value by month for filtered data"></canvas></div></div>' +
      '<div class="chart-box"><h3>By business</h3><div class="chart-canvas-wrap"><canvas id="chart-bar" role="img" aria-label="Bar chart: sum of values by business for filtered data"></canvas></div></div>' +
      '<div class="chart-box"><h3>Unit mix</h3><div class="chart-canvas-wrap"><canvas id="chart-doughnut" role="img" aria-label="Doughnut chart: row counts by unit type"></canvas></div></div>' +
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
      '<col class="col-s" /><col class="col-num" /><col class="col-num" />' +
      "</colgroup>" +
      "<thead><tr>" +
      '<th scope="col" data-sort="yearMonth" class="col-m">Month</th>' +
      '<th scope="col" data-sort="state" class="col-s">State</th>' +
      '<th scope="col" data-sort="businessName" class="col-s">Business</th>' +
      '<th scope="col" data-sort="kpiName" class="col-kpi">KPI</th>' +
      '<th scope="col" data-sort="unitType" class="col-s">Unit</th>' +
      '<th scope="col" data-sort="value" class="col-num">Value</th>' +
      '<th scope="col" data-sort="target" class="col-num">Target</th>' +
      "</tr></thead>" +
      '<tbody id="tbl-body"></tbody></table></div></div>';

    root.innerHTML = "";
    root.appendChild(wrap);

    document.getElementById("f-from").value = firstM;
    document.getElementById("f-to").value = lastM;
    if (cfg.showState) document.getElementById("f-state").value = "all";
    if (cfg.showBusiness) document.getElementById("f-biz").value = "all";
    if (cfg.showUnitType) document.getElementById("f-unit").value = "all";
    if (cfg.showKpi) document.getElementById("f-kpi").value = "all";

    document.getElementById("bc-home").addEventListener("click", (e) => {
      e.preventDefault();
      history.replaceState(
        null,
        "",
        window.location.pathname + window.location.search
      );
      renderHome();
    });

    function onFilterChange() {
      tableState.page = 0;
      refreshCategoryView(catKey);
    }

    ["f-kpi", "f-from", "f-to", "f-state", "f-biz", "f-unit"].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", onFilterChange);
    });

    document.getElementById("f-reset").addEventListener("click", () => {
      document.getElementById("f-from").value = firstM;
      document.getElementById("f-to").value = lastM;
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

  function renderHome() {
    currentCategoryKey = null;
    destroyCharts();
    history.replaceState(
      null,
      "",
      window.location.pathname + window.location.search
    );

    const box = document.createElement("div");
    box.className = "home-body";
    box.innerHTML =
      '<h2 id="home-h">Categories</h2>' +
      '<div class="home-grid" role="list" aria-labelledby="home-h"></div>';

    const grid = box.querySelector(".home-grid");
    DATA.categories.forEach((cat) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "category-card";
      btn.setAttribute("role", "listitem");
      btn.setAttribute(
        "aria-label",
        cat.categoryName + ", " + cat.kpiCount + " KPIs"
      );
      btn.innerHTML =
        '<span class="category-card__badge">' +
        cat.kpiCount +
        " KPIs</span>" +
        '<span class="category-card__name">' +
        escapeHtml(cat.categoryName) +
        "</span>" +
        '<span class="category-card__meta">' +
        escapeHtml(cat.uxNote) +
        "</span>";
      btn.addEventListener("click", () => {
        history.replaceState({}, "", "#cat=" + cat.categoryKey);
        renderCategory(cat.categoryKey);
      });
      btn.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          history.replaceState({}, "", "#cat=" + cat.categoryKey);
          renderCategory(cat.categoryKey);
        }
      });
      grid.appendChild(btn);
    });

    root.innerHTML = "";
    root.appendChild(box);
    const hh = document.getElementById("home-h");
    if (hh) hh.focus();
    else root.focus();
    announce(
      "Home. Nine safety categories. Select a card to drill KPIs, filters, and detail data."
    );
  }

  window.addEventListener("hashchange", () => {
    const m = location.hash.match(/cat=(\d+)/);
    if (m) renderCategory(parseInt(m[1], 10));
    else renderHome();
  });

  setHeaderTimestamp(meta.lastUpdateISO);

  const m0 = location.hash.match(/cat=(\d+)/);
  if (m0) renderCategory(parseInt(m0[1], 10));
  else renderHome();
})();
