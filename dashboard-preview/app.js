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
  let catSearchAnnounceTimer = null;

  function announce(msg) {
    if (liveRegion) liveRegion.textContent = msg;
  }

  /** IA: visible journey — aligns with header nav (Guide / Categories); no duplicate controls. */
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
      '<nav class="route-steps" aria-label="Journey: Guide, Categories, Explore KPIs">' +
      item(1, "Guide", step === 1) +
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

  function destroyCharts() {
    ["chart-line", "chart-bar", "chart-doughnut"].forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof Chart !== "undefined") {
        const c = Chart.getChart(el);
        if (c) c.destroy();
      }
    });
  }

  const chartPaletteColors = [
    "#00B16B",
    "#006DB6",
    "#8E278F",
    "#F04C23",
    "#00B16B",
    "#006DB6",
    "#8E278F",
    "#F04C23",
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
              borderColor: "#006DB6",
              backgroundColor: "rgba(0, 177, 107, 0.12)",
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
                (_, i) => chartPaletteColors[i % chartPaletteColors.length]
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
              backgroundColor: chartPaletteColors,
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
      renderCategories();
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
      journeyStepsHtml(3) +
      '<div class="cat-top-bar">' +
      '<div class="cat-top-bar__title">' +
      '<nav class="breadcrumb" aria-label="Breadcrumb">' +
      '<ol><li><a href="#categories" id="bc-cats">Categories</a></li><li aria-current="page">' +
      '<h2 class="cat-heading" id="cat-heading" tabindex="-1">' +
      escapeHtml(cat.categoryName) +
      "</h2></li></ol></nav>" +
      "</div>" +
      '<fieldset class="cat-toolbar cat-toolbar--compact" aria-label="Refine results">' +
      '<legend class="visually-hidden">Refine results</legend>' +
      '<div class="cat-toolbar__inner" role="group" aria-describedby="filter-hint">' +
      '<p id="filter-hint" class="visually-hidden">Narrow by KPI, month range, state, business, or unit when shown. KPI tiles, charts, and the detail table update for your selection.</p>' +
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
      '<div class="landing__bullet" role="listitem"><strong>What’s included</strong><span>Safety KPIs by category, filters, trend and business charts, unit mix, KPI tiles (with last month / last year), and a sortable detail table.</span></div>' +
      '<div class="landing__bullet" role="listitem"><strong>Who it’s for</strong><span>Anyone tracking group-wide safety performance across Adani businesses — from quick scans to deeper drill-down.</span></div>' +
      "</div>" +
      '<div class="landing__actions">' +
      '<button type="button" class="btn btn-primary" id="btn-start">Start now</button>' +
      '<button type="button" class="btn btn-ghost" id="btn-categories">Browse categories</button>' +
      "</div>" +
      "</div>" +
      '<div class="landing__graphic" aria-hidden="true">' +
      // Inline SVG inspired by the provided reference graphic (no external asset required for GitHub Pages)
      '<svg viewBox="0 0 820 420" class="landing-graphic" xmlns="http://www.w3.org/2000/svg">' +
      '<defs>' +
      '<filter id="ds" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="6" stdDeviation="8" flood-color="#0f172a" flood-opacity="0.18"/></filter>' +
      "</defs>" +
      '<rect x="0" y="0" width="820" height="420" rx="22" fill="#ffffff"/>' +
      '<g filter="url(#ds)">' +
      '<circle cx="410" cy="210" r="62" fill="#ffffff"/>' +
      '<circle cx="410" cy="210" r="62" fill="none" stroke="#006DB6" stroke-width="10"/>' +
      '<text x="410" y="205" text-anchor="middle" font-family="Segoe UI, system-ui, sans-serif" font-size="18" font-weight="700" fill="#0f172a">Adani</text>' +
      '<text x="410" y="228" text-anchor="middle" font-family="Segoe UI, system-ui, sans-serif" font-size="12" font-weight="600" fill="#475569">Group</text>' +
      "</g>" +
      '<g stroke="#94a3b8" stroke-dasharray="4 6" stroke-width="2" fill="none" opacity="0.9">' +
      '<path d="M348 190 C260 160, 230 120, 170 110"/>' +
      '<path d="M350 235 C260 260, 230 300, 170 310"/>' +
      '<path d="M472 190 C560 160, 590 120, 650 110"/>' +
      '<path d="M470 235 C560 260, 590 300, 650 310"/>' +
      "</g>" +
      '<g filter="url(#ds)">' +
      '<g transform="translate(90 60)">' +
      '<circle cx="90" cy="50" r="54" fill="#fff"/><circle cx="90" cy="50" r="54" fill="none" stroke="#93c5fd" stroke-width="10"/>' +
      '<text x="90" y="54" text-anchor="middle" font-family="Segoe UI, system-ui, sans-serif" font-size="16" font-weight="700" fill="#0f172a">Ports</text>' +
      "</g>" +
      '<g transform="translate(90 250)">' +
      '<circle cx="90" cy="50" r="54" fill=\"#fff\"/><circle cx=\"90\" cy=\"50\" r=\"54\" fill=\"none\" stroke=\"#a7f3d0\" stroke-width=\"10\"/>' +
      '<text x=\"90\" y=\"54\" text-anchor=\"middle\" font-family=\"Segoe UI, system-ui, sans-serif\" font-size=\"16\" font-weight=\"700\" fill=\"#0f172a\">Power</text>' +
      "</g>" +
      '<g transform="translate(560 60)">' +
      '<circle cx="90" cy="50" r="54" fill="#fff"/><circle cx="90" cy="50" r="54" fill="none" stroke="#fde68a" stroke-width="10"/>' +
      '<text x="90" y="54" text-anchor="middle" font-family="Segoe UI, system-ui, sans-serif" font-size="16" font-weight="700" fill="#0f172a">Airports</text>' +
      "</g>" +
      '<g transform="translate(560 250)">' +
      '<circle cx="90" cy="50" r="54" fill="#fff"/><circle cx="90" cy="50" r="54" fill="none" stroke="#8E278F" stroke-width="10"/>' +
      '<text x="90" y="54" text-anchor="middle" font-family="Segoe UI, system-ui, sans-serif" font-size="16" font-weight="700" fill="#0f172a">Energy</text>' +
      "</g>" +
      "</g>" +
      '<g opacity="0.9">' +
      '<text x="560" y="190" font-family="Segoe UI, system-ui, sans-serif" font-size="26" font-weight="800" fill="#006DB6">All About</text>' +
      '<text x="560" y="225" font-family="Segoe UI, system-ui, sans-serif" font-size="46" font-weight="900" fill="#00B16B">Safety</text>' +
      '<text x="560" y="258" font-family="Segoe UI, system-ui, sans-serif" font-size="22" font-weight="700" fill="#0f172a">for Adani Group</text>' +
      "</g>" +
      "</svg>" +
      "</div>" +
      "</div>" +
      '<section class="landing__guide" aria-labelledby="guide-h">' +
      '<h3 id="guide-h" class="landing__guide-title">How to use this dashboard</h3>' +
      '<ol class="landing__steps">' +
      "<li><strong>Start now</strong> (or <strong>Browse categories</strong>) to open the category list.</li>" +
      "<li><strong>Select a category</strong> to open that topic and explore its KPIs.</li>" +
      "<li><strong>Refine results</strong> using filters — core filters on every page, plus extra filters on some categories where they add insight.</li>" +
      "<li><strong>Review</strong> KPI tiles (with LM/LY comparisons), then charts, then the detail table for underlying rows.</li>" +
      "</ol>" +
      '<h4 class="landing__guide-sub">What’s on each category screen</h4>' +
      '<ul class="landing__guide-list">' +
      "<li><strong>KPI summary tiles</strong> — current values with last month (LM) and last year (LY) % changes where available.</li>" +
      "<li><strong>Charts</strong> — trend over time, by business, and unit mix for the filtered data.</li>" +
      "<li><strong>Detail data</strong> — sortable table; use column headers and pagination to scan rows.</li>" +
      "</ul>" +
      '<h4 class="landing__guide-sub">Design intent: research, IA, and inclusive UX</h4>' +
      '<p class="landing__guide-lede">This preview is built with a <strong>user-centered approach</strong> and <strong>user-centered design (UCD)</strong> — drawing on <strong>human–computer interaction (HCI)</strong> and <strong>customer experience (CX) design</strong> so the product stays understandable and trustworthy. We aim for strong <strong>usability</strong>, <strong>usefulness</strong>, <strong>desirability</strong>, and <strong>accessibility</strong> for every role.</p>' +
      '<ul class="landing__guide-list landing__guide-list--compact">' +
      "<li><strong>User research &amp; usability testing</strong> — Task-based review to spot confusion, validate flows, and capture feedback.</li>" +
      "<li><strong>Information architecture (IA)</strong> — One clear path: Guide → Categories → explore KPIs in a fixed order (tiles, then charts, then detail).</li>" +
      "<li><strong>Consistency &amp; hierarchy</strong> — Same header, filters, and content zones on each category screen so scanning is predictable.</li>" +
      "<li><strong>Accessibility</strong> — Keyboard navigation (Tab, Enter, Space), visible focus, and live updates when filters change for assistive technologies.</li>" +
      "<li><strong>Iterative process</strong> — This HTML preview supports cycles of review before the full Power BI experience.</li>" +
      "</ul>" +
      "</section>";

    root.innerHTML = "";
    root.appendChild(box);
    const start = document.getElementById("btn-start");
    const browse = document.getElementById("btn-categories");
    function goCats() {
      history.replaceState(null, "", "#categories");
      renderCategories();
    }
    if (start) start.addEventListener("click", goCats);
    if (browse) browse.addEventListener("click", goCats);
    const hh = document.getElementById("landing-h");
    if (hh) hh.focus();
    announce(
      "About this dashboard. Use Start now or Browse categories to continue."
    );
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
      '<p class="home-lede">Select a category to explore safety KPIs across Adani group businesses. Use filters inside each category to refine insights. Use <strong>Guide</strong> in the header to return to the overview.</p>' +
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
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "category-card";
        btn.setAttribute("role", "listitem");
        btn.setAttribute(
          "aria-label",
          cat.categoryName + ", " + cat.kpiCount + " KPIs"
        );
        btn.innerHTML =
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
