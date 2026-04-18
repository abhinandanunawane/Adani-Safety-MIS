$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$frag = [System.IO.File]::ReadAllText((Join-Path $root '_kpi-tables-fragment.html'), [System.Text.UTF8Encoding]::new($false))

$cssVer = '20260415a'

$kpiRefHtml = @"
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>KPI details &amp; measures &mdash; Adani Safety Performance Dashboard</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta
      name="description"
      content="Safety KPI catalogue: all categories, units, business definitions, and preview-model DAX-style formulas aligned to Dim_KPI."
    />
    <meta name="theme-color" content="#006db6" />
    <meta name="color-scheme" content="light" />
    <style>
      html {
        font-family: "Segoe UI", "Segoe UI Variable", system-ui, sans-serif;
      }
      .top-header {
        display: flex !important;
        flex-shrink: 0;
        align-items: stretch;
        justify-content: space-between;
        min-height: 58px;
        height: 58px;
        padding: 0 14px 0 0;
        background: #ffffff;
        color: #231f20;
        border-bottom: 3px solid #8e278f;
        box-sizing: border-box;
      }
      .top-header__logo-wrap {
        display: flex;
        align-items: center;
        justify-content: flex-start;
        flex-shrink: 0;
        height: 100%;
        padding: 0 12px 0 10px;
        box-sizing: border-box;
      }
      .top-header__titles {
        flex: 1;
        min-width: 0;
        display: flex;
        flex-direction: column;
        justify-content: center;
        padding: 0 10px 0 14px;
      }
      .top-header__title-line {
        margin: 0;
        font-size: calc(0.8125rem * 1.1);
        line-height: 1.25;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        color: #231f20;
      }
      .top-header__brand {
        font-weight: 700;
        color: #8e278f;
      }
      .top-header__logo-wrap img {
        height: auto;
        width: auto;
        max-height: max-content;
        flex-shrink: 0;
        display: block;
        object-fit: contain;
      }
      .top-header__meta--trailing {
        margin-left: auto;
        display: flex;
        align-items: center;
        padding-right: 4px;
        color: #6d6e71;
      }
    </style>
    <link rel="stylesheet" href="styles.css?v=$cssVer" />
  </head>
  <body class="shell--docs-body">
    <a class="skip-link" href="#doc-root">Skip to main content</a>

    <div class="shell shell--docs" data-app-shell>
      <header class="top-header" role="banner">
        <div class="top-header__logo-wrap">
          <img
            src="./assets/logo11.png"
            width="180"
            height="45"
            alt="Adani safetyverse"
            decoding="async"
          />
        </div>
        <div class="top-header__titles">
          <p class="top-header__title-line">
            <strong class="top-header__brand">Adani Safety Performance Dashboard</strong>
            <span class="header-sep" aria-hidden="true"> &middot; </span>
            <span class="top-header__desc">KPI details &amp; measures</span>
          </p>
        </div>
        <div class="top-header__meta top-header__meta--trailing top-header__meta--docs">
          <nav class="top-header__nav top-header__nav--static" aria-label="Go to dashboard">
            <a class="top-nav-link top-nav-link--primary" href="index.html#landing">Home</a>
            <a class="top-nav-link top-nav-link--primary" href="index.html#categories">Dashboard Categories</a>
          </nav>
          <nav class="top-header__docs" aria-label="Help and KPI reference">
            <a class="top-nav-link top-nav-link--doc top-nav-link--secondary" href="guide.html">How to use</a>
            <a class="top-nav-link top-nav-link--doc top-nav-link--secondary" href="kpi-reference.html" aria-current="page">KPI details</a>
          </nav>
          <span class="data-refreshed">Data refreshed: Daily</span>
          <span class="last-updated">Last updated <time datetime="2026-04-14">Tuesday, April 14, 2026</time></span>
        </div>
      </header>

      <main id="doc-root" class="viewport__main viewport__main--docs" role="main" tabindex="-1">
        <div class="doc-page doc-page--shell">
          <header class="doc-page__header">
            <h1>KPI details &amp; measures</h1>
            <p class="doc-page__back">
              <a href="index.html#landing">Dashboard home</a>
              <span class="header-sep" aria-hidden="true">&middot;</span>
              <a href="guide.html">How to use</a>
            </p>
            <p>
              Full wireframe catalogue from <code>PowerBI_Assets/Dim_KPI.csv</code> and
              <code>Dim_Category.csv</code>. Each row lists
              <strong>Category</strong>, <strong>KPI</strong>, <strong>Unit</strong>, <strong>Business definition</strong>, and
              <strong>Formula (preview model)</strong> &mdash; aggregate <code>Fact[Value]</code> for the listed
              <code>Dim_KPI[KPIKey]</code> in the active filter context. Extra sections document synthetic keys
              and map roll-ups used only in this HTML preview.
            </p>
          </header>

          <div class="doc-kpi-toolbar" role="search">
            <label class="doc-kpi-search">
              <span class="visually-hidden">Search KPIs by category, name, unit, or formula</span>
              <input
                type="search"
                id="kpi-doc-search"
                class="doc-kpi-search__input"
                placeholder="Search KPIs by category, name, or formula..."
                autocomplete="off"
                spellcheck="false"
              />
            </label>
            <p class="doc-kpi-toolbar__hint" id="kpi-doc-search-status" aria-live="polite"></p>
          </div>

$frag

          <section class="doc-kpi-cat doc-kpi-cat--source" id="source-files">
            <h2>Source files</h2>
            <p>
              Authoritative names and category mapping:
              <code>PowerBI_Assets/Dim_KPI.csv</code>,
              <code>PowerBI_Assets/Dim_Category.csv</code>, and
              <code>PowerBI_Assets/Master_KPI_Measures_Link.csv</code>. Align DAX with your enterprise semantic
              model and row-level security.
            </p>
            <p>
              Regenerate KPI tables by running
              <code>dashboard-preview/_generate-kpi-reference-fragment.ps1</code> after changing the CSVs.
            </p>
          </section>
        </div>
      </main>

      <footer class="app-footer" role="contentinfo" aria-label="Legal and data disclaimer">
        <p class="app-footer__text">
          <strong>Disclaimer:</strong> Sample data for user research, information architecture, usability testing, accessibility review, iterative UX validation, and UCD or HCI/CX design exploration &mdash; not operational reporting. Confirm all figures in your governed Power BI workspace.
        </p>
      </footer>
    </div>
    <script src="kpi-reference-search.js?v=20260415a" defer></script>
  </body>
</html>
"@

[System.IO.File]::WriteAllText((Join-Path $root 'kpi-reference.html'), $kpiRefHtml, [System.Text.UTF8Encoding]::new($false))
Write-Host 'Wrote kpi-reference.html'

$guideBody = [System.IO.File]::ReadAllText((Join-Path $root '_guide-body-inner.html'), [System.Text.UTF8Encoding]::new($false))

$guideHtml = @"
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>How to use this dashboard &mdash; Adani Safety Performance Dashboard</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta
      name="description"
      content="User guide: navigation, filters, Versus modes, reading charts, exporting PDF, JPEG, and CSV."
    />
    <meta name="theme-color" content="#006db6" />
    <meta name="color-scheme" content="light" />
    <style>
      html {
        font-family: "Segoe UI", "Segoe UI Variable", system-ui, sans-serif;
      }
      .top-header {
        display: flex !important;
        flex-shrink: 0;
        align-items: stretch;
        justify-content: space-between;
        min-height: 58px;
        height: 58px;
        padding: 0 14px 0 0;
        background: #ffffff;
        color: #231f20;
        border-bottom: 3px solid #8e278f;
        box-sizing: border-box;
      }
      .top-header__logo-wrap {
        display: flex;
        align-items: center;
        justify-content: flex-start;
        flex-shrink: 0;
        height: 100%;
        padding: 0 12px 0 10px;
        box-sizing: border-box;
      }
      .top-header__titles {
        flex: 1;
        min-width: 0;
        display: flex;
        flex-direction: column;
        justify-content: center;
        padding: 0 10px 0 14px;
      }
      .top-header__title-line {
        margin: 0;
        font-size: calc(0.8125rem * 1.1);
        line-height: 1.25;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        color: #231f20;
      }
      .top-header__brand {
        font-weight: 700;
        color: #8e278f;
      }
      .top-header__logo-wrap img {
        height: auto;
        width: auto;
        max-height: max-content;
        flex-shrink: 0;
        display: block;
        object-fit: contain;
      }
      .top-header__meta--trailing {
        margin-left: auto;
        display: flex;
        align-items: center;
        padding-right: 4px;
        color: #6d6e71;
      }
    </style>
    <link rel="stylesheet" href="styles.css?v=$cssVer" />
  </head>
  <body class="shell--docs-body">
    <a class="skip-link" href="#doc-root">Skip to main content</a>

    <div class="shell shell--docs" data-app-shell>
      <header class="top-header" role="banner">
        <div class="top-header__logo-wrap">
          <img
            src="./assets/logo11.png"
            width="180"
            height="45"
            alt="Adani safetyverse"
            decoding="async"
          />
        </div>
        <div class="top-header__titles">
          <p class="top-header__title-line">
            <strong class="top-header__brand">Adani Safety Performance Dashboard</strong>
            <span class="header-sep" aria-hidden="true"> &middot; </span>
            <span class="top-header__desc">How to use this dashboard</span>
          </p>
        </div>
        <div class="top-header__meta top-header__meta--trailing top-header__meta--docs">
          <nav class="top-header__nav top-header__nav--static" aria-label="Go to dashboard">
            <a class="top-nav-link top-nav-link--primary" href="index.html#landing">Home</a>
            <a class="top-nav-link top-nav-link--primary" href="index.html#categories">Dashboard Categories</a>
          </nav>
          <nav class="top-header__docs" aria-label="Help and KPI reference">
            <a class="top-nav-link top-nav-link--doc top-nav-link--secondary" href="guide.html" aria-current="page">How to use</a>
            <a class="top-nav-link top-nav-link--doc top-nav-link--secondary" href="kpi-reference.html">KPI details</a>
          </nav>
          <span class="data-refreshed">Data refreshed: Daily</span>
          <span class="last-updated">Last updated <time datetime="2026-04-14">Tuesday, April 14, 2026</time></span>
        </div>
      </header>

      <main id="doc-root" class="viewport__main viewport__main--docs" role="main" tabindex="-1">
        <div class="doc-page doc-page--shell">
$guideBody
        </div>
      </main>

      <footer class="app-footer" role="contentinfo" aria-label="Legal and data disclaimer">
        <p class="app-footer__text">
          <strong>Disclaimer:</strong> Sample data for user research, information architecture, usability testing, accessibility review, iterative UX validation, and UCD or HCI/CX design exploration &mdash; not operational reporting. Confirm all figures in your governed Power BI workspace.
        </p>
      </footer>
    </div>
  </body>
</html>
"@

[System.IO.File]::WriteAllText((Join-Path $root 'guide.html'), $guideHtml, [System.Text.UTF8Encoding]::new($false))
Write-Host 'Wrote guide.html'
