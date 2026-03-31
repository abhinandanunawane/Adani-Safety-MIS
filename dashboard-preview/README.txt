================================================================================
ADANI SAFETY MIS — BROWSER PREVIEW
================================================================================

VIEWPORT
--------
Layout is fixed at 1280 x 720 px with no page scrolling. Detail table uses
pagination and sortable columns (no body scroll). Resize browser window to see
the centered stage; content does not scroll at the document level.

OPEN
----
1) Double-click index.html (works offline; Chart.js loads from CDN — needs internet
   for first chart load), OR
2) From project folder, run a static server so team members can open a link:
      npx --yes serve . -p 3333
   Then browse to http://localhost:3333

SYNC WITH POWER BI CSVs
------------------------
After editing files in ..\PowerBI_Assets\ (*.csv), run:

  powershell -ExecutionPolicy Bypass -File Refresh-PreviewData.ps1

This refreshes embedded-data.json and embedded-data.js (timestamp updates).

FILES
-----
index.html          — Shell, header, meta bar (last updated)
styles.css          — Layout, focus states, responsive grid
app.js              — Home → category navigation, Chart.js charts, table
embedded-data.js    — KPI aggregates (generated)
assets/adani-logo.png — Logo copy for header

FEEDBACK
--------
Use this preview for usability testing and stakeholder review; mirror approved
layouts in Power BI Desktop using the theme and layout notes in PowerBI_Assets.

================================================================================
