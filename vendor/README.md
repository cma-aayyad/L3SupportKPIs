Vendored libraries live here so the page works without CDN access.

    npm install
    npm run vendor

This copies `xlsx.full.min.js` (SheetJS 0.20.3) and `chart.umd.js` (Chart.js 4.4.1). Commit both files.
Until they exist, index.html falls back to the same pinned versions on cdn.sheetjs.com and jsDelivr
(set `LIBS.allowCdnFallback = false` in index.html to forbid that).
