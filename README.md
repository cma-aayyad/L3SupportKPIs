# L2/L3 Support KPIs

A static, client-side dashboard for the monthly L2/L3 tracker. There is no backend: the page reads the published snapshot in `data/`, and any Excel you upload is processed in your browser only.

Open `index.html` from the hosted link or straight from the shared drive. Both work, because data files load as plain `<script>` tags rather than with `fetch()`.

## Views

Every view has its own link. Filters and the period are kept in the URL, for example `index.html#domain?domain=Shipping&status=Open&period=2026-10`, so you can share exactly what you see. Back and forward work.

| Link | What it's for |
|---|---|
| `#overview` (default) | Executive summary: KPI cards with change vs the previous month, domain cards, top 5 open issues by age, completion trend within the month |
| `#domain` | Domain cards, domain → application breakdown, status by domain, aging chart, task table |
| `#application` | Application summary, tasks by application and type, task table |
| `#team` | Workload per person (POC vs shadowing), aging of each person's open work, requesters and sources shown separately |
| `#trends` | Month-over-month charts across all published months: volume, completion rate, open issues, average age, volume per domain |
| `#quality` | Data-quality checks. Select one to list the affected tasks. The red badge in the header links here |
| `#print` | One-page A4 landscape layout for management. Use **Print / PDF** |

Click a domain card, a table row or a chart bar to drill down. **Export tasks CSV** and **Export KPI summary CSV** respect the current view and filters. The tasks export includes the normalized domain, the domain typed in the sheet, and the age in days.

`domain.html` still exists and redirects to `index.html#domain`, so old links keep working.

## Tracker template

File name: `L2_L3_Tracker - <Month> <Year>.xlsx`, for example `L2_L3_Tracker - October 2026.xlsx`. The month comes from the file name; if the name has no month, the latest date in TaskList is used.

**TaskList** columns are TASK, Type, Application, Date, POC, Shadowing, Status, Domain, Ticket ID. Extra columns are ignored and header case doesn't matter.

- **Date**: a real Excel date, `yyyy-mm-dd`, or `dd/mm/yyyy` (day first). Anything else is flagged, never guessed.
- **Status**: Requested, Pending, In Progress or Completed. Other wordings are mapped by the `statusBuckets` rules at the top of `core.js`. "Incomplete" is not treated as done.
- **POC, Shadowing and Ticket ID** can hold several values separated by `;`, `,` or ` / ` (slash with spaces). A bare `/` inside an ID, as in `HIP-2026/77`, is not split.
- **Domain** is only used when the application is missing from AppDomain.

**AppDomain** (Application, Domain) is the source of truth for domains. It is found by its name, or by any sheet with Application and Domain columns and no TASK column.

**Team** (optional) has two columns: Name and Aliases, with aliases separated by `;`.

| Name | Aliases |
|---|---|
| AYYAD Abdul Qadder | Abed; Abdul Qadder Ayyad |

With a Team sheet, a task counts for every team member found in POC or Shadowing, the workload chart splits POC from shadowing, and non-team POCs (requesters, PagerDuty Connector, DigiCert, …) go to "Requesters and sources". Without it, workload is POC-only, as before. Name variants such as "Carole Khattar" and "KHATTAR Carole" are merged automatically either way.

## Publishing a month (pick one)

**A. Inbox + GitHub Action (no tools needed)**
1. Put the workbook in `data/inbox/`.
2. Commit and push to `main`.
3. The **Publish tracker** Action updates `data/latest.js` and `data/history/`, runs the tests, and moves the workbook to `data/archive/`.

**B. Publish script**
1. Run `npm install` once.
2. Run `node scripts/publish.js "L2_L3_Tracker - October 2026.xlsx"`. Add `--period 2026-10` to override the detected month.
3. Run the `git add / commit / push` commands it prints.

**C. From the page**: load the file, then **Publish snapshot**. It downloads `latest.js`, `<yyyy-mm>.js` and `index.js`. Put them in `data/` and `data/history/`, then commit.

Publishing an older month only adds it to history: `latest.js` always holds the newest month.

## Local previews

**Upload tracker**, or dropping a file on the page, shows the month immediately without publishing it. The file is kept in this browser's IndexedDB, so it survives a reload. Uploading the same month again replaces it.

If your local copy of a month is newer than the published one, the page shows it with the badge **Local preview – not published**. Other people still see the published version.

**Clear local data** removes only your browser's uploads. The published snapshot is never touched.

## Development

```bash
npm install          # SheetJS 0.20.3 (from cdn.sheetjs.com) + Chart.js 4.4.1, both pinned
npm run vendor       # copies both libraries into vendor/; commit them
npm run fixture      # writes test/sample_tracker.xlsx (fake names, every edge case)
npm test             # parses the fixture with core.js and checks the KPI numbers
```

Code layout:

- **`core.js`**: config, parsing, normalization and aggregation. It has no DOM, so the page, the tests and the publish script all use it.
- **`index.html`**: persistence, routing, views, charts and export.

Settings at the top of `core.js`:

- `serviceNowBaseUrl` makes INC, RITM, CHG and SCTASK IDs clickable.
- `ageHighlightDays` (default 14) sets when the Age column turns red.
- `agingBuckets` defines the aging chart buckets.
- `issueRe` and `deploymentRe` decide which types count as issues and deployments.
- `kpiDirection` decides whether an increase in each KPI shows green or red.

Aging: for open tasks, age is the reference date minus Date. The reference date is today for the current month, and the last day of the month for past months.

The page makes no network calls except loading `vendor/*` and `data/*.js`. Until `vendor/` is populated it falls back to the same pinned library versions on a CDN; set `LIBS.allowCdnFallback = false` in `index.html` to forbid that.
