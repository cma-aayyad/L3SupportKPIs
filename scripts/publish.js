#!/usr/bin/env node
/* Publish a monthly tracker into data/.
 *   node scripts/publish.js "L2_L3_Tracker - October 2026.xlsx" [--period 2026-10]
 * Writes data/history/<yyyy-mm>.js, updates data/history/index.js, and replaces data/latest.js
 * when this month is the newest one. Period: --period, else the file name, else the max TaskList Date. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const C = require('../core.js');

const args = process.argv.slice(2);
const pIdx = args.indexOf('--period');
const forced = pIdx >= 0 ? args.splice(pIdx, 2)[1] : null;
const file = args[0];
if (!file) { console.error('Usage: node scripts/publish.js "<tracker>.xlsx" [--period yyyy-mm]'); process.exit(1); }
let XLSX;
try { XLSX = require('xlsx'); } catch (e) { console.error('SheetJS is missing: run `npm install` first.'); process.exit(1); }

const root = path.join(__dirname, '..');
const dataDir = path.join(root, 'data'), histDir = path.join(dataDir, 'history');
fs.mkdirSync(histDir, { recursive: true });

function readWindowVar(f, name) {
  if (!fs.existsSync(f)) return undefined;
  const sandbox = { window: {} };
  vm.runInNewContext(fs.readFileSync(f, 'utf8'), sandbox, { filename: f });
  return sandbox.window[name];
}

const fileName = path.basename(file);
const sheets = C.sheetsFromWorkbook(XLSX.read(fs.readFileSync(file), C.READ_OPTS), XLSX);
const ds = C.normalize(sheets);
if (!ds.meta.tasksSheet) { console.error('No task sheet (a sheet with a TASK column) in', fileName); process.exit(1); }
if (forced && !/^\d{4}-(0[1-9]|1[0-2])$/.test(forced)) { console.error('--period must look like 2026-10'); process.exit(1); }
const period = forced || C.periodFromFileName(fileName) || C.periodFromTasks(ds.tasks);
if (!period) { console.error('Could not detect the month. Name the file "... <Month> <Year>.xlsx", fill the Date column, or pass --period.'); process.exit(1); }

const snap = { fileName, period, publishedAt: new Date().toISOString(), sheets };
const index = C.mergeIndex(readWindowVar(path.join(histDir, 'index.js'), 'TRACKER_HISTORY_INDEX') || [],
  { period, fileName, publishedAt: snap.publishedAt, tasks: ds.tasks.length });
const latest = readWindowVar(path.join(dataDir, 'latest.js'), 'TRACKER_SNAPSHOT');
const isLatest = !latest || !latest.period || period >= latest.period;

const written = [];
const write = (f, txt) => { fs.writeFileSync(f, txt); written.push(path.relative(root, f)); };
write(path.join(histDir, `${period}.js`), C.historyJs(snap));
write(path.join(histDir, 'index.js'), C.historyIndexJs(index));
if (isLatest) write(path.join(dataDir, 'latest.js'), C.snapshotJs(snap));

const q = ds.quality.checks.filter(c => c.idx.length).map(c => `  ${String(c.idx.length).padStart(4)}  ${c.label}`);
console.log(`Published ${fileName} as ${C.periodLabel(period)} (${ds.tasks.length} tasks).`);
console.log('Wrote:\n' + written.map(w => '  ' + w).join('\n'));
if (!isLatest) console.log(`data/latest.js kept: ${C.periodLabel(latest.period)} is newer.`);
if (q.length) console.log('Data quality:\n' + q.join('\n'));
if (!process.env.GITHUB_ACTIONS) {
  console.log('\nTo share it:\n  git add data\n  git commit -m "Publish ' + period + '"\n  git push');
}
