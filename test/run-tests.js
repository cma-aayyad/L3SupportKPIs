#!/usr/bin/env node
/* node test/run-tests.js — runs core.js parsing + aggregation on the fixture and checks the numbers.
 * Uses test/sample_tracker.xlsx through SheetJS when xlsx is installed (npm install);
 * otherwise falls back to the same rows from fixture-data.js. */
const assert = require('assert/strict');
const fs = require('fs');
const path = require('path');
const C = require('../core.js');
const fx = require('./fixture-data');

let sheets, source;
try {
  const XLSX = require('xlsx');
  const file = path.join(__dirname, 'sample_tracker.xlsx');
  if (!fs.existsSync(file)) require('./make-fixture.js');
  sheets = C.sheetsFromWorkbook(XLSX.read(fs.readFileSync(file), C.READ_OPTS), XLSX);
  source = 'sample_tracker.xlsx via SheetJS ' + XLSX.version;
} catch (e) {
  if (e.code !== 'MODULE_NOT_FOUND') throw e;
  // same shape SheetJS produces: real date cells arrive as yyyy-mm-dd strings
  sheets = Object.fromEntries(Object.entries(fx.sheets).map(([n, rows]) => [n, rows.map(r =>
    Object.fromEntries(Object.entries(r).map(([k, v]) => [k, v instanceof Date ? C.parseDate(v) : v])))]));
  source = 'fixture-data.js (xlsx not installed — run npm install to test the real .xlsx path)';
}

let passed = 0;
const t = (name, fn) => { fn(); passed++; console.log('  ✓', name); };
console.log('Source:', source);

t('splitMulti', () => {
  assert.deepEqual(C.splitMulti('INC1 / INC2; INC3,INC4'), ['INC1', 'INC2', 'INC3', 'INC4']);
  assert.deepEqual(C.splitMulti('HIP-2026/77'), ['HIP-2026/77']);
});
t('parseDate day-first, ISO, serial, invalid', () => {
  assert.equal(C.parseDate('05/09/2026'), '2026-09-05');
  assert.equal(C.parseDate('2026-09-01'), '2026-09-01');
  assert.equal(C.parseDate('31/09/2026'), null);
  assert.equal(C.parseDate('Sept 5'), null);
  assert.equal(C.parseDate(46270), '2026-09-05');
});
t('status buckets use word boundaries', () => {
  assert.equal(C.bucketOf('Incomplete'), 'Pending');
  assert.equal(C.bucketOf('Completed'), 'Done');
  assert.equal(C.bucketOf(' in progress '), 'In Progress');
  assert.equal(C.bucketOf(''), 'Unknown');
});
t('name tokens ignore order and case', () => assert.equal(C.tokenKey('KHATTAR Carole'), C.tokenKey('Carole Khattar')));
t('period from file name', () => {
  assert.equal(C.periodFromFileName('L2_L3_Tracker - October 2026.xlsx'), '2026-10');
  assert.equal(C.periodFromFileName('tracker 2026-09.xlsx'), '2026-09');
  assert.equal(C.periodFromFileName('tracker.xlsx'), null);
});
t('aging reference date', () => {
  assert.equal(C.refDateFor('2026-09', '2026-10-30'), '2026-09-30');
  assert.equal(C.refDateFor('2026-10', '2026-10-30'), '2026-10-30');
});

const ds = C.normalize(sheets);
C.applyAges(ds.tasks, C.refDateFor('2026-09', '2026-10-30'));
const a = C.aggregate(ds.tasks, ds), k = a.kpis;
t('sheets detected', () => { assert.equal(ds.meta.mappingSheet, 'AppDomain'); assert.equal(ds.meta.teamSheet, 'Team'); });
t('core KPIs', () => {
  assert.equal(k.total, 10); assert.equal(k.done, 3); assert.equal(k.open, 7);
  assert.equal(Math.round(k.completionRate), 30);
  assert.equal(k.issues, 4); assert.equal(k.openIssues, 3);
  assert.equal(k.deployments, 2); assert.equal(k.pipeline, 1);
  assert.equal(k.activeApps, 5); assert.equal(k.totalApps, 5); assert.equal(k.unmappedActive, 1);
  assert.equal(k.avgAge.toFixed(2), '21.67');
});
t('domains follow the mapping', () => {
  const d = Object.fromEntries(a.byDomain.map(x => [x.domain, x.total]));
  assert.deepEqual(d, { Shipping: 3, Communication: 3, LineOps: 1, MyEup: 1, Unassigned: 2 });
  assert.equal(a.byDomain.find(x => x.domain === 'Shipping').apps[0].app, 'ONEWAY');
});
t('aging buckets', () => assert.deepEqual(a.aging.map(x => x.count), [1, 0, 4, 1]));
t('quality checks', () => {
  const q = Object.fromEntries(ds.quality.checks.map(c => [c.id, c.idx.length]));
  assert.deepEqual(q, { noTicket: 1, unmappedApp: 1, domainMismatch: 1, unassigned: 2, badDate: 1, blankStatus: 1, nameVariants: 9, dupTickets: 2 });
});
t('team workload with POC/Shadowing split', () => {
  const m = Object.fromEntries(a.team.members.map(x => [x.name, [x.poc, x.shadow]]));
  assert.deepEqual(m['Rami HADDAD'], [3, 1]); assert.deepEqual(m['Lina SAAB'], [2, 2]); assert.deepEqual(m['Nadim KHOURY'], [2, 1]);
  assert.deepEqual(m['Unassigned'], [1, 0]);
  assert.deepEqual(a.team.requesters.map(x => x.name).sort(), ['DigiCert', 'Group Application Security', 'PagerDuty Connector']);
});
t('multi-ticket cells', () => {
  assert.deepEqual(ds.tasks[3].tickets, ['INC0012346', 'INC0012347']);
  assert.deepEqual(ds.tasks[5].tickets, ['RITM0005001', 'HIP-12345']);
});
t('filters', () => {
  assert.equal(C.applyFilters(ds.tasks, { status: 'Open', domain: 'shipping' }, ds).length, 2);
  assert.equal(C.applyFilters(ds.tasks, { qc: 'badDate' }, ds).length, 1);
});
t('POC-only fallback without Team sheet', () => {
  const d2 = C.normalize({ TaskList: sheets.TaskList, AppDomain: sheets.AppDomain }); C.applyAges(d2.tasks, '2026-09-30');
  const w = C.aggregate(d2.tasks, d2).team;
  assert.equal(w.hasTeam, false); assert.equal(w.requesters.length, 0);
  assert.equal(w.members.reduce((s, x) => s + x.shadow, 0), 0);
});
t('mapping sheet found by columns when renamed', () => {
  const d3 = C.normalize({ Tasks: sheets.TaskList, Apps: sheets.AppDomain.map(r => ({ ...r, Owner: 'x' })) });
  assert.equal(d3.meta.mappingSheet, 'Apps');
});
console.log(`\n${passed} checks passed`);
