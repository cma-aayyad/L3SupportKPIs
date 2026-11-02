#!/usr/bin/env node
/* Writes test/sample_tracker.xlsx from fixture-data.js. Needs `npm install`. */
const path = require('path');
const XLSX = require('xlsx');
const { sheets } = require('./fixture-data');
const wb = XLSX.utils.book_new();
for (const [name, rows] of Object.entries(sheets)) {
  const ws = XLSX.utils.json_to_sheet(rows, { cellDates: true, dateNF: 'yyyy-mm-dd' });
  XLSX.utils.book_append_sheet(wb, ws, name);
}
const out = path.join(__dirname, 'sample_tracker.xlsx');
XLSX.writeFile(wb, out, { cellDates: true });
console.log('Wrote', out);
