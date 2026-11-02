#!/usr/bin/env node
/* Copies the pinned SheetJS and Chart.js builds from node_modules into vendor/. Run after `npm install`. */
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const files = [
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'vendor/xlsx.full.min.js'],
  ['node_modules/chart.js/dist/chart.umd.js', 'vendor/chart.umd.js'],
];
fs.mkdirSync(path.join(root, 'vendor'), { recursive: true });
for (const [from, to] of files) {
  fs.copyFileSync(path.join(root, from), path.join(root, to));
  console.log('vendored', to);
}
for (const pkg of ['xlsx', 'chart.js']) console.log(pkg, require(path.join(root, 'node_modules', pkg, 'package.json')).version);
