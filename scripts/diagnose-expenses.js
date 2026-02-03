import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';

function parseGermanNumber(str) {
  if (str === undefined || str === null) return 0;
  const s = String(str).trim();
  if (s === '') return 0;
  return parseFloat(s.replace(/"/g, '').replace(/\./g, '').replace(',', '.')) || 0;
}

const year = '2025';
const csvPath = path.join(process.cwd(), 'Daten', `gewinn-verlust-bericht_${year}.csv`);
const jsonPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${year}.json`);

if (!fs.existsSync(csvPath)) { console.error('CSV not found:', csvPath); process.exit(1); }
const csv = fs.readFileSync(csvPath, 'utf8');
const rows = parse(csv, { delimiter: ';', columns: true, skip_empty_lines: true, relax_column_count: true, quote: '"', escape: '"', bom: true });

const perCategory = {};
rows.forEach(r => {
  const name = (r.Name || r['Name'] || '').toString().trim();
  if (!name) return;
  const v = parseGermanNumber(r.Summe || r['Summe'] || 0);
  perCategory[name] = (perCategory[name] || 0) + v;
});

let json = {};
if (fs.existsSync(jsonPath)) {
  json = JSON.parse(fs.readFileSync(jsonPath, 'utf8'))[year] || {};
}

console.log(`Categories from CSV (${year}):`);
const entries = Object.entries(perCategory).sort((a,b)=>Math.abs(b[1])-Math.abs(a[1]));
entries.forEach(([k,v]) => {
  const inJson = Object.prototype.hasOwnProperty.call(json, k) ? json[k] : null;
  console.log(`${k.padEnd(40)} ${v.toFixed(2).padStart(12)}    inJson: ${inJson === null ? 'NO' : inJson}`);
});

const csvTotal = entries.reduce((s,[,v])=>s + Math.abs(v||0),0);
console.log('\nCSV authoritative total (abs) =', csvTotal);
const jsonAbsTotal = Object.values(json).reduce((s,v)=>s + Math.abs(v||0),0);
console.log('JSON abs total =', jsonAbsTotal);

console.log('\nEntries present in CSV but missing in JSON:');
entries.forEach(([k,v])=>{ if (!Object.prototype.hasOwnProperty.call(json, k)) console.log(`${k.padEnd(40)} ${Math.abs(v).toFixed(2)}`); });
