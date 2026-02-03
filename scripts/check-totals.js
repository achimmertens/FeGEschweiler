import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';

function parseGermanNumber(str) {
  if (str === undefined || str === null) return 0;
  const s = String(str).trim();
  if (s === '') return 0;
  return parseFloat(s.replace(/"/g, '').replace(/\./g, '').replace(',', '.')) || 0;
}

const reportYear = '2025';
const jsonPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${reportYear}.json`);
const csvPath = path.join(process.cwd(), 'Daten', `gewinn-verlust-bericht_${reportYear}.csv`);

const json = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
const csv = fs.readFileSync(csvPath, 'utf8');
const rows = parse(csv, { delimiter: ';', columns: true, skip_empty_lines: true, relax_column_count: true, quote: '"', escape: '"', bom: true });

const jsonSum = Object.values(json[reportYear]).reduce((s,v)=>s + (typeof v==='number'? Math.abs(v): (parseGermanNumber(v)||0)),0);
console.log('Sum from Ausgaben JSON (abs):', jsonSum);

let authoritative = 0;
rows.forEach(r => {
  const kk = (r.Kontoklasse || r['Kontoklasse'] || '').toString().toLowerCase();
  const v = parseGermanNumber(r.Summe || r['Summe']);
  if (v < 0 || kk.includes('aufwand') || kk.includes('personal')) {
    authoritative += Math.abs(v || 0);
  }
});
console.log('Authoritative sum from CSV (abs):', authoritative);

// list categories in json
console.log('Categories in JSON:', Object.keys(json[reportYear]));

// compute known sum excluding Sonstiges and Darlehensleistungen
const knownSum = Object.entries(json[reportYear]).reduce((s,[k,v])=> {
  if (k==='Sonstiges' || k==='Darlehensleistungen') return s; return s + Math.abs(v||0);
},0);
console.log('Known sum (excl Sonstiges & Darlehens):', knownSum);

console.log('Suggested Sonstiges to reconcile:', Math.max(0, authoritative - knownSum));
