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
const csv = fs.readFileSync(csvPath, 'utf8');
const rows = parse(csv, { delimiter: ';', columns: true, skip_empty_lines: true, relax_column_count: true, quote: '"', escape: '"', bom: true });

const yearCat = {};
const expenseClasses = ['betr. aufwendungen', 'sonstige aufwendungen', 'personalaufwand'];
rows.forEach(r => {
  const name = (r.Name || r['Name'] || '').toString();
  const kk = (r.Kontoklasse || r['Kontoklasse'] || '').toString().toLowerCase();
  const v = parseGermanNumber(r.Summe || r['Summe'] || 0);
  if (!name) return;
  const nameLower = name.toLowerCase();
  if (nameLower.includes('zinsen für darlehen')) return; // exclude
  if (v < 0 || expenseClasses.some(ic => kk.includes(ic))) {
    yearCat[name] = (yearCat[name] || 0) + Math.abs(v || 0);
  }
});

console.log('Included categories for year', year);
Object.entries(yearCat).sort((a,b)=>b[1]-a[1]).forEach(([k,v]) => console.log(k, v));

const total = Object.values(yearCat).reduce((s,v)=>s+v,0);
console.log('Total included sum:', total.toFixed(2));
