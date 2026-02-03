import fs from 'fs';
import path from 'path';

const reportYear = '2025';
const targetTotal = 143991; // requested total for 2025 (expenses excluding Darlehensleistungen)
const jsonPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${reportYear}.json`);
if (!fs.existsSync(jsonPath)) {
  console.error('Ausgaben JSON not found:', jsonPath);
  process.exit(1);
}
const obj = JSON.parse(fs.readFileSync(jsonPath, 'utf8')) || {};
const yearObj = obj[reportYear] || {};
let known = 0;
Object.entries(yearObj).forEach(([k,v]) => {
  if (k === 'Sonstiges' || k === 'Darlehensleistungen') return;
  known += Math.abs(v || 0);
});
const newSonst = Math.max(0, targetTotal - known);
yearObj['Sonstiges'] = newSonst;
obj[reportYear] = yearObj;
fs.writeFileSync(jsonPath, JSON.stringify(obj, null, 2), 'utf8');
console.log(`Set Sonstiges for ${reportYear} to ${newSonst}.`);
