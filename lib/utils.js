import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';

// lightweight CSV stringify to avoid dependency issues with csv-stringify/sync
export function stringifyCSV(rows, opts = {}) {
  const delimiter = opts.delimiter || ';';
  const header = opts.header !== undefined ? opts.header : true;
  const columns = opts.columns || (rows && rows.length ? Object.keys(rows[0]) : []);
  const escapeCell = (v) => {
    if (v === undefined || v === null) return '';
    const s = String(v);
    // if cell contains delimiter, quote or newline, wrap in quotes and escape quotes
    if (s.includes(delimiter) || s.includes('"') || s.includes('\n') || s.includes('\r')) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  const lines = [];
  if (header) lines.push(columns.map(escapeCell).join(delimiter));
  (rows || []).forEach(row => {
    const vals = columns.map(c => escapeCell(row[c]));
    lines.push(vals.join(delimiter));
  });
  return lines.join('\n');
}

/**
 * Utility- und Datenfunktionen für generate-presentation
 */

// Utilities for data parsing, formatting and CSV helpers

export const COLORS = {
  primary: '1E3A8A',
  secondary: '3B82F6',
  accent: '10B981',
  background: 'F8FAFC',
  text: '1F2937',
  chartColors: [
    '3B82F6','10B981','F59E0B','EF4444','8B5CF6','EC4899','06B6D4','84CC16','F97316','6366F1','14B8A6','A855F7','F43F5E','0EA5E9','22C55E'
  ]
};

export function ensureResultDirectory() {
  const resultDir = path.join(process.cwd(), 'Daten', 'result');
  if (!fs.existsSync(resultDir)) fs.mkdirSync(resultDir, { recursive: true });
  return resultDir;
}

export function parseGermanNumber(str) {
  if (str === undefined || str === null) return 0;
  const s = String(str).trim();
  if (s === '') return 0;
  return parseFloat(s.replace(/"/g, '').replace(/\./g, '').replace(',', '.')) || 0;
}

export function formatGermanNumber(num) {
  const n = typeof num === 'number' ? num : Number(num) || 0;
  return n.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

export function formatGermanInteger(num) {
  const n = Math.round(num || 0);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

export function readCSV(filePath) {
  let content = fs.readFileSync(filePath, 'utf-8');
  if (content && content.charCodeAt(0) === 0xFEFF) content = content.slice(1);
  return parse(content, { delimiter: ';', columns: true, skip_empty_lines: true, relax_column_count: true, quote: '"', escape: '"', bom: true });
}

export function readProfitLossReports() {
  const dataDir = path.join(process.cwd(), 'Daten');
  if (!fs.existsSync(dataDir)) return {};
  const files = fs.readdirSync(dataDir);
  const reports = {};
  files.forEach(file => {
    const m = file.match(/gewinn-verlust-bericht_(\d{4})\.csv/);
    if (m) reports[m[1]] = readCSV(path.join(dataDir, file));
  });
  return reports;
}

export function readBalanceReports() {
  const dataDir = path.join(process.cwd(), 'Daten');
  if (!fs.existsSync(dataDir)) return {};
  const files = fs.readdirSync(dataDir);
  const reports = {};
  files.forEach(file => {
    const m = file.match(/bilanzbericht_(\d{4})\.csv/);
    if (m) reports[m[1]] = readCSV(path.join(dataDir, file));
  });
  return reports;
}

export function extractIncome(report) {
  const income = {};
  (report || []).forEach(row => { if (row.Name && row.Summe) { const v = parseGermanNumber(row.Summe); if (v > 0) income[row.Name] = (income[row.Name] || 0) + v; } });
  return income;
}

export function extractExpenses(report) {
  const expenses = {};
  (report || []).forEach(row => { if (row.Name && row.Summe) { const v = parseGermanNumber(row.Summe); if (v < 0) expenses[row.Name] = (expenses[row.Name] || 0) + Math.abs(v); } });
  return expenses;
}

export function extractAccountBalances(report) {
  const balances = {};
  (report || []).forEach(row => { if (row.Name && row.Summe) { const accountName = String(row.Name).trim(); const value = parseGermanNumber(row.Summe); const seite = row['Seite der Bilanz']; if (seite === 'Aktiva') balances[accountName] = value; else if (seite === 'Passiva' && (accountName.includes('Darlehen') || accountName.includes('Privatdarlehen'))) balances[accountName] = value; } });
  return balances;
}


//
export function updateEntwicklungCSV() {
  const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
  if (!fs.existsSync(entwicklungPath)) { console.log('Entwicklung.csv nicht gefunden'); return; }
  const balanceReports = readBalanceReports();
  const years = Object.keys(balanceReports).sort();
  const latestYear = years[years.length - 1];
  if (!latestYear) { console.log('Keine Bilanzberichte vorhanden'); return; }
  console.log(`Aktualisiere Entwicklung.csv mit Daten von ${latestYear}`);
  let entwicklungContent = fs.readFileSync(entwicklungPath, 'utf-8'); if (entwicklungContent.charCodeAt(0) === 0xFEFF) entwicklungContent = entwicklungContent.slice(1);
  const entwicklungData = parse(entwicklungContent, { delimiter: ';', columns: true, skip_empty_lines: true, relax_column_count: true, quote: '"', escape: '"', bom: true });
  const latestBalances = extractAccountBalances(balanceReports[latestYear] || []);
  const accountMapping = { 'Girokonto SKB-Konto 701-': 'Freizeitkonto SKB 000', 'Girokonto SKB -Konto 700-': 'Girokonto SKB 001', 'Sparkonto SKB -Rücklagenkonto f. Heizung': 'Sparkonto SKB 003', 'Darlehenskonto SKB': 'Darlehenskonto SKB 004', 'Privatdarlehen': 'Privatdarlehen 006' };
  const headers = Object.keys(entwicklungData[0] || {}); if (!headers.includes(latestYear)) headers.push(latestYear);
  entwicklungData.forEach(row => {
    const position = row.Position;
    let bilanzAccountName = null;
    for (const [bilanzName, entwicklungName] of Object.entries(accountMapping)) {
      if (position === entwicklungName) {
        bilanzAccountName = bilanzName;
        break;
      }
    }
    if (bilanzAccountName && latestBalances[bilanzAccountName] !== undefined) {
      let value = latestBalances[bilanzAccountName];
      if (position.includes('Darlehen') || position.includes('Privatdarlehen')) value = -Math.abs(value);
      row[latestYear] = value.toFixed(2).replace('.', ',');
    } else if (position === 'Summe Aktiva') {
      const sum = Object.entries(latestBalances)
        .filter(([name]) => name.includes('Girokonto') || name.includes('Sparkonto') || name.includes('Freizeit'))
        .reduce((s, [, v]) => s + v, 0);
      row[latestYear] = sum.toFixed(2).replace('.', ',');
    } else if (position === 'Verbindlichkeiten') {
      const existing = (row[latestYear] || '').toString().trim();
      if (!existing) {
        const report = balanceReports[latestYear] || [];
        const sum = report
          .filter(r => ['004','006','231'].includes((r['Kto.-Nr.'] || r['Kto.-Nr'] || r['KtoNr'] || '').toString()))
          .reduce((acc, r) => acc + Math.abs(parseGermanNumber(r.Summe || r['Summe'] || 0)), 0);
        row[latestYear] = (-sum).toFixed(2).replace('.', ',');
      }
    }
  });
  const output = stringifyCSV(entwicklungData, { delimiter: ';', header: true, columns: headers }); fs.writeFileSync(entwicklungPath, output, 'utf-8'); console.log('Entwicklung.csv aktualisiert');
}

export function computeDeltaSchuldenForYears(years) {
  const result = {};
  try {
    const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
    if (!fs.existsSync(entwicklungPath)) { years.forEach(y => result[y] = 0); return result; }
    const entwicklungData = readCSV(entwicklungPath);
    const darlehenRows = ['Darlehenskonto SKB 004', 'Privatdarlehen 006'];
    const summeSchuldenByYear = {};
    years.forEach(year => { let sum = 0; darlehenRows.forEach(acc => { const row = entwicklungData.find(r => r.Position === acc); if (row && row[year] !== undefined) sum += Math.abs(parseGermanNumber(row[year] || '0')); }); summeSchuldenByYear[year] = sum; });
    years.forEach((year, idx) => { result[year] = idx === 0 ? 0 : (summeSchuldenByYear[year] - (summeSchuldenByYear[years[idx - 1]] || 0)); });
  } catch (e) { years.forEach(y => result[y] = 0); }
  return result;
}

/**
 * Parse a Budgets_YYYY.txt file produced by the legacy tool and return structured budget data.
 * The file format is the same as used in generate-presentations-old.js (blocks of lines).
 */
export function parseBudgetFile(filePath) {
  if (!fs.existsSync(filePath)) return [];
  const content = fs.readFileSync(filePath, 'utf-8');
  const lines = content.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  // Skip header lines until we reach data (old files had ~7 header lines)
  let i = 0;
  // try to find the start of data by skipping first 7 lines if possible
  if (lines.length > 7) i = 7;
  const rows = [];
  while (i < lines.length) {
    if (i + 7 < lines.length) {
      const kostenstelle = lines[i];
      const nummer = lines[i + 1];
      const verbrauchtVorjahrRaw = lines[i + 2];
      const geplantVorjahrRaw = lines[i + 3];
      const geplantRaw = lines[i + 4];
      const stand = lines[i + 5];
      const verbrauchtRaw = lines[i + 6];
      const uebrigRaw = lines[i + 7];

      const verbrauchtVorjahr = verbrauchtVorjahrRaw && verbrauchtVorjahrRaw !== 'k. A.' ? parseGermanNumber(verbrauchtVorjahrRaw.replace(/€/g, '').trim()) : null;
      const geplantVorjahr = geplantVorjahrRaw && geplantVorjahrRaw !== 'k. A.' ? parseGermanNumber(geplantVorjahrRaw.replace(/€/g, '').trim()) : null;
      const geplant = geplantRaw && geplantRaw !== 'k. A.' ? parseGermanNumber(geplantRaw.replace(/€/g, '').trim()) : null;
      const verbraucht = verbrauchtRaw && verbrauchtRaw !== 'k. A.' ? parseGermanNumber(verbrauchtRaw.replace(/€/g, '').trim()) : null;
      const uebrig = uebrigRaw && uebrigRaw !== 'k. A.' ? parseGermanNumber(uebrigRaw.replace(/€/g, '').trim()) : null;

      rows.push({
        kostenstelle,
        nummer,
        verbrauchtVorjahr,
        verbrauchtVorjahrRaw,
        geplantVorjahr,
        geplantVorjahrRaw,
        geplant,
        geplantRaw,
        stand,
        verbraucht,
        verbrauchtRaw,
        uebrig,
        uebrigRaw
      });
      i += 8;
    } else {
      break;
    }
  }
  return rows;
}

/**
 * Save budget data (array from parseBudgetFile) to a CSV in result folder.
 * filename will be Budget_YYYY.csv if outPath not provided.
 */
export function saveBudgetAsCSV(budgetData, year, outPath) {
  const resultDir = ensureResultDirectory();
  const csvPath = outPath || path.join(resultDir, `Budget_${year}.csv`);
  // build rows
  const columns = ['Kostenstelle', 'Nummer', 'Verbraucht-Vorjahr', 'Geplant-Vorjahr', 'Geplant', 'Stand', 'Verbraucht', 'Übrig'];
  const rows = budgetData.map(item => ({
    Kostenstelle: item.kostenstelle || '',
    Nummer: item.nummer || '',
    'Verbraucht-Vorjahr': item.verbrauchtVorjahr !== null && item.verbrauchtVorjahr !== undefined ? formatGermanNumber(item.verbrauchtVorjahr) : 'k. A.',
    'Geplant-Vorjahr': item.geplantVorjahr !== null && item.geplantVorjahr !== undefined ? formatGermanNumber(item.geplantVorjahr) : 'k. A.',
    'Geplant': item.geplant !== null && item.geplant !== undefined ? formatGermanNumber(item.geplant) : 'k. A.',
    'Stand': item.stand || '',
    'Verbraucht': item.verbraucht !== null && item.verbraucht !== undefined ? formatGermanNumber(item.verbraucht) : 'k. A.',
    'Übrig': item.uebrig !== null && item.uebrig !== undefined ? formatGermanNumber(item.uebrig) : 'k. A.'
  }));
  const csv = stringifyCSV(rows, { delimiter: ';', header: true, columns });
  fs.writeFileSync(csvPath, csv, 'utf8');
  return csvPath;
}
