import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';

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
  (report || []).forEach(row => { if (row.Name && row.Summe) { const v = parseGermanNumber(row.Summe); if (v > 0) income[row.Name] = v; } });
  return income;
}

export function extractExpenses(report) {
  const expenses = {};
  (report || []).forEach(row => { if (row.Name && row.Summe) { const v = parseGermanNumber(row.Summe); if (v < 0) expenses[row.Name] = Math.abs(v); } });
  return expenses;
}

export function extractAccountBalances(report) {
  const balances = {};
  (report || []).forEach(row => { if (row.Name && row.Summe) { const accountName = String(row.Name).trim(); const value = parseGermanNumber(row.Summe); const seite = row['Seite der Bilanz']; if (seite === 'Aktiva') balances[accountName] = value; else if (seite === 'Passiva' && (accountName.includes('Darlehen') || accountName.includes('Privatdarlehen'))) balances[accountName] = value; } });
  return balances;
}



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
  entwicklungData.forEach(row => { const position = row.Position; let bilanzAccountName = null; for (const [bilanzName, entwicklungName] of Object.entries(accountMapping)) { if (position === entwicklungName) { bilanzAccountName = bilanzName; break; } } if (bilanzAccountName && latestBalances[bilanzAccountName] !== undefined) { let value = latestBalances[bilanzAccountName]; if (position.includes('Darlehen') || position.includes('Privatdarlehen')) value = -Math.abs(value); row[latestYear] = value.toFixed(2).replace('.', ','); } else if (position === 'Summe Aktiva') { const sum = Object.entries(latestBalances).filter(([name]) => name.includes('Girokonto') || name.includes('Sparkonto') || name.includes('Freizeit')).reduce((s, [, v]) => s + v, 0); row[latestYear] = sum.toFixed(2).replace('.', ','); } else if (position === 'Verbindlichkeiten') { const sum = Object.entries(latestBalances).filter(([name]) => name.includes('Darlehen')).reduce((s, [, v]) => s + Math.abs(v), 0); row[latestYear] = (-sum).toFixed(2).replace('.', ','); } });
  const output = stringify(entwicklungData, { delimiter: ';', header: true, columns: headers }); fs.writeFileSync(entwicklungPath, output, 'utf-8'); console.log('Entwicklung.csv aktualisiert');
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
