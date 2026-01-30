import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';

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
