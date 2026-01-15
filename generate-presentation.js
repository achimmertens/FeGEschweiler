import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';
import PptxGenJS from 'pptxgenjs';

// Design-Farben basierend auf typischen Gemeindefarben (Blau/Grün-Töne)
const COLORS = {
  primary: '1E3A8A',      // Dunkelblau
  secondary: '3B82F6',     // Blau
  accent: '10B981',        // Grün
  background: 'F8FAFC',    // Hellgrau
  text: '1F2937',          // Dunkelgrau
  chartColors: [
    '3B82F6', '10B981', 'F59E0B', 'EF4444', '8B5CF6',
    'EC4899', '06B6D4', '84CC16', 'F97316', '6366F1',
    '14B8A6', 'A855F7', 'F43F5E', '0EA5E9', '22C55E'
  ]
};

/**
 * Konvertiert deutsche Zahlenformatierung (1.234,56) zu JavaScript-Zahl
 */
function parseGermanNumber(str) {
  if (!str || str.trim() === '') return 0;
  // Entferne Anführungszeichen und ersetze Komma durch Punkt, entferne Punkte (Tausender)
  const cleaned = str.toString().replace(/"/g, '').replace(/\./g, '').replace(',', '.');
  return parseFloat(cleaned) || 0;
}

/**
 * Formatiert Zahl als deutsches Format
 */
function formatGermanNumber(num) {
  return num.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

/**
 * Liest CSV-Datei und gibt Daten als Array zurück
 */
function readCSV(filePath) {
  let content = fs.readFileSync(filePath, 'utf-8');
  
  // Entferne UTF-8 BOM falls vorhanden
  if (content.charCodeAt(0) === 0xFEFF) {
    content = content.slice(1);
  }
  
  return parse(content, {
    delimiter: ';',
    columns: true,
    skip_empty_lines: true,
    relax_column_count: true,
    quote: '"',
    escape: '"',
    bom: true
  });
}

/**
 * Liest alle Gewinn-Verlust-Berichte
 */
function readProfitLossReports() {
  const dataDir = path.join(process.cwd(), 'Daten');
  const files = fs.readdirSync(dataDir);
  const reports = {};
  
  files.forEach(file => {
    const match = file.match(/gewinn-verlust-bericht_(\d{4})\.csv/);
    if (match) {
      const year = match[1];
      const filePath = path.join(dataDir, file);
      reports[year] = readCSV(filePath);
    }
  });
  
  return reports;
}

/**
 * Extrahiert Einnahmen aus einem Bericht
 */
function extractIncome(report) {
  const income = {};
  report.forEach(row => {
    if (row.Name && row.Summe) {
      const value = parseGermanNumber(row.Summe);
      if (value > 0) {
        income[row.Name] = value;
      }
    }
  });
  return income;
}

/**
 * Extrahiert Ausgaben aus einem Bericht
 */
function extractExpenses(report) {
  const expenses = {};
  report.forEach(row => {
    if (row.Name && row.Summe) {
      const value = parseGermanNumber(row.Summe);
      if (value < 0) {
        expenses[row.Name] = Math.abs(value); // Absolutwert für Ausgaben
      }
    }
  });
  return expenses;
}

/**
 * Liest Bilanzberichte
 */
function readBalanceReports() {
  const dataDir = path.join(process.cwd(), 'Daten');
  const files = fs.readdirSync(dataDir);
  const reports = {};
  
  files.forEach(file => {
    const match = file.match(/bilanzbericht_(\d{4})\.csv/);
    if (match) {
      const year = match[1];
      const filePath = path.join(dataDir, file);
      reports[year] = readCSV(filePath);
    }
  });
  
  return reports;
}

/**
 * Extrahiert Kontostände aus Bilanzbericht
 */
function extractAccountBalances(report) {
  const balances = {};
  report.forEach(row => {
    if (row.Name && row.Summe) {
      const accountName = row.Name.trim();
      const value = parseGermanNumber(row.Summe);
      const seite = row['Seite der Bilanz'];
      
      // Erfasse alle relevanten Konten
      // Aktiva-Konten (Girokonten, Sparkonten)
      if (seite === 'Aktiva') {
        balances[accountName] = value;
      }
      // Passiva-Konten (Darlehen)
      else if (seite === 'Passiva' && 
               (accountName.includes('Darlehen') || accountName.includes('Privatdarlehen'))) {
        balances[accountName] = value;
      }
    }
  });
  return balances;
}

/**
 * Aktualisiert Entwicklung.csv mit neuesten Daten
 */
function updateEntwicklungCSV() {
  const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
  const balanceReports = readBalanceReports();
  
  // Finde das neueste Jahr
  const years = Object.keys(balanceReports).sort();
  const latestYear = years[years.length - 1];
  
  if (!latestYear) {
    console.log('Keine Bilanzberichte gefunden');
    return;
  }
  
  console.log(`Aktualisiere Entwicklung.csv mit Daten von ${latestYear}`);
  
  // Lese aktuelle Entwicklung.csv
  let entwicklungContent = fs.readFileSync(entwicklungPath, 'utf-8');
  
  // Entferne UTF-8 BOM falls vorhanden
  if (entwicklungContent.charCodeAt(0) === 0xFEFF) {
    entwicklungContent = entwicklungContent.slice(1);
  }
  
  const entwicklungData = parse(entwicklungContent, {
    delimiter: ';',
    columns: true,
    skip_empty_lines: true,
    relax_column_count: true,
    quote: '"',
    escape: '"',
    bom: true
  });
  
  // Extrahiere Kontostände aus neuestem Bericht
  const latestBalances = extractAccountBalances(balanceReports[latestYear]);
  
  // Mappe Kontonamen zu Positionen in Entwicklung.csv
  const accountMapping = {
    'Girokonto SKB-Konto 701-': 'Freizeitkonto SKB 000',
    'Girokonto SKB -Konto 700-': 'Girokonto SKB 001',
    'Sparkonto SKB -Rücklagenkonto f. Heizung': 'Sparkonto SKB 003',
    'Darlehenskonto SKB': 'Darlehenskonto SKB 004',
    'Privatdarlehen': 'Privatdarlehen 006'
  };
  
  // Aktualisiere oder füge Spalte für neuestes Jahr hinzu
  const headers = Object.keys(entwicklungData[0] || {});
  if (!headers.includes(latestYear)) {
    headers.push(latestYear);
  }
  
  // Aktualisiere Daten
  entwicklungData.forEach(row => {
    const position = row.Position;
    let bilanzAccountName = null;
    
    // Finde passenden Kontonamen aus Bilanzbericht
    for (const [bilanzName, entwicklungName] of Object.entries(accountMapping)) {
      if (position === entwicklungName) {
        bilanzAccountName = bilanzName;
        break;
      }
    }
    
    if (bilanzAccountName && latestBalances[bilanzAccountName] !== undefined) {
      // Für Darlehen: negativer Wert (wie in Entwicklung.csv)
      let value = latestBalances[bilanzAccountName];
      if (position.includes('Darlehen')) {
        value = -Math.abs(value);
      }
      row[latestYear] = value.toFixed(2).replace('.', ',');
    } else if (position === 'Summe Aktiva') {
      // Berechne Summe Aktiva
      const sum = Object.entries(latestBalances)
        .filter(([name]) => name.includes('Girokonto') || name.includes('Sparkonto') || name.includes('Freizeit'))
        .reduce((sum, [, value]) => sum + value, 0);
      row[latestYear] = sum.toFixed(2).replace('.', ',');
    } else if (position === 'Verbindlichkeiten') {
      // Berechne Verbindlichkeiten
      const sum = Object.entries(latestBalances)
        .filter(([name]) => name.includes('Darlehen'))
        .reduce((sum, [, value]) => sum + Math.abs(value), 0);
      row[latestYear] = (-sum).toFixed(2).replace('.', ',');
    }
  });
  
  // Schreibe aktualisierte CSV
  const output = stringify(entwicklungData, {
    delimiter: ';',
    header: true,
    columns: headers
  });
  
  fs.writeFileSync(entwicklungPath, output, 'utf-8');
  console.log('Entwicklung.csv aktualisiert');
}

/**
 * Erstellt PowerPoint-Präsentation
 */
async function createPresentation() {
  const pres = new PptxGenJS();
  
  // Präsentationseinstellungen
  pres.layout = 'LAYOUT_WIDE';
  pres.defineLayout({ name: 'LAYOUT_WIDE', width: 10, height: 7.5 });
  
  // Lese Daten
  const profitLossReports = readProfitLossReports();
  const years = Object.keys(profitLossReports).sort();
  
  console.log(`Gefundene Jahre: ${years.join(', ')}`);
  
  // Slide 1: Einnahmen
  createIncomeSlide(pres, profitLossReports, years);
  
  // Slide 2: Ausgaben
  createExpensesSlide(pres, profitLossReports, years);
  
  // Slide 3: Kuchendiagramm Ausgaben
  createExpensesPieChartSlide(pres, profitLossReports, years);
  
  // Slide 4: Entwicklung
  createDevelopmentSlide(pres);
  
  // Speichere Präsentation
  const outputPath = path.join(process.cwd(), 'Finanzlage_FeG_Eschweiler.pptx');
  await pres.writeFile({ fileName: outputPath });
  console.log(`Präsentation erstellt: ${outputPath}`);
}

/**
 * Erstellt Slide für Einnahmen
 */
function createIncomeSlide(pres, reports, years) {
  const slide = pres.addSlide();
  
  // Titel
  slide.addText('Einnahmen', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 32,
    bold: true,
    color: COLORS.primary,
    align: 'center'
  });
  
  // Sammle alle Einnahmen-Kategorien über alle Jahre
  const allCategories = new Set();
  years.forEach(year => {
    const income = extractIncome(reports[year]);
    Object.keys(income).forEach(cat => allCategories.add(cat));
  });
  
  const categories = Array.from(allCategories).sort();
  
  // Bereite Daten für gestapeltes Balkendiagramm vor
  const chartData = [];
  categories.forEach((category, catIdx) => {
    const seriesData = years.map(year => {
      const income = extractIncome(reports[year]);
      return income[category] || 0;
    });
    chartData.push({
      name: category,
      labels: years,
      values: seriesData
    });
  });
  
  // Erstelle gestapeltes Balkendiagramm
  slide.addChart(pres.ChartType.bar, chartData, {
    x: 0.5,
    y: 1.2,
    w: 9,
    h: 5.5,
    barGrouping: 'stacked',
    catAxisTitle: 'Jahr',
    valAxisTitle: 'Betrag (EUR)',
    showLegend: true,
    legendPos: 'b',
    chartColors: COLORS.chartColors
  });
  
  // Füge Zahlenwerte als Tabelle hinzu
  const tableData = [
    ['Kategorie', ...years]
  ];
  
  categories.forEach(category => {
    const row = [category];
    years.forEach(year => {
      const income = extractIncome(reports[year]);
      const value = income[category] || 0;
      row.push(formatGermanNumber(value) + ' €');
    });
    tableData.push(row);
  });
  
  // Berechne Gesamtsummen
  const totals = ['Gesamt'];
  years.forEach(year => {
    const income = extractIncome(reports[year]);
    const total = Object.values(income).reduce((sum, val) => sum + val, 0);
    totals.push(formatGermanNumber(total) + ' €');
  });
  tableData.push(totals);
  
  slide.addTable(tableData, {
    x: 0.5,
    y: 6.8,
    w: 9,
    h: 0.7,
    fontSize: 8,
    colW: [3, 2, 2, 2],
    align: 'left',
    valign: 'middle',
    border: { type: 'solid', color: COLORS.text, pt: 1 }
  });
}

/**
 * Erstellt Slide für Ausgaben
 */
function createExpensesSlide(pres, reports, years) {
  const slide = pres.addSlide();
  
  // Titel
  slide.addText('Ausgaben', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 32,
    bold: true,
    color: COLORS.primary,
    align: 'center'
  });
  
  // Sammle alle Ausgaben-Kategorien über alle Jahre
  const allCategories = new Set();
  years.forEach(year => {
    const expenses = extractExpenses(reports[year]);
    Object.keys(expenses).forEach(cat => allCategories.add(cat));
  });
  
  const categories = Array.from(allCategories).sort();
  
  // Bereite Daten für gestapeltes Balkendiagramm vor
  const chartData = [];
  categories.forEach((category, catIdx) => {
    const seriesData = years.map(year => {
      const expenses = extractExpenses(reports[year]);
      return expenses[category] || 0;
    });
    chartData.push({
      name: category,
      labels: years,
      values: seriesData
    });
  });
  
  // Erstelle gestapeltes Balkendiagramm
  slide.addChart(pres.ChartType.bar, chartData, {
    x: 0.5,
    y: 1.2,
    w: 9,
    h: 5.5,
    barGrouping: 'stacked',
    catAxisTitle: 'Jahr',
    valAxisTitle: 'Betrag (EUR)',
    showLegend: true,
    legendPos: 'b',
    chartColors: COLORS.chartColors
  });
  
  // Füge Zahlenwerte als Tabelle hinzu
  const tableData = [
    ['Kategorie', ...years]
  ];
  
  categories.forEach(category => {
    const row = [category];
    years.forEach(year => {
      const expenses = extractExpenses(reports[year]);
      const value = expenses[category] || 0;
      row.push(formatGermanNumber(value) + ' €');
    });
    tableData.push(row);
  });
  
  // Berechne Gesamtsummen
  const totals = ['Gesamt'];
  years.forEach(year => {
    const expenses = extractExpenses(reports[year]);
    const total = Object.values(expenses).reduce((sum, val) => sum + val, 0);
    totals.push(formatGermanNumber(total) + ' €');
  });
  tableData.push(totals);
  
  slide.addTable(tableData, {
    x: 0.5,
    y: 6.8,
    w: 9,
    h: 0.7,
    fontSize: 8,
    colW: [3, 2, 2, 2],
    align: 'left',
    valign: 'middle',
    border: { type: 'solid', color: COLORS.text, pt: 1 }
  });
}

/**
 * Erstellt Slide für Kuchendiagramm der Ausgaben
 */
function createExpensesPieChartSlide(pres, reports, years) {
  const slide = pres.addSlide();
  
  // Titel
  slide.addText('Ausgaben nach Kategorien', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 32,
    bold: true,
    color: COLORS.primary,
    align: 'center'
  });
  
  // Verwende das neueste Jahr für das Kuchendiagramm
  const latestYear = years[years.length - 1];
  const expenses = extractExpenses(reports[latestYear]);
  
  // Bereite Daten für Kuchendiagramm vor
  const expenseEntries = Object.entries(expenses)
    .sort((a, b) => b[1] - a[1]); // Sortiere nach Wert
  
  const pieLabels = expenseEntries.map(([name]) => name);
  const pieValues = expenseEntries.map(([, value]) => value);
  
  const pieData = [{
    name: 'Ausgaben',
    labels: pieLabels,
    values: pieValues
  }];
  
  // Erstelle Kuchendiagramm
  slide.addChart(pres.ChartType.pie, pieData, {
    x: 1,
    y: 1.5,
    w: 4.5,
    h: 4.5,
    showLegend: true,
    legendPos: 'r',
    chartColors: COLORS.chartColors
  });
  
  // Füge Tabelle mit Details hinzu
  const tableData = [
    ['Kategorie', 'Betrag', 'Anteil']
  ];
  
  const total = Object.values(expenses).reduce((sum, val) => sum + val, 0);
  expenseEntries.forEach(([name, value]) => {
    const percentage = ((value / total) * 100).toFixed(1);
    tableData.push([
      name,
      formatGermanNumber(value) + ' €',
      percentage + ' %'
    ]);
  });
  
  slide.addTable(tableData, {
    x: 6,
    y: 1.5,
    w: 3.5,
    h: 4.5,
    fontSize: 9,
    colW: [2, 1, 0.5],
    align: 'left',
    valign: 'middle',
    border: { type: 'solid', color: COLORS.text, pt: 1 }
  });
  
  // Jahr anzeigen
  slide.addText(`Daten für ${latestYear}`, {
    x: 6,
    y: 6.2,
    w: 3.5,
    h: 0.4,
    fontSize: 10,
    italic: true,
    color: COLORS.text,
    align: 'center'
  });
}

/**
 * Erstellt Slide für Entwicklung der Kontostände
 */
function createDevelopmentSlide(pres) {
  const slide = pres.addSlide();
  
  // Titel
  slide.addText('Entwicklung der Kontostände', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 32,
    bold: true,
    color: COLORS.primary,
    align: 'center'
  });
  
  // Lese Entwicklung.csv
  const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
  const entwicklungData = readCSV(entwicklungPath);
  
  // Extrahiere Jahre aus Headern
  const firstRow = entwicklungData[0];
  const years = Object.keys(firstRow)
    .filter(key => key !== 'Position' && /^\d{4}$/.test(key))
    .sort();
  
  // Wichtige Konten für die Darstellung
  const importantAccounts = [
    'Girokonto SKB 001',
    'Sparkonto SKB 003',
    'Freizeitkonto SKB 000',
    'Darlehenskonto SKB 004',
    'Privatdarlehen 006'
  ];
  
  // Bereite Daten für Liniendiagramm vor
  const chartData = [];
  importantAccounts.forEach(account => {
    const row = entwicklungData.find(r => r.Position === account);
    if (row) {
      const values = years.map(year => {
        const value = parseGermanNumber(row[year] || '0');
        // Für Darlehen: negativer Wert für bessere Darstellung
        if (account.includes('Darlehen')) {
          return Math.abs(value);
        }
        return value;
      });
      chartData.push({
        name: account,
        labels: years,
        values: values
      });
    }
  });
  
  // Erstelle Liniendiagramm
  slide.addChart(pres.ChartType.line, chartData, {
    x: 0.5,
    y: 1.2,
    w: 9,
    h: 3.5,
    catAxisTitle: 'Jahr',
    valAxisTitle: 'Betrag (EUR)',
    showLegend: true,
    legendPos: 'b',
    chartColors: COLORS.chartColors
  });
  
  // Erstelle Tabelle mit Kontoständen
  const tableData = [
    ['Konto', ...years]
  ];
  
  importantAccounts.forEach(account => {
    const row = entwicklungData.find(r => r.Position === account);
    if (row) {
      const tableRow = [account];
      years.forEach(year => {
        const value = parseGermanNumber(row[year] || '0');
        tableRow.push(formatGermanNumber(value) + ' €');
      });
      tableData.push(tableRow);
    }
  });
  
  slide.addTable(tableData, {
    x: 0.5,
    y: 4.8,
    w: 9,
    h: 2.5,
    fontSize: 9,
    colW: [2.5, ...years.map(() => 0.8)],
    align: 'left',
    valign: 'middle',
    border: { type: 'solid', color: COLORS.text, pt: 1 }
  });
}

// Hauptfunktion
async function main() {
  try {
    console.log('Starte Generierung der Präsentation...');
    
    // Aktualisiere Entwicklung.csv
    updateEntwicklungCSV();
    
    // Erstelle PowerPoint-Präsentation
    await createPresentation();
    
    console.log('Fertig!');
  } catch (error) {
    console.error('Fehler:', error);
    process.exit(1);
  }
}

main();
