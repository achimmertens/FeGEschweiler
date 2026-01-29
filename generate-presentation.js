import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';
import PptxGenJS from 'pptxgenjs';
import ExcelJS from 'exceljs';
import QuickChart from 'quickchart-js';

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

// Speichere sortierte Daten für PowerPoint
let sortedIncomeData = null;
let sortedExpensesData = null;
let sortedExpensesPieData = null;
let developmentData = null;

/**
 * Erstellt den Ergebnisordner falls er nicht existiert
 */
function ensureResultDirectory() {
  const resultDir = path.join(process.cwd(), 'Daten', 'result');
  if (!fs.existsSync(resultDir)) {
    fs.mkdirSync(resultDir, { recursive: true });
  }
  return resultDir;
}

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
 * Formatiert Zahl ohne Dezimalstellen mit deutschen Tausenderpunkten
 */
function formatGermanInteger(num) {
  const n = Math.round(num || 0);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
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
 * Berechnet das Delta der Summe Schulden (Darlehen) für gegebene Jahre
 * Liefert ein Objekt { year: delta }
 */
function computeDeltaSchuldenForYears(years) {
  const result = {};
  try {
    const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
    const entwicklungData = readCSV(entwicklungPath);
    const darlehenRows = ['Darlehenskonto SKB 004', 'Privatdarlehen 006'];
    const summeSchuldenByYear = {};
    years.forEach(year => {
      let sum = 0;
      darlehenRows.forEach(acc => {
        const row = entwicklungData.find(r => r.Position === acc);
        if (row && row[year] !== undefined) {
          sum += Math.abs(parseGermanNumber(row[year] || '0'));
        }
      });
      summeSchuldenByYear[year] = sum;
    });
    years.forEach((year, idx) => {
      if (idx === 0) result[year] = 0;
      else {
        const prev = years[idx - 1];
        result[year] = (summeSchuldenByYear[year] || 0) - (summeSchuldenByYear[prev] || 0);
      }
    });
  } catch (err) {
    years.forEach(year => { result[year] = 0; });
  }
  return result;
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
      if (position.includes('Darlehen') || position.includes('Privatdarlehen')) {
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
 * Erstellt Excel-Datei mit sortierten Tabellen und Grafiken
 */
async function createExcelFile(profitLossReports, years) {
  const workbook = new ExcelJS.Workbook();
  
  // Slide 1: Einnahmen
  const incomeSheet = workbook.addWorksheet('Einnahmen');
  
  // Sammle alle Einnahmen-Kategorien über alle Jahre
  const allIncomeCategories = new Set();
  years.forEach(year => {
    const income = extractIncome(profitLossReports[year]);
    Object.keys(income).forEach(cat => allIncomeCategories.add(cat));
  });
  // Berechne Gesamtsummen pro Kategorie über alle Jahre (für Sortierung)
  const incomeTotals = {};
  allIncomeCategories.forEach(category => {
    const total = years.reduce((sum, year) => {
      const income = extractIncome(profitLossReports[year]);
      return sum + (income[category] || 0);
    }, 0);
    incomeTotals[category] = total;
  });

  // Sortiere nach Gesamtsumme (absteigend)
  const sortedIncomeCategories = Array.from(allIncomeCategories)
    .sort((a, b) => incomeTotals[b] - incomeTotals[a]);

  // Erstelle Tabelle
  incomeSheet.columns = [
    { header: 'Kategorie', key: 'category', width: 40 },
    ...years.map(year => ({ header: year, key: year, width: 15 }))
  ];

  // Fülle Zeilen
  sortedIncomeCategories.forEach(category => {
    const row = { category };
    years.forEach(year => {
      const income = extractIncome(profitLossReports[year]);
      row[year] = income[category] || 0;
    });
    incomeSheet.addRow(row);
  });

  // Gesamtsumme
  const totalRow = { category: 'Gesamt' };
  years.forEach(year => {
    const income = extractIncome(profitLossReports[year]);
    totalRow[year] = Object.values(income).reduce((sum, val) => sum + val, 0);
  });
  incomeSheet.addRow(totalRow);

  // Formatiere Zahlen
  incomeSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      years.forEach((year, idx) => {
        const cell = row.getCell(idx + 2);
        if (cell.value !== null && typeof cell.value === 'number') {
          cell.numFmt = '#,##0.00';
        }
      });
    }
  });

  // Header-Formatierung
  incomeSheet.getRow(1).font = { bold: true };
  incomeSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF' + COLORS.primary }
  };
  incomeSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  // Speichere sortierte Daten für PowerPoint
  sortedIncomeData = {
    categories: sortedIncomeCategories,
    data: sortedIncomeCategories.map(category => {
      const row = {};
      years.forEach(year => {
        const income = extractIncome(profitLossReports[year]);
        row[year] = income[category] || 0;
      });
      return { category, ...row };
    }),
    totals: totalRow
  };

  // Slide 2: Ausgaben
  const expensesSheet = workbook.addWorksheet('Ausgaben');
  
  // Sammle alle Ausgaben-Kategorien über alle Jahre
  const allExpenseCategories = new Set();
  years.forEach(year => {
    const expenses = extractExpenses(profitLossReports[year]);
    Object.keys(expenses).forEach(cat => allExpenseCategories.add(cat));
  });
  
  
  
  // Berechne Gesamtsummen pro Kategorie über alle Jahre
  const expenseTotals = {};
  allExpenseCategories.forEach(category => {
    const total = years.reduce((sum, year) => {
      const expenses = extractExpenses(profitLossReports[year]);
      return sum + (expenses[category] || 0);
    }, 0);
    expenseTotals[category] = total;
  });
  
  // Sortiere nach Gesamtsumme (absteigend)
  const sortedExpenseCategories = Array.from(allExpenseCategories)
    .sort((a, b) => expenseTotals[b] - expenseTotals[a]);
  
  // Erstelle Tabelle
  expensesSheet.columns = [
    { header: 'Kategorie', key: 'category', width: 40 },
    ...years.map(year => ({ header: year, key: year, width: 15 }))
  ];
  
  sortedExpenseCategories.forEach(category => {
    const row = { category };
    years.forEach(year => {
      const expenses = extractExpenses(profitLossReports[year]);
      row[year] = expenses[category] || 0;
    });
    expensesSheet.addRow(row);
  });
  
  // Gesamtsumme
  const expenseTotalRow = { category: 'Gesamt' };
  years.forEach(year => {
    const expenses = extractExpenses(profitLossReports[year]);
    expenseTotalRow[year] = Object.values(expenses).reduce((sum, val) => sum + val, 0);
  });
  expensesSheet.addRow(expenseTotalRow);
  
  // Formatiere Zahlen
  expensesSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      years.forEach((year, idx) => {
        const cell = row.getCell(idx + 2);
        if (cell.value !== null && typeof cell.value === 'number') {
          cell.numFmt = '#,##0.00';
        }
      });
    }
  });
  
  // Header-Formatierung
  expensesSheet.getRow(1).font = { bold: true };
  expensesSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF' + COLORS.primary }
  };
  expensesSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  
  // Speichere sortierte Daten für PowerPoint
  sortedExpensesData = {
    categories: sortedExpenseCategories,
    data: sortedExpenseCategories.map(category => {
      const row = {};
      years.forEach(year => {
        const expenses = extractExpenses(profitLossReports[year]);
        row[year] = expenses[category] || 0;
      });
      return { category, ...row };
    }),
    totals: expenseTotalRow
  };
  
  // Slide 3: Kuchendiagramm Ausgaben
  const pieSheet = workbook.addWorksheet('Ausgaben Kuchendiagramm');
  
  // Verwende das neueste Jahr für das Kuchendiagramm
  const latestYear = years[years.length - 1];
  const expenses = extractExpenses(profitLossReports[latestYear]);
  
  // Sortiere nach Wert (absteigend)
  const sortedPieData = Object.entries(expenses)
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({ name, value }));
  
  // Erstelle Tabelle
  pieSheet.columns = [
    { header: 'Kategorie', key: 'category', width: 40 },
    { header: 'Betrag', key: 'amount', width: 15 },
    { header: 'Anteil %', key: 'percentage', width: 12 }
  ];
  
  const totalExpenses = Object.values(expenses).reduce((sum, val) => sum + val, 0);
  sortedPieData.forEach(({ name, value }) => {
    const percentage = (value / totalExpenses) * 100;
    pieSheet.addRow({
      category: name,
      amount: value,
      percentage: percentage
    });
  });
  
  // Formatiere Zahlen
  pieSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.getCell(2).numFmt = '#,##0.00';
      row.getCell(3).numFmt = '0.0';
    }
  });
  
  // Header-Formatierung
  pieSheet.getRow(1).font = { bold: true };
  pieSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF' + COLORS.primary }
  };
  pieSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  
  // Speichere sortierte Daten für PowerPoint
  sortedExpensesPieData = sortedPieData;
  
  // Slide 4: Entwicklung
  const developmentSheet = workbook.addWorksheet('Entwicklung');
  
  const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
  const entwicklungData = readCSV(entwicklungPath);
  
  // Extrahiere Jahre aus Headern
  const firstRow = entwicklungData[0];
  const devYears = Object.keys(firstRow)
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
  
  // Erstelle Tabelle
  developmentSheet.columns = [
    { header: 'Konto', key: 'account', width: 30 },
    ...devYears.map(year => ({ header: year, key: year, width: 15 }))
  ];
  
  // Sammle Daten für Summenberechnung
  const accountData = [];
  
  importantAccounts.forEach(account => {
    const row = entwicklungData.find(r => r.Position === account);
    if (row) {
      const tableRow = { account };
      devYears.forEach(year => {
        tableRow[year] = parseGermanNumber(row[year] || '0');
      });
      developmentSheet.addRow(tableRow);
      // Sammle nur die Original-Konten (ohne Delta-Zeilen)
      accountData.push({ account, ...tableRow });

      // Falls Darlehen-Konten: Delta zum Vorjahr direkt darunter einfügen
      if (account === 'Darlehenskonto SKB 004' || account === 'Privatdarlehen 006') {
        const deltaRow = { account: `${account} Δ` };
        devYears.forEach((year, idx) => {
          if (idx === 0) {
            deltaRow[year] = 0; // kein Vorjahr vorhanden
          } else {
            const prevYear = devYears[idx - 1];
            deltaRow[year] = tableRow[year] - (tableRow[prevYear] || 0);
          }
        });
        const inserted = developmentSheet.addRow(deltaRow);
        // Grauen Hintergrund für Delta-Zeilen
        inserted.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD9D9D9' }
          };
        });
      }
    }
  });
  
  // Berechne Summe Guthaben (Aktiva-Konten)
  const guthabenAccounts = ['Girokonto SKB 001', 'Sparkonto SKB 003', 'Freizeitkonto SKB 000'];
  const summeGuthabenRow = { account: 'Summe Guthaben' };
  devYears.forEach(year => {
    const sum = accountData
      .filter(item => guthabenAccounts.includes(item.account))
      .reduce((sum, item) => sum + (item[year] || 0), 0);
    summeGuthabenRow[year] = sum;
  });
  // Füge Summe-Zeile ein und markiere sie fett
  const sumGRow = developmentSheet.addRow(summeGuthabenRow);
  sumGRow.getCell(1).font = { bold: true };
  devYears.forEach((year, idx) => {
    sumGRow.getCell(idx + 2).font = { bold: true };
  });

  // Delta-Zeile für Summe Guthaben
  const sumGDelta = { account: 'Summe Guthaben Δ' };
  devYears.forEach((year, idx) => {
    if (idx === 0) sumGDelta[year] = 0;
    else {
      const prevYear = devYears[idx - 1];
      sumGDelta[year] = (summeGuthabenRow[year] || 0) - (summeGuthabenRow[prevYear] || 0);
    }
  });
  const sumGDeltaRow = developmentSheet.addRow(sumGDelta);
  sumGDeltaRow.eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9D9D9' }
    };
  });
  
  // Berechne Summe Schulden (Passiva-Konten)
  const schuldenAccounts = ['Darlehenskonto SKB 004', 'Privatdarlehen 006'];
  const summeSchuldenRow = { account: 'Summe Schulden' };
  devYears.forEach(year => {
    const sum = accountData
      .filter(item => schuldenAccounts.includes(item.account))
      .reduce((sum, item) => sum + Math.abs(item[year] || 0), 0);
    summeSchuldenRow[year] = sum;
  });
  // Füge Summe-Schulden-Zeile ein und markiere sie fett
  const sumSRow = developmentSheet.addRow(summeSchuldenRow);
  sumSRow.getCell(1).font = { bold: true };
  devYears.forEach((year, idx) => {
    sumSRow.getCell(idx + 2).font = { bold: true };
  });

  // Delta-Zeile für Summe Schulden
  const sumSDelta = { account: 'Summe Schulden Δ' };
  devYears.forEach((year, idx) => {
    if (idx === 0) sumSDelta[year] = 0;
    else {
      const prevYear = devYears[idx - 1];
      sumSDelta[year] = (summeSchuldenRow[year] || 0) - (summeSchuldenRow[prevYear] || 0);
    }
  });
  const sumSDeltaRow = developmentSheet.addRow(sumSDelta);
  sumSDeltaRow.eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9D9D9' }
    };
  });
  
  // Formatiere Zahlen
  developmentSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      devYears.forEach((year, idx) => {
        const cell = row.getCell(idx + 2);
        if (cell.value !== null && typeof cell.value === 'number') {
          cell.numFmt = '#,##0.00';
        }
      });
      // Fette Formatierung für Summen-Zeilen
      if (rowNumber > importantAccounts.length + 1) {
        row.getCell(1).font = { bold: true };
        devYears.forEach((year, idx) => {
          row.getCell(idx + 2).font = { bold: true };
        });
      }
    }
  });
  
  // Header-Formatierung
  developmentSheet.getRow(1).font = { bold: true };
  developmentSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF' + COLORS.primary }
  };
  developmentSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  
  // Speichere Daten für PowerPoint
  developmentData = {
    accounts: importantAccounts,
    years: devYears,
    data: importantAccounts.map(account => {
      const row = entwicklungData.find(r => r.Position === account);
      if (row) {
        const tableRow = { account };
        devYears.forEach(year => {
          tableRow[year] = parseGermanNumber(row[year] || '0');
        });
        return tableRow;
      }
      return null;
    }).filter(r => r !== null)
  };
  
  // Slide 5: Budget
  await createBudgetSheet(workbook);
  
  // Speichere Excel-Datei
  ensureResultDirectory();
  const excelPath = path.join(process.cwd(), 'Daten', 'result', 'Finanzlage_FeG_Eschweiler.xlsx');
  await workbook.xlsx.writeFile(excelPath);
  console.log(`Excel-Datei erstellt: ${excelPath}`);
}

/**
 * Erstellt gestapeltes Balkendiagramm der Einnahmen als PNG
 */
async function createIncomeChart(profitLossReports, years) {
  // Finde das neueste Jahr aus Bilanzberichten
  const balanceReports = readBalanceReports();
  const balanceYears = Object.keys(balanceReports).sort();
  const latestYear = balanceYears[balanceYears.length - 1];
  
  if (!latestYear) {
    console.log('Keine Bilanzberichte gefunden, kann Einnahmen-Diagramm nicht erstellen');
    return;
  }
  
  console.log(`Erstelle Einnahmen-Diagramm für Jahr ${latestYear}`);
  
  // Sammle alle Einnahmen-Kategorien über alle Jahre
  const allIncomeCategories = new Set();
  years.forEach(year => {
    const income = extractIncome(profitLossReports[year]);
    Object.keys(income).forEach(cat => allIncomeCategories.add(cat));
  });
  
  // Berechne Gesamtsummen pro Kategorie über alle Jahre (für Sortierung)
  const incomeTotals = {};
  allIncomeCategories.forEach(category => {
    const total = years.reduce((sum, year) => {
      const income = extractIncome(profitLossReports[year]);
      return sum + (income[category] || 0);
    }, 0);
    incomeTotals[category] = total;
  });
  
  // Sortiere nach Gesamtsumme (absteigend) - größte zuerst
  const sortedIncomeCategories = Array.from(allIncomeCategories)
    .sort((a, b) => incomeTotals[b] - incomeTotals[a]);
  
  // Bereite Daten für Chart.js vor
  const datasets = sortedIncomeCategories.map((category, index) => {
    const values = years.map(year => {
      const income = extractIncome(profitLossReports[year]);
      return income[category] || 0;
    });
    
    // Farben basierend auf Index (verwende verschiedene Farbtöne)
    const colorIndex = index % COLORS.chartColors.length;
    const color = COLORS.chartColors[colorIndex];
    
    return {
      label: category,
      data: values,
      backgroundColor: `#${color}`,
      borderColor: `#${color}`,
      borderWidth: 1,
      stack: 'stack1'
    };
  });
  
  // Berechne maximale Gesamtsumme für Y-Achse
  const maxTotal = Math.max(...years.map(year => {
    const income = extractIncome(profitLossReports[year]);
    return Object.values(income).reduce((sum, val) => sum + val, 0);
  }));
  
  // Runde auf nächste 20.000 für Y-Achse
  const yAxisMax = Math.ceil(maxTotal / 20000) * 20000;
  
  // Erstelle Chart-Konfiguration für QuickChart
  const configuration = {
    type: 'bar',
    data: {
      labels: years,
      datasets: datasets
    },
    options: {
      indexAxis: 'x',
      plugins: {
        title: {
          display: true,
          text: 'Einnahmen',
          font: {
            size: 20,
            weight: 'bold'
          },
          padding: {
            top: 10,
            bottom: 20
          }
        },
        legend: {
          display: true,
          position: 'right',
          labels: {
            boxWidth: 15,
            padding: 8,
            font: {
              size: 10
            }
          }
        },
        datalabels: {
          color: '#ffffff',
          font: {
            weight: 'bold',
            size: 10
          },
          formatter: function(value, context) {
            return context.dataset.label;
          },
          anchor: 'center',
          align: 'center',
          display: true,
          clip: false,
          clamp: false
        },
        tooltip: {
          mode: 'index',
          intersect: false,
          callbacks: {
            label: function(context) {
              const value = context.parsed.y;
              return `${context.dataset.label}: ${formatGermanNumber(value)} €`;
            }
          }
        }
      },
      scales: {
        x: {
          stacked: true,
          title: {
            display: false
          },
          grid: {
            display: true,
            color: 'rgba(0, 0, 0, 0.1)'
          }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          max: yAxisMax,
          ticks: {
            stepSize: 20000,
            callback: function(value) {
              return formatGermanNumber(value);
            }
          },
          title: {
            display: false
          },
          grid: {
            display: true,
            color: 'rgba(0, 0, 0, 0.1)'
          }
        }
      },
      responsive: false,
      maintainAspectRatio: false
    }
  };
  
  // Erstelle QuickChart-Instanz
  const chart = new QuickChart();
  chart.setConfig(configuration);
  chart.setWidth(1200);
  chart.setHeight(1000);
  chart.setFormat('png');
  chart.setBackgroundColor('white');
  
  // Speichere direkt als PNG-Datei
  ensureResultDirectory();
  const outputPath = path.join(process.cwd(), 'Daten', 'result', `Einnahmen_${latestYear}.png`);
  await chart.toFile(outputPath);
  console.log(`Einnahmen-Diagramm gespeichert: ${outputPath}`);
}

/**
 * Erstellt gestapeltes Balkendiagramm der Ausgaben als PNG
 */
async function createExpensesChart(profitLossReports, years) {
  // Finde das neueste Jahr aus Bilanzberichten
  const balanceReports = readBalanceReports();
  const balanceYears = Object.keys(balanceReports).sort();
  const latestYear = balanceYears[balanceYears.length - 1];
  
  if (!latestYear) {
    console.log('Keine Bilanzberichte gefunden, kann Ausgaben-Diagramm nicht erstellen');
    return;
  }
  
  console.log(`Erstelle Ausgaben-Diagramm für Jahr ${latestYear}`);
  
  // Sammle alle Ausgaben-Kategorien über alle Jahre
  const allExpenseCategories = new Set();
  years.forEach(year => {
    const expenses = extractExpenses(profitLossReports[year]);
    Object.keys(expenses).forEach(cat => allExpenseCategories.add(cat));
  });
  
  // Berechne Gesamtsummen pro Kategorie über alle Jahre (für Sortierung)
  const expenseTotals = {};
  allExpenseCategories.forEach(category => {
    const total = years.reduce((sum, year) => {
      const expenses = extractExpenses(profitLossReports[year]);
      return sum + (expenses[category] || 0);
    }, 0);
    expenseTotals[category] = total;
  });

  // Berechne absolute Gesamtsummen pro Kategorie über alle Jahre (für Sortierung nach Größe)
  const expenseTotalsAbs = {};
  allExpenseCategories.forEach(category => {
    const totalAbs = years.reduce((sum, year) => {
      const expenses = extractExpenses(profitLossReports[year]);
      return sum + Math.abs(expenses[category] || 0);
    }, 0);
    expenseTotalsAbs[category] = totalAbs;
  });

  // Sortiere nach absoluter Gesamtsumme (Größe) - größte zuerst. So werden auch negative Kategorien berücksichtigt.
  const sortedExpenseCategories = Array.from(allExpenseCategories)
    .sort((a, b) => expenseTotalsAbs[b] - expenseTotalsAbs[a]);

  // Baue Datasets aus existierenden Kategorien
  // Option B: Zeige nur Top-10 Kategorien einzeln, aggregiere Rest in 'Sonstige'
  const darlehensLabel = 'Darlehensleistung';

  // Berechne Delta der Summe Schulden als Fallback für Darlehensleistung
  const deltaSchuldenByYearLocal = computeDeltaSchuldenForYears(years);
  const darlehensValues = years.map(year => {
    const expenses = extractExpenses(profitLossReports[year]);
    if (expenses[darlehensLabel] !== undefined) return expenses[darlehensLabel];
    return deltaSchuldenByYearLocal[year] || 0;
  });

  // Falls Darlehensleistung nicht in der Liste, ergänze sie mit berechneten Werten
  if (!sortedExpenseCategories.includes(darlehensLabel)) {
    const totalDarlehenAbs = darlehensValues.reduce((s, v) => s + Math.abs(v || 0), 0);
    // Füge nur hinzu, wenn tatsächlich Werte vorhanden
    if (totalDarlehenAbs > 0.0001) {
      sortedExpenseCategories.push(darlehensLabel);
      expenseTotals[darlehensLabel] = darlehensValues.reduce((s, v) => s + (v || 0), 0);
      expenseTotalsAbs[darlehensLabel] = totalDarlehenAbs;
      // Resort list by abs size including new entry
      sortedExpenseCategories.sort((a, b) => (expenseTotalsAbs[b] || 0) - (expenseTotalsAbs[a] || 0));
    }
  }

  // Bestimme Top-10 Kategorien
  const topCategories = sortedExpenseCategories.slice(0, 10);
  const otherCategories = sortedExpenseCategories.slice(10);

  const datasets = [];
  // Datasets für Top-10
  topCategories.forEach((category, index) => {
    const values = years.map(year => {
      if (category === darlehensLabel) return darlehensValues[years.indexOf(year)] || 0;
      const expenses = extractExpenses(profitLossReports[year]);
      return expenses[category] || 0;
    });
    const colorIndex = index % COLORS.chartColors.length;
    const color = COLORS.chartColors[colorIndex];
    const ds = {
      label: category,
      data: values,
      backgroundColor: `#${color}`,
      borderColor: `#${color}`,
      borderWidth: 1,
      stack: 'stack1'
    };
    // Wenn Kategorie in Top-10, aktiviere Datalabels nur für dieses Dataset
    ds.datalabels = {
      // Nur anzeigen, wenn der Wert != 0 (vermeidet z.B. "Darlehensleistungen 0 €")
      display: function(context) {
        try {
          const v = context && context.dataset && context.dataset.data && typeof context.dataIndex === 'number'
            ? context.dataset.data[context.dataIndex]
            : (context && context.parsed ? context.parsed : 0);
          return Math.abs(v || 0) > 0.5; // Schwelle: >0.5
        } catch (e) {
          return false;
        }
      },
      color: '#ffffff',
      font: { weight: 'bold', size: 22 },
      formatter: function(value, context) {
        try {
          const label = context && context.dataset && context.dataset.label ? context.dataset.label : '';
          return label + ' ' + formatGermanInteger(value) + ' €';
        } catch (e) {
          return (context && context.dataset && context.dataset.label ? context.dataset.label + ' ' : '') + Math.round(value) + ' €';
        }
      }
    };
    datasets.push(ds);
  });

  // Aggregiere Rest in 'Sonstiges'
  if (otherCategories.length > 0) {
    const othersValues = years.map(year => {
      return otherCategories.reduce((sum, category) => {
        if (category === darlehensLabel) return sum + (darlehensValues[years.indexOf(year)] || 0);
        const expenses = extractExpenses(profitLossReports[year]);
        return sum + (expenses[category] || 0);
      }, 0);
    });
    const anyNonZero = othersValues.some(v => Math.abs(v) > 0.0001);
    if (anyNonZero) {
      const color = 'A9A9A9'; // grau für Sonstige
      datasets.push({
        label: 'Sonstiges',
        data: othersValues,
        backgroundColor: `#${color}`,
        borderColor: `#${color}`,
        borderWidth: 1,
        stack: 'stack1'
      });
    }
  }

  // Bestimme maximale Gesamtsumme pro Jahr für Y-Achse
  const maxTotal = Math.max(...years.map((year, yi) => {
    return datasets.reduce((sum, ds) => sum + (ds.data[yi] || 0), 0);
  }).concat([0]));
  const yAxisMax = Math.max(20000, Math.ceil(maxTotal / 20000) * 20000);
  
  // Bestimme die 10 größten Kategorien (nach Gesamtsumme über alle Jahre) für Datapoints-Beschriftungen
  const totalsPerDataset = datasets.map(ds => ({
    label: ds.label,
    total: (ds.data || []).reduce((s, v) => s + (v || 0), 0)
  }));
  totalsPerDataset.sort((a, b) => b.total - a.total);
  const top10Labels = totalsPerDataset.slice(0, 10).map(d => d.label);
  
  // Erstelle Chart-Konfiguration für QuickChart
  const configuration = {
    type: 'bar',
    data: {
      labels: years,
      datasets: datasets
    },
    options: {
      indexAxis: 'x',
      plugins: {
        title: {
          display: true,
          text: 'Ausgaben',
          font: {
            size: 40,
            weight: 'bold'
          },
          padding: {
            top: 20,
            bottom: 40
          }
        },
        legend: {
          display: true,
          position: 'right',
          labels: {
            boxWidth: 30,
            padding: 16,
            font: {
              size: 20
            }
          }
        },
        // Standard: keine globalen Datalabels, wir steuern dies pro Dataset
        datalabels: {
          display: false
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const value = context.parsed.y;
              return `${context.dataset.label}: ${formatGermanNumber(value)} €`;
            }
          }
        }
      },
      scales: {
        x: {
          stacked: true,
          title: {
            display: false
          },
          grid: {
            display: true,
            color: 'rgba(0, 0, 0, 0.1)'
          }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          max: yAxisMax,
          ticks: {
            stepSize: 20000,
            callback: function(value) {
              return formatGermanNumber(value);
            }
          },
          title: {
            display: false
          },
          grid: {
            display: true,
            color: 'rgba(0, 0, 0, 0.1)'
          }
        }
      },
      responsive: false,
      maintainAspectRatio: false
    }
  };
  
  // Erstelle QuickChart-Instanz
  const chart = new QuickChart();
  chart.setConfig(configuration);
  chart.setWidth(1200);
  chart.setHeight(1200);
  chart.setFormat('png');
  chart.setBackgroundColor('white');
  
  // Speichere direkt als PNG-Datei
  ensureResultDirectory();
  const outputPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${latestYear}.png`);
  await chart.toFile(outputPath);
  console.log(`Ausgaben-Diagramm gespeichert: ${outputPath}`);
}

/**
 * Erstellt Kuchendiagramm der Ausgaben als PNG
 */
async function createExpensesPieChart(profitLossReports, years) {
  // Finde das neueste Jahr aus Bilanzberichten
  const balanceReports = readBalanceReports();
  const balanceYears = Object.keys(balanceReports).sort();
  const latestYear = balanceYears[balanceYears.length - 1];
  
  if (!latestYear) {
    console.log('Keine Bilanzberichte gefunden, kann Ausgaben-Kuchendiagramm nicht erstellen');
    return;
  }
  
  console.log(`Erstelle Ausgaben-Kuchendiagramm für Jahr ${latestYear}`);
  
  // Verwende das neueste Jahr für das Kuchendiagramm
  const expenses = extractExpenses(profitLossReports[latestYear]);
  
  // Sortiere nach Wert (absteigend)
  const sortedPieData = Object.entries(expenses)
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({ name, value }));
  
  // Bereite Daten für Chart.js vor
  const labels = sortedPieData.map(item => item.name);
  const data = sortedPieData.map(item => item.value);
  const backgroundColors = sortedPieData.map((item, index) => {
    const colorIndex = index % COLORS.chartColors.length;
    return `#${COLORS.chartColors[colorIndex]}`;
  });
  
  // Erstelle Chart-Konfiguration für QuickChart
  const configuration = {
    type: 'pie',
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: backgroundColors,
        borderColor: '#ffffff',
        borderWidth: 2
      }]
    },
    options: {
      plugins: {
        title: {
          display: true,
          text: `Ausgaben nach Kategorien (${latestYear})`,
          font: {
            size: 20,
            weight: 'bold'
          },
          padding: {
            top: 10,
            bottom: 20
          }
        },
        legend: {
          display: true,
          position: 'right',
          labels: {
            boxWidth: 15,
            padding: 8,
            font: {
              size: 10
            }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const label = context.label || '';
              const value = context.parsed || 0;
              const total = context.dataset.data.reduce((a, b) => a + b, 0);
              const percentage = ((value / total) * 100).toFixed(1);
              return `${label}: ${formatGermanNumber(value)} € (${percentage}%)`;
            }
          }
        }
      },
      responsive: false,
      maintainAspectRatio: false
    }
  };
  
  // Erstelle QuickChart-Instanz
  const chart = new QuickChart();
  chart.setConfig(configuration);
  chart.setWidth(1200);
  chart.setHeight(800);
  chart.setFormat('png');
  chart.setBackgroundColor('white');
  
  // Speichere direkt als PNG-Datei
  ensureResultDirectory();
  const outputPath = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_Kuchendiagramm_${latestYear}.png`);
  await chart.toFile(outputPath);
  console.log(`Ausgaben-Kuchendiagramm gespeichert: ${outputPath}`);
}

/**
 * Erstellt gestapeltes Balkendiagramm der Entwicklung als PNG
 */
async function createDevelopmentChart() {
  // Finde das neueste Jahr aus Bilanzberichten
  const balanceReports = readBalanceReports();
  const balanceYears = Object.keys(balanceReports).sort();
  const latestYear = balanceYears[balanceYears.length - 1];
  
  if (!latestYear) {
    console.log('Keine Bilanzberichte gefunden, kann Entwicklungs-Diagramm nicht erstellen');
    return;
  }
  
  console.log(`Erstelle Entwicklungs-Diagramm für Jahr ${latestYear}`);
  
  // Lese Entwicklung.csv
  const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
  const entwicklungData = readCSV(entwicklungPath);
  
  // Extrahiere Jahre aus Headern
  const firstRow = entwicklungData[0];
  const devYears = Object.keys(firstRow)
    .filter(key => key !== 'Position' && /^\d{4}$/.test(key))
    .sort();
  
  // Wichtige Konten für die Darstellung (Zeilen 2-6, keine Summen)
  const importantAccounts = [
    'Girokonto SKB 001',
    'Sparkonto SKB 003',
    'Freizeitkonto SKB 000',
    'Darlehenskonto SKB 004',
    'Privatdarlehen 006'
  ];
  
  // Sammle Daten für die Konten
  const accountData = [];
  importantAccounts.forEach(account => {
    const row = entwicklungData.find(r => r.Position === account);
    if (row) {
      const accountRow = { account };
      devYears.forEach(year => {
        accountRow[year] = parseGermanNumber(row[year] || '0');
      });
      accountData.push(accountRow);
    }
  });
  
  // Bereite Daten für Chart.js vor - alle Konten im gleichen Stack für korrektes Stacking
  const datasets = accountData.map((accountRow, index) => {
    const values = devYears.map(year => {
      let value = accountRow[year] || 0;
      // Werte behalten ihr Vorzeichen: Guthaben positiv (oben), Schulden negativ (unten)
      return value;
    });
    
    // Farben basierend auf Index
    const colorIndex = index % COLORS.chartColors.length;
    const color = COLORS.chartColors[colorIndex];
    
    return {
      label: accountRow.account,
      data: values,
      backgroundColor: `#${color}`,
      borderColor: `#${color}`,
      borderWidth: 1,
      stack: 'stack1'
    };
  });
  
  // Berechne maximale/minimale Werte für Y-Achse
  const allValues = accountData.flatMap(row => 
    devYears.map(year => {
      let value = row[year] || 0;
      // Behalte das Vorzeichen für korrekte Darstellung
      return value;
    })
  );
  
  const maxValue = Math.max(...allValues);
  const minValue = Math.min(...allValues);
  const range = Math.abs(maxValue) + Math.abs(minValue);
  
  // Runde Y-Achse sinnvoll
  let yAxisMax, yAxisMin;
  if (range > 0) {
    const step = Math.pow(10, Math.floor(Math.log10(range)));
    yAxisMax = Math.ceil(maxValue / step) * step;
    yAxisMin = Math.floor(minValue / step) * step;
  } else {
    yAxisMax = 10000;
    yAxisMin = -10000;
  }
  
  // Erstelle Chart-Konfiguration für QuickChart
  const configuration = {
    type: 'bar',
    data: {
      labels: devYears,
      datasets: datasets
    },
    options: {
      indexAxis: 'x',
      plugins: {
        title: {
          display: true,
          text: 'Entwicklung der Kontostände',
          font: {
            size: 420,
            weight: 'bold'
          },
          padding: {
            top: 210,
            bottom: 420
          }
        },
        legend: {
          display: true,
          position: 'right',
          labels: {
            boxWidth: 315,
            padding: 168,
            font: {
              size: 210
            }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const value = context.parsed.y;
              return `${context.dataset.label}: ${formatGermanNumber(value)} €`;
            }
          }
        }
      },
      scales: {
        x: {
          stacked: true,
          title: {
            display: false
          },
          grid: {
            display: true,
            color: 'rgba(0, 0, 0, 0.1)'
          }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          min: yAxisMin,
          max: yAxisMax,
          ticks: {
            font: {
              size: 168
            },
            callback: function(value) {
              return formatGermanNumber(value);
            }
          },
          title: {
            display: false
          },
          grid: {
            display: true,
            color: 'rgba(0, 0, 0, 0.1)'
          }
        }
      },
      responsive: false,
      maintainAspectRatio: false
    }
  };
  
  // Erstelle QuickChart-Instanz
  const chart = new QuickChart();
  chart.setConfig(configuration);
  chart.setWidth(800);
  chart.setHeight(600);
  chart.setFormat('png');
  chart.setBackgroundColor('white');
  
  // Speichere direkt als PNG-Datei
  ensureResultDirectory();
  const outputPath = path.join(process.cwd(), 'Daten', 'result', `Entwicklung_${latestYear}.png`);
  await chart.toFile(outputPath);
  console.log(`Entwicklungs-Diagramm gespeichert: ${outputPath}`);
}

/**
 * Liest die neueste Budget-Datei
 */
function findLatestBudgetFile() {
  const dataDir = path.join(process.cwd(), 'Daten');
  const files = fs.readdirSync(dataDir);
  const budgetFiles = files.filter(file => file.startsWith('Budgets_') && file.endsWith('.txt'));
  
  if (budgetFiles.length === 0) {
    return null;
  }
  
  // Sortiere nach Jahr (neueste zuerst)
  budgetFiles.sort((a, b) => {
    const yearA = parseInt(a.match(/\d{4}/)?.[0] || '0');
    const yearB = parseInt(b.match(/\d{4}/)?.[0] || '0');
    return yearB - yearA;
  });
  
  return path.join(dataDir, budgetFiles[0]);
}

/**
 * Konvertiert Budget-Wert zu Zahl (behandelt "k. A." und deutsche Zahlenformate)
 */
function parseBudgetValue(value) {
  if (!value || value.trim() === '' || value === 'k. A.') {
    return null;
  }
  
  // Entferne €-Symbol und Leerzeichen
  let cleaned = value.toString().trim().replace(/€/g, '').trim();
  
  // Wenn leer nach Entfernen, return null
  if (cleaned === '' || cleaned === 'k. A.') {
    return null;
  }
  
  // Parse deutsche Zahlenformatierung (1.234,56 -> 1234.56)
  return parseGermanNumber(cleaned);
}

/**
 * Parst Budget-Datei und konvertiert in strukturierte Daten
 */
function parseBudgetFile(filePath) {
  const content = fs.readFileSync(filePath, 'utf-8');
  const lines = content.split('\n').map(line => line.trim()).filter(line => line.length > 0);
  
  // Die ersten 7 Zeilen sind Header
  // Dann kommen die Daten in Blöcken von 8 Zeilen:
  // 0: Kostenstellenname
  // 1: Kostenstellennummer
  // 2: Verbraucht-Vorjahr
  // 3: Geplant-Vorjahr
  // 4: Geplant
  // 5: Stand (Prozent)
  // 6: Verbraucht
  // 7: Übrig
  
  const budgetData = [];
  let i = 7; // Starte nach den Headern
  
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
      
      // Parse Zahlenwerte (C, D, E, G, H)
      const verbrauchtVorjahr = parseBudgetValue(verbrauchtVorjahrRaw);
      const geplantVorjahr = parseBudgetValue(geplantVorjahrRaw);
      const geplant = parseBudgetValue(geplantRaw);
      const verbraucht = parseBudgetValue(verbrauchtRaw);
      const uebrig = parseBudgetValue(uebrigRaw);
      
      budgetData.push({
        kostenstelle: kostenstelle,
        nummer: nummer,
        verbrauchtVorjahr: verbrauchtVorjahr,
        verbrauchtVorjahrRaw: verbrauchtVorjahrRaw, // Für CSV
        geplantVorjahr: geplantVorjahr,
        geplantVorjahrRaw: geplantVorjahrRaw, // Für CSV
        geplant: geplant,
        geplantRaw: geplantRaw, // Für CSV
        stand: stand,
        verbraucht: verbraucht,
        verbrauchtRaw: verbrauchtRaw, // Für CSV
        uebrig: uebrig,
        uebrigRaw: uebrigRaw // Für CSV
      });
      
      i += 8;
    } else {
      break;
    }
  }
  
  return budgetData;
}

/**
 * Konvertiert Budget-Daten in CSV und speichert sie
 */
function saveBudgetAsCSV(budgetData, year) {
  ensureResultDirectory();
  const csvPath = path.join(process.cwd(), 'Daten', 'result', `Budgets_${year}.csv`);
  
  const csvRows = [
    ['Kostenstelle', 'Nummer', 'Verbraucht-Vorjahr', 'Geplant-Vorjahr', 'Geplant', 'Stand', 'Verbraucht', 'Übrig']
  ];
  
  budgetData.forEach(item => {
    csvRows.push([
      item.kostenstelle,
      item.nummer,
      item.verbrauchtVorjahr !== null ? formatGermanNumber(item.verbrauchtVorjahr) : 'k. A.',
      item.geplantVorjahr !== null ? formatGermanNumber(item.geplantVorjahr) : 'k. A.',
      item.geplant !== null ? formatGermanNumber(item.geplant) : 'k. A.',
      item.stand,
      item.verbraucht !== null ? formatGermanNumber(item.verbraucht) : 'k. A.',
      item.uebrig !== null ? formatGermanNumber(item.uebrig) : 'k. A.'
    ]);
  });
  
  const csvContent = csvRows.map(row => 
    row.map(cell => `"${cell}"`).join(';')
  ).join('\n');
  
  fs.writeFileSync(csvPath, csvContent, 'utf-8');
  console.log(`Budget-CSV erstellt: ${csvPath}`);
  
  return csvPath;
}

/**
 * Erstellt Budget-Reiter in Excel
 */
async function createBudgetSheet(workbook) {
  const budgetFilePath = findLatestBudgetFile();
  
  if (!budgetFilePath) {
    console.log('Keine Budget-Datei gefunden');
    return;
  }
  
  // Extrahiere Jahr aus Dateinamen
  const yearMatch = budgetFilePath.match(/Budgets_(\d{4})/);
  const year = yearMatch ? yearMatch[1] : '2025';
  
  console.log(`Lese Budget-Datei: ${budgetFilePath}`);
  
  // Parse Budget-Datei
  const budgetData = parseBudgetFile(budgetFilePath);
  
  // Speichere als CSV
  saveBudgetAsCSV(budgetData, year);
  
  // Erstelle Excel-Reiter
  const budgetSheet = workbook.addWorksheet(`Budget ${year}`);
  
  // Erstelle Spalten
  budgetSheet.columns = [
    { header: 'Kostenstelle', key: 'kostenstelle', width: 40 },
    { header: 'Nummer', key: 'nummer', width: 10 },
    { header: 'Verbraucht-Vorjahr', key: 'verbrauchtVorjahr', width: 18 },
    { header: 'Geplant-Vorjahr', key: 'geplantVorjahr', width: 18 },
    { header: 'Geplant', key: 'geplant', width: 15 },
    { header: 'Stand', key: 'stand', width: 10 },
    { header: 'Verbraucht', key: 'verbraucht', width: 15 },
    { header: 'Übrig', key: 'uebrig', width: 15 }
  ];
  
  // Füge Daten hinzu
  budgetData.forEach(item => {
    budgetSheet.addRow({
      kostenstelle: item.kostenstelle,
      nummer: item.nummer,
      verbrauchtVorjahr: item.verbrauchtVorjahr, // Spalte C - Zahl
      geplantVorjahr: item.geplantVorjahr, // Spalte D - Zahl
      geplant: item.geplant, // Spalte E - Zahl
      stand: item.stand,
      verbraucht: item.verbraucht, // Spalte G - Zahl
      uebrig: item.uebrig // Spalte H - Zahl
    });
  });
  
  // Formatiere Zahlen-Spalten (C, D, E, G, H)
  budgetSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Spalte C: Verbraucht-Vorjahr
      const cellC = row.getCell(3);
      if (cellC.value !== null && typeof cellC.value === 'number') {
        cellC.numFmt = '#,##0.00';
      }
      
      // Spalte D: Geplant-Vorjahr
      const cellD = row.getCell(4);
      if (cellD.value !== null && typeof cellD.value === 'number') {
        cellD.numFmt = '#,##0.00';
      }
      
      // Spalte E: Geplant
      const cellE = row.getCell(5);
      if (cellE.value !== null && typeof cellE.value === 'number') {
        cellE.numFmt = '#,##0.00';
      }
      
      // Spalte G: Verbraucht
      const cellG = row.getCell(7);
      if (cellG.value !== null && typeof cellG.value === 'number') {
        cellG.numFmt = '#,##0.00';
      }
      
      // Spalte H: Übrig
      const cellH = row.getCell(8);
      if (cellH.value !== null && typeof cellH.value === 'number') {
        cellH.numFmt = '#,##0.00';
      }
    }
  });
  
  // Header-Formatierung
  budgetSheet.getRow(1).font = { bold: true };
  budgetSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF' + COLORS.primary }
  };
  budgetSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  
  console.log(`Budget-Reiter erstellt: Budget ${year}`);
}

/**
 * Erstellt PowerPoint-Präsentation mit sortierten Daten
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
  
  // Erstelle Excel-Datei zuerst (sortiert die Daten)
  await createExcelFile(profitLossReports, years);
  
  // Erstelle Diagramme als PNG
  await createIncomeChart(profitLossReports, years);
  await createExpensesChart(profitLossReports, years);
  await createExpensesPieChart(profitLossReports, years);
  await createDevelopmentChart();
  
  // Slide 1: Einnahmen
  createIncomeSlide(pres, years);
  
  // Slide 2: Ausgaben
  createExpensesSlide(pres, years);
  
  // Slide 3: Kuchendiagramm Ausgaben
  createExpensesPieChartSlide(pres, years);
  
  // Slide 4: Entwicklung
  createDevelopmentSlide(pres);
  
  // Speichere Präsentation
  ensureResultDirectory();
  const outputPath = path.join(process.cwd(), 'Daten', 'result', 'Finanzlage_FeG_Eschweiler.pptx');
  await pres.writeFile({ fileName: outputPath });
  console.log(`Präsentation erstellt: ${outputPath}`);
}

/**
 * Erstellt Slide für Einnahmen mit sortierten Daten
 */
function createIncomeSlide(pres, years) {
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
  
  if (!sortedIncomeData) return;
  
  // Bereite Daten für gestapeltes Balkendiagramm vor
  const chartData = [];
  sortedIncomeData.categories.forEach(category => {
    const seriesData = years.map(year => {
      const item = sortedIncomeData.data.find(d => d.category === category);
      return item ? item[year] : 0;
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
    h: 4.5,
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
  
  sortedIncomeData.categories.forEach(category => {
    const item = sortedIncomeData.data.find(d => d.category === category);
    const row = [category];
    years.forEach(year => {
      row.push(formatGermanNumber(item ? item[year] : 0) + ' €');
    });
    tableData.push(row);
  });
  
  // Gesamtsumme
  const totals = ['Gesamt'];
  years.forEach(year => {
    totals.push(formatGermanNumber(sortedIncomeData.totals[year]) + ' €');
  });
  tableData.push(totals);
  
  slide.addTable(tableData, {
    x: 0.5,
    y: 5.8,
    w: 9,
    h: 1.5,
    fontSize: 8,
    colW: [3, 2, 2, 2],
    align: 'left',
    valign: 'middle',
    border: { type: 'solid', color: COLORS.text, pt: 1 }
  });
}

/**
 * Erstellt Slide für Ausgaben mit sortierten Daten
 */
function createExpensesSlide(pres, years) {
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
  
  if (!sortedExpensesData) return;
  
  // Bereite Daten für gestapeltes Balkendiagramm vor
  const chartData = [];
  sortedExpensesData.categories.forEach(category => {
    const seriesData = years.map(year => {
      const item = sortedExpensesData.data.find(d => d.category === category);
      return item ? item[year] : 0;
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
    h: 4.5,
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
  
  sortedExpensesData.categories.forEach(category => {
    const item = sortedExpensesData.data.find(d => d.category === category);
    const row = [category];
    years.forEach(year => {
      row.push(formatGermanNumber(item ? item[year] : 0) + ' €');
    });
    tableData.push(row);
  });
  
  // Gesamtsumme
  const totals = ['Gesamt'];
  years.forEach(year => {
    totals.push(formatGermanNumber(sortedExpensesData.totals[year]) + ' €');
  });
  tableData.push(totals);
  
  slide.addTable(tableData, {
    x: 0.5,
    y: 5.8,
    w: 9,
    h: 1.5,
    fontSize: 8,
    colW: [3, 2, 2, 2],
    align: 'left',
    valign: 'middle',
    border: { type: 'solid', color: COLORS.text, pt: 1 }
  });
}

/**
 * Erstellt Slide für Kuchendiagramm der Ausgaben mit sortierten Daten
 */
function createExpensesPieChartSlide(pres, years) {
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
  
  if (!sortedExpensesPieData) return;
  
  const latestYear = years[years.length - 1];
  
  // Bereite Daten für Kuchendiagramm vor
  const pieLabels = sortedExpensesPieData.map(({ name }) => name);
  const pieValues = sortedExpensesPieData.map(({ value }) => value);
  
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
  
  const total = pieValues.reduce((sum, val) => sum + val, 0);
  sortedExpensesPieData.forEach(({ name, value }) => {
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
  
  if (!developmentData) return;
  
  // Bereite Daten für Liniendiagramm vor
  const chartData = [];
  developmentData.data.forEach(({ account, ...values }) => {
    const chartValues = developmentData.years.map(year => {
      const value = values[year] || 0;
      // Für Darlehen: absoluter Wert für bessere Darstellung
      if (account.includes('Darlehen')) {
        return Math.abs(value);
      }
      return value;
    });
    chartData.push({
      name: account,
      labels: developmentData.years,
      values: chartValues
    });
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
    ['Konto', ...developmentData.years]
  ];
  
  // Sammle Daten für Summenberechnung
  const accountRows = [];
  
  developmentData.data.forEach(({ account, ...values }) => {
    const tableRow = [account];
    developmentData.years.forEach(year => {
      tableRow.push(formatGermanNumber(values[year] || 0) + ' €');
    });
    tableData.push(tableRow);
    accountRows.push({ account, ...values });
  });
  
  // Berechne Summe Guthaben (Aktiva-Konten)
  const guthabenAccounts = ['Girokonto SKB 001', 'Sparkonto SKB 003', 'Freizeitkonto SKB 000'];
  const summeGuthabenRow = ['Summe Guthaben'];
  developmentData.years.forEach(year => {
    const sum = accountRows
      .filter(item => guthabenAccounts.includes(item.account))
      .reduce((sum, item) => sum + (item[year] || 0), 0);
    summeGuthabenRow.push(formatGermanNumber(sum) + ' €');
  });
  tableData.push(summeGuthabenRow);
  
  // Berechne Summe Schulden (Passiva-Konten)
  const schuldenAccounts = ['Darlehenskonto SKB 004', 'Privatdarlehen 006'];
  const summeSchuldenRow = ['Summe Schulden'];
  developmentData.years.forEach(year => {
    const sum = accountRows
      .filter(item => schuldenAccounts.includes(item.account))
      .reduce((sum, item) => sum + Math.abs(item[year] || 0), 0);
    summeSchuldenRow.push(formatGermanNumber(sum) + ' €');
  });
  tableData.push(summeSchuldenRow);
  
  slide.addTable(tableData, {
    x: 0.5,
    y: 4.8,
    w: 9,
    h: 2.8,
    fontSize: 9,
    colW: [2.5, ...developmentData.years.map(() => 0.8)],
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
    
    // Erstelle PowerPoint-Präsentation (erstellt auch Excel-Datei)
    await createPresentation();
    
    console.log('Fertig!');
  } catch (error) {
    console.error('Fehler:', error);
    process.exit(1);
  }
}

main();
