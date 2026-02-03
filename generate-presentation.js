import { updateEntwicklungCSV } from './lib/utils.js';
import fs from 'fs';
import path from 'path';
// charts helper will be imported dynamically in createPresentation

const logFilePath = path.join(process.cwd(), 'debug.log');

function logToFile(message) {
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] ${message}\n`;
  
  // Log to console
  console.log(message);
  
  // Log to file
  fs.appendFileSync(logFilePath, logMessage, 'utf8');
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
  
  // Call the new JSON generation function
  // JSON generation is performed in createPresentation(), do not call here.

  // Remove the old image placeholder as we are now generating JSON
  // slide.addImage({
  //   path: chartImagePath,
  //   x: 0.5,
  //   y: 1.2,
  //   w: 9,
  //   h: 4.5,
  // });

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
  
  // Log chartData before it's used
  logToFile('chartData before chart creation: ' + JSON.stringify(chartData, null, 2));
  
// Erstelle gestapeltes Balkendiagramm für Einnahmen
      slide.addChart(pres.ChartType.bar, incomeChartData, {
        x: 0.5,
        y: 1.2,
        w: 9,
        h: 4.5,
        barGrouping: 'stacked',
        catAxisTitle: 'Jahr',
        valAxisTitle: 'Betrag (EUR)',
        showLegend: true,
        legendPos: 'b',
        chartColors: COLORS.chartColors,
        datalabels: {
          display: true,
          formatter: function(value, context) {
            try {
              // Access category from the data point
              const dataPoint = context.dataset.data[context.dataIndex];
              const category = dataPoint.__cat;
              const formattedValue = formatGermanInteger(value);
              return category + ': ' + formattedValue + ' €';
            } catch (e) {
              return formatGermanInteger(value) + ' €';
        }
      }
    }
  });
  
  // Füge Zahlenwerte als Tabelle hinzu
  const tableData = [
    ['Kategorie', ...years]
  ];
  
  sortedExpensesData.categories.forEach(category => {
    const item = sortedExpensesData.data.find(d => d.category === category);
    const row = [category];
    years.forEach(year => {
      row.push(category + ' ' + formatGermanNumber(item ? item[year] : 0) + ' €');
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
    chartColors: COLORS.chartColors,
    datalabels: {
      display: true // Enable datalabels globally, per-dataset config will take precedence
    }
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
    chartColors: COLORS.chartColors,
    datalabels: {
      display: true // Enable datalabels globally, per-dataset config will take precedence
    }
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

async function createPresentation() {
  const { readProfitLossReports } = await import('./lib/utils.js');
  const { createExcelFile } = await import('./lib/excel.js');
  const { createPPT } = await import('./lib/ppt.js');
  // import both chart modules
  const chartsModule = {};
  chartsModule.ausgaben = await import('./lib/generate-ausgaben-chart.js');
  chartsModule.einnahmen = await import('./lib/generate-einnahmen-chart.js');

  const profitLossReports = readProfitLossReports();
  const years = Object.keys(profitLossReports).sort();
  if (years.length === 0) {
    logToFile('Keine Gewinn-Verlust-Berichte gefunden.');
    return;
  }

  logToFile(`Gefundene Jahre: ${years.join(', ')}`);

  // Erstelle Excel (liefert sortierte Daten für PPT zurück)
  const excelResult = await createExcelFile(profitLossReports, years);

  // Debug dump of sortedExpensesData to help diagnose empty JSON output
  try {
    const debugOut = path.join(process.cwd(), 'Daten', 'result', 'debug_sortedExpensesData.json');
    fs.writeFileSync(debugOut, JSON.stringify(excelResult.sortedExpensesData || {}, null, 2), 'utf8');
    logToFile(`Wrote debug file: ${debugOut}`);
  } catch (e) {
    logToFile('Fehler beim Schreiben debug_sortedExpensesData: ' + e.message);
  }

  // Erstelle JSON-Datei für Ausgaben (Top 9 + Sonstiges) und PPT aus den Ergebnissen
  // current year used for naming the Ausgaben_YYYY.json (YYYY = currentYear - 1)
  const currentYear = new Date().getFullYear();
  // Debug: inspect sortedExpensesData
  try {
    logToFile('sortedExpensesData rows: ' + ((excelResult.sortedExpensesData && excelResult.sortedExpensesData.data) ? excelResult.sortedExpensesData.data.length : 0));
    if (excelResult.sortedExpensesData && excelResult.sortedExpensesData.data && excelResult.sortedExpensesData.data.length > 0) {
      logToFile('sortedExpensesData sample row: ' + JSON.stringify(excelResult.sortedExpensesData.data[0]));
    }
  } catch (e) {
    logToFile('Error logging sortedExpensesData: ' + e.message);
  }

    // Generate Ausgaben JSON
  if (chartsModule.generateAusgabenJsonFromSorted && excelResult.sortedExpensesData && Array.isArray(excelResult.sortedExpensesData.data) && excelResult.sortedExpensesData.data.length > 0) {
    const outPath = chartsModule.generateAusgabenJsonFromSorted(excelResult.sortedExpensesData, years, currentYear);
    logToFile(`Ausgaben JSON erstellt: ${outPath}`);
  } else if (chartsModule.generateAusgabenJson) {
    await chartsModule.generateAusgabenJson(currentYear);
    logToFile('Ausgaben JSON erstellt (fallback raw reports)');
  }

  // Generate Einnahmen JSON (analog)
  if (chartsModule.generateEinnahmenJsonFromSorted && excelResult.sortedIncomeData && Array.isArray(excelResult.sortedIncomeData.data) && excelResult.sortedIncomeData.data.length > 0) {
    const outPathInc = chartsModule.generateEinnahmenJsonFromSorted(excelResult.sortedIncomeData, years, currentYear);
    logToFile(`Einnahmen JSON erstellt: ${outPathInc}`);
  } else if (chartsModule.generateEinnahmenJson) {
    await chartsModule.generateEinnahmenJson(currentYear);
    logToFile('Einnahmen JSON erstellt (fallback raw reports)');
  }
  // Generate Ausgaben PNG from JSON using Playwright renderer if available
  try {
    const reportYear = currentYear - 1;
      if (chartsModule.ausgaben && chartsModule.ausgaben.generateAusgabenChartPlaywright) {
        const pngPath = await chartsModule.ausgaben.generateAusgabenChartPlaywright(reportYear);
        logToFile(`Ausgaben PNG erstellt (Playwright): ${JSON.stringify(pngPath)}`);
      } else if (chartsModule.generateAusgabenChartFromJson) {
        const pngPath = await chartsModule.generateAusgabenChartFromJson(reportYear);
        logToFile(`Ausgaben PNG erstellt (QuickChart): ${pngPath}`);
      }
  } catch (e) {
    logToFile('Fehler beim Erzeugen Ausgaben-PNG: ' + e.message);
  }
  // Generate Einnahmen PNG from JSON using Playwright renderer
  try {
    const reportYear = currentYear - 1;
    if (chartsModule.einnahmen && chartsModule.einnahmen.generateEinnahmenChartPlaywright) {
        const pngPath = await chartsModule.einnahmen.generateEinnahmenChartPlaywright(reportYear);
        logToFile(`Einnahmen PNG erstellt (Playwright): ${JSON.stringify(pngPath)}`);
      }
  } catch (e) {
    logToFile('Fehler beim Erzeugen Einnahmen-PNG: ' + e.message);
  }
  await createPPT(excelResult.sortedIncomeData, excelResult.sortedExpensesData, excelResult.sortedPieData);
}

// Hauptfunktion
async function main() {
  try {
    logToFile('Starte Generierung der Präsentation...');
    
    // Aktualisiere Entwicklung.csv
    updateEntwicklungCSV();
    
    // Erstelle PowerPoint-Präsentation (erstellt auch Excel-Datei)
    await createPresentation();
    
    // charts are generated inside createPresentation(), no need to run scripts/run-chart.js here
    
    logToFile('Fertig!');
  } catch (error) {
    logToFile('Fehler: ' + error);
    process.exit(1);
  }
}

main();
