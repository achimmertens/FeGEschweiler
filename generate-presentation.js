import { updateEntwicklungCSV, formatGermanInteger, parseGermanNumber, readCSV, extractExpenses, computeDeltaSchuldenForYears } from './lib/utils.js';
import fs from 'fs';
import path from 'path';
import QuickChart from 'quickchart-js';
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
  // Excel generation not required anymore - skip creating Finanzlage_FeG_Eschweiler.xlsx
  const { createPPT } = await import('./lib/ppt.js');
  // import both chart modules
  const chartsModule = {};
  chartsModule.ausgaben = await import('./lib/generate-ausgaben-chart.js');
  chartsModule.einnahmen = await import('./lib/generate-einnahmen-chart.js');
  // also import shared charts utilities (generates JSON from sorted data)
  try {
    const chartsLib = await import('./lib/charts.js');
    // attach helpers if available so existing calls work
    if (chartsLib.generateAusgabenJsonFromSorted) chartsModule.generateAusgabenJsonFromSorted = chartsLib.generateAusgabenJsonFromSorted;
    if (chartsLib.generateEinnahmenJsonFromSorted) chartsModule.generateEinnahmenJsonFromSorted = chartsLib.generateEinnahmenJsonFromSorted;
    if (chartsLib.generateAusgabenJson) chartsModule.generateAusgabenJson = chartsLib.generateAusgabenJson;
    if (chartsLib.generateEinnahmenJson) chartsModule.generateEinnahmenJson = chartsLib.generateEinnahmenJson;
  } catch (e) {
    logToFile('Warnung: lib/charts.js konnte nicht geladen werden: ' + (e && e.message ? e.message : String(e)));
  }

  const profitLossReports = readProfitLossReports();
  const years = Object.keys(profitLossReports).sort();
  if (years.length === 0) {
    logToFile('Keine Gewinn-Verlust-Berichte gefunden.');
    return;
  }

  logToFile(`Gefundene Jahre: ${years.join(', ')}`);

  // Excel generation skipped - provide empty structures used downstream
  logToFile('Excel-Erzeugung übersprungen (Finanzlage_FeG_Eschweiler.xlsx wird nicht erstellt).');
  const excelResult = {
    sortedExpensesData: { data: [], categories: [], totals: {} },
    sortedIncomeData: { data: [], categories: [], totals: {} },
    sortedPieData: []
  };

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
  // If a Budget CSV for the current year exists in /Daten, copy it to /Daten/result
  try {
    const budgetSrc = path.join(process.cwd(), 'Daten', `Budget_${currentYear}.csv`);
    const resultDir = path.join(process.cwd(), 'Daten', 'result');
    if (fs.existsSync(budgetSrc)) {
      if (!fs.existsSync(resultDir)) fs.mkdirSync(resultDir, { recursive: true });
      const budgetDst = path.join(resultDir, `Budget_${currentYear}.csv`);
      fs.copyFileSync(budgetSrc, budgetDst);
      logToFile(`Budget CSV für aktuelles Jahr kopiert: ${budgetDst}`);
    }
  } catch (e) {
    logToFile('Fehler beim Kopieren Budget CSV aktuelles Jahr: ' + (e && e.message ? e.message : String(e)));
  }
  // Also try to parse the latest Budgets_YYYY.txt in the Daten folder and save as CSV in result
  try {
    const dataDir = path.join(process.cwd(), 'Daten');
    const files = fs.existsSync(dataDir) ? fs.readdirSync(dataDir) : [];
    const budgetFiles = files.map(f => {
      const m = f.match(/^Budgets_(\d{4})\.txt$/);
      return m ? { file: path.join(dataDir, f), year: Number(m[1]) } : null;
    }).filter(Boolean).sort((a,b) => b.year - a.year);
    if (budgetFiles.length > 0) {
      const best = budgetFiles[0];
      const { parseBudgetFile, saveBudgetAsCSV } = await import('./lib/utils.js');
      const { generateBudgetHtml } = await import('./lib/budget.js');
      const parsed = parseBudgetFile(best.file);
      const outCsv = saveBudgetAsCSV(parsed, best.year);
      logToFile(`Budget CSV erstellt aus ${best.file}: ${outCsv}`);
      // Also generate a simple HTML table view of the budget CSV for inspection
      try {
        const csvRows = readCSV(outCsv);
        // collect headers from CSV and remove unwanted previous-year columns
        let headers = csvRows && csvRows.length ? Object.keys(csvRows[0]) : [];
        // Remove the two columns 'Verbraucht-Vorjahr' and 'Geplant-Vorjahr' from display
        headers = headers.filter(h => h !== 'Verbraucht-Vorjahr' && h !== 'Geplant-Vorjahr');
        const escapeHtml = s => String(s === undefined || s === null ? '' : s)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

        // Try to enrich with plan column from Budget_<currentYear>.csv if available
        const planYear = currentYear;
        const planCandidates = [];
        const planSrcA = path.join(process.cwd(), 'Daten', `Budget_${planYear}.csv`);
        const planSrcB = path.join(process.cwd(), 'Daten', 'result', `Budget_${planYear}.csv`);
        if (fs.existsSync(planSrcA)) planCandidates.push(planSrcA);
        if (fs.existsSync(planSrcB) && planSrcB !== outCsv) planCandidates.push(planSrcB);

        let planHeader = null;
        const planMap = {}; // key -> plan value (key = Nummer or Kostenstelle)
        try {
          for (const p of planCandidates) {
            const planRows = readCSV(p);
            if (!planRows || !planRows.length) continue;
            // Use the last column of the budget CSV as the plan for the current year
            const planCols = Object.keys(planRows[0] || {});
            if (!planCols || planCols.length === 0) continue;
            const ph = planCols[planCols.length - 1];
            planHeader = ph;
            planRows.forEach(r => {
              const key = (r['Nummer'] || r['Nummer'] === 0) ? String(r['Nummer']).trim() : String(r['Kostenstelle'] || '').trim();
              if (key) planMap[key] = r[planHeader];
            });
            if (planHeader) break;
          }
        } catch (e) {
          logToFile('Fehler beim Lesen Plan-CSV: ' + (e && e.message ? e.message : String(e)));
        }

        // Build table rows, append plan cell when available
        const tableRows = csvRows.map(r => {
          const base = headers.map(h => `<td>${escapeHtml(r[h]===undefined? '': r[h])}</td>`).join('');
          const key = (r['Nummer'] || r['Nummer'] === 0) ? String(r['Nummer']).trim() : String(r['Kostenstelle'] || '').trim();
          const planCell = planHeader ? `<td>${escapeHtml(planMap[key] !== undefined ? planMap[key] : '')}</td>` : '';
          return '<tr>' + base + planCell + '</tr>';
        }).join('\n');

        const finalHeaders = headers.slice();
        if (planHeader) finalHeaders.push(`Plan ${planYear}`);

        const html = `<!doctype html>\n<html lang="de">\n<head>\n<meta charset="utf-8"/>\n<meta name="viewport" content="width=device-width,initial-scale=1"/>\n<title>Budget ${best.year}</title>\n<style>body{font-family:Arial,sans-serif;margin:12px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left;font-size:13px}th{background:#f3f4f6}</style>\n</head>\n<body>\n<h1>Budget ${best.year} + ${best.year+1}</h1>\n<table>\n<thead>\n<tr>\n${finalHeaders.map(h=>`<th>${escapeHtml(h)}</th>`).join('\n')}\n</tr>\n</thead>\n<tbody>\n${tableRows}\n</tbody>\n</table>\n</body>\n</html>`;

        // Delegate HTML generation to lib/budget.js
        try {
          const generated = await generateBudgetHtml(outCsv, best.year, planYear);
          logToFile(`Budget HTML erstellt: ${generated}`);
        } catch (e) {
          // fallback: write the original html if helper fails
          const outHtml = path.join(process.cwd(), 'Daten', 'result', `Budget_${planYear}.html`);
          fs.writeFileSync(outHtml, html, 'utf8');
          logToFile(`Budget HTML (Fallback) erstellt: ${outHtml}`);
        }
      } catch (e) { logToFile('Fehler beim Erzeugen Budget-HTML: ' + (e && e.message ? e.message : String(e))); }
    } else {
      logToFile('Keine Budgets_YYYY.txt Datei im Daten-Ordner gefunden.');
    }
  } catch (e) { logToFile('Fehler beim Verarbeiten Budget-Datei: ' + (e && e.message ? e.message : String(e))); }
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
    try {
      const outPath = chartsModule.generateAusgabenJsonFromSorted(excelResult.sortedExpensesData, years, currentYear);
      if (outPath && fs.existsSync(outPath)) {
        logToFile(`Ausgaben JSON erstellt: ${outPath}`);
      } else {
        logToFile(`Ausgaben JSON: Funktion lieferte keinen gültigen Pfad oder Datei wurde nicht gefunden: ${outPath}`);
      }
    } catch (e) {
      logToFile('Fehler beim Erzeugen Ausgaben-JSON aus sortierten Daten: ' + (e && e.message ? e.message : String(e)));
    }
  } else if (chartsModule.generateAusgabenJson) {
    try {
      await chartsModule.generateAusgabenJson(currentYear);
      logToFile('Ausgaben JSON erstellt (fallback raw reports)');
    } catch (e) {
      logToFile('Fehler beim Erzeugen Ausgaben-JSON (fallback): ' + (e && e.message ? e.message : String(e)));
    }
  }

  // Generate Einnahmen JSON (analog)
  if (chartsModule.generateEinnahmenJsonFromSorted && excelResult.sortedIncomeData && Array.isArray(excelResult.sortedIncomeData.data) && excelResult.sortedIncomeData.data.length > 0) {
    try {
      const outPathInc = chartsModule.generateEinnahmenJsonFromSorted(excelResult.sortedIncomeData, years, currentYear);
      if (outPathInc && fs.existsSync(outPathInc)) {
        logToFile(`Einnahmen JSON erstellt: ${outPathInc}`);
      } else {
        logToFile(`Einnahmen JSON: Funktion lieferte keinen gültigen Pfad oder Datei wurde nicht gefunden: ${outPathInc}`);
      }
    } catch (e) {
      logToFile('Fehler beim Erzeugen Einnahmen-JSON aus sortierten Daten: ' + (e && e.message ? e.message : String(e)));
    }
  } else if (chartsModule.generateEinnahmenJson) {
    try {
      await chartsModule.generateEinnahmenJson(currentYear);
      logToFile('Einnahmen JSON erstellt (fallback raw reports)');
    } catch (e) {
      logToFile('Fehler beim Erzeugen Einnahmen-JSON (fallback): ' + (e && e.message ? e.message : String(e)));
    }
  }
  // Overwrite (or create) authoritative Einnahmen JSON from CSV reports to ensure totals match source
  try {
    const reportYear = currentYear - 1;
    const yearsToProcess = [currentYear - 3, currentYear - 2, currentYear - 1];
    const reports = readProfitLossReports();
    const incomesOut = {};
    for (const y of yearsToProcess) {
      const rows = reports[String(y)] || [];
      const map = {};
      let total = 0;
      rows.forEach(r => {
        try {
          const v = parseGermanNumber(r.Summe || r['Summe'] || 0);
          const kk = (r.Kontoklasse || r['Kontoklasse'] || '').toString().toLowerCase();
          const name = (r.Name || r['Name'] || '').toString();
          const isRevenueClass = /betr\.? erträge/i.test(r.Kontoklasse || r['Kontoklasse'] || '') || /umsatzerlöse/i.test(r.Kontoklasse || r['Kontoklasse'] || '');
          const isInterest = /haben-?zinsen|zinsen/i.test(name.toLowerCase());
          if (v > 0 || isRevenueClass || isInterest) {
            const key = name || '(unnamed)';
            map[key] = (map[key] || 0) + Number(v || 0);
            total += Number(v || 0);
          }
        } catch (e) { /* ignore row parse errors */ }
      });
      // normalize and round
      Object.keys(map).forEach(k => map[k] = Number(Number(map[k] || 0).toFixed(2)));
      const sumMap = Object.values(map).reduce((s,n)=>s+Number(n||0),0);
      const gap = Number((total - sumMap).toFixed(2));
      if (Math.abs(gap) > 0.001) map['Sonstiges'] = Number((map['Sonstiges'] || 0) + gap);
      else if (!map.hasOwnProperty('Sonstiges')) map['Sonstiges'] = 0;
      incomesOut[y] = map;
      logToFile(`Einnahmen authoritative ${y}: total=${total} listed=${sumMap} gap=${gap}`);
    }
    const outPath = path.join(process.cwd(), 'Daten', 'result', `Einnahmen_${reportYear}.json`);
    fs.writeFileSync(outPath, JSON.stringify(incomesOut, null, 2), 'utf8');
    logToFile(`Einnahmen JSON (authoritative) geschrieben: ${outPath}`);
  } catch (e) { logToFile('Fehler beim Erzeugen authoritativer Einnahmen-JSON: ' + (e && e.message ? e.message : String(e))); }
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
  // Generate Entwicklung HTML (fallback) and PNG preview optional
  try {
    // read Entwicklung.csv produced by updateEntwicklungCSV
    const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
    if (fs.existsSync(entwicklungPath)) {
      const entwicklung = readCSV(entwicklungPath);
      const firstRow = entwicklung[0] || {};
      const devYears = Object.keys(firstRow).filter(k => k !== 'Position' && /^\d{4}$/.test(k)).sort();
      const accounts = ['Girokonto SKB 001','Sparkonto SKB 003','Freizeitkonto SKB 000','Darlehenskonto SKB 004','Privatdarlehen 006'];
      const datasets = accounts.map((account, idx) => {
        const row = entwicklung.find(r => r.Position === account) || {};
        // parse values as numbers (keep sign as in CSV)
        const vals = devYears.map(y => parseGermanNumber(row[y] || '0'));
        const colors = ['#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6'];
        return { label: account, data: vals, color: colors[idx % colors.length] };
      });

      // Build a nicely formatted HTML matching the requested template
      const labelsJSON = JSON.stringify(devYears, null, 4);
      const datasetsForHtml = datasets.map(d => ({ label: d.label, data: d.data, backgroundColor: d.color, stack: 'stack1' }));
      const datasetsJSON = JSON.stringify(datasetsForHtml, null, 4);

      const html = `<!doctype html>
<html lang="de">
    <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width,initial-scale=1" />
        <title>Entwicklung Kontostände ${devYears[devYears.length-1]}</title>
        <style>
            body { margin: 0; font-family: Arial, sans-serif; background: #fff; }
            .container { padding: 12px; }
            canvas { display: block; }
            h1 { font-size: 18px; margin: 0 0 8px 0; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Entwicklung Kontostände (bis ${devYears[devYears.length-1]})</h1>
            <canvas id="chart" width="1200" height="800"></canvas>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <script>
            const labels = ${labelsJSON};
            const datasets = ${datasetsJSON};

            const config = {
                type: 'bar',
                data: { labels, datasets },
                options: {
                    indexAxis: 'x',
                    plugins: {
                        title: { display: true, text: 'Entwicklung der Kontostände', font: { size: 20 } },
                        legend: { position: 'top' }
                    },
                    scales: {
                        x: { stacked: true },
                        y: {
                            stacked: true,
                            ticks: {
                                callback: function(value) {
                                    return value.toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, '.');
                                }
                            }
                        }
                    },
                    responsive: false,
                    maintainAspectRatio: false
                }
            };

            const ctx = document.getElementById('chart').getContext('2d');
            new Chart(ctx, config);
        </script>
    </body>
</html>`;
      const outHtml = path.join(process.cwd(), 'Daten', 'result', `Entwicklung_${devYears[devYears.length-1]}.html`);
      fs.writeFileSync(outHtml, html, 'utf8');
      logToFile(`Entwicklung HTML erstellt: ${outHtml}`);
      // Try to render the HTML to a PNG using Playwright (if available)
      try {
        const pw = await import('playwright');
        const browser = await pw.chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1200, height: 800 } });
        // Use file:// URL to load the generated HTML
        const fileUrl = 'file://' + outHtml.replace(/\\/g, '/');
        await page.goto(fileUrl, { waitUntil: 'networkidle' }).catch(()=>{});
        // wait a short moment for Chart.js to render
        await page.waitForSelector('canvas', { timeout: 2000 }).catch(()=>{});
        await page.waitForTimeout(500);
        const canvas = await page.$('canvas#chart') || await page.$('canvas');
        const outPng = path.join(process.cwd(), 'Daten', 'result', `Entwicklung_${devYears[devYears.length-1]}.png`);
        if (canvas) {
          const box = await canvas.boundingBox();
          if (box && box.width > 0 && box.height > 0) {
            await page.screenshot({ path: outPng, clip: { x: Math.max(0, Math.floor(box.x)), y: Math.max(0, Math.floor(box.y)), width: Math.ceil(box.width), height: Math.ceil(box.height) } });
          } else {
            await page.screenshot({ path: outPng, fullPage: true });
          }
        } else {
          await page.screenshot({ path: outPng, fullPage: true });
        }
        await browser.close();
        logToFile(`Entwicklung PNG erstellt (Playwright): ${outPng}`);
      } catch (e) {
        logToFile('Playwright-Render für Entwicklung-PNG fehlgeschlagen: ' + (e && e.message ? e.message : String(e)));
      }
    }
  } catch (e) { logToFile('Fehler beim Erzeugen Entwicklung-HTML: ' + e.message); }
  // Generate Entwicklung chart (based on Entwicklung.csv) as PNG via QuickChart
  try {
    const balanceReports = await import('./lib/utils.js');
    const balanceYears = Object.keys(balanceReports.readBalanceReports()).sort();
    const latestYear = balanceYears[balanceYears.length - 1];
    if (latestYear) {
      // read Entwicklung.csv and prepare datasets
      const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
      if (fs.existsSync(entwicklungPath)) {
        const entwicklung = (await import('./lib/utils.js')).readCSV(entwicklungPath);
        const firstRow = entwicklung[0] || {};
        const devYears = Object.keys(firstRow).filter(k => k !== 'Position' && /^\d{4}$/.test(k)).sort();
        const accounts = ['Girokonto SKB 001','Sparkonto SKB 003','Freizeitkonto SKB 000','Darlehenskonto SKB 004','Privatdarlehen 006'];
        const datasets = accounts.map((account, idx) => {
          const row = entwicklung.find(r => r.Position === account) || {};
          const vals = devYears.map(y => parseGermanNumber(row[y] || '0'));
          return { label: account, data: vals };
        });
        const configuration = {
          type: 'bar',
          data: { labels: devYears, datasets: datasets.map((d, i) => ({ label: d.label, data: d.data, backgroundColor: ['#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6'][i%5], stack: 'stack1' })) },
          options: { indexAxis: 'x', plugins: { title: { display: true, text: `Entwicklung ${latestYear}`, font: { size: 20 } } }, scales: { x: { stacked: true }, y: { stacked: true, ticks: { callback: function(v){ return formatGermanInteger(v); } } } }, responsive: false, maintainAspectRatio: false }
        };
        const qc = new QuickChart(); qc.setConfig(configuration); qc.setWidth(1200); qc.setHeight(800); qc.setFormat('png'); qc.setBackgroundColor('white');
        const outDev = path.join(process.cwd(), 'Daten', 'result', `Entwicklung_${latestYear}.png`);
        await qc.toFile(outDev);
        logToFile(`Entwicklung PNG erstellt: ${outDev}`);
      }
    }
  } catch (e) { logToFile('Fehler beim Erzeugen Entwicklung-PNG: ' + e.message); }
  
  // Generate index.html for navigation
  try {
    const indexHtml = `<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FeG Eschweiler - Finanzberichte</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: #f5f5f5;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 2rem;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        .header h1 {
            font-size: 2rem;
            margin-bottom: 0.5rem;
        }

        .header p {
            opacity: 0.9;
            font-size: 1rem;
        }

        .nav {
            background: white;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            position: sticky;
            top: 0;
            z-index: 100;
        }

        .nav-container {
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            flex-wrap: wrap;
            padding: 0;
        }

        .nav-item {
            flex: 1;
            min-width: 200px;
            text-align: center;
            border-right: 1px solid #e5e5e5;
        }

        .nav-item:last-child {
            border-right: none;
        }

        .nav-item a {
            display: block;
            padding: 1.2rem 1.5rem;
            color: #333;
            text-decoration: none;
            font-weight: 500;
            transition: all 0.3s ease;
            position: relative;
        }

        .nav-item a:hover {
            background: #667eea;
            color: white;
        }

        .nav-item a.active {
            background: #667eea;
            color: white;
        }

        .nav-item a::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 50%;
            transform: translateX(-50%);
            width: 0;
            height: 3px;
            background: #764ba2;
            transition: width 0.3s ease;
        }

        .nav-item a:hover::after,
        .nav-item a.active::after {
            width: 80%;
        }

        .content {
            max-width: 1400px;
            margin: 2rem auto;
            padding: 0 1rem;
        }

        .iframe-container {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
            min-height: 600px;
        }

        iframe {
            width: 100%;
            height: 850px;
            border: none;
            display: block;
        }

        .welcome {
            background: white;
            border-radius: 8px;
            padding: 3rem;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            text-align: center;
        }

        .welcome h2 {
            color: #667eea;
            font-size: 2rem;
            margin-bottom: 1rem;
        }

        .welcome p {
            color: #666;
            font-size: 1.1rem;
            line-height: 1.6;
            max-width: 600px;
            margin: 0 auto;
        }

        .card-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 1.5rem;
            margin-top: 2rem;
        }

        .card {
            background: white;
            border-radius: 8px;
            padding: 2rem;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            cursor: pointer;
        }

        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 20px rgba(0,0,0,0.15);
        }

        .card h3 {
            color: #667eea;
            margin-bottom: 0.5rem;
            font-size: 1.3rem;
        }

        .card p {
            color: #666;
            line-height: 1.5;
        }

        .footer {
            text-align: center;
            padding: 2rem;
            color: #999;
            margin-top: 3rem;
        }

        @media (max-width: 768px) {
            .nav-container {
                flex-direction: column;
            }

            .nav-item {
                border-right: none;
                border-bottom: 1px solid #e5e5e5;
            }

            .nav-item:last-child {
                border-bottom: none;
            }

            .header h1 {
                font-size: 1.5rem;
            }

            .welcome {
                padding: 2rem 1rem;
            }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>FeG Eschweiler - Finanzberichte</h1>
        <p>Übersicht der Finanzberichte und Budgets</p>
    </div>

    <nav class="nav">
        <div class="nav-container">
            <div class="nav-item">
                <a href="#" onclick="showWelcome(); return false;" class="active" id="nav-welcome">Startseite</a>
            </div>
            <div class="nav-item">
              <a href="#" onclick="loadPage('Checkliste.html', this); return false;">Checkliste</a>
            </div>
            <div class="nav-item">
                <a href="#" onclick="loadPage('Ausgaben_2025.html', this); return false;">Ausgaben 2025</a>
            </div>
            <div class="nav-item">
              <a href="#" onclick="loadPage('AusgabenTab_${currentYear-1}.html', this); return false;">AusgabenTabelle</a>
            </div>
            <div class="nav-item">
                <a href="#" onclick="loadPage('Einnahmen_2025.html', this); return false;">Einnahmen 2025</a>
            </div>
            <div class="nav-item">
              <a href="#" onclick="loadPage('EinnahmenTab_${currentYear-1}.html', this); return false;">EinnahmenTabelle</a>
            </div>
            <div class="nav-item">
              <a href="#" onclick="loadPage('Entwicklung_2025.html', this); return false;">Entwicklung 2025</a>
            </div>
            <div class="nav-item">
                <a href="#" onclick="loadPage('EntwicklungTab_${currentYear-1}.html', this); return false;">EntwicklungTab</a>
            </div>
            <div class="nav-item">
              <a href="#" onclick="loadPage('Budget_2026.html', this); return false;">Budget 2026</a>
            </div>
            <div class="nav-item">
              <a href="#" onclick="loadPage('Sonderspenden.html', this); return false;">Sonderspenden</a>
            </div>
            <div class="nav-item">
              <a href="#" onclick="loadPage('JahresabschlussUnterlagen_${currentYear-1}.html', this); return false;">JahresabschlussUnterlagen ${currentYear-1}</a>
            </div>
        </div>
    </nav>

    <div class="content">
        <div id="welcome-section" class="welcome">
            <h2>Willkommen</h2>
            <p>Hier finden Sie eine Übersicht über die Finanzberichte der FeG Eschweiler. Wählen Sie einen Bericht aus dem Menü oben aus, um die Details anzuzeigen.</p>
            
            <div class="card-grid">
                <div class="card" onclick="loadPage('Ausgaben_2025.html', document.querySelector('[onclick*=Ausgaben]'))">
                    <h3>📊 Ausgaben 2025</h3>
                    <p>Detaillierte Übersicht der Ausgaben für das Jahr 2025 im Vergleich zu den Vorjahren.</p>
                </div>
                <div class="card" onclick="loadPage('Einnahmen_2025.html', document.querySelector('[onclick*=Einnahmen]'))">
                  <h3>💰 Einnahmen 2025</h3>
                  <p>Übersicht der Einnahmen und Erträge für das Jahr 2025.</p>
                </div>
                <div class="card" onclick="loadPage('EinnahmenTab_${currentYear-1}.html', document.querySelector('[onclick*=EinnahmenTabelle]'))">
                  <h3>📋 EinnahmenTabelle</h3>
                  <p>Tabellarische Darstellung der Einnahmen für ${currentYear-1}.</p>
                </div>
                <div class="card" onclick="loadPage('Checkliste.html', document.querySelector('[onclick*=Checkliste]'))">
                  <h3>📝 Checkliste</h3>
                  <p>Jährliche Schritte zur Aktualisierung der Daten.</p>
                </div>
                <div class="card" onclick="loadPage('Entwicklung_2025.html', document.querySelector('[onclick*=Entwicklung]'))">
                    <h3>📈 Entwicklung 2025</h3>
                    <p>Entwicklung der Finanzen über mehrere Jahre hinweg.</p>
                </div>
                <div class="card" onclick="loadPage('EntwicklungTab_${currentYear-1}.html', document.querySelector('[onclick*=EntwicklungTab]'))">
                  <h3>🗂 EntwicklungTab</h3>
                  <p>Tabellarische Darstellung von Entwicklung.csv für ${currentYear-1}.</p>
                </div>
                <div class="card" onclick="loadPage('Budget_2026.html', document.querySelector('[onclick*=Budget]'))">
                  <h3>📋 Budget 2026</h3>
                  <p>Budgetplanung und -übersicht für das Jahr 2026.</p>
                </div>
                <div class="card" onclick="loadPage('Sonderspenden.html', document.querySelector('[onclick*=Sonderspenden]'))">
                  <h3>🎁 Sonderspenden</h3>
                  <p>Übersicht zu anstehenden Sonderspenden und Terminen.</p>
                </div>
                <div class="card" onclick="loadPage('JahresabschlussUnterlagen_${currentYear-1}.html', document.querySelector('[onclick*=JahresabschlussUnterlagen]'))">
                  <h3>📋 Jahresabschlussunterlagen ${currentYear-1}</h3>
                  <p>Bilanz und Gewinn-Verlust-Rechnung für ${currentYear-1}.</p>
                </div>
            </div>
        </div>

        <div id="iframe-section" class="iframe-container" style="display: none;">
            <iframe id="content-frame" src=""></iframe>
        </div>
    </div>

    <div class="footer">
        <p>&copy; ${currentYear} FeG Eschweiler - Finanzberichte</p>
    </div>

    <script>
        function loadPage(page, linkElement) {
            // Update active state
            document.querySelectorAll('.nav-item a').forEach(a => a.classList.remove('active'));
            if (linkElement) {
                linkElement.classList.add('active');
            }

            // Hide welcome, show iframe
            document.getElementById('welcome-section').style.display = 'none';
            document.getElementById('iframe-section').style.display = 'block';

            // Load page
            document.getElementById('content-frame').src = page;
        }

        function showWelcome() {
            // Update active state
            document.querySelectorAll('.nav-item a').forEach(a => a.classList.remove('active'));
            document.getElementById('nav-welcome').classList.add('active');

            // Show welcome, hide iframe
            document.getElementById('welcome-section').style.display = 'block';
            document.getElementById('iframe-section').style.display = 'none';

            // Clear iframe
            document.getElementById('content-frame').src = '';
        }
    </script>
</body>
</html>`;
    const indexPath = path.join(process.cwd(), 'Daten', 'result', 'index.html');
    fs.writeFileSync(indexPath, indexHtml, 'utf8');
    logToFile(`Index HTML erstellt: ${indexPath}`);
  } catch (e) {
    logToFile('Fehler beim Erzeugen index.html: ' + (e && e.message ? e.message : String(e)));
  }
  
  // Also create a simple Sonderspenden page if Sonderspenden CSV exists
  try {
    const { generateSonderspendenPage } = await import('./lib/sonderspenden.js');
    const sOut = await generateSonderspendenPage();
    if (sOut) logToFile(`Sonderspenden HTML erstellt: ${sOut}`);
  } catch (e) { logToFile('Fehler beim Erzeugen Sonderspenden-Seite: ' + (e && e.message ? e.message : String(e))); }

  // Create Einnahmen tabellarische Seite using Einnahmen JSON if available
  try {
    const reportYear = currentYear - 1;
    const einJson = path.join(process.cwd(), 'Daten', 'result', `Einnahmen_${reportYear}.json`);
    if (fs.existsSync(einJson)) {
      const data = JSON.parse(fs.readFileSync(einJson, 'utf8')) || {};
      // build table: categories x values
      const years = Object.keys(data).sort();
      const categories = Object.keys(years.length ? data[years[0]] : {});
      const filteredCategories = categories.filter(cat => {
        if (cat === 'Sonstiges') {
          return years.some(year => Number((data[year] && data[year][cat]) || 0) !== 0);
        }
        return true;
      });
      const latestYear = years[years.length - 1];
      const sortedCategories = filteredCategories.slice().sort((a, b) => {
        const valueA = Number((data[latestYear] && data[latestYear][a]) || 0);
        const valueB = Number((data[latestYear] && data[latestYear][b]) || 0);
        return valueA - valueB;
      });
      const headers = ['Kategorie', ...years];
      const escapeHtml = s => String(s === undefined || s === null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      let formatGermanNumber = n => (Number(n||0).toFixed(2)).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
      try { const _utils = await import('./lib/utils.js'); if (_utils.formatGermanNumber) formatGermanNumber = _utils.formatGermanNumber; } catch (e) {}
      const formatGermanInteger = value => {
        const raw = formatGermanNumber(value);
        return raw.endsWith(',00') ? raw.slice(0, -3) : raw;
      };
      // rows for categories
      const rows = sortedCategories.map(cat => {
        const cells = years.map(y => {
          const value = Number((data[y] && data[y][cat]) || 0);
          return `<td>${escapeHtml(formatGermanInteger(value))}</td>`;
        }).join('');
        return `<tr><td>${escapeHtml(cat)}</td>${cells}</tr>`;
      }).join('\n');
      // totals per year (sum of categories)
      const totals = years.map(y => sortedCategories.reduce((s,cat) => s + (Number((data[y] && data[y][cat]) || 0)), 0));
      const totalRow = '<tr>' + [`<td><strong>Gesamt</strong></td>`, ...totals.map(t => {
        const raw = formatGermanInteger(t);
        const withoutCents = raw.endsWith(',00') ? raw.slice(0, -3) : raw;
        return `<td><strong>${escapeHtml(withoutCents)} €</strong></td>`;
      })].join('') + '</tr>';
      const tableHtml = `<!doctype html><html lang="de"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>Einnahmen ${reportYear}</title><style>body{font-family:Arial,sans-serif;margin:12px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left;font-size:13px}th{background:#f3f4f6}</style></head><body><h1>Einnahmen ${reportYear}</h1><table><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${rows}\n${totalRow}</tbody></table></body></html>`;
      const outPath = path.join(process.cwd(), 'Daten', 'result', `EinnahmenTab_${reportYear}.html`);
      fs.writeFileSync(outPath, tableHtml, 'utf8');
      logToFile(`Einnahmen-Tabelle erstellt: ${outPath}`);
    }
  } catch (e) { logToFile('Fehler beim Erzeugen Einnahmen-Tabelle: ' + (e && e.message ? e.message : String(e))); }

  // Create Ausgaben-Tab similarly
  try {
    const reportYear = currentYear - 1;
    const einYear = currentYear - 1;
    const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
    if (fs.existsSync(entwicklungPath)) {
      const entwicklungRows = readCSV(entwicklungPath);
      if (entwicklungRows && entwicklungRows.length > 0) {
        const columnsOrig = Object.keys(entwicklungRows[0] || {});
        const positionCol = 'Position';
        const orderedColumns = [positionCol, ...columnsOrig.filter(col => col !== positionCol)];
        const escapeHtml = s => String(s === undefined || s === null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        let formatGermanNumber = n => (Number(n||0).toFixed(2)).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
        try { const _utils = await import('./lib/utils.js'); if (_utils.formatGermanNumber) formatGermanNumber = _utils.formatGermanNumber; } catch (e) {}
        const formatGermanInteger = (value) => {
          const raw = formatGermanNumber(value);
          return raw.endsWith(',00') ? raw.slice(0, -3) : raw;
        };
        const tableRows = entwicklungRows.map(row => {
          const cells = orderedColumns.map(col => {
            const cellValue = row[col];
            if (col === positionCol) {
              return `<td>${escapeHtml(cellValue)}</td>`;
            }
            const numeric = typeof cellValue === 'string' && /-?\d/.test(cellValue.replace(/\./g, '').replace(',', '.'));
            if (!numeric) {
              return `<td>${escapeHtml(cellValue)}</td>`;
            }
            const parsed = parseGermanNumber(cellValue);
            return `<td>${escapeHtml(formatGermanInteger(parsed))} €</td>`;
          }).join('');
          return `<tr>${cells}</tr>`;
        }).join('\n');
        const headers = orderedColumns.map(h => h === positionCol ? 'Position' : h);
        const tableHtml = `<!doctype html><html lang="de"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>Entwicklung ${reportYear}</title><style>body{font-family:Arial,sans-serif;margin:12px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left;font-size:13px}th{background:#f3f4f6}</style></head><body><h1>Entwicklung ${reportYear}</h1><table><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${tableRows}</tbody></table></body></html>`;
        const outPath = path.join(process.cwd(), 'Daten', 'result', `EntwicklungTab_${reportYear}.html`);
        fs.writeFileSync(outPath, tableHtml, 'utf8');
        logToFile(`Entwicklung-Tab erstellt: ${outPath}`);
      }
    }
  } catch (e) {
    logToFile('Fehler beim Erzeugen Entwicklung-Tab: ' + (e && e.message ? e.message : String(e)));
  }

  // Create a simple annual checklist page (delegated to lib/checklist.js)
  try {
    const { generateDefaultChecklist } = await import('./lib/checklist.js');
    const chkOut = generateDefaultChecklist(currentYear);
    logToFile(`Checkliste HTML erstellt: ${chkOut}`);
  } catch (e) { logToFile('Fehler beim Erzeugen Checkliste-Seite: ' + (e && e.message ? e.message : String(e))); }

  // Create Jahresabschluss page
  try {
    const reportYear = currentYear - 1;
    const bilanzPath = path.join(process.cwd(), 'Daten', `bilanzbericht_${reportYear}.csv`);
    const gvPath = path.join(process.cwd(), 'Daten', `gewinn-verlust-bericht_${reportYear}.csv`);
    const einJson = path.join(process.cwd(), 'Daten', 'result', `Einnahmen_${reportYear}.json`);
    const ausgabenJson = path.join(process.cwd(), 'Daten', 'result', `Ausgaben_${reportYear}.json`);
    const entwicklungPath = path.join(process.cwd(), 'Daten', 'Entwicklung.csv');
    if (fs.existsSync(bilanzPath) && fs.existsSync(gvPath) && fs.existsSync(einJson) && fs.existsSync(ausgabenJson) && fs.existsSync(entwicklungPath)) {
      const bilanzRows = readCSV(bilanzPath);
      const gvRows = readCSV(gvPath);
      const einData = JSON.parse(fs.readFileSync(einJson, 'utf8')) || {};
      const ausgabenData = JSON.parse(fs.readFileSync(ausgabenJson, 'utf8')) || {};
      const entwicklungRows = readCSV(entwicklungPath);
      const escapeHtml = s => String(s === undefined || s === null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      const formatGermanNumber = n => (Number(n||0).toFixed(2)).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
      const formatGermanInteger = (value) => {
        const raw = formatGermanNumber(value);
        return raw.endsWith(',00') ? raw.slice(0, -3) : raw;
      };

      // Einnahmen Tabelle für reportYear
      let einnahmenTable = '<h2>Einnahmen</h2><table><thead><tr><th>Kategorie</th><th>Betrag</th></tr></thead><tbody>';
      const einYearData = einData[reportYear] || {};
      const einCategories = Object.keys(einYearData).sort((a, b) => Number(einYearData[b] || 0) - Number(einYearData[a] || 0));
      einCategories.forEach(cat => {
        const value = Number(einYearData[cat] || 0);
        einnahmenTable += `<tr><td>${escapeHtml(cat)}</td><td>${escapeHtml(formatGermanInteger(value))} €</td></tr>`;
      });
      const einTotal = einCategories.reduce((s, cat) => s + Number(einYearData[cat] || 0), 0);
      einnahmenTable += `<tr><td><strong>Gesamt</strong></td><td><strong>${escapeHtml(formatGermanInteger(einTotal))} €</strong></td></tr></tbody></table>`;

      // Ausgaben Tabelle für reportYear
      let ausgabenTable = '<h2>Ausgaben</h2><table><thead><tr><th>Kategorie</th><th>Betrag</th></tr></thead><tbody>';
      const ausYearData = ausgabenData[reportYear] || {};
      const ausCategories = Object.keys(ausYearData).sort((a, b) => Number(ausYearData[b] || 0) - Number(ausYearData[a] || 0));
      ausCategories.forEach(cat => {
        const value = Number(ausYearData[cat] || 0);
        ausgabenTable += `<tr><td>${escapeHtml(cat)}</td><td>${escapeHtml(formatGermanInteger(value))} €</td></tr>`;
      });
      const ausTotal = ausCategories.reduce((s, cat) => s + Number(ausYearData[cat] || 0), 0);
      ausgabenTable += `<tr><td><strong>Gesamt</strong></td><td><strong>${escapeHtml(formatGermanInteger(ausTotal))} €</strong></td></tr></tbody></table>`;

      // Vermögensaufbau: Erste und letzte Spalte aus Entwicklung
      const columnsOrig = Object.keys(entwicklungRows[0] || {});
      const positionCol = 'Position';
      const yearCols = columnsOrig.filter(col => col !== positionCol && /^\d{4}$/.test(col)).sort();
      const firstYear = yearCols[0];
      const lastYear = yearCols[yearCols.length - 1];
      let vermoegenTable = '<h2>Vermögensaufbau</h2><table><thead><tr><th>Position</th><th>' + firstYear + '</th><th>' + lastYear + '</th></tr></thead><tbody>';
      entwicklungRows.forEach(row => {
        const pos = row[positionCol] || '';
        const firstVal = parseGermanNumber(row[firstYear] || '0');
        const lastVal = parseGermanNumber(row[lastYear] || '0');
        vermoegenTable += `<tr><td>${escapeHtml(pos)}</td><td>${escapeHtml(formatGermanInteger(firstVal))} €</td><td>${escapeHtml(formatGermanInteger(lastVal))} €</td></tr>`;
      });
      vermoegenTable += '</tbody></table>';

      // Anzahl Gemeindemitglieder
      const mitgliederField = '<h2>Anzahl Gemeindemitglieder</h2><input type="text" placeholder="Anzahl eintragen" style="padding: 8px; font-size: 16px; width: 200px;">';

      // Unterschriften
      const unterschriften = '<h2>Unterschriften der Gemeindeleitung</h2><p>Digital unterschrieben mit <a href="https://yousign.app" target="_blank">https://yousign.app</a></p>';

      const html = `<!doctype html><html lang="de"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>Jahresabschlussunterlagen ${reportYear}</title><style>body{font-family:Arial,sans-serif;margin:20px}table{border-collapse:collapse;width:100%;margin-bottom:20px}th,td{border:1px solid #ddd;padding:8px;text-align:left;font-size:14px}th{background:#f3f4f6}h2{margin-top:30px;border-bottom:2px solid #333;padding-bottom:5px}</style></head><body><h1>Jahresabschlussunterlagen ${reportYear}</h1>${einnahmenTable}${ausgabenTable}${vermoegenTable}${mitgliederField}${unterschriften}</body></html>`;
      const outPath = path.join(process.cwd(), 'Daten', 'result', `JahresabschlussUnterlagen_${reportYear}.html`);
      fs.writeFileSync(outPath, html, 'utf8');
      logToFile(`Jahresabschluss HTML erstellt: ${outPath}`);
    }
  } catch (e) { logToFile('Fehler beim Erzeugen Jahresabschluss-Seite: ' + (e && e.message ? e.message : String(e))); }
  
  // Create PPT only if there is meaningful data
  try {
    const hasIncome = excelResult.sortedIncomeData && Array.isArray(excelResult.sortedIncomeData.data) && excelResult.sortedIncomeData.data.length > 0;
    const hasExpenses = excelResult.sortedExpensesData && Array.isArray(excelResult.sortedExpensesData.data) && excelResult.sortedExpensesData.data.length > 0;
    const hasPie = excelResult.sortedPieData && Array.isArray(excelResult.sortedPieData) && excelResult.sortedPieData.length > 0;
    if (hasIncome || hasExpenses || hasPie) {
      await createPPT(excelResult.sortedIncomeData, excelResult.sortedExpensesData, excelResult.sortedPieData);
    } else {
      logToFile('Keine Daten für PPT vorhanden, PPT-Erzeugung übersprungen.');
    }
  } catch (e) {
    logToFile('Fehler beim Erzeugen PPT: ' + (e && e.message ? e.message : String(e)));
  }
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
