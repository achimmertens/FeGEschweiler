import { updateEntwicklungCSV, formatGermanInteger, parseGermanNumber, readCSV } from './lib/utils.js';
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
      const parsed = parseBudgetFile(best.file);
      const outCsv = saveBudgetAsCSV(parsed, best.year);
      logToFile(`Budget CSV erstellt aus ${best.file}: ${outCsv}`);
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
    const outPath = chartsModule.generateAusgabenJsonFromSorted(excelResult.sortedExpensesData, years, currentYear);
    logToFile(`Ausgaben JSON erstellt: ${outPath}`);
    // Insert Schuldenabbau into the generated Ausgaben JSON using Bilanzberichte
    try {
      const { readBalanceReports, extractAccountBalances } = await import('./lib/utils.js');
      const balanceReports = readBalanceReports();
      const jsonPath = outPath;
      if (jsonPath && fs.existsSync(jsonPath)) {
        const j = JSON.parse(fs.readFileSync(jsonPath, 'utf8')) || {};
        const yrs = Object.keys(j).sort();
        let prevSum = null;
        for (const y of yrs) {
          const bal = extractAccountBalances(balanceReports[y] || []);
          let dsum = 0;
          Object.entries(bal || {}).forEach(([k,v]) => {
            const key = String(k || '').toLowerCase();
            if (key.includes('darleh') || key.includes('privatdarlehen') || key.includes('privat')) dsum += Math.abs(Number(v) || 0);
          });
          if (prevSum === null) {
            j[y]['Schuldenabbau'] = 0;
          } else {
            j[y]['Schuldenabbau'] = Number((dsum - prevSum).toFixed(2));
          }
          prevSum = dsum;
        }
        fs.writeFileSync(jsonPath, JSON.stringify(j, null, 2), 'utf8');
        logToFile(`Schuldenabbau in ${jsonPath} eingetragen`);
      }
    } catch (e) { logToFile('Fehler beim Eintragen Schuldenabbau: ' + e.message); }
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
