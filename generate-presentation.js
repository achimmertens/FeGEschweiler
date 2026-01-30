import { updateEntwicklungCSV } from './lib/utils.js';

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

async function createPresentation() {
  const { readProfitLossReports } = await import('./lib/utils.js');
  const { createExcelFile } = await import('./lib/excel.js');
  const { createPPT } = await import('./lib/ppt.js');

  const profitLossReports = readProfitLossReports();
  const years = Object.keys(profitLossReports).sort();
  if (years.length === 0) {
    console.log('Keine Gewinn-Verlust-Berichte gefunden.');
    return;
  }

  console.log(`Gefundene Jahre: ${years.join(', ')}`);

  // Erstelle Excel (liefert sortierte Daten für PPT zurück)
  const excelResult = await createExcelFile(profitLossReports, years);

  // Erstelle PPT aus den Ergebnissen
  await createPPT(excelResult.sortedIncomeData, excelResult.sortedExpensesData, excelResult.sortedPieData);
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
