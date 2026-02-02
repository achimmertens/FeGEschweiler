import QuickChart from 'quickchart-js';
import path from 'path';
import { ensureResultDirectory, formatGermanNumber, formatGermanInteger, COLORS } from './utils.js';

export async function saveChartAsPNG(configuration, filename, width = 1000, height = 700) {
  const chart = new QuickChart();
  chart.setConfig(configuration);
  chart.setWidth(width);
  chart.setHeight(height);
  chart.setFormat('png');
  chart.setBackgroundColor('white');
  ensureResultDirectory();
  const outputPath = path.join(process.cwd(), 'Daten', 'result', filename);
  await chart.toFile(outputPath);
  console.log(`${filename} gespeichert: ${outputPath}`);
}

export function commonPieConfig(title, labels, data, colors) {
  return {
    type: 'pie',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: '#ffffff', borderWidth: 2 }] },
    options: {
      plugins: {
        title: { display: true, text: title, font: { size: 20, weight: 'bold' }, padding: { top: 10, bottom: 20 } },
        legend: { display: true, position: 'right', labels: { boxWidth: 15, padding: 8, font: { size: 10 } } },
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
}

export function commonBarConfig(title, labels, datasets, options = {}) {
  return {
    type: 'bar',
    data: { labels, datasets },
    options: Object.assign({
      indexAxis: 'x',
      plugins: {
        title: { display: true, text: title, font: { size: 40, weight: 'bold' }, padding: { top: 20, bottom: 40 } },
        legend: { display: true, position: 'right', labels: { boxWidth: 30, padding: 16, font: { size: 40 } } },
        datalabels: { display: false },
        tooltip: { callbacks: { label: function(context) { const value = context.parsed.y; return `${context.dataset.label}: ${formatGermanNumber(value)} €`; } } }
      },
      elements: { bar: { barThickness: 80, maxBarThickness: 140 } },
      scales: {
        x: { stacked: true, title: { display: false }, ticks: { font: { size: 20 } }, grid: { display: true, color: 'rgba(0, 0, 0, 0.1)' } },
        y: { stacked: true, beginAtZero: true, ticks: { stepSize: 20000, font: { size: 20 }, callback: function(value) { return formatGermanInteger(value); } }, title: { display: false }, grid: { display: true, color: 'rgba(0, 0, 0, 0.1)' } }
      },
      responsive: false,
      maintainAspectRatio: false
    }, options)
  };
}

// Erzeuge für jede Jahres-Auswertung eine PNG-Datei für Einnahmen (ein Jahr pro PNG)
export async function saveIncomePNGs(sortedIncomeData) {
  // sortedIncomeData: [{ year, data: { category: value, ... } }, ...]
  for (const entry of sortedIncomeData) {
    const year = entry.year;
    const dataObj = entry.data || {};

    // sortiere nach Betrag absteigend
    const items = Object.entries(dataObj).sort((a, b) => b[1] - a[1]);
    const labels = items.map(i => i[0]);
    const values = items.map(i => i[1]);

    // Wenn keine Daten, überspringen
    if (labels.length === 0) continue;

    // Farben pro Kategorie: Sonstiges -> grau, sonst aus Palette
    const colors = labels.map((lbl, idx) => (lbl === 'Sonstiges' ? '#A9A9A9' : `#${COLORS.chartColors[idx % COLORS.chartColors.length]}`));
    // Single dataset with per-bar datalabels showing category + value
    const datasets = [{
      label: 'Einnahmen',
      data: values,
      backgroundColor: colors,
      datalabels: {
        display: function(context) {
          try {
            const v = context && typeof context.dataIndex === 'number' ? context.dataset.data[context.dataIndex] : (context && context.parsed ? context.parsed : 0);
            return Math.abs(v || 0) > 0.5;
          } catch (e) {
            return false;
          }
        },
        color: '#ffffff',
        font: { weight: 'bold', size: 12 },
        formatter: function(value, context) {
          try {
            const label = context && context.chart && context.chart.data && context.chart.data.labels && typeof context.dataIndex === 'number'
              ? context.chart.data.labels[context.dataIndex]
              : '';
            return label + ' ' + formatGermanInteger(value) + ' €';
          } catch (e) {
            return Math.round(value) + ' €';
          }
        }
      }
    }];

    const config = commonBarConfig(`Einnahmen ${year}`, labels, datasets, {
      // gleiche Optik/Skalierung wie für Ausgaben
      options: {
        scales: {
          y: { beginAtZero: true }
        }
      }
    });

    const filename = `Einnahmen_${year}.png`;
    await saveChartAsPNG(config, filename);
  }
}

// Erzeuge ein gestapeltes Balkendiagramm über die letzten drei Jahre (analog zu Ausgaben_YYYY.png)
export async function saveIncomeStackedPNG(sortedIncomeData) {
  // sortedIncomeData: [{ year, data: { category: value, ... } }, ...]
  const yearsAll = (sortedIncomeData || []).map(e => String(e.year)).sort();
  if (yearsAll.length === 0) return;

  // Wähle die letzten drei Jahre (oder alle, falls weniger)
  const selectedYears = yearsAll.slice(-3);
  const latestYear = selectedYears[selectedYears.length - 1];

  // Sammle alle Kategorien über die ausgewählten Jahre
  const categoriesSet = new Set();
  sortedIncomeData.forEach(entry => {
    if (!selectedYears.includes(String(entry.year))) return;
    Object.keys(entry.data || {}).forEach(k => categoriesSet.add(k));
  });
  const categories = Array.from(categoriesSet);
  if (categories.length === 0) return;

  // Erstelle Datasets: für jede Kategorie Werte über die Jahre
  const datasets = categories.map((cat, idx) => {
    const values = selectedYears.map(y => {
      const entry = sortedIncomeData.find(e => String(e.year) === String(y));
      return entry && entry.data && entry.data[cat] ? entry.data[cat] : 0;
    });
    return {
      label: cat,
      data: values,
      backgroundColor: (cat === 'Sonstiges' ? '#A9A9A9' : `#${COLORS.chartColors[idx % COLORS.chartColors.length]}`),
      borderColor: '#ffffff',
      borderWidth: 1,
      stack: 'stack1',
      barPercentage: 1.0,
      categoryPercentage: 1.0
    };
  });
  // Add datalabels per dataset for stacked income
  datasets.forEach(ds => {
    ds.datalabels = {
      display: function(context) {
        try {
          const v = context && context.dataset && context.dataset.data && typeof context.dataIndex === 'number'
            ? context.dataset.data[context.dataIndex]
            : (context && context.parsed ? context.parsed : 0);
          return Math.abs(v || 0) > 0.5;
        } catch (e) {
          return false;
        }
      },
      color: '#ffffff',
      font: { weight: 'bold', size: 14 },
      formatter: function(value, context) {
        try {
          const label = context && context.dataset && context.dataset.label ? context.dataset.label : '';
          return label + ' ' + formatGermanInteger(value) + ' €';
        } catch (e) {
          return Math.round(value) + ' €';
        }
      }
    };
  });

  // Konfiguration: gestapeltes Balkendiagramm, gleiche Optik wie commonBarConfig
  const config = commonBarConfig(`Einnahmen ${latestYear}`, selectedYears, datasets, {
    // keine zusätzlichen Optionen nötig; commonBarConfig nutzt stacked scales
  });

  const filename = `Einnahmen_${latestYear}.png`;
  await saveChartAsPNG(config, filename);
}

// Wiederherstellung: Erzeuge ein gestapeltes Balkendiagramm für Ausgaben wie früher
export async function saveExpensesStackedPNG(sortedExpensesData) {
  // sortedExpensesData: [{ year, data: { category: value, ... } }, ...]
  const yearsAll = (sortedExpensesData || []).map(e => String(e.year)).sort();
  if (yearsAll.length === 0) return;

  const selectedYears = yearsAll.slice(-3);
  const latestYear = selectedYears[selectedYears.length - 1];

  const categoriesSet = new Set();
  sortedExpensesData.forEach(entry => {
    if (!selectedYears.includes(String(entry.year))) return;
    Object.keys(entry.data || {}).forEach(k => categoriesSet.add(k));
  });
  const categories = Array.from(categoriesSet);
  if (categories.length === 0) return;

  const datasets = categories.map((cat, idx) => {
    const values = selectedYears.map(y => {
      const entry = sortedExpensesData.find(e => String(e.year) === String(y));
      return entry && entry.data && entry.data[cat] ? entry.data[cat] : 0;
    });
    return {
      label: cat,
      data: values,
      backgroundColor: (cat === 'Sonstiges' ? '#A9A9A9' : `#${COLORS.chartColors[idx % COLORS.chartColors.length]}`),
      borderColor: '#ffffff',
      borderWidth: 1,
      stack: 'stack1',
      barPercentage: 1.0,
      categoryPercentage: 1.0
    };
  });
  // Add datalabels per dataset for stacked expenses (show label+value)
  datasets.forEach(ds => {
    ds.datalabels = {
      display: function(context) {
        try {
          const v = context && context.dataset && context.dataset.data && typeof context.dataIndex === 'number'
            ? context.dataset.data[context.dataIndex]
            : (context && context.parsed ? context.parsed : 0);
          return Math.abs(v || 0) > 0.5;
        } catch (e) {
          return false;
        }
      },
      color: '#ffffff',
      font: { weight: 'bold', size: 18 },
      formatter: function(value, context) {
        try {
          const label = context && context.dataset && context.dataset.label ? context.dataset.label : '';
          return label + ' ' + formatGermanInteger(value) + ' €';
        } catch (e) {
          return Math.round(value) + ' €';
        }
      }
    };
  });

  const config = commonBarConfig(`Ausgaben ${latestYear}`, selectedYears, datasets, {});
  const filename = `Ausgaben_${latestYear}.png`;
  await saveChartAsPNG(config, filename);
}

/**
 * Erstellung aller relevanten PNGs (Einnahmen per Jahr, gestapelte Einnahmen, gestapelte Ausgaben, Ausgaben-Pie)
 * Erwartet die neuen Excel-Objekte: { categories, data: [{ category, '2023': val, ... }], totals }
 */
export async function createCharts(sortedIncomeObj, sortedExpensesObj, sortedPieData, years) {
  // Transform income object (rows) into array [{ year, data: { cat: val } }, ...]
  const incomeByYear = (years || []).map(year => {
    const data = {};
    (sortedIncomeObj && sortedIncomeObj.data || []).forEach(row => {
      data[row.category] = row[year] || 0;
    });
    return { year: String(year), data };
  });

  // Transform expenses object similarly
  const expensesByYear = (years || []).map(year => {
    const data = {};
    (sortedExpensesObj && sortedExpensesObj.data || []).forEach(row => {
      data[row.category] = row[year] || 0;
    });
    return { year: String(year), data };
  });

  // Save per-year income PNGs (one file per year)
  try {
    await saveIncomePNGs(incomeByYear);
  } catch (e) {
    console.error('Fehler beim Speichern der Einnahmen-PNGs:', e);
  }

  // Save stacked income (last 3 years)
  try {
    await saveIncomeStackedPNG(incomeByYear);
  } catch (e) {
    console.error('Fehler beim Speichern des gestapelten Einnahmen-PNGs:', e);
  }

  // Save stacked expenses (last 3 years)
  try {
    await saveExpensesStackedPNG(expensesByYear);
  } catch (e) {
    console.error('Fehler beim Speichern des gestapelten Ausgaben-PNGs:', e);
  }

  // Save expenses pie for latest year using sortedPieData if available
  try {
    if (sortedPieData && sortedPieData.length > 0) {
      const labels = sortedPieData.map(p => p.name);
      const data = sortedPieData.map(p => p.value);
      const colors = labels.map((lbl, idx) => (lbl === 'Sonstiges' ? '#A9A9A9' : `#${COLORS.chartColors[idx % COLORS.chartColors.length]}`));
      const cfg = commonPieConfig(`Ausgaben nach Kategorien (${years && years.length ? years[years.length - 1] : ''})`, labels, data, colors);
      await saveChartAsPNG(cfg, `Ausgaben_Kuchendiagramm_${years && years.length ? years[years.length - 1] : ''}.png`, 1000, 1000);
    }
  } catch (e) {
    console.error('Fehler beim Speichern des Ausgaben-Pie-PNGs:', e);
  }
}
