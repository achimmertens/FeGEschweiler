import QuickChart from 'quickchart-js';
import path from 'path';
import fs from 'fs';
import { ensureResultDirectory, formatGermanNumber, formatGermanInteger, COLORS } from './utils.js';

export async function saveChartAsPNG(configuration, filename, width = 1000, height = 700) {
  const chart = new QuickChart();
  // Log configuration (serialize functions as strings) to debug.log for troubleshooting
  try {
    const debugPath = path.join(process.cwd(), 'debug.log');
    function serialize(obj) {
      return JSON.stringify(obj, function(key, value) {
        if (typeof value === 'function') return value.toString();
        return value;
      }, 2);
    }
    const serialized = serialize(configuration);
    const header = `[${new Date().toISOString()}] saveChartAsPNG: ${filename} - configuration:\n`;
    fs.appendFileSync(debugPath, header + serialized + '\n\n', 'utf8');
  } catch (e) {
    // ignore logging errors
  }

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
        // Default datalabels formatter: prefer dataset label or data.__cat, fallback to chart.labels
        datalabels: {
          display: true,
          color: '#ffffff',
          font: { weight: 'bold', size: 12 },
          anchor: 'center',
          align: 'center',
          formatter: function(value, context) {
            try {
              // prefer label attached to data point
              const labelFromData = context && context.dataset && context.dataset.data && typeof context.dataIndex === 'number' && context.dataset.data[context.dataIndex]
                ? (context.dataset.data[context.dataIndex].__cat || context.dataset.data[context.dataIndex].label || '')
                : '';
              const labelFromDataset = context && context.dataset && context.dataset.label ? context.dataset.label : '';
              const idx = (typeof context.dataIndex === 'number') ? context.dataIndex : (typeof context.index === 'number' ? context.index : null);
              const labelFromLabels = idx !== null && context && context.chart && context.chart.data && context.chart.data.labels && context.chart.data.labels[idx]
                ? context.chart.data.labels[idx]
                : '';
              const label = labelFromData || labelFromDataset || labelFromLabels || '';
              const n = (typeof value === 'number' ? Math.round(value) : parseInt(value) || 0);
              const numFmt = n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
              return (label ? label + ' ' : '') + numFmt + ' €';
            } catch (e) {
              return (typeof value === 'number' ? Math.round(value) : String(value)) + ' €';
            }
          }
        },
        tooltip: { callbacks: { label: function(context) { const value = context.parsed.y; const s = (function(v){ return v.toFixed(2).replace('.',',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); })(value); return context.dataset.label + ': ' + s + ' €'; } } }
      },
      elements: { bar: { barThickness: 80, maxBarThickness: 140 } },
      scales: {
        x: { stacked: true, title: { display: false }, ticks: { font: { size: 20 } }, grid: { display: true, color: 'rgba(0, 0, 0, 0.1)' } },
        y: { stacked: true, beginAtZero: true, ticks: { stepSize: 20000, font: { size: 20 }, callback: function(value) { const n = Math.round(value); return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); } }, title: { display: false }, grid: { display: true, color: 'rgba(0, 0, 0, 0.1)' } }
      },
      responsive: false,
      maintainAspectRatio: false
    }, options)
  };
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
  let categories = Array.from(categoriesSet);
  if (categories.length === 0) return;

  // Aggregate data for each year and category
  const aggregatedData = {};
  selectedYears.forEach(year => {
    aggregatedData[year] = {};
    const entry = sortedIncomeData.find(e => String(e.year) === String(year));
    if (entry && entry.data) {
      Object.assign(aggregatedData[year], entry.data);
    }
  });

  // Combine all items and sort to determine top 10 and "Sonstiges"
  let allItems = [];
  selectedYears.forEach(year => {
    Object.entries(aggregatedData[year]).forEach(([cat, value]) => {
      allItems.push({ category: cat, value: value, year: year });
    });
  });

  // Sum values per category across selected years
  const categoryTotals = {};
  allItems.forEach(item => {
    categoryTotals[item.category] = (categoryTotals[item.category] || 0) + item.value;
  });

  // Sort categories by total value to find top 10
  const sortedCategories = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]);

  const top10Categories = sortedCategories.slice(0, 10).map(item => item[0]);
  const remainingCategories = sortedCategories.slice(10).map(item => item[0]);

  let displayCategories = top10Categories;
  let sonstigeValueTotal = 0;
  if (remainingCategories.length > 0) {
    sonstigeValueTotal = remainingCategories.reduce((sum, cat) => sum + categoryTotals[cat], 0);
    if (sonstigeValueTotal > 0) {
      displayCategories.push('Sonstiges');
    }
  }

  // Erstelle Datasets: für jede Kategorie Werte über die Jahre
  const datasets = displayCategories.map((cat, idx) => {
    const values = selectedYears.map(y => {
      const entry = sortedIncomeData.find(e => String(e.year) === String(y));
      let val = 0;
      if (cat === 'Sonstiges') {
        // Sum up values from remaining categories for this year
        let yearSonstigeValue = 0;
        remainingCategories.forEach(remCat => {
          if (entry && entry.data && entry.data[remCat]) {
            yearSonstigeValue += entry.data[remCat];
          }
        });
        val = yearSonstigeValue;
      } else {
        // Get value for the specific category
        val = entry && entry.data && entry.data[cat] ? entry.data[cat] : 0;
      }
      return { y: val, __cat: cat };
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
      anchor: 'center',
      align: 'center',
      formatter: function(value, context) {
        try {
          const label = context && context.dataset && context.dataset.label ? context.dataset.label : '';
          // show category and value on same line
          return (label ? label + ' ' : '') + (function(v){const n=Math.round(v); return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');})(value) + ' €';
        } catch (e) {
          return Math.round(value) + ' €';
        }
      }
    };
  });

  // Konfiguration: gestapeltes Balkendiagramm, gleiche Optik wie commonBarConfig
  const config = commonBarConfig(`Einnahmen ${latestYear}`, selectedYears, datasets, {
    plugins: {
      datalabels: {
        display: true,
        color: '#ffffff',
        font: { weight: 'bold', size: 14 },
        anchor: 'center',
        align: 'center',
        formatter: function(value, context) {
          try {
            const label = context && context.dataset && context.dataset.label ? context.dataset.label : '';
            return (label ? label + ' ' : '') + (function(v){const n=Math.round(v); return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');})(value) + ' €';
          } catch (e) {
            return Math.round(value) + ' €';
          }
        }
      }
    }
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
      const val = entry && entry.data && entry.data[cat] ? entry.data[cat] : 0;
      return { y: val, __cat: cat };
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
      anchor: 'center',
      align: 'center',
      formatter: function(value, context) {
        try {
          const label = context && context.dataset && context.dataset.label ? context.dataset.label : '';
          // show category and value on same line
          return (label ? label + ' ' : '') + (function(v){const n=Math.round(v); return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');})(value) + ' €';
        } catch (e) {
          return Math.round(value) + ' €';
        }
      }
    };
  });

  const config = commonBarConfig(`Ausgaben ${latestYear}`, selectedYears, datasets, {
    plugins: {
      datalabels: {
        display: true,
        color: '#ffffff',
        font: { weight: 'bold', size: 18 },
        anchor: 'center',
        align: 'center',
        formatter: function(value, context) {
          try {
            const label = context && context.dataset && context.dataset.label ? context.dataset.label : '';
              // If data point is object, use its y as value and show category and value on same line
              const dp = context && context.dataset && context.dataset.data && typeof context.dataIndex === 'number' ? context.dataset.data[context.dataIndex] : null;
              const v = dp && dp.y !== undefined ? dp.y : value;
              const n = Math.round(v);
              const numFmt = n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
              // Display category name and value
              return (label ? label + ' ' : '') + numFmt + ' €';
          } catch (e) {
            return Math.round(value) + ' €';
          }
        }
      }
    }
  });
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
