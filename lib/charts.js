import QuickChart from 'quickchart-js';
import path from 'path';
import { ensureResultDirectory, formatGermanNumber, formatGermanInteger } from './utils.js';

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
