import { readProfitLossReports } from './lib/utils.js';
import { createExcelFile } from './lib/excel.js';
import { createPPT } from './lib/ppt.js';

export default async function main() {
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

if (require.main === module) {
  main().catch(err => { console.error(err); process.exit(1); });
}
