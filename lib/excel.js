import ExcelJS from 'exceljs';
import { formatGermanNumber, formatGermanInteger, extractIncome, extractExpenses, ensureResultDirectory, readCSV } from './utils.js';

// Minimal, functional excel creator that returns structured data for PPT
export async function createExcelFile(profitLossReports, years) {
   ensureResultDirectory();
   const wb = new ExcelJS.Workbook();
   const sheet = wb.addWorksheet('Übersicht');
   sheet.columns = [{header:'Position', key:'pos', width:40}, {header:'Wert', key:'val', width:20}];

   // summarize income/expenses per year
   const sortedIncomeData = [];
   const sortedExpensesData = [];
   const sortedPieData = [];

   years.forEach(year => {
      const report = profitLossReports[year] || [];
      const inc = extractIncome(report);
      const exp = extractExpenses(report);
      sortedIncomeData.push({ year, data: inc });
      sortedExpensesData.push({ year, data: exp });
      // pie: top 6 expenses
      const pie = Object.entries(exp).sort((a,b)=>b[1]-a[1]).slice(0,6).map(x=>({label:x[0], value:x[1]}));
      sortedPieData.push({ year, data: pie });
   });

   // write a tiny summary to Excel
   sheet.addRow(['Jahre', years.join(', ')]);
   const outPath = `${process.cwd()}/Daten/result/Budgets_Übersicht.xlsx`;
   await wb.xlsx.writeFile(outPath);
   console.log('Excel gespeichert:', outPath);
   return { sortedIncomeData, sortedExpensesData, sortedPieData };
}

