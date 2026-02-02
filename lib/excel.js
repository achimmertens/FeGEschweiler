import ExcelJS from 'exceljs';
import { formatGermanNumber, extractIncome, extractExpenses, ensureResultDirectory } from './utils.js';

// Create Excel file and return structured, sorted data for PPT
export async function createExcelFile(profitLossReports, years) {
   ensureResultDirectory();
   const wb = new ExcelJS.Workbook();
   const incomeSheet = wb.addWorksheet('Einnahmen');
   const expensesSheet = wb.addWorksheet('Ausgaben');

   // Collect all categories across years
   const allIncomeCategories = new Set();
   const allExpenseCategories = new Set();
   years.forEach(year => {
      const report = profitLossReports[year] || [];
      const inc = extractIncome(report);
      const exp = extractExpenses(report);
      Object.keys(inc).forEach(k => allIncomeCategories.add(k));
      Object.keys(exp).forEach(k => allExpenseCategories.add(k));
   });

   // Compute totals per category across all years for sorting
   const incomeTotals = {};
   Array.from(allIncomeCategories).forEach(category => {
      incomeTotals[category] = years.reduce((s, year) => {
         const inc = extractIncome(profitLossReports[year] || []);
         return s + (inc[category] || 0);
      }, 0);
   });

   const expenseTotals = {};
   Array.from(allExpenseCategories).forEach(category => {
      expenseTotals[category] = years.reduce((s, year) => {
         const exp = extractExpenses(profitLossReports[year] || []);
         return s + (exp[category] || 0);
      }, 0);
   });

   // Sort categories by total (desc)
   const sortedIncomeCategories = Array.from(allIncomeCategories).sort((a, b) => (incomeTotals[b] || 0) - (incomeTotals[a] || 0));
   const sortedExpenseCategories = Array.from(allExpenseCategories).sort((a, b) => (expenseTotals[b] || 0) - (expenseTotals[a] || 0));

   // Keep top 10 categories, aggregate rest into 'Sonstiges'
   const TOP_N = 10;
   function buildTopAndAggregate(sortedCats, totalsMap, extractFn) {
      const top = sortedCats.slice(0, TOP_N);
      const others = sortedCats.slice(TOP_N);
      const dataRows = top.map(category => {
         const row = { category };
         years.forEach(year => {
            const values = extractFn(year) || {};
            row[year] = values[category] || 0;
         });
         return row;
      });

      if (others.length > 0) {
         const otherRow = { category: 'Sonstiges' };
         years.forEach(year => {
            otherRow[year] = others.reduce((s, c) => {
               const values = extractFn(year) || {};
               return s + (values[c] || 0);
            }, 0);
         });
         // Only include Sonstiges when non-zero
         const anyNonZero = years.some(y => Math.abs(otherRow[y] || 0) > 1e-8);
         if (anyNonZero) dataRows.push(otherRow);
      }

      // Totals row
      const totalsRow = { category: 'Gesamt' };
      years.forEach(year => {
         const values = extractFn(year) || {};
         totalsRow[year] = Object.values(values).reduce((s, v) => s + (v || 0), 0);
      });

      const categories = dataRows.map(r => r.category);
      return { categories, data: dataRows, totals: totalsRow };
   }

   const sortedIncomeData = buildTopAndAggregate(sortedIncomeCategories, incomeTotals, (year) => extractIncome(profitLossReports[year] || []));
   const sortedExpensesData = buildTopAndAggregate(sortedExpenseCategories, expenseTotals, (year) => extractExpenses(profitLossReports[year] || []));

   // Pie data: use latest year and top categories for pie chart (sorted by value desc)
   const latestYear = years[years.length - 1];
   const latestExpenses = extractExpenses(profitLossReports[latestYear] || []);
   const pieEntries = Object.entries(latestExpenses).sort((a, b) => b[1] - a[1]);
   const sortedPieData = pieEntries.map(([name, value]) => ({ name, value }));

   // Write a minimal excel summary (keeps compatibility)
   incomeSheet.columns = [{ header: 'Kategorie', key: 'category', width: 40 }, ...years.map(y => ({ header: y, key: y, width: 15 }))];
   sortedIncomeData.data.forEach(row => {
      const r = [row.category];
      years.forEach(y => r.push(row[y] || 0));
      incomeSheet.addRow(r);
   });

   expensesSheet.columns = [{ header: 'Kategorie', key: 'category', width: 40 }, ...years.map(y => ({ header: y, key: y, width: 15 }))];
   sortedExpensesData.data.forEach(row => {
      const r = [row.category];
      years.forEach(y => r.push(row[y] || 0));
      expensesSheet.addRow(r);
   });

   const outPath = `${process.cwd()}/Daten/result/Finanzlage_FeG_Eschweiler.xlsx`;
   await wb.xlsx.writeFile(outPath);
   console.log('Excel gespeichert:', outPath);

   return { sortedIncomeData, sortedExpensesData, sortedPieData };
}

