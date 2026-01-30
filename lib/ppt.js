import PptxGenJS from 'pptxgenjs';
import { COLORS, ensureResultDirectory } from './utils.js';

export function createPPT(sortedIncomeData, sortedExpensesData, sortedPieData) {
  const pptx = new PptxGenJS();
  // Minimal example: create title slide
  const slide = pptx.addSlide();
  slide.addText('Finanzübersicht', { x: 1, y: 1, fontSize: 36, color: '#' + COLORS.primary });
  ensureResultDirectory();
  const output = `${process.cwd()}/Daten/result/Präsentation.pptx`;
  pptx.writeFile(output).then(() => console.log(`PPTX gespeichert: ${output}`));
}
