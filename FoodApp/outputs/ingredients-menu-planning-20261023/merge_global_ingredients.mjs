import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const dir = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023';
const outputPath = `${dir}/Chips_Condiments_Sauces_Bottled_Products.xlsx`;
const sourcePath = `${dir}/Ingredients_Menu_Planning.xlsx`;
const outWb = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const sourceWb = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const global = outWb.worksheets.add('Global Ingredients');
const cols = ['Library section','Ingredient / variant','Family / type','Origin & cuisine associations','Taste profile','Texture / behavior','Dish & menu uses','Nutrition highlights','Prep / handling notes','Diet / allergen notes'];
const sections = ['Vegetables','Fruit & Citrus','Grains & Pulses','Spices & Herbs','Protein & Dairy','Pantry & Condiments','Origin Variants','Specialty Foods'];
const rows = [];
for (const section of sections) {
  const values = sourceWb.worksheets.getItem(section).getUsedRange().values;
  if (section === 'Origin Variants') {
    for (const r of values.slice(1)) if (r[0]) rows.push([section,r[0],'Named regional / cultivar variant',r[1],r[2],'',r[3],'',r[4],'']);
  } else {
    for (const r of values.slice(1)) if (r[0]) rows.push([section,...r]);
  }
}
global.getRange('A1:J1').merge();
global.getRange('A1').values = [['GLOBAL FOOD INGREDIENT LIBRARY | CATEGORY, ORIGIN & TASTE VARIANTS']];
global.getRange('A2:J2').merge();
global.getRange('A2').values = [[`Expanded with all ${rows.length} source ingredient and named-variant records from the current ingredient library. The separate Packaged Products sheet holds snack chips and additional packaged styles. This is a growing reference, not a complete inventory of every edible organism, local cultivar, preparation, or commercial product on Earth. Global foods include under-documented and locally named species; use specialist/local sources to add them carefully.`]];
global.getRange('A3:J3').values = [cols];
global.getRange(`A4:J${rows.length + 3}`).values = rows;
global.tables.add(`A3:J${rows.length + 3}`, true, 'GlobalIngredientsTable');
global.freezePanes.freezeRows(3);
global.getRange('A1:J1').format.fill = '#123B43';
global.getRange('A1:J1').format.font = { bold: true, color: '#FFFFFF', size: 16 };
global.getRange('A2:J2').format.fill = '#E8F2F0';
global.getRange('A2:J2').format.wrapText = true;
global.getRange('A3:J3').format.fill = '#155E63';
global.getRange('A3:J3').format.font = { bold: true, color: '#FFFFFF' };
global.getRange(`A4:J${rows.length + 3}`).format.wrapText = true;
global.getRange(`A4:J${rows.length + 3}`).format.verticalAlignment = 'top';
const widths = [23,30,27,38,31,34,40,34,42,33];
widths.forEach((width, i) => global.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth = width);
global.getRange(`A4:J${rows.length + 3}`).format.rowHeight = 54;
global.getRange('A1:J1').format.rowHeight = 30;
global.getRange('A2:J2').format.rowHeight = 58;

const source = sourceWb.worksheets.getItem('Sources & Notes').getUsedRange().values;
const refs = outWb.worksheets.add('Coverage & Sources');
refs.getRange(`A1:C${source.length}`).values = source;
refs.tables.add(`A1:C${source.length}`, true, 'CoverageSourcesTable');
refs.freezePanes.freezeRows(1);
refs.getRange(`A1:C${source.length}`).format.wrapText = true;
refs.getRange(`A2:C${source.length}`).format.verticalAlignment = 'top';
[28,70,72].forEach((width,i)=>refs.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth=width);
refs.getRange(`A2:C${source.length}`).format.rowHeight = 48;

const exported = await SpreadsheetFile.exportXlsx(outWb);
await exported.save(outputPath);
console.log(`Added Global Ingredients with ${rows.length} records across ${sections.length} library sections, plus Coverage & Sources.`);
