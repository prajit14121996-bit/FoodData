import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const path = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const links = {
  'https://www.fao.org/giahs/giahs-around-the-world/peru-andean-agriculture/20th-anniversary-celebrations-of-the-globally-important-agricultural-heritage-systems-(giahs)-programme/en': 'FAO — Andean Agriculture, Peru',
  'https://www.fao.org/newsroom/detail/new-compendium-identifies-100-promising-forgotten-foods-for-africa/': 'FAO — Forgotten Foods in Africa',
  'https://www.fao.org/newsroom/story/Biodiversity-reveals-the-treasures-all-around-us/en': 'FAO — Biodiversity Reveals Local Foods',
  'https://www.fao.org/4/i3022e/i3022e.pdf': 'FAO — Sustainable Diets and Biodiversity',
  'https://www.fao.org/4/xii/0186-a1.htm': 'FAO — Indigenous Foods and Biodiversity, India',
};
const sh = wb.worksheets.getItem('Biodiversity Foods');
const used = sh.getUsedRange();
const refs = used.values.slice(3).map(r => [links[r[6]] || r[6]]);
sh.getRange(`G4:G${refs.length + 3}`).values = refs;
sh.getRange('G:G').format.columnWidth = 43;
sh.getRange(`G4:G${refs.length + 3}`).format.wrapText = true;

const source = wb.worksheets.add('Research Sources');
const sourceRows = [
  ['FAO — Andean Agriculture, Peru',links['https://www.fao.org/giahs/giahs-around-the-world/peru-andean-agriculture/20th-anniversary-celebrations-of-the-globally-important-agricultural-heritage-systems-(giahs)-programme/en'],'Named Andean crops and regional associations: oca, olluco, mashua, kañiwa, and lupine.'],
  ['FAO — Forgotten Foods in Africa',links['https://www.fao.org/newsroom/detail/new-compendium-identifies-100-promising-forgotten-foods-for-africa/'],'Regional crops highlighted include Bambara groundnut, fonio, baobab fruit, amaranth, and taro; FAO/FARA describe a 100-food scoping compendium.'],
  ['FAO — Biodiversity Reveals Local Foods',links['https://www.fao.org/newsroom/story/Biodiversity-reveals-the-treasures-all-around-us/en'],'Pohnpei cultivar examples, Karat banana, yellow-fleshed giant swamp taro, and edible plants reported in the West African Sahel.'],
  ['FAO — Sustainable Diets and Biodiversity',links['https://www.fao.org/4/i3022e/i3022e.pdf'],'Regional plant examples reported in food-system studies, including jute mallow, Crotalaria brevidens, and tamarillo.'],
  ['FAO — Indigenous Foods and Biodiversity, India',links['https://www.fao.org/4/xii/0186-a1.htm'],'Reported named rice landraces and locally used wild fruits, leaves, seeds, roots, and tubers in India.'],
];
source.getRange('A1:C1').values = [['Source','Reference URL','Use in this catalog']];
source.getRange(`A2:C${sourceRows.length + 1}`).values = sourceRows;
source.tables.add(`A1:C${sourceRows.length + 1}`, true, 'ResearchSourcesTable');
source.freezePanes.freezeRows(1);
source.getRange(`A1:C${sourceRows.length + 1}`).format.wrapText = true;
source.getRange(`A2:C${sourceRows.length + 1}`).format.verticalAlignment = 'top';
[43,82,76].forEach((w,i)=>source.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth=w);
source.getRange(`A2:C${sourceRows.length + 1}`).format.rowHeight = 56;
const x = await SpreadsheetFile.exportXlsx(wb);
await x.save(path);
console.log(`Updated ${refs.length} short source labels and added ${sourceRows.length} research references.`);
