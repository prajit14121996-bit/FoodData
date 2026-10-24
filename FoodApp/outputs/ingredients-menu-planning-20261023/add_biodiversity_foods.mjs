import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const path = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const old = wb.worksheets.getItem('Global Ingredients').getRange('A4:J997').values;
const known = new Set(old.map(r => String(r[1] || '').trim().toLowerCase()));
const andes = 'https://www.fao.org/giahs/giahs-around-the-world/peru-andean-agriculture/20th-anniversary-celebrations-of-the-globally-important-agricultural-heritage-systems-(giahs)-programme/en';
const africa = 'https://www.fao.org/newsroom/detail/new-compendium-identifies-100-promising-forgotten-foods-for-africa/';
const pohnpei = 'https://www.fao.org/newsroom/story/Biodiversity-reveals-the-treasures-all-around-us/en';
const india = 'https://www.fao.org/4/xii/0186-a1.htm';
const items = [
['Karat banana','Banana cultivar','Pohnpei, Federated States of Micronesia','Yellow-fleshed local banana; food cultivar','Sweet when ripe; local cultivar profile should be recorded','Document ripe stage and local preparation before use',pohnpei],
['Giant swamp taro, yellow-fleshed local variety','Taro cultivar','Pohnpei, Federated States of Micronesia','Yellow-fleshed giant swamp taro; local food variety','Taste varies by variety; local sensory profile needed','Follow locally established preparation',pohnpei],
['Oca','Andean tuber crop','Peruvian Andes','Tuber','Taste varies by cultivar and preparation; document locally','Cook according to local practice',andes],
['Olluco','Andean tuber crop','Peruvian Andes','Tuber','Taste varies by cultivar and preparation; document locally','Cook according to local practice',andes],
['Mashua','Andean tuber crop','Peruvian Andes','Tuber','Taste varies by cultivar and preparation; document locally','Cook according to local practice',andes],
['Kañiwa / cañihua','Andean grain','Peruvian and Bolivian highlands','Seed grain','Nutty notes are common; cultivar-specific profile needed','Rinse and cook as grain',andes],
['Tarwi / Andean lupine','Andean pulse','Andean highlands','Pulse seed','Nutty, mildly bitter before processing; local types vary','Requires appropriate traditional processing before eating',andes],
['Bambara groundnut','African pulse','West and Central Africa','Underground pulse seed','Earthy, nutty; local cultivar profile varies','Cook thoroughly; dried seeds often need soaking',africa],
['Fonio','African small-grain cereal','West Africa','Small cereal grain','Mild, lightly nutty; varieties vary','Rinse and cook; grain size and water ratio vary',africa],
['Baobab fruit','Wild / tree fruit','Africa; especially Sahel regions','Fruit pulp; locally used leaves also documented','Tart, citrus-like pulp; tree and processing vary','Separate pulp from seeds and fibers',africa],
['Bush mango / ogbono','Tree fruit and seed','West and Central Africa','Fruit pulp and seed/kernel are used in different ways','Fruit and kernel have distinct profiles; document separately','Identify edible part and local processing before use',africa],
['Amaranth, grain forms','Pseudocereal / seed crop','Africa; also cultivated elsewhere','Seed grain; leafy forms are separate foods','Mild, earthy; grain and leaf forms differ','Use grain and leafy forms as separate records',africa],
['Taro, regional forms','Aroid tuber crop','Africa; Pacific; Asia','Corm and cormels; forms differ by species and variety','Mild, starchy; texture varies by type','Cook thoroughly; species and preparation matter',africa],
['Tiger nut tuber','Tuber crop','West African Sahel; cultivated elsewhere','Edible tuber','Sweet, nutty','Eat fresh or dried; processing varies by use',pohnpei],
['Red sorrel leaves / roselle greens','Leaf vegetable','West African Sahel','Leaves; calyces are a separate food form','Tart, green; leaf and calyx profiles differ','Record leaves separately from dried calyces',pohnpei],
['Kapok leaves','Wild / tree leaf vegetable','West African Sahel','Leaves used as a locally gathered food','Taste not specified in reference; local profile needed','Use only locally identified edible species and preparation',pohnpei],
['Jute mallow leaves','Leaf vegetable','East Africa; Kenya','Leaves; used as a green vegetable','Mild to earthy; local cultivar and cooking style vary','Cook according to local practice', 'https://www.fao.org/4/i3022e/i3022e.pdf'],
['Mito leaves','Leaf vegetable','Uganda; East Africa','Crotalaria brevidens leaves','Taste not specified in reference; local profile needed','Use locally established identification and preparation','https://www.fao.org/4/i3022e/i3022e.pdf'],
['Tamarillo / tree tomato','Fruit','Uganda; Andean origin and wider cultivation','Fruit','Tart-sweet; cultivar and ripeness affect profile','Use ripe fruit; local varieties differ','https://www.fao.org/4/i3022e/i3022e.pdf'],
['Bel / bael fruit','Tree fruit','South Asia; documented in India','Fruit','Aromatic, sweet-tart when ripe; maturity matters','Reference reports fruit roasted and eaten; preparations vary',india],
['Suran / elephant foot yam','Tuber crop','Central and South India','Corm / tuber; edible part and preparation vary locally','Earthy, starchy; local preparation affects profile','Use established local processing; do not assume raw edible',india],
['Keolar bhaji / Bauhinia purpurea','Tree vegetable','Central India','Leaves, flowers, and seeds are locally used as vegetables','Different plant parts have different profiles; document separately','Harvest and prepare the specific edible part using local practice',india],
['Sehar / Bauhinia vahlii leaves','Wild / tree leaf vegetable','Central India','Leaves used as vegetable','Taste not specified in reference; local profile needed','Use locally established identification and preparation',india],
['Dudhia aru / Dioscorea alata','Yam cultivar / local food name','India; local name reported in tribal food list','Tuber','Taste varies by species, cultivar and preparation','Cook thoroughly; use established local preparation',india],
['Kali musli / Curculigo orchioides','Wild tuber / root food','India; local food use reported','Root and tuber','Taste not specified in reference; local profile needed','Confirm food use and preparation with local source',india],
['Jambu / Xylia xylocarpa seeds','Wild seed food','India; local food use reported','Seeds used as vegetable','Taste not specified in reference; local profile needed','Confirm species and preparation with local source',india],
['Entada pursaetha seeds','Wild seed food','India; local food use reported','Seeds used as vegetable','Taste not specified in reference; local profile needed','Confirm species and processing with local source',india],
['Ratalu / air potato','Yam species','India; South Asian local food name','Dioscorea bulbifera tuber','Taste varies by type; local profile needed','Some forms require careful processing; follow authoritative local food guidance',india],
['Pattambi rice','Named rice landrace','Kerala, India','Rice cultivar / landrace','Aroma and grain qualities vary; record cultivar-specific sensory notes', 'Cook using cultivar-specific water ratio',india],
['Champara rice','Named rice landrace','Northeast India','Rice cultivar / landrace','Aroma and grain qualities vary; record cultivar-specific sensory notes','Cook using cultivar-specific water ratio',india],
['Valsana rice','Named rice landrace','Northeast India','Rice cultivar / landrace','Aroma and grain qualities vary; record cultivar-specific sensory notes','Cook using cultivar-specific water ratio',india],
];
const added = [];
for (const item of items) {
  const key = item[0].trim().toLowerCase();
  if (known.has(key)) continue;
  known.add(key);
  added.push(item);
}
if (added.length) {
  const sh = wb.worksheets.add('Biodiversity Foods');
  sh.getRange('A1:G1').merge();
  sh.getRange('A1').values = [['REGIONAL & BIODIVERSITY FOODS | SOURCED ADDITIONS']];
  sh.getRange('A2:G2').merge();
  sh.getRange('A2').values = [['A sourced batch of regional foods, wild edible species, and named landraces. Taste details are marked as needing local documentation when the cited source does not describe them. Do not use this list alone to identify or prepare wild foods.']];
  sh.getRange('A3:G3').values = [['Ingredient / variant','Family / type','Origin','Edible part / form','Taste profile','Prep / handling notes','Reference']];
  sh.getRange(`A4:G${added.length + 3}`).values = added;
  sh.tables.add(`A3:G${added.length + 3}`, true, 'BiodiversityFoodsTable');
  sh.freezePanes.freezeRows(3);
  sh.getRange('A1:G1').format.fill = '#123B43';
  sh.getRange('A1:G1').format.font = { bold: true, color: '#FFFFFF', size: 15 };
  sh.getRange('A2:G2').format.fill = '#E8F2F0';
  sh.getRange('A2:G2').format.wrapText = true;
  sh.getRange('A3:G3').format.fill = '#155E63';
  sh.getRange('A3:G3').format.font = { bold: true, color: '#FFFFFF' };
  sh.getRange(`A4:G${added.length + 3}`).format.wrapText = true;
  sh.getRange(`A4:G${added.length + 3}`).format.verticalAlignment = 'top';
  [34,27,31,40,42,58,65].forEach((width,i)=>sh.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth=width);
  sh.getRange('A1:G1').format.rowHeight = 28;
  sh.getRange('A2:G2').format.rowHeight = 48;
  sh.getRange(`A4:G${added.length + 3}`).format.rowHeight = 62;
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(path);
console.log(`Added ${added.length} sourced regional / biodiversity food records; skipped ${items.length-added.length} exact name duplicates.`);
