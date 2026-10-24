import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const path = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g,'');
const known = new Set();
for (const [sheetName, col, firstDataRow] of [['Global Ingredients',1,4],['Biodiversity Foods',0,4],['African Food Diversity',0,4]]) {
  const sheet = wb.worksheets.getItem(sheetName);
  for (const r of sheet.getUsedRange().values.slice(firstDataRow-1)) known.add(norm(r[col]));
}
const comp = 'FAO/FARA — Compendium of Forgotten Foods in Africa';
const gnetum = 'FAO — Gnetum domestication and edible leaves';
const enset = 'FAO Feedipedia — Enset corms and pseudostems';
const ensetReview = 'Enset food forms — kocho, bulla and amicho';
const items = [
  ['Bologi / worow','Leafy vegetable; Crassocephalum biafrae','West Africa; moist lowland tropics','Edible vines and leaves','Spinach-like','Leaves are steamed; also used as a tea substitute; juice, wine, pickles and jam are reported',comp],
  ['African winged bean','Pulse and leafy vegetable; Psophocarpus scandens','Tropical Africa and Asia; lowland and riverine habitats','Leaves, pods, sprouts and seeds are edible','Taste not described in source; local sensory notes needed','Leaves and young sprouts are cooked as potherb; seeds may be roasted and milled to flour',comp],
  ['Country potato / frafra potato','Tuber crop; Solenostemon rotundifolius','Tropical African lowlands','Tuber','Taste not described in source; record local cultivar notes','Tubers are boiled, baked or fried; dried tubers can be milled into flour for chips and baked foods',comp],
  ['Egusi melon','Seed crop; Citrullus colocynthis','West Africa; cultivated in semi-arid tropical areas and elsewhere','Seeds; fruit is not grown for sweet red flesh','Taste not described in source; document seed roast and cultivar','Seeds are ground as soup/stew thickener, made into dumplings or eaten whole/popped; seed oil is also used',comp],
  ['African star apple / agbalumo','Tropical fruit tree; Chrysophyllum albidum','Tropical Africa; Yoruba name agbalumo is documented in Nigeria','Fruit pulp','Taste not specified in cited profile; local ripeness notes needed','Eaten fresh, in fruit salad or dessert; juice and jam are reported',comp],
  ['Eru / okazi / ukasi / afang (Gnetum africanum)','Forest leaf vegetable; Gnetum africanum','Central and West Africa, including Nigeria, Cameroon, Gabon, Congo, DRC and Angola','Leaves','Described as a valued/tasty vegetable; standardized taste descriptors are not given','Finely shred into soups and stews; some communities also eat leaves raw in salads',gnetum],
  ['Okok / koko (Gnetum buchholzianum)','Forest leaf vegetable; Gnetum buchholzianum','Central Africa; local naming overlaps with Gnetum africanum and varies by place','Leaves','Species-specific taste notes not reported; market forms may be grouped with Gnetum africanum','Leaves are used as vegetables; verify species locally because the two Gnetum species are similar',gnetum],
  ['Enset / false banana','Starch crop; Ensete ventricosum','Southern and southwestern Ethiopia','Corm and pseudostem; processed foods include kocho, bulla and amicho','Taste and texture depend on product and processing','Corm and pseudostem are scraped and processed; forms and fermentation periods differ',enset],
  ['Kocho','Fermented enset food; Ensete ventricosum','Central, southern and southwestern Ethiopia','Fermented mixture of scraped pseudostem and corm','Fermented flavour; sourness and aroma vary with landrace, processing and fermentation time','Fermented staple or costaple; often prepared as a flatbread',ensetReview],
  ['Bulla','Enset starch product; Ensete ventricosum','Southern Ethiopia','Starch-rich liquid separated from processed enset pulp, then concentrated/dried','Starchy; precise sensory profile varies and is not standardized in cited sources','Used in foods such as porridge, pancake or dumpling; product definitions vary across references',ensetReview],
  ['Amicho','Cooked enset corm; Ensete ventricosum','Southern Ethiopia','Boiled inner corm, often from younger plants','Reported to taste similar to cooked potato','Boil and eat as a root/tuber food; locally selected landraces are used',ensetReview],
];
const added = items.filter(r => !known.has(norm(r[0])));
const sh = wb.worksheets.add('Regional Crop Forms');
sh.getRange('A1:G1').merge();
sh.getRange('A1').values = [['REGIONAL CROP FORMS | AFRICA AND ETHIOPIA']];
sh.getRange('A2:G2').merge();
sh.getRange('A2').values = [['Distinct species, edible forms and enset products from the sourced batch. Local names can overlap across species; confirm wild-food identification and preparation with local expertise. Taste details are marked as unspecified where the cited sources do not provide a reliable descriptor.']];
sh.getRange('A3:G3').values = [['Ingredient / variant','Family / type','Origin / regional range','Edible part / form','Taste profile','Prep / handling notes','Reference']];
sh.getRange(`A4:G${added.length+3}`).values = added;
sh.tables.add(`A3:G${added.length+3}`, true, 'RegionalCropFormsTable');
sh.freezePanes.freezeRows(3);
sh.getRange('A1:G1').format.fill = '#123B43';
sh.getRange('A1:G1').format.font = { bold: true, color: '#FFFFFF', size: 15 };
sh.getRange('A2:G2').format.fill = '#E8F2F0';
sh.getRange('A2:G2').format.wrapText = true;
sh.getRange('A3:G3').format.fill = '#155E63';
sh.getRange('A3:G3').format.font = { bold: true, color: '#FFFFFF' };
sh.getRange(`A4:G${added.length+3}`).format.wrapText = true;
sh.getRange(`A4:G${added.length+3}`).format.verticalAlignment = 'top';
[38,42,48,47,48,60,47].forEach((w,i)=>sh.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth=w);
sh.getRange('A1:G1').format.rowHeight = 28;
sh.getRange('A2:G2').format.rowHeight = 52;
sh.getRange(`A4:G${added.length+3}`).format.rowHeight = 64;

const sources = [
  ['FAO/FARA — Compendium of Forgotten Foods in Africa','https://openknowledge.fao.org/3/cc5044en/cc5044en.pdf','Bologi, African winged bean, country potato, egusi melon and African star apple profiles.'],
  ['FAO — Gnetum domestication and edible leaves','https://www.fao.org/4/XII/0671-B5.htm','Gnetum africanum and Gnetum buchholzianum species, regional names, leaf uses and culinary selection traits.'],
  ['FAO — Gnetum foods in the Igbo food system','https://www.fao.org/4/i0370e/i0370e.pdf','Okazi and afang leaf use, including raw salad and cooked vegetable preparations.'],
  ['FAO Feedipedia — Enset corms and pseudostems','https://feedipedia.review.fao.org/node/21251','Enset food crop, edible corm/pseudostem, and kocho, bulla and amicho forms.'],
  ['Enset product characteristics and uses','https://www.sciencedirect.com/science/article/pii/S0889157508000744','Kocho fermentation, bulla starch fraction and amicho corm form; amicho taste compared with potato.'],
  ['Microbial and physicochemical dynamics of kocho','https://pmc.ncbi.nlm.nih.gov/articles/PMC10602703/','Kocho as fermented staple/costaple food and regional use in Ethiopia.'],
];
const ref = wb.worksheets.add('Regional Crop Sources');
ref.getRange('A1:C1').values = [['Source','Reference URL','Supports these entries']];
ref.getRange(`A2:C${sources.length+1}`).values = sources;
ref.tables.add(`A1:C${sources.length+1}`, true, 'RegionalCropSourcesTable');
ref.freezePanes.freezeRows(1);
ref.getRange(`A1:C${sources.length+1}`).format.wrapText = true;
ref.getRange(`A2:C${sources.length+1}`).format.verticalAlignment = 'top';
[62,85,88].forEach((w,i)=>ref.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth=w);
ref.getRange(`A2:C${sources.length+1}`).format.rowHeight=54;

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(path);
console.log(`Added ${added.length} distinct regional crop and food-form entries and ${sources.length} references.`);
