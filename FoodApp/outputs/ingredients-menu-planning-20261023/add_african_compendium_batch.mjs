import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const path = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const compendium = 'FAO/FARA — Compendium of Forgotten Foods in Africa';
const yamBeanRef = 'FAO — Pulses and their By-products';
const locustRef = 'FAO Feedipedia — African locust bean';
const locustNamesRef = 'FAO/WHO Codex — Fermented locust beans';
const ogiriRef = 'FAO — Traditional food species and products';

const norm = x => String(x || '').toLowerCase().replace(/[^a-z0-9]/g,'');
const globalNames = new Set(wb.worksheets.getItem('Global Ingredients').getRange('B4:B997').values.map(r => norm(r[0])));
const biodiversityNames = new Set(wb.worksheets.getItem('Biodiversity Foods').getRange('A4:A28').values.map(r => norm(r[0])));
const candidateFoods = [
  ['African yam bean','Pulse / tuber crop; Sphenostylis stenocarpa','Originated in Ethiopia; cultivated in West, Central, East and Southern Africa','Seeds and underground tubers; seed and tuber forms vary','Taste varies by seed form and preparation; local sensory profile needed','Seed colours and sizes vary; seeds require processing such as soaking, cooking, dehulling or roasting before use',yamBeanRef],
  ['African nightshade','Leaf vegetable and fruit; Solanum nigrum','Widespread across Africa; tropical and subtropical zones','Leaves and shoots cooked as vegetables; ripe berries used in jam and pies','Taste not described in cited profile; record local bitter/green notes by form','Source lists blanching, boiling, frying, and cooking leaves/shoots',compendium],
  ['African spiderflower','Leaf vegetable; Gynandropsis gynandra','Indigenous to Africa; cultivated and gathered across multiple regions','Leaves, young shoots and occasional flowers eaten as potherb, relish, stew or side','Taste not described in cited profile; local sensory profile needed','Boil or cook; leaves and shoots may be dried for storage',compendium],
  ['Sweet bitterleaf','Leaf vegetable; Vernonia hymenolepis','Wild in Cameroon, Uganda, Kenya, Tanzania and Ethiopia; cultivated in Cameroon','Tender leafy shoots','Taste not described in cited profile; local sensory profile needed','Boiled or added to soup; cut and dried leaves can garnish other dishes',compendium],
  ['Bitter leaf','Leaf vegetable; Vernonia amygdalina','Tropical Africa; grown in shaded environments','Leaves cooked as vegetables or used in soup','Bitter character is reflected in the common name; bitterness varies by preparation','Dried leaves are also used for later food preparation',compendium],
  ['Sweet berry / miracle fruit','Fruit; Synsepalum dulcificum','Tropical West Africa','Fruit pulp used as a sweetener and flavouring agent','Sweetening use documented; detailed cultivar taste notes needed','Used in palm wine, soured cornbread, beverages, beer, vinegar and pickles',compendium],
  ['White fonio','Small-grain cereal; Digitaria exilis','Cultivated from Senegal to Chad; black fonio is reported mainly in Nigeria and northern Togo and Benin','Seeds used as staple; flour, bread, couscous, porridges and popped whole grain','Taste varies by grain and preparation; local sensory profile needed','Rinse and cook as grain; flour use and fermentation alter texture',compendium],
  ['African locust bean / néré','Tree pulse and condiment crop; Parkia biglobosa','West Africa, southern Sudan and northern Uganda','Seeds fermented into seasonings; fruit pulp used in beverages','Fermented seed product is strongly pungent and savoury; pulp is sweet','Seeds are dried and fermented; fruit pulp and other tree parts have separate uses',locustRef],
];
const existingNames = new Set([...globalNames,...biodiversityNames]);
const foods = candidateFoods.filter(r => !existingNames.has(norm(r[0])));
const pkg = wb.worksheets.getItem('Packaged Products');
const pkgRows = pkg.getUsedRange().values;
const existingProducts = new Set(pkgRows.slice(3).map(r => norm(r[0])));
const condiments = [
  ['Dawadawa / daddawa','Fermented African locust bean condiment','Hausa-speaking communities; Nigeria and Ghana','Pungent, strongly savoury, fermented','Paste, balls or dried seed product; local form varies','Soup and sauce seasoning','Fermented legume; nutrition and sodium vary by product','Seeds are boiled, fermented and often dried; methods differ by locality','Legume; check packaged ingredients and cross-contact'],
  ['Iru','Fermented African locust bean condiment','Yoruba communities; southwest Nigeria; also Benin','Pungent, savoury, umami-rich','Soft paste to whole fermented seeds; style varies','Soups, stews, sauces','Fermented legume; nutrition and sodium vary by product','Fermentation methods differ among local producers','Legume; check packaged ingredients and cross-contact'],
  ['Soumbala / sumbala','Fermented African locust bean condiment','Mali, Burkina Faso and Côte d’Ivoire','Strongly pungent, savoury, aromatic','Paste, cakes or dried forms','Sauces and soups','Fermented legume; nutrition and sodium vary by product','Seed fermentation and drying vary by producer','Legume; check packaged ingredients and cross-contact'],
  ['Nététou / netetu','Fermented African locust bean condiment','Senegal','Strongly pungent, savoury, aromatic','Fermented seed condiment; local texture varies','Sauces and soups','Fermented legume; nutrition and sodium vary by product','Traditional fermentation; forms vary by producer','Legume; check packaged ingredients and cross-contact'],
  ['Afitin / sonru','Fermented African locust bean condiment','Benin; Fon, Yoruba and Bariba production traditions','Strongly pungent, savoury; fermentation profile varies','Fermented seed condiment; local texture varies','Soups and sauces','Fermented legume; nutrition and sodium vary by product','Traditional fermentation technologies differ among localities','Legume; check packaged ingredients and cross-contact'],
  ['Ogiri ugba','Fermented African locust bean condiment','Igbo communities; Nigeria','Fermented, savoury, pungent','Fermented product; local form varies','Traditional soup seasoning','Fermented legume; nutrition and sodium vary by product','Fermentation method and final form vary locally','Legume; check packaged ingredients and cross-contact'],
].filter(r => !existingProducts.has(norm(r[0])));
const productStart = pkgRows.length + 1;
if (condiments.length) {
  pkg.getRange(`A${productStart}:I${productStart+condiments.length-1}`).values = condiments;
  pkg.getRange(`A${productStart}:I${productStart+condiments.length-1}`).format.wrapText = true;
  pkg.getRange(`A${productStart}:I${productStart+condiments.length-1}`).format.verticalAlignment = 'top';
  pkg.getRange(`A${productStart}:I${productStart+condiments.length-1}`).format.rowHeight = 58;
  pkg.tables.add(`A3:I${productStart+condiments.length-1}`, true, 'PackagedProductsTable');
}

if (foods.length) {
  const sh = wb.worksheets.add('African Food Diversity');
  sh.getRange('A1:G1').merge();
  sh.getRange('A1').values = [['AFRICAN FOOD DIVERSITY | SOURCE-BASED CROP PROFILES']];
  sh.getRange('A2:G2').merge();
  sh.getRange('A2').values = [['Records are drawn from FAO/FARA profiles. Regional sensory detail is not always documented in the cited profile; such gaps are labeled instead of inferred. Use local food knowledge for species identification and preparation, especially for wild foods.']];
  sh.getRange('A3:G3').values = [['Ingredient / variant','Family / type','Origin / regional range','Edible part / form','Taste profile','Prep / handling notes','Reference']];
  sh.getRange(`A4:G${foods.length+3}`).values = foods;
  sh.tables.add(`A3:G${foods.length+3}`, true, 'AfricanFoodDiversityTable');
  sh.freezePanes.freezeRows(3);
  sh.getRange('A1:G1').format.fill = '#123B43';
  sh.getRange('A1:G1').format.font = { bold: true, color: '#FFFFFF', size: 15 };
  sh.getRange('A2:G2').format.fill = '#E8F2F0';
  sh.getRange('A2:G2').format.wrapText = true;
  sh.getRange('A3:G3').format.fill = '#155E63';
  sh.getRange('A3:G3').format.font = { bold: true, color: '#FFFFFF' };
  sh.getRange(`A4:G${foods.length+3}`).format.wrapText = true;
  sh.getRange(`A4:G${foods.length+3}`).format.verticalAlignment = 'top';
  [32,38,46,47,48,56,43].forEach((w,i)=>sh.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth=w);
  sh.getRange('A1:G1').format.rowHeight = 28;
  sh.getRange('A2:G2').format.rowHeight = 48;
  sh.getRange(`A4:G${foods.length+3}`).format.rowHeight = 64;
}

const sourceRefs = [
  ['FAO/FARA — Compendium of Forgotten Foods in Africa','https://openknowledge.fao.org/3/cc5044en/cc5044en.pdf','African nightshade, African spiderflower, sweet bitterleaf, bitter leaf, sweet berry and white fonio profiles.'],
  ['FAO — Pulses and their By-products','https://openknowledge.fao.org/3/i7779en/i7779en.pdf','African yam bean distribution, edible seed/tuber forms, seed-colour variation and processing.'],
  ['FAO Feedipedia — African locust bean','https://feedipedia.review.fao.org/node/268','African locust bean edible parts, fermented condiment names and strong pungent flavour.'],
  ['FAO/WHO Codex — Regional standard proposal for fermented locust beans','https://www.fao.org/fao-who-codexalimentarius/sh-proxy/zh/?lnk=1&url=https%253A%252F%252Fworkspace.fao.org%252Fsites%252Fcodex%252FMeetings%252FCX-707-25%252FCRDs%252Fca25_CRD03.pdf','Dawadawa, iru, ogiri and other locust-bean condiment names; preparation and flavour context.'],
  ['FAO — Traditional food species and products','https://www.fao.org/4/i0370e/i0370e.pdf','Ogiri ugba and Parkia biglobosa as a fermented traditional soup seasoning.'],
  ['Fermented locust-bean styles across West Africa','https://www.sciencedirect.com/science/article/pii/S0168160505005611','Regional name and production differences for dawadawa/iru, soumbala, netetu and afitin/sonru.'],
];
const refs = wb.worksheets.add('African Food Sources');
refs.getRange('A1:C1').values = [['Source','Reference URL','Supports these entries']];
refs.getRange(`A2:C${sourceRefs.length+1}`).values = sourceRefs;
refs.tables.add(`A1:C${sourceRefs.length+1}`, true, 'AfricanFoodSourcesTable');
refs.freezePanes.freezeRows(1);
refs.getRange(`A1:C${sourceRefs.length+1}`).format.wrapText = true;
refs.getRange(`A2:C${sourceRefs.length+1}`).format.verticalAlignment = 'top';
[55,88,82].forEach((w,i)=>refs.getRange(`${String.fromCharCode(65+i)}:${String.fromCharCode(65+i)}`).format.columnWidth=w);
refs.getRange(`A2:C${sourceRefs.length+1}`).format.rowHeight = 54;

const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(path);
console.log(`Added ${foods.length} African crop profiles and ${condiments.length} regional Parkia condiment forms; added a product filter and ${sourceRefs.length} source references.`);
