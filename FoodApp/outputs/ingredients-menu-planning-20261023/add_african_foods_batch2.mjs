import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const path = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g,'');
const known = new Set();
for (const [name,col] of [['Global Ingredients',1],['Biodiversity Foods',0],['African Food Diversity',0],['Regional Crop Forms',0]]) {
  const sh=wb.worksheets.getItem(name);
  for (const row of sh.getUsedRange().values.slice(name==='Global Ingredients'?3:3)) if(row[col]) known.add(norm(row[col]));
}
const ref='FAO/FARA — Compendium of Forgotten Foods in Africa';
const items=[
 ['Aerial yam','Tuber crop; Dioscorea bulbifera','Tropical Africa and Asia; cultivated and gathered across varied habitats','Underground tubers and aerial bulbils','Often bitter; bitterness varies by form and processing','Edible forms require appropriate processing/cooking to reduce bitterness and antinutritional factors; cooked, dried or milled forms are reported',ref],
 ['Bitter kola','Forest tree nut; Garcinia kola','West and Central Africa','Seeds / kernels','Bitter (common-name descriptor; intensity and use vary locally)','Chewed fresh or used in local food preparations; cited profile gives limited culinary processing detail',ref],
 ['Cocoyam','Tuber crop; Xanthosoma sagittifolium','Tropical America origin; widely grown in tropical Africa and elsewhere','Corms, leaves and petioles','Mild starchy corm; leafy parts are used as greens','Corms peeled and boiled, steamed, baked or fried; leaves cooked; flour and chips are reported',ref],
 ['Cowpea','Pulse and vegetable; Vigna unguiculata','African domestication; now cultivated worldwide in warm regions','Dry seeds, green pods and leaves','Mild/earthy pulse; cultivar and preparation vary','Seeds boiled, steamed or fried; pods and leaves used as vegetables; flour and other processed uses reported',ref],
 ['Dandelion','Leaf and root vegetable; Taraxacum officinale','Native across Eurasia; naturalized and used globally','Leaves, flowers and roots','Leaves can be bitter; bitterness varies with age and growing conditions','Leaves eaten raw or cooked; roots used as a coffee substitute; flowers used in food and beverage preparations',ref],
 ['Garden huckleberry','Leaf and fruit vegetable; Solanum scabrum','Tropical Africa; also cultivated elsewhere','Ripe fruit and leaves','Fruit flavour is mild and preparation-dependent; source does not standardize taste','Fruit cooked into preserves, jams or pies; leaves cooked in soups and other dishes',ref],
 ['Horned melon / kiwano','Fruit; Cucumis metuliferus','Southern and central Africa; now cultivated in other regions','Ripe fruit pulp and seeds','Ripe pulp described as refreshing, mildly sweet-tart; cultivar/ripeness affects taste','Eaten fresh or used in fruit preparations; use only correctly identified ripe fruit',ref],
 ['Jambolan / Java plum','Fruit tree; Syzygium cumini','South Asia origin; widely cultivated in tropical regions','Ripe fruit','Sweet-sour and astringent; varies by cultivar and ripeness','Eaten fresh or used for juice, jam and wine',ref],
 ['Milk thistle (compendium entry)','Leaf vegetable; Sonchus oleraceus','Widespread in Africa and other warm temperate/tropical regions','Leaves and roots','Leaf taste not specified in cited profile; record local descriptors','Leaves eaten raw or cooked like spinach; roots reported as a coffee substitute. Taxon retained as printed in compendium; common-name mapping may vary.',ref],
 ['Redflower ragleaf','Leafy vegetable; Crassocephalum crepidioides','Tropical Africa and other humid tropical regions','Leaves','Taste not specified in cited profile; local sensory notes needed','Lightly blanch then cook in soups/stews with peppers, onions, tomatoes, melon and sometimes fish or meat; groundnut-paste sauce is reported in Sierra Leone',ref],
 ['Turkey berry','Fruit vegetable; Solanum torvum','Tropical and subtropical regions, including Africa and Asia','Fruits','Taste not specified in cited profile; maturity and local preparation may affect flavour','Fruit eaten directly or cooked; used in soups and sauces; processed into sauces',ref],
 ['Water yam','Tuber crop; Dioscorea alata','Southeast Asian origin; widely cultivated in tropical regions','Tuber','Starchy; texture and flavour vary by cultivar','Boiled, mashed, fried, roasted or baked; flour is used in baked foods',ref],
 ['Wild lettuce','Leafy vegetable; Launaea taraxacifolia','West Africa','Leaves','Taste not specified in cited profile; local sensory notes needed','Leaves used as a vegetable; confirm local species and preparation practice',ref],
];
const added=items.filter(r=>!known.has(norm(r[0])));
const sh=wb.worksheets.add('African Foods Batch 2');
sh.getRange('A1:G1').merge(); sh.getRange('A1').values=[['AFRICAN FOOD DIVERSITY | BATCH 2']];
sh.getRange('A2:G2').merge(); sh.getRange('A2').values=[['Additional distinct edible species and forms from the FAO/FARA compendium. Taste descriptors are recorded only where supported or plainly qualified; verify wild-food identity and local preparation. “Milk thistle” retains the scientific name printed in the compendium and should be taxonomically checked before using as a canonical mapping.']];
sh.getRange('A3:G3').values=[['Ingredient / variant','Family / type','Origin / regional range','Edible part / form','Taste profile','Prep / handling notes','Reference']];
sh.getRange(`A4:G${added.length+3}`).values=added;
sh.tables.add(`A3:G${added.length+3}`,true,'AfricanFoodsBatch2Table'); sh.freezePanes.freezeRows(3);
sh.getRange('A1:G1').format.fill='#123B43'; sh.getRange('A1:G1').format.font={bold:true,color:'#FFFFFF',size:15};
sh.getRange('A2:G2').format.fill='#E8F2F0'; sh.getRange('A2:G2').format.wrapText=true;
sh.getRange('A3:G3').format.fill='#155E63'; sh.getRange('A3:G3').format.font={bold:true,color:'#FFFFFF'};
sh.getRange(`A4:G${added.length+3}`).format.wrapText=true; sh.getRange(`A4:G${added.length+3}`).format.verticalAlignment='top';
[38,42,48,47,48,60,47].forEach((w,i)=>{const c=String.fromCharCode(65+i);sh.getRange(`${c}:${c}`).format.columnWidth=w;});
sh.getRange('A1:G1').format.rowHeight=28; sh.getRange('A2:G2').format.rowHeight=55; sh.getRange(`A4:G${added.length+3}`).format.rowHeight=68;
const sources=[[ref,'https://openknowledge.fao.org/3/cc5044en/cc5044en.pdf','Profiles covering aerial yam, bitter kola, cocoyam, cowpea, dandelion, garden huckleberry, horned melon, jambolan, milk thistle (taxon as printed), redflower ragleaf, turkey berry, water yam and wild lettuce.']];
const ss=wb.worksheets.add('African Batch 2 Sources'); ss.getRange('A1:C1').values=[['Source','Reference URL','Supports these entries']]; ss.getRange('A2:C2').values=sources; ss.tables.add('A1:C2',true,'AfricanBatch2SourcesTable'); ss.freezePanes.freezeRows(1); ss.getRange('A1:C2').format.wrapText=true; ss.getRange('A2:C2').format.verticalAlignment='top'; [62,85,100].forEach((w,i)=>{const c=String.fromCharCode(65+i);ss.getRange(`${c}:${c}`).format.columnWidth=w;}); ss.getRange('A2:C2').format.rowHeight=55;
wb.recalculate();
const output=await SpreadsheetFile.exportXlsx(wb); await output.save(path);
console.log(`Added ${added.length} new records (deduplicated against four existing ingredient tabs); source sheet added.`);
