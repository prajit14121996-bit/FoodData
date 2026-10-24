import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const path='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const norm=s=>String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');
const names=new Set();
for(const n of ['Global Ingredients','Biodiversity Foods','African Food Diversity','Regional Crop Forms','African Foods Batch 2']){
 const sh=wb.worksheets.getItem(n); for(const r of sh.getUsedRange().values.slice(3)){const v=n==='Global Ingredients'?r[1]:r[0];if(v)names.add(norm(v));}
}
const entries=[
 ['Tarwi / chocho / Andean lupin','Pulse; Lupinus mutabilis','Andean highlands of Peru, Bolivia and Ecuador','Seeds','Strongly bitter before debittering because of quinolizidine alkaloids; taste after processing varies by dish','Seeds used in traditional dishes after proper washing/processing to remove bitter alkaloids; never treat raw bitter seed as ready to eat','FAO EcoCrop; FAO AGRIS'],
 ['Cañihua / kañiwa / kaniwa','Pseudocereal; Chenopodium pallidicaule','High Andes, especially Peru and Bolivia','Seeds','Review reports it lacks the characteristic bitter taste associated with quinoa and tarwi; cultivar notes vary','Seeds are used as a food grain and milled; regional preparation and naming vary','Kañihua review; Local Food Plants for Nutrition'],
 ['Maca','Andean root crop; Lepidium meyenii','Central Peruvian Andes, high-elevation plateaus','Hypocotyl/root','Specific flavour not standardized in cited profile; locally valued forms include yellow, red and black','Traditionally naturally dried and then boiled/softened for juice or other preparations; CIP lists baked, salad and juice uses','CIP Andean crops; Ethnobotany of Lepidium meyenii'],
 ['Achira / edible canna / Queensland arrowroot','Starch root crop; Canna indica','Andean South America; cultivated in other regions as a starch crop','Starchy rhizome; leaves also used in some Andean preparations','Mild starchy profile not specified by cited source','Rhizome boiled or baked; starch extracted for flour, noodles and baked goods; forms vary regionally','CIP Achira; CIP Andean crops; Slow Food Apurimac Achira'],
 ['Arracacha / racacha / zanahoria blanca','Root crop; Arracacia xanthorrhiza','Andean region from Venezuela southward; also grown in Brazil','Fleshy storage roots','Taste descriptors vary by cultivar; cited sources emphasize culinary use rather than a standardized profile','Boiled, baked, fried; used in soups, stews, purees, puddings and other foods','CIP Andean crops; FAO Traditional High Andean Cuisine'],
 ['Ahipa / ashipa','Root vegetable; Pachyrhizus ahipa','Andean South America, especially Bolivia','Hypocotyl / swollen storage root','Specific taste descriptor not stated in cited crop profile','CIP records salad use; crop profile describes the root as eaten raw like an apple; check cultivar and local practice','CIP Andean crops'],
 ['Mauka / chago','Root crop; Mirabilis expansa','Andean highlands of Bolivia, Peru and Ecuador','Storage roots','Can be unpleasantly bitter before sun exposure; local form and processing matter','CIP recommends exposing roots to sunlight before consumption to reduce bitterness; also boiled, fried or used in stews and soups','CIP Andean crops; CIP Achira and Mauka'],
 ['Chaya / Mayan spinach / tree spinach','Leafy vegetable; Cnidoscolus aconitifolius','Mesoamerica, especially Mexico and the Yucatán; now grown in tropical regions','Young leaves and tender shoots','ECHO profile describes a mild leaf that takes on seasoning flavours; cooked texture is dense','Cook thoroughly before eating; raw leaves are unsafe. Cooked leaves used like spinach in soups, stews, tamales and regional meals','NParks Chaya; ECHO Asia Chaya Fact Sheet'],
];
const added=entries.filter(r=>!names.has(norm(r[0])));
const sh=wb.worksheets.add('Andean & Mesoamerican Foods');
sh.getRange('A1:G1').merge();sh.getRange('A1').values=[['ANDEAN & MESOAMERICAN FOOD DIVERSITY | NEW FORMS']];
sh.getRange('A2:G2').merge();sh.getRange('A2').values=[['Distinct crops and local forms added from FAO, the International Potato Center and specialist crop references. Taste is recorded only where a source supports it. Local names can refer to different crops across regions; verify identity and required processing, especially for bitter seeds and chaya leaves.']];
sh.getRange('A3:G3').values=[['Ingredient / local names','Family / type','Origin / regional range','Edible part / form','Taste profile','Use / preparation notes','References']];
sh.getRange(`A4:G${added.length+3}`).values=added;sh.tables.add(`A3:G${added.length+3}`,true,'AndeanMesoamericanFoodsTable');sh.freezePanes.freezeRows(3);
sh.getRange('A1:G1').format.fill='#123B43';sh.getRange('A1:G1').format.font={bold:true,color:'#FFFFFF',size:15};
sh.getRange('A2:G2').format.fill='#E8F2F0';sh.getRange('A2:G2').format.wrapText=true;
sh.getRange('A3:G3').format.fill='#155E63';sh.getRange('A3:G3').format.font={bold:true,color:'#FFFFFF'};
sh.getRange(`A4:G${added.length+3}`).format.wrapText=true;sh.getRange(`A4:G${added.length+3}`).format.verticalAlignment='top';
[42,42,49,45,49,67,55].forEach((w,i)=>{const c=String.fromCharCode(65+i);sh.getRange(`${c}:${c}`).format.columnWidth=w;});
sh.getRange('A1:G1').format.rowHeight=28;sh.getRange('A2:G2').format.rowHeight=54;sh.getRange(`A4:G${added.length+3}`).format.rowHeight=66;
const src=[
 ['International Potato Center (CIP) Genebank — Andean Root and Tuber Crops','https://cipotato.org/genebankcip/process/andean-root-and-tuber-crops-artcs/','Crop origins/ranges, edible parts and current/traditional uses for achira, arracacha, maca, mauka and ahipa.'],
 ['International Potato Center (CIP) — Achira','https://cipotato.org/genebankcip/process/achira/','Achira identity and starchy rhizome uses.'],
 ['FAO EcoCrop — Lupinus mutabilis','https://ecocrop.apps.fao.org/ecocrop/srv/en/cropView?id=7435','Tarwi/Andean lupin common names and need to remove bitter taste before use.'],
 ['FAO AGRIS — Characterising Tarwi Nutritional and Alkaloid Profiles','https://agris.fao.org/search/en/providers/122413/records/687a92d3464818917f335f62','Tarwi seeds contain alkaloids that must be removed before consumption.'],
 ['Kañihua review (PMC)','https://pmc.ncbi.nlm.nih.gov/articles/PMC11305323/','Cañihua crop and comparison of bitter taste with quinoa and tarwi.'],
 ['Ethnobotany and Ethnopharmacology of Lepidium meyenii (PMC)','https://pmc.ncbi.nlm.nih.gov/articles/PMC3184420/','Maca edible hypocotyl and traditional dried/boiled preparations.'],
 ['FAO — Traditional High Andean Cuisine','https://www.fao.org/4/i1466e/i1466e.pdf','Arracacha domestication and Andean distribution.'],
 ['NParks Flora & Fauna Web — Chaya','https://www.nparks.gov.sg/florafaunaweb/flora/4/8/4839','Chaya edible parts and requirement to cook thoroughly; raw leaves are unsafe.'],
 ['ECHO Asia — Chaya Seed Fact Sheet','https://assets.echocommunity.org/publication_issue/9d9ed624-8c1b-4d34-b133-0796a9dfb2d1/en/cnidoscolus_aconitifolius.pdf','Chaya origin, cooked use, mild taste and texture.'],
 ['Slow Food Foundation — Apurimac River Valley Achira','https://www.fondazioneslowfood.com/en/ark-of-taste-slow-food/apurimac-river-valley-achira/','Regional Achira forms and culinary processing in the Apurimac Valley.'],
];
const ss=wb.worksheets.add('Andean & Mesoamerican Sources');ss.getRange('A1:C1').values=[['Source','Reference URL','Supports these entries']];ss.getRange(`A2:C${src.length+1}`).values=src;ss.tables.add(`A1:C${src.length+1}`,true,'AndeanMesoamericanSourcesTable');ss.freezePanes.freezeRows(1);ss.getRange(`A1:C${src.length+1}`).format.wrapText=true;ss.getRange(`A2:C${src.length+1}`).format.verticalAlignment='top';[66,84,95].forEach((w,i)=>{const c=String.fromCharCode(65+i);ss.getRange(`${c}:${c}`).format.columnWidth=w;});ss.getRange(`A2:C${src.length+1}`).format.rowHeight=50;
wb.recalculate();const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);console.log(`Added ${added.length} distinct foods and ${src.length} references after checking existing ingredient tabs.`);
