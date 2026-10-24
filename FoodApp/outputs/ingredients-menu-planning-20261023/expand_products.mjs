import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const path = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const input = await FileBlob.load(path);
const wb = await SpreadsheetFile.importXlsx(input);
const ws = wb.worksheets.getItem('Packaged Products');
const additions = [
['Potato chips, salted','Potato snack chip','Global; industrial and local styles','Salty, toasted potato','Thin, crisp','Snack; serve with dips','Starch; fat; sodium varies','Keep sealed and dry','Usually vegan; check oils and flavorings'],
['Potato chips, vinegar','Potato snack chip','United Kingdom; global commercial style','Sharp, sour, salty','Thin, crisp','Snack; pair with creamy dips','Starch; fat; sodium varies','Acidity from vinegar seasoning','Usually vegan; check flavorings'],
['Potato chips, paprika','Potato snack chip','Central European; global commercial style','Sweet pepper, smoky, savory','Thin, crisp','Snack; sandwich accompaniment','Starch; fat; sodium varies','Paprika seasoning varies','Check dairy or other flavorings'],
['Potato chips, chili-lime','Potato snack chip','Mexico-inspired; global commercial style','Hot, tart, salty','Thin, crisp','Snack; garnish for bowls','Starch; fat; sodium varies','Chili and citrus blends vary','Check seasoning allergens'],
['Potato chips, seaweed','Potato snack chip','East Asian-inspired; global commercial style','Salty, marine umami','Thin, crisp','Snack; serve with rice dishes','Starch; fat; iodine and sodium vary','Seaweed seasoning varies','May contain sesame or fish'],
['Potato chips, sour cream and onion','Potato snack chip','North American; global commercial style','Tangy, oniony, savory','Thin, crisp','Snack; sandwich accompaniment','Starch; fat; sodium varies','Dairy-based flavoring common','Usually contains milk'],
['Potato chips, truffle-seasoned','Potato snack chip','European-inspired; global commercial style','Earthy, garlicky, savory','Thin, crisp','Snack; serve with sandwiches','Starch; fat; sodium varies','Truffle flavor may be flavoring rather than truffle','Check dairy and flavoring ingredients'],
['Tortilla chips, salted','Corn snack chip','Mesoamerican / Mexican; global','Toasted corn, salty','Rigid, crisp','Nachos, salsa, guacamole','Starch; fat; sodium varies','Often made from nixtamalized corn','Usually gluten-free; check cross-contact'],
['Tortilla chips, blue corn','Corn snack chip','Southwestern US; Mexican culinary influence','Toasted, earthy, nutty','Rigid, crisp','Salsa, dips, nachos','Starch; fiber; anthocyanins vary','Blue color comes from corn variety','Usually gluten-free; check cross-contact'],
['Tortilla chips, lime','Corn snack chip','Mexican-inspired; global commercial style','Corn, citrus, salt','Rigid, crisp','Salsa, ceviche-style dishes','Starch; fat; sodium varies','Lime may be juice or seasoning','Check seasoning allergens'],
['Plantain chips, green','Plantain snack chip','Caribbean; West African; Latin American','Starchy, lightly sweet, salty','Crisp, firm','Snack; serve with dips','Starch; fat; sodium varies','Green plantains are less sweet','Usually vegan; check frying oil'],
['Plantain chips, ripe','Plantain snack chip','Caribbean; West African; Latin American','Sweet, caramelized, salty','Crisp, sometimes brittle','Snack; sweet-savory side','Starch; sugars; fat varies','Riper fruit browns faster','Usually vegan; check frying oil'],
['Cassava chips','Root snack chip','West Africa; South America; South Asia','Mild, earthy, salty','Light, crisp','Snack; serve with chutney or salsa','Starch; fat; sodium varies','Commercial seasoning varies','Usually gluten-free; check cross-contact'],
['Taro chips','Root snack chip','East Asia; Pacific Islands; Southeast Asia','Nutty, earthy, mildly sweet','Crisp, delicate','Snack; serve with dips','Starch; fat; sodium varies','Cook thoroughly; commercial chips are cooked','Check oils and seasoning'],
['Sweet potato chips','Root snack chip','Global; contemporary packaged snack','Sweet, earthy, salty or spiced','Crisp to dense','Snack; side for sandwiches','Starch; beta-carotene; fat varies','Color and sweetness vary by cultivar','Usually vegan; check seasonings'],
['Beet chips','Root snack chip','European; global packaged snack','Earthy, sweet, salty','Crisp, light','Snack; garnish for salads','Fiber; folate; fat and sodium vary','Color can stain; product may be fried or baked','Usually vegan; check seasonings'],
['Kale chips','Leafy vegetable snack','North American; global contemporary snack','Green, savory, sometimes cheesy','Light, brittle','Snack; salad topping','Vitamins vary; sodium and fat depend on recipe','Baked or dehydrated; seasoning varies','May contain dairy, nuts, or seeds'],
['Prawn crackers','Starch and seafood snack','Indonesia; Malaysia; Southeast Asia','Savory, briny, lightly sweet','Puffed, airy, crisp','Snack; accompaniment to meals','Starch; protein; sodium varies','Often fried; uncooked crackers expand when fried','Contains crustacean; may contain egg'],
['Rice crackers, senbei','Rice snack cracker','Japan','Toasty, salty, sweet-soy, or seaweed','Crisp, brittle','Snack; tea accompaniment','Mostly carbohydrate; sodium varies','Styles include baked, grilled, and fried','May contain soy, wheat, sesame, or fish'],
['Rice crackers, shrimp','Rice snack cracker','Japan; East Asia','Savory, briny, lightly sweet','Crisp, airy','Snack; soup garnish','Mostly carbohydrate; sodium varies','Shrimp content and flavor strength vary','May contain crustacean'],
['Chickpea puffs','Legume snack','South Asia; global packaged snack','Nutty, savory, spiced','Puffed, crunchy','Snack; salad topper','Protein; fiber; sodium varies','Baked or extruded; flavoring varies','Legume sensitivity possible'],
['Roasted chickpeas, masala','Legume snack','South Asia','Spiced, savory, hot','Crunchy, dense','Snack; chaat topping','Protein; fiber; sodium varies','Seasoning blends vary regionally','Check spice blend and cross-contact'],
['Wasabi peas','Legume snack','Japan-inspired; global commercial snack','Hot, sweet, pungent','Crunchy, coated','Snack; salad or rice topping','Protein; fiber; sodium varies','Coating often includes starch','May contain soy or wheat'],
['Peanut chikki','Nut and sugar confection','India','Roasted, caramel-like, sweet','Hard, brittle crunch','Snack; confectionery','Protein; unsaturated fat; sugars','Traditional brittle made with jaggery or sugar','Contains peanut'],
['Murukku','Rice and pulse snack','South India','Savory, sesame, spice','Crisp, ridged','Snack; tea accompaniment','Starch; protein; fat varies','Rice flour and pulse flour blends vary','May contain sesame; fried'],
['Banana chips, Kerala style','Fruit snack chip','Kerala, India','Savory, coconut-oil aroma; ripe styles sweeter','Crisp, dense','Snack; meal accompaniment','Starch; potassium; fat varies','Raw banana and ripe banana versions differ','Often fried; check oil'],
['Chili oil','Bottled chili condiment','China; East Asia; global','Hot, aromatic, toasted','Oil with chili sediment','Noodles, dumplings, eggs','Fat; sodium varies','Heat and solids vary; shake if directed','May contain sesame, soy, or peanuts'],
['Extra-hot fermented chili sauce','Fermented bottled sauce','Caribbean; Africa; Americas; global','Hot, sour, fruity','Pourable, thin to thick','Marinade, table sauce, stew','Sodium varies','Pepper, vinegar, and fermentation styles vary','Check allergens and acidity'],
['Green chili sauce','Bottled chili sauce','South Asia; Latin America; global','Fresh, sharp, hot','Pourable puree','Sandwiches, snacks, marinades','Sodium varies','Fresh green chilies, herbs, and vinegar vary','Check preservatives and allergens'],
['Peri-peri sauce','Bottled chili sauce','Southern African; Portuguese culinary influence','Hot, citrusy, garlicky','Pourable sauce','Grill marinade, table sauce','Sodium varies','Heat and lemon-vinegar balance vary by brand','Check allergens'],
['Sriracha-style chili sauce','Bottled chili sauce','Thai-inspired; global commercial style','Garlicky, sweet-hot, tangy','Thick pourable sauce','Noodles, eggs, sandwiches','Added sugar and sodium vary','Brand formulas differ substantially','Check preservatives and allergens'],
['Ketchup, tomato','Bottled tomato condiment','Global; modern commercial form','Sweet, tangy, tomato-rich','Smooth, thick','Fries, sandwiches, sauces','Added sugar and sodium vary','Styles differ in sweetness and spice','Check allergens'],
['Mustard, yellow','Bottled mustard condiment','North American; European','Mild, sharp, vinegary','Smooth, pourable paste','Sandwiches, dressings, marinades','Sodium varies','Prepared from mustard seed, acid, and spices','Mustard allergen'],
['Mustard, wholegrain','Bottled mustard condiment','French; German; European','Sharp, nutty, tangy','Coarse, seeded paste','Dressings, cheese, roasts','Sodium varies','Seed size and wine/vinegar base vary','Mustard allergen'],
['Chutney, tamarind','Bottled fruit condiment','South Asia','Sour-sweet, fruity, spiced','Smooth to pulpy sauce','Chaat, snacks, samosa, marinades','Added sugar and sodium vary','Sweetness and spice vary by region','Check dates, spices, and allergens'],
['Chutney, coconut','Bottled or jarred condiment','South India; Sri Lanka','Fresh, nutty, savory','Coarse to smooth paste','Idli, dosa, rice, snacks','Fat; sodium varies','Fresh and shelf-stable versions differ','May contain coconut; check preservatives'],
['Chutney, tomato','Bottled fruit-vegetable condiment','South Asia','Sweet, sour, hot, savory','Smooth or chunky','Snacks, sandwiches, rice','Added sugar and sodium vary','Cooked and fresh styles differ','Check seasoning allergens'],
['Salsa roja','Bottled tomato-chili sauce','Mexico','Roasted, tart, savory, hot','Chunky to smooth','Tacos, eggs, beans, chips','Sodium varies','Roasted, boiled, and raw versions differ','Usually vegan; check product'],
['Salsa verde, tomatillo','Bottled tomatillo sauce','Mexico','Tart, herbal, green chili','Chunky to smooth','Tacos, enchiladas, beans','Sodium varies','Tomatillo and chili ratios vary','Usually vegan; check product'],
['Aji amarillo sauce','Bottled chili sauce','Peru','Fruity, aromatic, medium-hot','Creamy to pourable','Potatoes, sandwiches, grilled foods','Sodium varies','Often blended with dairy or oil','May contain milk or egg'],
['Peanut satay sauce','Bottled nut sauce','Indonesia; Malaysia; Southeast Asia','Roasted, sweet-savory, chili','Thick, creamy','Satay, noodles, vegetables','Fat, sugars, sodium vary','Thin with water or citrus for dipping','Contains peanut; may contain soy'],
['Sesame dressing','Bottled dressing','Japan; East Asia; global','Toasted, sweet-savory, tangy','Creamy or pourable','Salad, noodles, vegetables','Fat, sugars, sodium vary','Roasted sesame and emulsifier levels vary','Contains sesame; may contain soy'],
['Caesar dressing','Bottled dressing','North American; Italian-American','Creamy, salty, garlicky, umami','Emulsified, creamy','Salad, wraps, sandwiches','Fat, sodium vary','Anchovy and cheese are common','May contain egg, milk, fish'],
['Ranch dressing','Bottled dressing','North American','Creamy, herby, tangy','Emulsified, pourable','Salads, vegetables, snacks','Fat, sodium vary','Dairy and herb blends vary','Often contains milk and egg'],
['Vinaigrette, balsamic','Bottled dressing','Italian-inspired; global','Sweet-tart, woody','Emulsified or separated liquid','Salads, roasted vegetables','Oil and sugar vary','Shake before use if separated','Check mustard and other allergens'],
['Coconut vinegar','Bottled vinegar','Philippines; Southeast Asia','Mild, fruity, gently sour','Thin liquid','Dipping sauces, pickles, marinades','Acidity varies','Made from coconut sap or coconut water','Usually vegan; check label'],
['Black vinegar, Chinkiang','Bottled aged vinegar','China; Jiangsu','Smoky, malty, tart, savory','Dark, thin liquid','Dumplings, noodles, braises','Acidity and sodium vary','Aged grain vinegar; formulas differ','May contain wheat'],
['Malt vinegar, chip shop style','Bottled malt vinegar','United Kingdom; Ireland','Sharp, malty, tangy','Thin liquid','Chips, fish, pickles','Negligible nutrients','Malt extract and acidity vary','Contains barley; may contain gluten'],
['Pomegranate syrup, grenadine style','Bottled fruit syrup','Middle East; global bar and food use','Sweet-tart, floral','Pourable syrup','Drinks, desserts, glaze','Mostly sugars','Commercial products may rely on added color/flavor','Check additives'],
['Rose syrup','Bottled floral syrup','South Asia; Middle East','Floral, sweet','Pourable syrup','Drinks, desserts, shaved ice','Mostly sugars','Rose intensity and added color vary','Check additives'],
['Soy sauce, reduced sodium','Bottled fermented condiment','East Asia; global','Salty umami, milder salt impact','Thin liquid','Seasoning, marinade, dipping','Sodium lower than standard versions; varies by brand','Compare labels; reduced sodium is not low sodium','Soy; often wheat'],
['Vegan Worcestershire-style sauce','Bottled fermented condiment','Global commercial adaptation','Sour-sweet, savory, spiced','Thin liquid','Marinade, stew, tomato drinks','Sodium and sugars vary','Anchovy-free formulas vary','Check soy, wheat, and other allergens'],
];

const startRow = 116;
ws.getRange(`A${startRow}:I${startRow + additions.length - 1}`).values = additions;
ws.getRange(`A${startRow}:I${startRow + additions.length - 1}`).format.wrapText = true;
ws.getRange(`A${startRow}:I${startRow + additions.length - 1}`).format.verticalAlignment = 'top';
ws.getRange(`A${startRow}:I${startRow + additions.length - 1}`).format.rowHeight = 56;
const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(path);
console.log(`Added ${additions.length} entries; workbook now has ${additions.length + 112} product rows.`);
