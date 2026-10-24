import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outDir = '/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023';
const headers = ['Ingredient', 'Family / type', 'Origin & cuisine associations*', 'Taste profile', 'Texture / behavior', 'Dish & menu uses', 'Nutrition highlights†', 'Prep / handling notes', 'Diet / allergen notes'];
const parse = (s) => s.trim().split('\n').filter(Boolean).map(line => line.split('|'));

const data = {
  'Vegetables': parse(`
Artichoke|Flower bud|Mediterranean; Italian; French|Earthy, nutty, mildly sweet|Tender heart, leafy bracts|Roast, grill, braise, dip|Fiber; folate; vitamin C; potassium|Trim tips; acidulate cut surfaces|Usually vegan; check dips
Asparagus|Young shoot|Mediterranean; Northern European|Green, grassy, sweet|Crisp-tender; roasts quickly|Side, pasta, salad, soup|Folate; vitamins K and C; fiber|Snap woody ends; blanch or roast|Vegan
Beetroot|Root|Eastern European; Mediterranean|Earthy, sweet|Dense raw; silky when cooked|Roast, pickle, soup, salads|Folate; manganese; fiber|Wear gloves; roast whole for sweetness|Vegan
Bell pepper|Fruit vegetable|Americas; Mediterranean; East Asian|Sweet, grassy; red is sweeter|Crisp raw; softens and chars|Salads, stir-fries, roast, stuff|Vitamin C; carotenoids; fiber|Remove ribs/seeds as desired|Vegan
Bok choy|Leafy brassica|Chinese; East Asian|Mild, sweet, mustardy|Juicy stems; tender leaves|Stir-fry, soup, steam, dumplings|Vitamins C and K; folate; calcium|Separate stems and leaves for even cooking|Vegan
Broccoli|Brassica|Italian; global|Green, mildly bitter, sweet when roasted|Florets soften; stems stay firm|Roast, steam, stir-fry, soup|Vitamins C and K; folate; fiber|Peel tough stem; cook briefly for bite|Vegan
Brussels sprouts|Brassica|Belgian; Northern European|Nutty, cabbage-like, gently bitter|Dense; caramelizes when roasted|Roast, shave for salad, braise|Vitamins C and K; folate; fiber|Halve larger sprouts; avoid overcooking|Vegan
Cabbage, green|Leafy brassica|Central/Eastern European; East Asian|Mild, peppery, sweet when cooked|Crisp raw; silky when braised|Slaw, stir-fry, rolls, soup|Vitamin C; vitamin K; fiber|Slice thin for quick cooking|Vegan
Cabbage, red|Leafy brassica|Central/Eastern European|Peppery, mildly sweet|Crunchy raw; softens in braise|Slaw, pickle, braise, tacos|Vitamin C; vitamin K; anthocyanins|Acid helps retain bright color|Vegan
Carrot|Root|Central Asian heritage; global|Sweet, earthy|Crunchy raw; tender cooked|Soup, roast, salad, mirepoix|Beta-carotene; fiber; potassium|Roast to concentrate sweetness|Vegan
Cauliflower|Brassica|South Asian; Mediterranean|Mild, nutty, slightly sweet|Tender; can crisp or puree|Roast, curry, mash, rice substitute|Vitamin C; folate; fiber|Dry well before roasting for browning|Vegan
Celery|Stem|Mediterranean; French; North American|Savory, herbal, salty-green|Juicy crunch; softens in stocks|Stocks, salads, soups, braises|Vitamin K; folate; potassium|Use leaves for extra aroma|Vegan; common allergen in EU
Celery root|Root|French; Central European|Earthy, celery-like, nutty|Firm; creamy when pureed|Remoulade, mash, soup, roast|Fiber; potassium; vitamin C|Peel thick skin; use lemon to limit browning|Vegan
Chayote|Gourd|Mexican; Central American; Caribbean|Mild, cucumber-like, lightly sweet|Crisp to tender; holds shape|Stew, saute, slaw, curry|Vitamin C; folate; fiber|Peel if skin is tough; remove seed if desired|Vegan
Corn, sweet|Grain vegetable|Mesoamerican; Americas|Sweet, milky, grassy|Juicy kernels; browns well|Grill, salsa, chowder, polenta-style dishes|Carbohydrate; fiber; folate; carotenoids|Cook soon after shucking for sweetness|Vegan
Cucumber|Gourd|South Asian heritage; global|Cool, mild, lightly sweet|Watery, crisp|Salads, pickle, raita, sandwiches|Water; vitamin K; small amounts of fiber|Salt and drain for less watery salads|Vegan
Eggplant|Nightshade|South Asian; Middle Eastern; Mediterranean|Mild, savory, subtly bitter|Spongy raw; silky when cooked|Roast, curry, stew, grill, dips|Fiber; manganese; anthocyanins in purple skin|Salt only if desired; cook through for creamy texture|Vegan
Fennel bulb|Stem/bulb|Italian; Mediterranean|Anise, sweet, fresh|Crisp raw; soft and sweet when cooked|Salad, roast, braise, seafood pairing|Fiber; vitamin C; potassium|Fronds are edible garnish|Vegan
Garlic|Allium bulb|Central Asian heritage; Mediterranean; East Asian|Pungent, sulfurous, sweet when cooked|Juicy raw; creamy when roasted|Base aromatics, sauces, marinades, rubs|Manganese; vitamin B6; small amounts of micronutrients|Burns quickly; add at suitable stage|Vegan
Green beans|Pod vegetable|Americas; European|Green, sweet, vegetal|Crisp-tender|Saute, steam, casserole, pickle|Fiber; folate; vitamins C and K|Trim ends; blanch for vivid color|Vegan
Kale|Leafy brassica|Mediterranean; Northern European|Peppery, earthy, slightly bitter|Sturdy; softens with massage or cooking|Salad, soup, saute, chips|Vitamins K, A and C; folate; fiber|Remove thick stems; massage for raw salads|Vegan
Leek|Allium stem|Welsh; French; Northern European|Mild onion, sweet|Tender layers; melts in braise|Soup, tart, risotto, roast|Vitamin K; folate; fiber|Wash between layers to remove grit|Vegan
Lettuce, romaine|Leafy green|Mediterranean; global|Mild, fresh, slightly bitter|Crisp ribs; tender leaves|Salad, wraps, grilled hearts|Folate; vitamins A and K; water|Dry well for dressing adhesion|Vegan
Mushroom, button|Fungus|European; global|Umami, earthy|Juicy; browns when not crowded|Saute, soup, pasta, fillings|B vitamins; selenium; potassium|Cook off moisture before seasoning|Vegan; not botanically a vegetable
Mushroom, shiitake|Fungus|Chinese; Japanese|Deep umami, woody|Chewy cap; savory broth|Broth, stir-fry, noodles, dumplings|B vitamins; copper; fiber|Remove tough stems; dried form is concentrated|Vegan; not botanically a vegetable
Okra|Pod vegetable|West African; South Asian; Southern US|Green, grassy|Tender pod; mucilage thickens liquids|Gumbo, curry, fry, pickle|Fiber; folate; vitamin C|Dry before frying; acid can reduce sliminess|Vegan
Onion, red|Allium bulb|Central Asian heritage; Mediterranean|Pungent raw, sweet cooked|Crisp raw; silky caramelized|Salad, pickle, curry, soup, base|Vitamin C; fiber; small amounts of folate|Rinse or pickle for gentler raw flavor|Vegan
Onion, yellow|Allium bulb|Global; French mirepoix|Pungent, sweet when browned|Softens; caramelizes|Aromatic base, soup, roast, sauce|Vitamin C; fiber; small amounts of folate|Cook low for sweetness; high heat for browning|Vegan
Parsnip|Root|Northern European; British|Sweet, earthy, nutty|Firm; softens and browns|Roast, mash, soup|Fiber; folate; potassium|Peel larger roots; roast evenly|Vegan
Peas, green|Legume vegetable|Mediterranean; European; South Asian|Sweet, grassy|Tender, starchy|Soup, pilaf, curry, puree|Protein; fiber; folate; vitamins C and K|Add late to preserve color|Vegan
Potato, russet|Tuber|Andean; global|Earthy, neutral|Fluffy inside; crisp outside|Bake, mash, fry, roast|Carbohydrate; potassium; vitamin B6|Store cool and dark; don't eat green portions|Vegan
Potato, waxy|Tuber|Andean; European|Mild, buttery|Holds shape when cooked|Salad, stew, gratin, roast|Carbohydrate; potassium; vitamin C|Choose for soups and salads that need intact pieces|Vegan
Pumpkin / winter squash|Gourd|Americas; East Asian|Sweet, nutty, earthy|Dense; creamy when cooked|Soup, roast, curry, pie|Beta-carotene; fiber; potassium|Roast to deepen flavor; seeds are edible|Vegan
Radish, daikon|Root|East Asian|Fresh, peppery, lightly sweet|Juicy raw; tender braised|Pickle, soup, stew, salad|Vitamin C; fiber; water|Salt and rinse for crisp quick pickle|Vegan
Radish, red|Root|European; global|Peppery, crisp|Juicy and crunchy|Salad, pickle, taco garnish|Vitamin C; fiber; water|Leaves can be cooked like greens|Vegan
Spinach|Leafy green|Persian heritage; South Asian; Mediterranean|Mineral, mildly sweet|Tender; wilts rapidly|Saag, salad, pasta, soup, filling|Folate; vitamins K and A; iron; magnesium|Wash grit thoroughly; volume shrinks substantially|Vegan
Sweet potato|Tuber|Central/South American heritage; global|Sweet, earthy|Creamy; crisps when roasted|Roast, mash, curry, pie|Beta-carotene; fiber; potassium|Bake or roast to concentrate sweetness|Vegan
Tomato|Fruit vegetable|Andean/Mesoamerican heritage; Italian; global|Acidic, sweet, umami|Juicy raw; breaks down into sauce|Salad, curry, salsa, stew, sauce|Vitamin C; potassium; lycopene|Ripe tomatoes suit raw use; canned suits long cooks|Vegan
Turnip|Root|European; Middle Eastern|Peppery, mildly sweet|Firm; softens in braise|Roast, mash, stew, pickle|Vitamin C; fiber; potassium|Young greens are edible|Vegan
Zucchini|Gourd|Mesoamerican heritage; Italian; Mediterranean|Mild, sweet, grassy|Tender; high water content|Grill, saute, fritter, bake|Vitamin C; potassium; fiber|Use high heat and avoid crowding to brown|Vegan
Arugula / rocket|Leafy green|Mediterranean; Italian|Peppery, mustardy|Tender leaves, crisp stems|Salads, pizza, pesto, sandwich|Vitamin K; folate; calcium|Add late; heat softens the peppery bite|Vegan
Collard greens|Leafy brassica|West African; Southern US|Earthy, cabbage-like|Sturdy leaves; tender when braised|Braise, soup, wraps, saute|Vitamins K, A and C; folate; fiber|Remove thick stems; cook slowly for tenderness|Vegan
Swiss chard|Leafy green|Mediterranean; European|Mineral, earthy, mildly bitter|Soft leaves; crisp colorful stems|Saute, soup, tart, grain bowl|Vitamins K, A and C; magnesium|Cook stems before leaves|Vegan
Mustard greens|Leafy brassica|South Asian; East Asian; Southern US|Peppery, pungent|Sturdy; wilts with heat|Saag, stir-fry, braise, soup|Vitamins K, A and C; folate|Young leaves suit salads; mature leaves benefit from cooking|Vegan
Watercress|Leafy green|European; East Asian|Peppery, fresh|Delicate, crisp|Salad, soup, sandwich, garnish|Vitamins K and C; folate|Wash well; use soon after purchase|Vegan
Endive|Leafy chicory|Belgian; French|Bitter, crisp|Firm leaf cups; tender when cooked|Salad, braise, dip vessels|Fiber; folate; vitamin K|Pair bitter leaves with fruit, cheese or dressing|Vegan
Radicchio|Leafy chicory|Italian|Bitter, peppery|Crisp raw; softens on grill|Salad, grill, risotto, roast|Fiber; vitamin K; vitamin C|Salt and acid balance bitterness|Vegan
Escarole|Leafy chicory|Italian; Mediterranean|Mildly bitter, green|Tender inner leaves; sturdy outer leaves|Soup, beans, braise, salad|Folate; vitamins A and K; fiber|Wash leaves carefully; cook outer leaves longer|Vegan
Napa cabbage|Leafy brassica|Chinese; Korean|Mild, sweet, juicy|Crisp ribs; tender leaves|Kimchi, soup, dumplings, stir-fry|Vitamins C and K; folate|Separate leaves for quick cooking|Vegan
Kohlrabi|Brassica stem|Central European; South Asian|Sweet, turnip-like, peppery|Crisp raw; tender cooked|Slaw, pickle, roast, curry|Vitamin C; fiber; potassium|Peel fibrous outer layer|Vegan
Rutabaga / swede|Root|Northern European; Scandinavian|Sweet, earthy, cabbage-like|Dense; creamy when cooked|Mash, roast, soup, stew|Vitamin C; fiber; potassium|Peel thick skin; cut evenly for roasting|Vegan
Jicama|Root|Mexican; Central American|Mild, juicy, sweet|Very crisp, water-rich|Slaw, snack, salad, stir-fry|Fiber; vitamin C; water|Peel before eating; commonly served raw|Vegan
Cassava / yuca|Tuber|South American; West African; Caribbean|Mild, starchy|Dense; floury when cooked|Boil, fry, mash, tapioca|Carbohydrate; vitamin C; resistant starch varies|Peel and cook thoroughly; never eat raw cassava|Vegan
Taro|Corm|South Asian; East Asian; Pacific Islands|Nutty, earthy|Starchy, sticky-creamy|Stew, curry, mash, dessert|Carbohydrate; fiber; potassium|Cook fully; raw taro contains irritating calcium oxalate crystals|Vegan
Lotus root|Rhizome|East Asian; South Asian|Mild, sweet, slightly nutty|Crisp with distinctive holes|Stir-fry, soup, pickle, chips|Fiber; vitamin C; potassium|Rinse cut pieces; acidulate to limit browning|Vegan
Bamboo shoots|Young shoot|East Asian; Southeast Asian|Mild, earthy|Crisp, fibrous|Stir-fry, curry, soup, dumplings|Fiber; potassium; low energy density|Use commercially prepared shoots; boil fresh shoots thoroughly|Vegan
Bean sprouts, mung|Germinated legume|East Asian; Southeast Asian|Fresh, mild, nutty|Juicy crunch|Stir-fry, noodle soup, salad, filling|Vitamin C; folate; fiber|Wash and cook safely; sprouts are perishable|Vegan
Bitter melon|Gourd|South Asian; East Asian; Caribbean|Distinctly bitter, grassy|Firm raw; softens when cooked|Stir-fry, curry, pickle, stuffed dishes|Vitamin C; fiber; folate|Salt briefly or pair with acid/sweetness to soften bitterness|Vegan
Bottle gourd|Gourd|South Asian; African|Mild, lightly sweet|Watery, tender|Curry, soup, kofta, stew|Water; vitamin C; fiber|Taste a small raw piece; discard if unusually bitter|Vegan
Plantain|Starchy fruit|West African; Caribbean; Central American|Starchy green; sweet when ripe|Firm; softens and caramelizes|Fry, roast, mash, stew, chips|Carbohydrate; potassium; vitamin C|Green and ripe plantains have different uses|Vegan
Jerusalem artichoke / sunchoke|Tuber|North American; French|Sweet, nutty, artichoke-like|Crisp raw; creamy roasted|Soup, roast, puree, salad|Fiber (inulin); potassium; iron|Can cause digestive discomfort for some; cook until tender|Vegan
Kabocha squash|Winter squash|Japanese; East Asian|Sweet, chestnut-like|Dense, creamy flesh|Roast, tempura, curry, soup|Beta-carotene; fiber; potassium|Skin is edible when cooked tender|Vegan
Acorn squash|Winter squash|North American|Mild, sweet, nutty|Tender, slightly fibrous|Roast, stuff, soup, mash|Beta-carotene; fiber; potassium|Cut carefully; roast halves for easy serving|Vegan
Delicata squash|Winter squash|North American|Sweet, mild, corn-like|Tender skin; creamy flesh|Roast rings, stuff, soup|Fiber; potassium; carotenoids|Skin usually edible after cooking|Vegan
Spaghetti squash|Winter squash|North American|Mild, lightly sweet|Flesh separates into strands|Roast, stuff, grain bowl, pasta-style dish|Carbohydrate; fiber; vitamin C|Roast cut-side down; scrape strands after cooling slightly|Vegan
Tomatillo|Fruit vegetable|Mexican; Central American|Tart, citrusy, green|Firm flesh under papery husk|Salsa verde, stew, roast, sauce|Vitamin C; fiber; potassium|Remove husk and rinse sticky coating|Vegan
Poblano pepper|Fresh chili pepper|Mexican|Mild, earthy, raisin-like|Thick flesh; softens when roasted|Relleno, roast, sauce, stew|Vitamin C; carotenoids|Roast and peel for silky sauce if desired|Vegan
Jalapeño|Fresh chili pepper|Mexican; Latin American|Fresh, grassy, medium heat|Crisp raw; softens when cooked|Salsa, pickle, poppers, stew|Vitamin C; capsaicin|Remove ribs/seeds for less heat; wear gloves if sensitive|Vegan
Serrano pepper|Fresh chili pepper|Mexican|Bright, sharp, hotter than jalapeño|Small, crisp|Salsa, garnish, stew, pickle|Vitamin C; capsaicin|Add gradually; heat varies by pepper|Vegan
Habanero pepper|Fresh chili pepper|Caribbean; Yucatecan|Fruity, floral, very hot|Thin-walled, aromatic|Hot sauce, salsa, marinade|Vitamin C; capsaicin|Use gloves; add in tiny amounts|Vegan
Shallot|Allium bulb|French; Southeast Asian|Sweet, delicate onion-garlic|Fine layers; browns quickly|Vinaigrette, sauce, curry base, roast|Fiber; vitamin C; small amounts of potassium|Cook gently to avoid bitterness|Vegan
Scallion / green onion|Allium stem|East Asian; global|Fresh onion, mild|Crisp green tops; tender white base|Garnish, stir-fry, pancake, salsa|Vitamin K; vitamin C; folate|Use whites early and greens late|Vegan
Garlic scapes|Allium flower stem|East Asian; North American|Mild garlic, grassy|Tender, curly stem|Pesto, saute, pickle, grill|Vitamin C; fiber; small amounts of minerals|Trim tough ends; cook briefly|Vegan
Romaine hearts|Leafy green|Mediterranean; global|Mild, fresh, sweet|Very crisp ribs|Salad, wrap, grill, sandwich|Folate; vitamins A and K; water|Keep chilled and dry before dressing|Vegan
Iceberg lettuce|Leafy green|North American; global|Mild, sweet, watery|Very crisp, dense head|Salad, burger, lettuce wraps|Water; vitamin K; small amounts of folate|Chill for maximum crunch|Vegan
Radicchio, Treviso|Leafy chicory variety|Veneto, Italy|Bitter, peppery, lightly sweet|Tight leaves; holds shape on grill|Grill, salad, risotto, roast|Fiber; vitamin K; vitamin C|Balance bitterness with sweet dressing or cheese|Vegan
Tomato, cherry|Fruit vegetable|Americas; Mediterranean; global|Sweet, tangy, concentrated|Juicy, thin skin|Salad, roast, pasta, skewer|Vitamin C; lycopene; potassium|Roast whole to concentrate sweetness|Vegan
Tomato, tomatillo-style canned|Prepared fruit vegetable|Mexican; global pantry|Tart, bright, savory|Soft chunks in juice|Salsa, chili, stew, braise|Vitamin C; sodium varies|Drain only if recipe needs less liquid|Vegan
Mushroom, oyster|Fungus|East Asian; global cultivation|Mild, savory, lightly sweet|Velvety, tender; edges crisp|Stir-fry, roast, soup, pasta|B vitamins; copper; potassium|Tear into strips; avoid washing until just before use|Vegan; not botanically a vegetable
Mushroom, enoki|Fungus|East Asian|Mild, delicate, lightly fruity|Thin strands, crisp-tender|Hot pot, soup, noodle dishes, salad|B vitamins; fiber; potassium|Trim root base; cook thoroughly|Vegan; not botanically a vegetable
`),
  'Fruit & Citrus': parse(`
Apple|Pome fruit|Central Asian heritage; European; global|Sweet-tart, aromatic|Crisp; softens in baking|Snack, salad, chutney, dessert, sauce|Fiber; vitamin C; polyphenols|Acidulated water slows browning|Vegan
Apricot|Stone fruit|Central Asian heritage; Middle Eastern|Sweet-tart, floral|Soft, juicy; dried form chewy|Jam, tagine, baking, snack|Vitamin A precursors; fiber; potassium|Dried fruit is concentrated in sugars|Vegan
Banana|Tropical fruit|Southeast Asian heritage; tropical regions|Sweet, creamy|Soft; starchier when green|Breakfast, baking, smoothie, curry|Potassium; vitamin B6; carbohydrate|Green and ripe stages behave differently|Vegan
Blueberry|Berry|North American; European|Sweet-tart, floral|Juicy, delicate|Breakfast, dessert, sauce, salad|Fiber; vitamin C; anthocyanins|Wash just before use|Vegan
Cherry|Stone fruit|West Asian heritage; European|Sweet or tart, floral|Juicy, tender|Dessert, compote, sauce, salad|Vitamin C; potassium; polyphenols|Remove pits before serving|Vegan
Coconut|Drupe|South Asian; Southeast Asian; Caribbean|Sweet, nutty, rich|Flesh firm; milk creamy|Curries, desserts, drinks, baking|Fat; fiber; manganese|Coconut milk varies widely in fat|Vegan; tree-nut labeling varies by jurisdiction
Date|Drupe fruit|Middle Eastern; North African|Caramel, honey, very sweet|Chewy, sticky|Snack, baking, sauces, tagine|Carbohydrate; fiber; potassium|Pit before chopping; very concentrated sweetness|Vegan
Fig|Multiple fruit|Mediterranean; Middle Eastern|Honeyed, berry-like|Soft, seed-crunch|Cheese boards, salad, jam, dessert|Fiber; potassium; calcium|Very ripe figs are delicate|Vegan
Grape|Berry|Mediterranean; West Asian heritage|Sweet, tart, floral|Juicy; skins add snap|Snack, salad, roast, preserve|Carbohydrate; vitamin K; polyphenols|Seedless varieties simplify service|Vegan
Grapefruit|Citrus|Caribbean; Mediterranean|Tart, bitter, aromatic|Juicy segments|Breakfast, salad, sauce, drink|Vitamin C; fiber|Can interact with some medicines; check personal advice|Vegan
Guava|Tropical fruit|Central American; Caribbean; South Asian|Perfumed, sweet-tart|Soft flesh; seedy center|Chutney, juice, dessert, snack|Vitamin C; fiber; folate|Edible seeds and skin on many varieties|Vegan
Lemon|Citrus|South Asian heritage; Mediterranean|Strong sour, floral zest|Juicy; zest is aromatic|Acid, finishing, curd, pickle, marinade|Vitamin C; small amounts of folate|Zest before juicing; avoid bitter pith|Vegan
Lime|Citrus|South Asian heritage; Southeast Asian; Latin American|Tart, sharp, aromatic|Juicy; fragrant zest|Ceviche-style acid, salsa, curry, drinks|Vitamin C|Add juice near finish for fresh aroma|Vegan
Mango|Tropical drupe|South Asian; Southeast Asian; Latin American|Sweet, floral, tangy|Juicy ripe; firm green|Chutney, salsa, dessert, curry, salad|Vitamin C; folate; beta-carotene|Ripeness changes sweetness and acidity|Vegan
Melon, cantaloupe|Melon|North African; Mediterranean|Sweet, musky|Soft, juicy|Salad, dessert, chilled soup|Vitamins A and C; water|Chill and handle safely after cutting|Vegan
Orange|Citrus|East Asian heritage; Mediterranean|Sweet-tart, fragrant|Juicy segments|Snack, zest, sauce, baking|Vitamin C; folate; fiber|Zest only colored peel|Vegan
Papaya|Tropical fruit|Central American; Caribbean; Southeast Asian|Mild, musky, sweet|Soft ripe; crisp green|Salad, salsa, smoothie, dessert|Vitamin C; folate; beta-carotene|Green and ripe papaya are used differently|Vegan
Peach|Stone fruit|Chinese heritage; Mediterranean|Floral, sweet-tart|Juicy, tender|Dessert, grill, salad, preserve|Vitamin C; potassium; fiber|Ripen at room temperature|Vegan
Pear|Pome fruit|European; East Asian|Sweet, delicate, floral|Juicy; some varieties grainy|Salad, poach, bake, cheese pairing|Fiber; vitamin C; copper|Choose firm for cooking, ripe for eating|Vegan
Pineapple|Tropical fruit|South American; Caribbean; Southeast Asian|Sweet, tart, aromatic|Juicy, fibrous|Salsa, grill, curry, dessert|Vitamin C; manganese|Core is edible but fibrous; enzyme tenderizes protein|Vegan
Plum|Stone fruit|East Asian heritage; European|Sweet-tart, floral|Juicy; skin may be tart|Sauce, jam, roast, dessert|Vitamin C; fiber; polyphenols|Cook firm tart varieties for sauces|Vegan
Pomegranate|Multiple fruit|Persian; Middle Eastern; South Asian|Sweet-tart, tannic|Juicy arils with crunch|Salad, rice, sauce, garnish|Fiber; vitamin C; polyphenols|Cut and tap out arils over a bowl|Vegan
Raspberry|Berry|European; North American|Bright tart, floral|Delicate, seedy|Dessert, jam, sauce, breakfast|Fiber; vitamin C; manganese|Very delicate; wash gently|Vegan
Strawberry|Aggregate fruit|European; North American; global|Sweet, aromatic, lightly tart|Juicy, tender|Dessert, salad, jam, breakfast|Vitamin C; manganese; folate|Keep stems on while washing|Vegan
Tamarind|Legume pod fruit|South Asian; African; Latin American|Sour, fruity, caramel-like|Sticky pulp|Chutney, pad thai-style sauces, stews, drinks|Carbohydrate; potassium; tartaric acid|Soak pulp and strain seeds/fibers|Vegan
Avocado|Berry|Mesoamerican; Mexican; global|Buttery, grassy, mild|Creamy flesh|Guacamole, salad, toast, sauce|Monounsaturated fat; fiber; folate; potassium|Cut near service; acid slows browning|Vegan
Blackberry|Berry|European; North American|Sweet-tart, wine-like|Juicy, seedy|Dessert, jam, sauce, breakfast|Fiber; vitamin C; manganese|Wash gently just before use|Vegan
Cranberry|Berry|North American|Tart, astringent|Firm raw; bursts when cooked|Sauce, relish, baking, drinks|Fiber; vitamin C; polyphenols|Usually sweetened for sauces due to tartness|Vegan
Kiwi|Berry|Chinese heritage; New Zealand cultivation|Bright sweet-tart, tropical|Juicy, tender with edible seeds|Snack, fruit salad, salsa, dessert|Vitamin C; fiber; potassium|Peel or halve and scoop; enzyme can affect gelatin|Vegan
Lychee|Tropical fruit|South Chinese; Southeast Asian|Floral, sweet, perfumed|Juicy, translucent flesh|Dessert, drinks, fruit salad, sauces|Vitamin C; carbohydrate|Remove shell and pit; canned versions may be syrup-packed|Vegan
Passion fruit|Tropical fruit|South American|Intensely tart, floral|Juicy pulp with crunchy seeds|Curd, drink, dessert, sauce|Fiber; vitamin C; vitamin A precursors|Spoon out pulp; strain if seedless texture is desired|Vegan
Dragon fruit / pitaya|Cactus fruit|Central American; Southeast Asian|Mild, lightly sweet|Juicy with tiny crunchy seeds|Fruit bowls, drinks, dessert, garnish|Fiber; vitamin C; magnesium|Ripe fruit gives slightly under finger pressure|Vegan
Persimmon|Pome-like fruit|East Asian|Honey-sweet, floral|Firm crisp or soft jelly-like by variety|Snack, salad, baking, preserve|Fiber; vitamin A precursors; manganese|Astringent varieties must fully ripen before eating|Vegan
Quince|Pome fruit|West Asian; Mediterranean|Floral, tart, woody raw|Hard raw; turns rosy and soft when cooked|Poach, paste, jam, braise|Fiber; vitamin C; copper|Usually cooked; peel and core before poaching|Vegan
Starfruit / carambola|Tropical fruit|Southeast Asian; South Asian|Tart-sweet, citrusy|Crisp, juicy|Salad, garnish, pickle, chutney|Vitamin C; fiber; potassium|Avoid if medically advised with kidney disease|Vegan
Jackfruit|Multiple fruit|South Asian; Southeast Asian|Sweet ripe; mild and savory unripe|Ripe pods are fibrous; young fruit shreds|Dessert, curry, stew, plant-based filling|Carbohydrate; fiber; potassium|Canned young jackfruit is usually brined; rinse before seasoning|Vegan
Durian|Tropical fruit|Southeast Asian|Very rich, sweet, sulfurous aroma|Custardy, creamy|Dessert, sticky rice, frozen treats|Carbohydrate; fat; potassium|Strong aroma; use ripe flesh and follow local handling guidance|Vegan
Rambutan|Tropical fruit|Southeast Asian|Sweet, floral, lightly acidic|Juicy, grape-like flesh|Snack, fruit salad, dessert|Vitamin C; carbohydrate|Peel shell and remove seed|Vegan
Longan|Tropical fruit|South China; Southeast Asian|Sweet, musky, floral|Translucent juicy flesh|Snack, soup, dessert, drinks|Vitamin C; carbohydrate|Remove shell and seed; canned versions may be syrup-packed|Vegan
Kumquat|Citrus|East Asian; Mediterranean|Sweet peel, tart pulp|Small, juicy, edible peel|Marmalade, salad, garnish, preserve|Vitamin C; fiber|Wash well; often eaten whole|Vegan
Pomelo|Citrus|Southeast Asian|Mild, sweet-bitter, floral|Large dry-ish segments|Salad, snack, dessert, seafood pairing|Vitamin C; fiber|Remove thick pith and membranes|Vegan
Watermelon|Melon|African heritage; global|Sweet, refreshing|Juicy, crisp|Snack, salad, grill, chilled soup|Water; lycopene; vitamin C|Chill after cutting and handle safely|Vegan
Melon, honeydew|Melon|African heritage; Mediterranean|Mild, sweet|Juicy, smooth|Fruit salad, dessert, prosciutto pairing|Vitamin C; potassium; water|Choose fragrant fruit with slight stem-end give|Vegan
Mulberry|Berry|West Asian; East Asian; Mediterranean|Sweet, tart, wine-like|Soft, delicate|Jam, dessert, snack, sauce|Fiber; vitamin C; iron|Handle gently; stains readily|Vegan
Black currant|Berry|Northern European; Central Asian|Tart, intense, cassis-like|Small, juicy berries|Jam, cordial, sauce, baking|Vitamin C; fiber; anthocyanins|Strong flavor suits sweetened preserves|Vegan
Red currant|Berry|European|Tart, bright, lightly floral|Small translucent berries|Jelly, garnish, sauce, dessert|Vitamin C; fiber|Strip from stems with a fork|Vegan
Gooseberry|Berry|European; South Asian|Tart green; sweet when ripe|Firm, juicy|Jam, chutney, pie, compote|Fiber; vitamin C; potassium|Cook tart green berries or eat ripe ones fresh|Vegan
Elderberry|Berry|European; North American|Tart, earthy, floral|Small dark berries|Syrup, jam, cooked sauce|Fiber; vitamin C; anthocyanins|Cook before eating; avoid raw stems/leaves|Vegan
Prune / dried plum|Dried fruit|European; West Asian|Caramel, sweet, tangy|Soft, chewy|Braising, baking, snack, compote|Fiber; potassium; vitamin K|Pit if necessary; hydrates well in warm liquid|Vegan
Raisin|Dried grape|Mediterranean; West Asian|Sweet, fruity|Chewy|Baking, pilaf, salad, snack|Carbohydrate; fiber; potassium|Soak briefly to plump for baking|Vegan
Sultana / golden raisin|Dried grape|Mediterranean; West Asian|Sweet, honeyed|Soft, chewy|Baking, pilaf, salad, snack|Carbohydrate; fiber; potassium|May be oil-treated to prevent sticking|Vegan
Yuzu|Citrus|Japanese; East Asian|Floral, tart, complex citrus|Very aromatic zest; modest juice|Ponzu-style sauce, dressing, dessert, garnish|Vitamin C; small amounts of fiber|Use zest and juice; bottled juice varies|Vegan
Bergamot|Citrus|Mediterranean|Floral, bitter, fragrant|Aromatic peel|Tea, marmalade, baking, dressing|Vitamin C; small amounts of fiber|Used mainly for aromatic zest and oil|Vegan
Blood orange|Citrus variety|Mediterranean; Italian|Sweet-tart, berry-like|Juicy, red-flecked flesh|Salad, dessert, juice, sauce|Vitamin C; anthocyanins; fiber|Color intensity varies by variety and season|Vegan
Mandarin|Citrus|East Asian; Mediterranean|Sweet, floral|Easy-peel segments|Snack, salad, dessert, sauce|Vitamin C; folate; fiber|Remove seeds as needed|Vegan
Clementine|Citrus hybrid|Mediterranean; North African|Sweet, mild|Easy-peel, juicy|Snack, salad, dessert, garnish|Vitamin C; fiber|Peel shortly before serving to preserve aroma|Vegan
Nectarine|Stone fruit|Chinese heritage; Mediterranean|Floral, sweet-tart|Smooth skin; juicy flesh|Snack, grill, salad, baking|Vitamin C; potassium; fiber|Ripen at room temperature|Vegan
Dried apricot|Dried stone fruit|West Asian; Middle Eastern|Sweet-tart, concentrated|Chewy, dense|Tagine, baking, snack, chutney|Fiber; potassium; beta-carotene|Sulfite-treated and unsulfured varieties differ|May contain sulfites
Golden kiwi|Berry|New Zealand; Chinese heritage|Sweet, tropical, low tartness|Soft, juicy|Snack, fruit bowl, dessert, salsa|Vitamin C; fiber; potassium|Peel or eat skin if washed and desired|Vegan
Pomelo, pink|Citrus variety|Southeast Asian|Mild sweet-tart, floral|Large juicy segments|Salad, fruit plate, seafood pairing|Vitamin C; fiber|Remove thick membranes for best texture|Vegan
Candied citrus peel|Preserved citrus|European; Middle Eastern|Sweet-bitter, fragrant|Chewy, glossy|Baking, confectionery, garnish|Mostly carbohydrate; citrus compounds|Chop finely; rinse if overly sweet|Vegan
`),
  'Grains & Pulses': parse(`
Amaranth|Pseudocereal|Mesoamerican; South Asian|Earthy, nutty|Tiny; porridge-like when cooked|Porridge, pilaf, popped garnish, flour|Protein; fiber; iron; magnesium|Rinse; use extra liquid for porridge|Naturally gluten-free; cross-contact possible
Barley|Whole grain|West Asian heritage; European|Nutty, malty|Chewy; thickens soups|Soup, pilaf, salad, malt|Fiber (beta-glucan); selenium; manganese|Pearled cooks faster but has less bran|Contains gluten
Buckwheat|Pseudocereal|East European; Japanese|Earthy, nutty|Tender; toasted groats stay distinct|Kasha, soba, pancakes, porridge|Fiber; magnesium; manganese|Toast groats for deeper flavor|Naturally gluten-free; soba may contain wheat
Bulgur|Parboiled wheat|Middle Eastern; Mediterranean|Wheaty, nutty|Quick-cooking, fluffy|Tabbouleh, pilaf, salads|Fiber; manganese; magnesium|Fine grade can steep; coarse needs simmering|Contains gluten
Chickpeas|Pulse|Middle Eastern; South Asian; Mediterranean|Nutty, earthy|Creamy inside; crisp when roasted|Hummus, curry, salad, stew, falafel|Protein; fiber; folate; iron|Soaking reduces cooking time; canned is convenient|Vegan; legume sensitivity possible
Cornmeal / polenta|Milled grain|Mesoamerican; Italian|Sweet, earthy|Creamy porridge or firm when cooled|Polenta, breading, grits, baking|Carbohydrate; fiber; thiamin|Stir gradually to limit lumps|Naturally gluten-free; check cross-contact
Couscous|Wheat pasta granule|Maghrebi; North African|Mild, wheaty|Light, fluffy|Tagine base, salads, quick side|Carbohydrate; selenium; protein|Steam or hydrate per package|Contains gluten
Farro|Ancient wheat|Italian; Mediterranean|Nutty, pleasantly sour|Chewy, intact grains|Salad, soup, risotto-style dishes|Fiber; protein; magnesium|Soak if desired; cooking time varies by type|Contains gluten
Lentils, brown/green|Pulse|Middle Eastern; South Asian; Mediterranean|Earthy, peppery|Hold shape when simmered|Soup, salad, stew, dal|Protein; fiber; folate; iron|Rinse; acid late if slow to soften|Vegan; legume sensitivity possible
Lentils, red|Pulse|South Asian; Middle Eastern|Sweet, earthy|Break down to creamy puree|Dal, soup, thickener, puree|Protein; fiber; folate; iron|Rinse; cook quickly without soaking|Vegan; legume sensitivity possible
Millet|Whole grain|African; South Asian|Mild, sweet, nutty|Fluffy or porridge-like|Pilaf, porridge, flatbread, side|Magnesium; phosphorus; fiber|Toast before cooking for nuttiness|Naturally gluten-free
Oats|Whole grain|Northern European; global|Creamy, toasted, mild|Soft, sticky to chewy|Porridge, granola, baking, savory oats|Soluble fiber (beta-glucan); manganese; iron|Steel-cut, rolled and instant have different cook times|Naturally gluten-free; often cross-contaminated
Quinoa|Pseudocereal|Andean|Earthy, nutty|Light, fluffy, slight pop|Salad, bowls, pilaf, breakfast|Protein; fiber; magnesium; folate|Rinse well to remove bitter saponins|Naturally gluten-free
Rice, basmati|Rice|South Asian; Persian|Fragrant, nutty|Long, separate grains|Biryani, pilaf, curry accompaniment|Carbohydrate; small amounts of B vitamins|Rinse; soak for even long-grain cooking|Naturally gluten-free
Rice, brown|Whole grain rice|East Asian; global|Nutty, earthy|Chewy, bran layer intact|Bowls, pilaf, sides|Fiber; manganese; magnesium|Needs more water and time than white rice|Naturally gluten-free
Rice, jasmine|Rice|Thai; Southeast Asian|Floral, lightly sweet|Soft, slightly sticky|Thai curries, rice plates|Carbohydrate; small amounts of B vitamins|Rinse lightly; avoid over-water|Naturally gluten-free
Rice, short-grain|Rice|Japanese; Korean|Mild, sweet|Clings when cooked|Sushi, rice bowls, congee|Carbohydrate; small amounts of B vitamins|Wash until water is less cloudy; use correct water ratio|Naturally gluten-free
Rye|Whole grain|Northern/Eastern European|Sour, earthy|Dense, chewy|Bread, crispbread, porridge|Fiber; manganese; selenium|Flavor pairs well with caraway|Contains gluten
Sorghum|Whole grain|African; South Asian|Mild, sweet, nutty|Chewy whole grain; pops|Porridge, flour, pilaf, popped snack|Fiber; iron; antioxidants|Soak overnight for faster cooking|Naturally gluten-free
Teff|Whole grain|Ethiopian; Eritrean|Earthy, cocoa-like|Tiny; creamy porridge|Injera, porridge, flatbread|Iron; calcium; fiber|Traditionally fermented for injera|Naturally gluten-free; fermentation varies
Wheat flour, all-purpose|Milled grain|West Asian heritage; global|Mild, wheaty|Forms gluten; structure in dough|Bread, pastry, noodles, thickening|Carbohydrate; protein; fortified micronutrients vary|Measure by weight for baking accuracy|Contains gluten
Wheat flour, whole wheat|Whole grain flour|West Asian heritage; global|Nutty, earthy|Absorbs more water; denser crumb|Bread, flatbread, pasta, baking|Fiber; manganese; magnesium|Blend with refined flour for lighter texture|Contains gluten
White beans|Pulse|Mediterranean; Americas|Mild, creamy|Soft, mashable|Soup, cassoulet-style stew, dip|Protein; fiber; folate; iron|Cook until fully tender|Vegan; legume sensitivity possible
Black beans|Pulse|Central American; Caribbean|Earthy, savory|Creamy, hold shape|Tacos, bowls, soup, rice dishes|Protein; fiber; folate; iron|Rinse canned beans; retain liquid for soup|Vegan; legume sensitivity possible
Kidney beans|Pulse|Central American; South Asian|Earthy, robust|Firm, holds shape|Chili, rajma, salads, stew|Protein; fiber; folate; potassium|Boil raw soaked beans thoroughly; never slow-cook from raw alone|Vegan; legume sensitivity possible
Peanuts|Legume seed|South American; West African; South Asian|Roasty, savory, sweet|Crunchy or creamy as paste|Sauce, snack, garnish, stew, sweets|Protein; unsaturated fat; niacin; vitamin E|Roast gently; peanut oil can be refined or unrefined|Major allergen
Black-eyed peas|Pulse|West African; Southern US; South Asian|Earthy, mild|Creamy, holds shape|Stew, fritters, rice dishes, salad|Protein; fiber; folate; iron|Soak optional; simmer until tender|Vegan; legume sensitivity possible
Pigeon peas|Pulse|South Asian; Caribbean; East African|Nutty, earthy|Creamy, hearty|Rice dishes, dal, curry, stew|Protein; fiber; folate; potassium|Dried peas need soaking; frozen/fresh cook faster|Vegan; legume sensitivity possible
Split peas, green|Pulse|European; South Asian|Earthy, sweet|Breaks down to thick puree|Soup, stew, dal, puree|Protein; fiber; folate; potassium|Rinse; simmer until soft; no soak usually needed|Vegan; legume sensitivity possible
Split peas, yellow|Pulse|South Asian; Middle Eastern; European|Mild, earthy|Creamy puree|Dal, soup, stew, fritters|Protein; fiber; folate; iron|Rinse and simmer until tender|Vegan; legume sensitivity possible
Mung beans|Pulse|South Asian; East Asian|Mild, sweet, nutty|Whole beans hold; split mung softens|Dal, sprouts, porridge, dessert|Protein; fiber; folate; magnesium|Soaking shortens cooking; split form cooks quickly|Vegan; legume sensitivity possible
Adzuki beans|Pulse|East Asian|Sweet, nutty|Small, firm; paste becomes smooth|Sweet bean paste, rice, soup, stew|Protein; fiber; folate; potassium|Soak and cook until fully tender|Vegan; legume sensitivity possible
Fava / broad beans|Pulse|Mediterranean; Middle Eastern; North African|Sweet, grassy, earthy|Creamy inside; skin may be firm|Puree, stew, salad, pasta|Protein; fiber; folate; iron|Large beans may need outer skin removed|Vegan; legume allergy possible
Lima beans|Pulse|Andean; Caribbean; Southern US|Mild, buttery|Creamy, starchy|Succotash, soup, stew, puree|Protein; fiber; folate; potassium|Cook dried beans thoroughly; avoid raw|Vegan; legume sensitivity possible
Pinto beans|Pulse|Mexican; Central American; Southern US|Earthy, creamy|Soft; mashes easily|Refried beans, chili, burritos, soup|Protein; fiber; folate; iron|Rinse canned beans; simmer dried beans until tender|Vegan; legume sensitivity possible
Navy beans|Pulse|North American; European|Mild, creamy|Small, soft; thickens broth|Baked beans, soup, cassoulet-style dishes|Protein; fiber; folate; iron|Soak dried beans for more even cooking|Vegan; legume sensitivity possible
Soybeans, mature|Pulse|East Asian|Nutty, savory|Firm cooked bean|Soy milk, tofu, natto, stew|Protein; fiber; iron; folate|Cook thoroughly; raw dried beans are not edible|Soy allergen
Freekeh|Roasted green wheat|Levantine; Middle Eastern|Smoky, grassy, nutty|Chewy, intact grain|Pilaf, salad, soup, stuffing|Fiber; protein; manganese|Rinse and simmer; cooking time varies by cracked/whole form|Contains gluten
Wheat berries|Whole wheat kernel|West Asian heritage; global|Nutty, sweet|Very chewy, intact|Grain salad, pilaf, soup|Fiber; protein; manganese|Soak overnight for faster cooking|Contains gluten
Semolina|Durum wheat product|Italian; North African; South Asian|Wheaty, nutty|Coarse flour; sets firmly|Pasta, couscous, porridge, pudding|Carbohydrate; protein; selenium|Different grinds suit pasta or porridge|Contains gluten
Rice noodles|Rice-based noodle|East Asian; Southeast Asian|Mild, neutral|Slippery, tender-chewy|Pho-style soups, stir-fry, salads|Mostly carbohydrate; low fat|Soak or boil according to thickness and package|Usually gluten-free; check cross-contact and sauces
Soba noodles|Buckwheat noodle|Japanese|Nutty, earthy|Firm, slightly chewy|Cold noodles, soup, stir-fry|Carbohydrate; protein; manganese|Rinse after cooking for cold service|Many commercial soba noodles contain wheat
Wheat pasta|Wheat product|Italian; global|Mild, wheaty|Al dente chew|Pasta mains, soups, baked dishes|Carbohydrate; protein; B vitamins vary|Salt water; reserve cooking water for sauce|Contains gluten
Rice paper wrappers|Rice starch wrapper|Vietnamese; Southeast Asian|Mild, neutral|Soft and pliable when hydrated|Fresh rolls, fried rolls, wraps|Mostly carbohydrate|Dip briefly in warm water; avoid over-soaking|Usually gluten-free; check label
Corn tortillas|Masa flatbread|Mesoamerican; Mexican|Sweet corn, toasted|Soft, pliable; chars well|Tacos, enchiladas, tostadas|Carbohydrate; fiber; calcium if lime-treated|Warm before filling to prevent cracking|Usually gluten-free; cross-contact possible
Sourdough bread|Fermented wheat bread|European; global craft baking|Tangy, wheaty|Chewy crumb, crisp crust|Toast, sandwiches, sides, croutons|Carbohydrate; protein; fiber varies|Long fermentation changes flavor, not necessarily gluten status|Usually contains gluten
Rice, arborio|Short-grain rice|Northern Italian|Mild, starchy|Creamy outside, firm center|Risotto, rice pudding, arancini|Carbohydrate; small amounts of protein|Add warm liquid gradually for risotto texture|Naturally gluten-free
Rice, sticky / glutinous|Glutinous rice|East Asian; Southeast Asian|Mild, lightly sweet|Very sticky when cooked|Mochi, sticky rice, dumplings, desserts|Carbohydrate; small amounts of B vitamins|Soak before steaming; glutinous refers to texture, not gluten protein|Naturally gluten-free
Rice, wild|Aquatic grass seed|North American; Great Lakes region|Earthy, nutty|Long, chewy, often splits open|Pilaf, salad, soup, stuffing|Protein; fiber; manganese; phosphorus|Rinse; longer cooking than most rice|Naturally gluten-free
Rice, black / forbidden|Whole grain rice|Chinese; East Asian|Nutty, earthy, lightly sweet|Chewy; releases purple color|Rice bowl, porridge, dessert, salad|Fiber; anthocyanins; iron|Rinse well; may need longer soak/cook|Naturally gluten-free
Rice, parboiled|Processed rice|South Asian; global|Mild, nutty|Separate, firm grains|Pilaf, biryani, meal prep|Carbohydrate; some B vitamins|Cook according to package; holds shape well|Naturally gluten-free
Rice flour|Milled grain flour|East Asian; South Asian|Mild, neutral|Fine, starch-rich|Noodles, batters, baking, thickener|Carbohydrate; small amounts of protein|White and brown rice flour differ in fiber and texture|Naturally gluten-free; check cross-contact
Chickpea flour / besan|Pulse flour|South Asian; Mediterranean|Nutty, earthy|Dense batter; binds when cooked|Pakora, socca, flatbread, thickener|Protein; fiber; folate; iron|Whisk with liquid and rest for smoother batter|Naturally gluten-free; legume sensitivity possible
Corn flour / masa harina|Nixtamalized corn flour|Mesoamerican; Mexican|Sweet corn, toasted|Fine flour; forms pliable dough with water|Tortillas, tamales, arepas, dumplings|Carbohydrate; fiber; calcium varies with nixtamalization|Masa harina is treated differently from plain cornmeal|Naturally gluten-free; check cross-contact
Potato starch|Starch|European; East Asian|Neutral|Light, crisp coating; thickens glossy|Baking, noodle making, frying, sauces|Mostly carbohydrate|Add slurry near end; high heat can thin the gel|Usually gluten-free; verify packaging
Tapioca pearls / boba|Cassava starch pearls|Taiwanese; Southeast Asian|Mild, sweet when soaked|Chewy, springy|Milk tea, pudding, dessert soups|Mostly carbohydrate|Cook and sweeten per package; best fresh|Usually gluten-free; check additives
Ramen noodles|Wheat noodle|Japanese; Chinese origins|Wheaty, savory|Springy, alkaline chew|Soup, stir-fry, cold noodle dishes|Carbohydrate; protein; sodium varies by pack|Fresh, dried and instant products differ greatly|Usually contains wheat/gluten; instant soup can be high sodium
Udon noodles|Thick wheat noodle|Japanese|Mild, wheaty|Thick, soft-chewy|Soup, hot pot, stir-fry|Carbohydrate; protein|Frozen or vacuum-packed noodles heat quickly|Contains gluten
Egg noodles|Wheat-and-egg pasta|Central European; East Asian|Rich, mild|Tender, springy|Soup, stir-fry, casserole, buttered side|Carbohydrate; protein; choline|Cooking time varies by width and freshness|Contains wheat/gluten and egg
Glass noodles|Starch noodle|Chinese; Korean; Southeast Asian|Neutral, absorbs sauce|Translucent, slippery-chewy|Japchae-style dishes, hot pot, soup, salad|Mostly carbohydrate|Soak or boil by starch type; mung and sweet potato varieties differ|Often gluten-free; check seasoning packets
Phyllo pastry|Layered wheat pastry|Greek; Turkish; Middle Eastern|Buttery when baked, mild|Paper-thin, flaky-crisp|Baklava, spanakopita, savory pies|Carbohydrate; fat depends on brushing|Cover unused sheets; brush each layer with fat|Contains wheat/gluten
Puff pastry|Laminated pastry|French; European|Buttery, rich|Crisp, flaky layers|Tart, pie, vol-au-vent, pastry|Carbohydrate; fat varies by butter/shortening|Keep cold before baking for lift|Contains wheat/gluten; often dairy
Pita bread|Leavened flatbread|Levantine; Middle Eastern|Mild, wheaty|Soft pocket bread|Wrap, dip, sandwich, pizza base|Carbohydrate; protein; fiber varies|Warm briefly to soften|Usually contains wheat/gluten
Naan|Leavened flatbread|South Asian|Soft, lightly tangy, often buttery|Tender, blistered|Curry accompaniment, wrap, pizza-style base|Carbohydrate; protein; fat varies|Reheat in hot pan or oven|Usually wheat/gluten; often dairy
Lavash|Thin flatbread|Armenian; West Asian; Middle Eastern|Mild, wheaty|Soft when fresh; crisp when baked|Wrap, sandwich, cracker, flatbread|Carbohydrate; fiber varies|Keep covered to prevent drying|Usually contains wheat/gluten
Arepa / precooked cornmeal|Corn dough|Venezuelan; Colombian|Sweet corn, toasted|Crisp outside, tender inside|Stuffed cake, side, sandwich|Carbohydrate; fiber; calcium varies|Use precooked corn flour, not raw cornmeal|Naturally gluten-free; cross-contact possible
`),
  'Spices & Herbs': parse(`
Ajwain|Seed spice|South Asian|Thyme-like, sharp, warm|Tiny seed, potent aroma|Paratha, pakora, lentils, breads|Small culinary quantities; not a major nutrient source|Crush or bloom in hot fat|Vegan
Allspice|Dried berry|Caribbean; Central American|Clove, cinnamon, nutmeg-like|Warm, rounded|Jerk seasoning, pickles, baking, stews|Small culinary quantities; manganese trace|Use whole for infusions, ground for rubs|Vegan
Anise seed|Seed spice|Mediterranean; Middle Eastern|Sweet licorice|Fragrant, slightly woody|Baking, tea, sausage, braise|Small culinary quantities|Toast lightly; use sparingly|Vegan
Asafoetida (hing)|Resin spice|South Asian|Allium-like, savory after cooking|Powder/resin; blooms in fat|Dal, legumes, vegetarian tempering|Negligible in pinch amounts|Cook briefly in oil; raw smell is strong|Some blends contain wheat; check label
Bay leaf|Leaf spice|Mediterranean; Indian|Herbal, woody, tea-like|Infuses slowly; remove before serving|Stocks, beans, soups, braise|Small culinary quantities|Add whole early; remove before plating|Vegan
Black cardamom|Seed pod|Himalayan; North Indian|Smoky, resinous, camphor|Large pod, very potent|Biryani, slow-cooked meat, dals|Small culinary quantities|Crack pod; use for background aroma|Vegan
Black pepper|Berry spice|South Indian; global|Pungent, woody, warm|Fresh grind is aromatic|Finishing, marinades, sauces, spice blends|Small culinary quantities; piperine|Grind near service for lively aroma|Vegan
Cardamom, green|Seed pod|South Asian; Middle Eastern|Floral, citrus, sweet spice|Fragrant pod and seeds|Biryani, chai, desserts, rice pudding|Small culinary quantities|Crush pods; use seeds ground in pastry|Vegan
Caraway|Seed spice|Central/Eastern European|Earthy, anise, citrus|Firm seed|Rye bread, cabbage, cheese, stew|Small culinary quantities|Toast or crush to release oils|Vegan
Cayenne|Ground chili|Americas; global|Hot, sharp|Fine powder disperses quickly|Hot sauces, rubs, curries|Capsaicin; tiny nutrient amount per use|Add gradually; heat varies|Vegan
Celery seed|Seed spice|European; North American|Concentrated celery, bitter|Small seeds, strong aroma|Pickles, coleslaw, spice rubs|Small culinary quantities|Use sparingly; grind for dressing|Vegan
Chili flakes|Dried chili|Italian; global|Hot, fruity, variable|Flakes add visible heat|Pizza, oil, pasta, sauces|Capsaicin; tiny nutrient amount per use|Bloom in oil for broader aroma|Vegan
Chili powder, ancho|Dried chili|Mexican|Mild heat, raisin, cocoa|Soft, fruity powder|Mole-style sauce, rub, beans|Capsaicin; tiny nutrient amount per use|Toast briefly; do not scorch|Vegan
Cinnamon|Bark spice|Sri Lankan; South Asian; Middle Eastern|Sweet, woody, warm|Powder or rolled bark|Desserts, chai, tagine, savory braise|Manganese trace; small culinary quantities|Ceylon and cassia differ in flavor|Vegan
Clove|Flower bud|Indonesian; Indian; Middle Eastern|Intense, sweet, numbing|Very potent whole bud|Baking, rice, ham, spice blends|Small culinary quantities|Use whole for easy removal|Vegan
Coriander seed|Seed spice|South Asian; Middle Eastern|Citrus, nutty, warm|Round seed; toasts well|Curries, pickles, rubs, breads|Small culinary quantities|Toast and grind for fresher flavor|Vegan
Cumin|Seed spice|South Asian; Middle Eastern; Mexican|Earthy, warm, toasted|Seed or powder|Dal, tacos, rice, spice blends|Iron trace; small culinary quantities|Bloom seeds in oil or toast dry|Vegan
Curry leaves|Leaf herb/spice|South Indian; Sri Lankan|Citrus, nutty, herbal|Delicate leaf crisps in oil|Tempering, chutney, soups, rice|Small culinary quantities|Add to hot oil; fresh leaves freeze well|Vegan
Dill seed|Seed spice|Eastern European; Nordic|Caraway-like, grassy|Firm seed|Pickles, bread, fish, potato|Small culinary quantities|Toast or steep in pickle brine|Vegan
Fennel seed|Seed spice|Indian; Italian; Mediterranean|Sweet anise, minty|Crunchy seed|Sausage, curry, tea, bread|Small culinary quantities|Toast; often paired with cumin|Vegan
Fenugreek seed|Seed spice|South Asian; Ethiopian|Bitter, maple-like when cooked|Hard seed; softens after soaking|Curry blends, pickles, breads|Fiber; iron trace in small amounts|Use sparingly; toast to mellow bitterness|Vegan; legume-family cross-reactivity possible
Fenugreek leaf, dried|Leaf spice|North Indian; Pakistani|Maple, celery, bitter-herbal|Crumbly leaf|Butter-style curries, flatbreads, lentils|Small culinary quantities|Crush between palms before adding|Vegan
Garam masala|Spice blend|North Indian; South Asian|Warm, fragrant, blend-dependent|Fine powder|Finish curries, lentils, roasted vegetables|Nutrient mix depends on recipe; use as seasoning|Blend varies; add late for aroma|Check blend for allergens
Ginger, fresh|Rhizome|South Asian; East Asian|Hot, citrusy, fresh|Fibrous, juicy|Stir-fries, curry base, tea, marinades|Small amounts of vitamin B6 and manganese|Grate against fibers; freeze for easy grating|Vegan
Ginger, ground|Dried rhizome|South Asian; baking traditions|Warm, dry, sweet-hot|Fine powder|Baking, spice rub, drinks|Small culinary quantities|Ground form is warmer and less bright than fresh|Vegan
Kashmiri chili|Dried chili|Kashmir; North Indian|Mild-medium heat, fruity, red color|Fine powder or whole chili|Rogan josh-style dishes, color, marinades|Capsaicin; tiny nutrient amount per use|Color is strong; heat varies by source|Vegan
Mustard seed|Seed spice|South Asian; European|Pungent, sharp when crushed|Pops and blooms in oil|Tempering, pickles, dressings, rubs|Small culinary quantities|Heat until popping; avoid burning|Mustard is a major allergen in some regions
Nutmeg|Seed spice|Indonesian; European|Sweet, woody, warm|Finely grated powder|Baking, cream sauces, spinach, drinks|Small culinary quantities|Grate fresh; excess is undesirable|Vegan
Paprika, smoked|Dried pepper|Spanish|Sweet, smoky, mild|Fine powder|Rub, stew, potatoes, beans|Carotenoids; small amounts per use|Add away from scorching heat|Vegan
Paprika, sweet|Dried pepper|Hungarian; Spanish|Sweet, earthy, mild|Fine powder|Goulash, eggs, rub, sauces|Carotenoids; small amounts per use|Store sealed and away from light|Vegan
Saffron|Flower stigma|Iranian; Kashmiri; Spanish|Honeyed, floral, hay-like|Fine threads, high impact|Rice, broth, dessert, seafood|Negligible nutrient contribution per serving|Bloom in warm liquid before adding|Vegan
Star anise|Fruit spice|Chinese; Vietnamese|Licorice, sweet, woody|Hard star-shaped pod|Broth, braise, five-spice, poached fruit|Small culinary quantities|Use whole and remove before serving|Vegan
Sumac|Dried berry spice|Levantine; Middle Eastern|Tart, fruity, lemony|Coarse red powder|Salad, grilled meat, dips, onions|Small culinary quantities|Add at finish to preserve tang|Vegan
Turmeric, fresh|Rhizome|South Asian; Southeast Asian|Earthy, peppery, bitter|Firm rhizome, stains strongly|Curry base, rice, soup, pickle|Manganese; iron; curcuminoids (small food amounts)|Wear gloves; pair with fat for flavor dispersion|Vegan
Turmeric, ground|Dried rhizome|South Asian; global|Earthy, warm, mildly bitter|Fine powder|Curries, rice, lentils, marinades|Small culinary quantities|Bloom briefly in oil; avoid large amounts|Vegan
Basil, sweet|Leaf herb|Italian; Mediterranean|Sweet, clove-like, peppery|Soft leaf; bruises easily|Pesto, tomato dishes, salads, garnish|Vitamin K; small amounts of vitamin A|Add late; avoid prolonged heat|Vegan
Cilantro / coriander leaf|Leaf herb|South Asian; Latin American; Middle Eastern|Citrusy, green; flavor perception varies|Tender leaf and stems|Chutney, salsa, curry, garnish|Vitamin K; vitamin C; folate|Use stems in blends; wash grit well|Vegan
Mint|Leaf herb|Middle Eastern; South Asian; Mediterranean|Cool, sweet, menthol|Tender, aromatic|Chutney, tea, salad, lamb, dessert|Vitamin A; small amounts of iron|Bruise gently; add late|Vegan
Oregano|Leaf herb|Greek; Italian; Mexican|Peppery, earthy, floral|Dried leaf is concentrated|Tomato sauce, beans, roast, marinades|Small culinary quantities|Dried is stronger than fresh|Vegan
Parsley|Leaf herb|Mediterranean; European|Fresh, grassy, lightly peppery|Crisp stems, tender leaf|Gremolata, salad, soup, garnish|Vitamin K; vitamin C; folate|Stems flavor stocks; add leaf late|Vegan
Epazote|Leaf herb|Mexican; Central American|Resinous, citrusy, oregano-like|Tender aromatic leaf|Beans, quesadillas, soups, salsa|Small culinary quantities|Use a modest amount; flavor is assertive|Vegan
Shiso / perilla leaf|Leaf herb|Japanese; Korean; East Asian|Minty, basil-like, anise notes|Soft aromatic leaf|Sushi, wraps, pickles, garnish|Small amounts of vitamins and minerals|Use fresh leaves near service; green and red varieties differ|Vegan
Pandan leaf|Aromatic leaf|Southeast Asian|Vanilla-like, grassy, toasted rice aroma|Fibrous leaf, infuses liquids|Rice, custard, cake, coconut desserts|Small culinary quantities|Tie into knots for infusion; remove before eating|Vegan
Chervil|Leaf herb|French; European|Mild parsley, anise|Delicate frilled leaf|Eggs, cream sauce, salad, garnish|Small culinary quantities|Add at the end; delicate flavor fades with heat|Vegan
Summer savory|Leaf herb|Balkan; Mediterranean|Peppery, thyme-like|Fine leaf|Beans, sausage, roast vegetables|Small culinary quantities|Add dried early or fresh near finish|Vegan
Winter savory|Leaf herb|Mediterranean; Central European|Peppery, piney, robust|Sturdier leaf|Beans, braise, sausage, stuffing|Small culinary quantities|Stands up to longer cooking|Vegan
Lemon balm|Leaf herb|European; West Asian|Lemony, minty, floral|Soft, delicate leaf|Tea, fruit salad, sauce, garnish|Small culinary quantities|Bruise gently; add late|Vegan
Shallot powder|Dried allium seasoning|European; global|Sweet onion, mild garlic|Fine powder|Rub, dip, dressing, quick sauce|Small culinary quantities|Use less than fresh shallot; check for fillers|Vegan
Aleppo pepper|Dried chili|Syrian; Turkish|Fruity, tangy, mild-medium heat|Coarse flakes, slightly oily|Eggs, kebab, hummus, roast vegetables|Capsaicin; small culinary quantities|Add at finish or bloom in oil|Vegan
Urfa biber|Dried chili|Southeastern Turkish|Smoky, raisin-like, mild heat|Dark, flaky flakes|Lamb, yogurt, roasted vegetables, eggs|Capsaicin; small culinary quantities|Often used as a finishing spice|Vegan
Chipotle chili|Smoked dried jalapeño|Mexican|Smoky, earthy, medium-hot|Whole dried pepper or adobo paste|Salsa, beans, marinade, stew|Capsaicin; small culinary quantities|Rehydrate dried pepper or mince canned chipotles|Vegan; canned adobo label may vary
Guajillo chili|Dried chili|Mexican|Mild, tangy, berry-like|Smooth leathery pod|Adobo, mole-style sauce, stew|Capsaicin; small culinary quantities|Toast briefly, soak, then blend|Vegan
Chile de árbol|Dried chili|Mexican|Clean, sharp, hot|Small thin pod|Salsa, chili oil, soup|Capsaicin; small culinary quantities|Use sparingly; heat is strong|Vegan
Annatto / achiote|Seed spice|Caribbean; Central American; Filipino|Earthy, peppery, lightly floral|Hard seeds or paste; adds color|Rice, marinades, recado, stews|Carotenoids; trace amounts in small servings|Infuse in oil; whole seeds are hard to grind|Check commercial paste for allergens
Ají amarillo|Fresh/dried chili|Peruvian; Andean|Fruity, tropical, medium heat|Thick-walled pepper or paste|Sauces, ceviche accompaniments, stew|Vitamin C; capsaicin|Frozen and jarred pastes are common substitutes|Vegan; check paste labels
Ají panca|Dried chili|Peruvian|Smoky, fruity, mild|Wrinkled dried pod or paste|Marinade, stew, grilled meats, sauces|Capsaicin; small culinary quantities|Rehydrate dried pods before blending|Vegan
Mahlab|Stone-fruit kernel spice|Eastern Mediterranean; Middle Eastern|Almond, cherry, floral|Fine ground seed|Sweet breads, cookies, pastries|Small culinary quantities|Grind shortly before use; aroma fades|Tree-nut cross-reactivity possible; use caution for allergies
Nigella seed|Seed spice|South Asian; Middle Eastern|Onion, oregano, pepper|Small black seeds|Flatbread, curry, pickle, salad|Fiber; small amounts of minerals|Toast lightly for aroma|Vegan
White sesame seed|Oilseed|East Asian; Middle Eastern|Nutty, mild|Small pale seeds|Baking, sauces, garnish|Unsaturated fat; calcium; copper|Toast gently; burns quickly|Sesame allergen
Poppy seed|Seed spice|Central/Eastern European; South Asian|Nutty, mild|Tiny crunchy seeds|Baking, dressing, noodles, filling|Unsaturated fat; calcium; manganese|Toast or grind for richer flavor|Vegan; legal/import rules vary by location
Mustard powder|Ground seed spice|British; South Asian|Sharp, pungent|Fine powder|Rub, dressing, cheese sauce, pickles|Small culinary quantities|Mix with liquid and rest for stronger heat|Mustard allergen
Furikake|Japanese seasoning blend|Japanese|Savory, salty, seaweed, sesame|Dry sprinkle|Rice, noodles, eggs, vegetables|Sodium; iodine and calcium may be present|Blend varies; use as finishing seasoning|May contain fish, sesame, soy
Berbere|Ethiopian spice blend|Ethiopian; Eritrean|Hot, warm, aromatic|Fine red blend|Lentils, stews, roasted vegetables, meat|Nutrient profile depends on blend; seasoning amounts small|Bloom in oil; recipe varies by maker|Check blend for allergens
Rosemary|Leaf herb|Mediterranean|Piney, resinous, peppery|Woody needle-like leaf|Roast potatoes, bread, lamb, beans|Small culinary quantities|Use sparingly; strip from woody stems|Vegan
Thyme|Leaf herb|Mediterranean; French|Earthy, floral, lemony|Small leaf; sturdy in braise|Stocks, roast, beans, sauces|Small culinary quantities|Can withstand longer cooking|Vegan
Chives|Allium herb|European; East Asian|Mild onion, fresh|Fine hollow leaves|Eggs, potatoes, soup, garnish|Vitamin K; small culinary quantities|Snip with scissors; add near serving|Vegan
Tarragon|Leaf herb|French; Central Asian heritage|Sweet anise, herbal|Soft leaf, potent aroma|Vinaigrette, chicken, fish, eggs|Small culinary quantities|Use sparingly; dried is more concentrated|Vegan
Sage|Leaf herb|Mediterranean; Italian|Earthy, peppery, eucalyptus-like|Velvety leaf; crisps in fat|Pasta, stuffing, beans, browned butter|Small culinary quantities|Fry leaves briefly for crisp garnish|Vegan
Marjoram|Leaf herb|Mediterranean; West Asian|Sweet, floral, oregano-like|Delicate leaf|Beans, roast vegetables, poultry, tomato|Small culinary quantities|Add late for delicate aroma|Vegan
Lemongrass|Aromatic stalk|Southeast Asian|Lemon, floral, ginger-like|Fibrous stalk; infuses liquid|Thai-style curry, soup, marinade, tea|Small amounts of vitamin C|Bruise and remove before serving, or mince tender core|Vegan
Galangal|Rhizome|Thai; Indonesian; Southeast Asian|Piney, citrusy, peppery|Hard, fibrous|Curry paste, soup, braise|Small culinary quantities|Slice thin or pound; often removed after infusion|Vegan
Makrut lime leaf|Citrus leaf|Thai; Southeast Asian|Intensely citrus, floral|Tough leaf; infuses broth|Curry, soup, rice, marinade|Small culinary quantities|Tear to release aroma; remove before serving|Vegan
Green chili|Fresh chili|South Asian; Latin American; global|Fresh, grassy heat|Crisp when raw; softens in cooking|Chutney, curry, salsa, pickle|Vitamin C; capsaicin|Heat varies by variety and seed/pith|Vegan
Gochugaru|Dried chili flakes|Korean|Fruity, smoky, medium heat|Coarse flakes|Kimchi, stew, marinades, dipping sauces|Capsaicin; tiny nutrient amount per use|Color and heat vary; store sealed|Vegan
Sichuan pepper|Dried husk spice|Chinese, especially Sichuan|Citrusy, floral, numbing|Tingly husk, little heat itself|Mapo tofu-style dishes, chili oil, dry rubs|Small culinary quantities|Toast lightly and remove hard seeds|Vegan
White pepper|Pepper berry spice|Chinese; Southeast Asian; European|Earthy, sharp, fermented notes|Fine powder or whole berry|White sauces, soups, stir-fries|Small culinary quantities|Add gradually; aroma is distinct from black pepper|Vegan
Mace|Seed covering spice|Indonesian; South Asian; European|Warm, floral, nutmeg-like|Lacy aril, ground or whole|Baking, rice, cream sauce, spice blends|Small culinary quantities|Use less than nutmeg by volume|Vegan
Vanilla|Orchid pod spice|Mesoamerican; global pastry|Floral, sweet, creamy|Pod, extract or paste|Custard, baking, drinks, desserts|Negligible nutrient contribution per use|Use seeds and pod; extract strength varies|Vegan
Cacao / cocoa powder|Processed seed product|Mesoamerican; global|Bitter, roasted, chocolate|Fine powder|Baking, drinks, rubs, sauces|Fiber; magnesium; iron; flavanols vary|Natural and alkalized cocoa differ in acidity|Vegan; check facility cross-contact
Curry powder|Spice blend|British South Asian-inspired blend|Warm, earthy, blend-dependent|Powder|Quick curries, soups, marinades|Nutrients depend on blend; seasoning amounts small|Bloom in oil; recipes vary by brand|Check blend for mustard, wheat or other allergens
Za'atar|Spice blend|Levantine|Herbal, tart, sesame, savory|Dry mix|Flatbread, dip, roast vegetables, eggs|Small amounts of calcium/fat from sesame|Blend varies; often thyme, sumac and sesame|Sesame allergen common
Ras el hanout|Spice blend|Moroccan; North African|Warm, floral, complex|Dry blend|Tagine, couscous, roast, rub|Small culinary quantities|Blend can include many spices; recipes vary|Check label for allergens
Chinese five-spice|Spice blend|Chinese|Sweet, anise, warm, peppery|Fine powder|Duck, pork, tofu, braise, roast|Small culinary quantities|Use lightly; strong aroma|Vegan
Black lime / loomi|Dried citrus|Persian Gulf; Middle Eastern|Sour, earthy, fermented citrus|Hard dried fruit|Stew, rice, fish, broth|Negligible nutrient contribution per use|Pierce or grind; simmer in liquid|Vegan
Vanilla extract|Flavor extract|Global baking; vanilla growing regions|Floral, sweet|Liquid extract|Baking, custard, drinks|Negligible nutrient contribution per use|Alcohol-free and pure extracts differ|Check label
Salt, kosher / coarse|Mineral seasoning|Global|Salty, clean|Large flakes or grains|Seasoning, brining, crust|Sodium chloride; no meaningful energy|Crystal size varies; season by taste, not spoon conversion|Vegan
Salt, sea salt|Mineral seasoning|Coastal regions worldwide|Salty; trace mineral notes vary|Flakes or crystals|Finishing, cooking, brining|Sodium chloride|Mineral trace amounts are nutritionally minor|Vegan
`),
  'Protein & Dairy': parse(`
Chicken breast|Poultry|Global; widely used in South Asian, Mediterranean, East Asian cooking|Mild, savory|Lean; dries if overcooked|Grill, roast, curry, sandwich, salad|Protein; niacin; vitamin B6|Cook to safe internal temperature; rest before slicing|Animal product
Chicken thigh|Poultry|Global|Richer, savory|Juicy; forgiving in braise|Curry, roast, grill, stew, skewers|Protein; iron; zinc|Bone-in suits braising; boneless cooks quickly|Animal product
Duck|Poultry|Chinese; French; Southeast Asian|Rich, gamey|Fatty skin; tender meat|Roast, confit, noodle soup, curry|Protein; iron; B vitamins|Render skin slowly for crispness|Animal product
Egg, chicken|Egg|Global|Rich, savory|Sets, emulsifies, foams|Breakfast, baking, binders, sauces|Protein; choline; vitamin B12|Use fresh eggs; cook as needed for food safety|Egg allergen
Lamb|Red meat|Middle Eastern; South Asian; Mediterranean|Robust, grassy|Tender roast or slow-cook shoulder|Kebab, curry, roast, stew|Protein; iron; zinc; B12|Match cut to cooking method|Animal product
Pork|Red meat|East Asian; European; Latin American|Savory, mildly sweet|Varies by cut; fat adds moisture|Roast, stir-fry, braise, sausage|Protein; thiamin; B vitamins|Cook safely; choose cut for method|Animal product
Beef, ground|Red meat|Global|Savory, umami|Crumble or bind into patties|Ragù, kebab, burger, chili, dumplings|Protein; iron; zinc; B12|Brown in batches for color|Animal product
Tofu, firm|Soybean curd|Chinese; Japanese; Southeast Asian|Mild, absorbs seasoning|Firm, browns when pressed|Stir-fry, grill, curry, scramble|Protein; iron; calcium if calcium-set|Press for crisp browning; check coagulant|Soy allergen; vegan
Tofu, silken|Soybean curd|Chinese; Japanese|Mild, creamy|Delicate, custard-like|Soup, dressing, dessert, smoothie|Protein; iron; calcium varies|Handle gently; add near end to soup|Soy allergen; vegan
Tempeh|Fermented soy cake|Indonesian|Nutty, savory, fermented|Firm, chewy|Grill, stir-fry, sandwich, curry|Protein; fiber; iron|Steam or marinate before searing if desired|Soy allergen; vegan
Edamame|Immature soybean|Japanese; East Asian|Sweet, grassy, nutty|Firm, tender bean|Snack, salad, rice bowl, stir-fry|Protein; fiber; folate|Cook from frozen according to package|Soy allergen; vegan
Salmon|Oily fish|North Atlantic; Pacific|Rich, buttery|Flaky, moist|Roast, grill, curry, rice bowl|Protein; omega-3 fats; vitamin D; B12|Cook gently; check bones|Fish allergen
Cod|White fish|North Atlantic; European|Mild, clean|Lean, flaky|Bake, stew, batter-fry, curry|Protein; iodine; B12|Avoid overcooking; flakes when done|Fish allergen
Sardine|Oily fish|Mediterranean; Atlantic|Salty, savory, bold|Soft, oily|Toast, pasta, salad, grill|Protein; omega-3; calcium when bones eaten|Canned product is convenient; watch sodium|Fish allergen
Shrimp / prawn|Shellfish|South Asian; Southeast Asian; global coastal|Sweet, briny|Firm, quick-cooking|Curry, stir-fry, grill, pasta|Protein; iodine; selenium|Cook briefly until opaque|Crustacean shellfish allergen
Mussels|Mollusk|European coastal; Southeast Asian|Briny, sweet, mineral|Tender; broth is aromatic|Steamed, pasta, curry, chowder|Protein; iron; B12; selenium|Discard cracked or unopened cooked shells|Mollusk allergen; shellfish
Greek yogurt|Cultured dairy|Eastern Mediterranean|Tangy, creamy|Thick; tenderizes marinades|Dip, breakfast, sauce, marinade|Protein; calcium; B12; cultures|Use plain unsweetened for savory dishes|Milk allergen
Milk|Dairy|Global|Mild, creamy, subtly sweet|Liquid; browns in baking|Sauce, baking, drinks, porridge|Protein; calcium; B12; riboflavin|Fat level affects body and richness|Milk allergen
Cream|Dairy fat|French; European|Rich, lightly sweet|Thickens; whips depending on fat|Sauce, soup, dessert, curry finish|Fat; some vitamin A|Temper into hot acidic dishes to reduce curdling|Milk allergen
Butter|Dairy fat|European; global|Creamy, nutty when browned|Solid fat; browns and foams|Baking, sauce, saute, finishing|Fat; vitamin A|Clarified butter tolerates higher heat|Milk allergen; ghee may retain trace milk proteins
Cheddar|Aged cheese|British; global|Sharp, savory, nutty|Firm; melts well|Sandwich, gratin, sauce, topping|Protein; calcium; fat|Aging increases sharpness|Milk allergen
Feta|Brined cheese|Greek; Eastern Mediterranean|Salty, tangy|Crumbly, creamy|Salad, pastry, grain bowl, dip|Protein; calcium; sodium|Rinse if too salty; brine keeps it moist|Milk allergen
Paneer|Fresh cheese|South Asian|Mild, milky|Firm; holds shape in curry|Saag paneer, skewers, curry, wraps|Protein; calcium; fat|Sear before adding to sauce for texture|Milk allergen; vegetarian
Turkey|Poultry|North American; European; global|Mild, savory|Lean breast; richer dark meat|Roast, mince, stew, sandwich, meatballs|Protein; niacin; vitamin B6; selenium|Cook safely; breast dries if overcooked|Animal product
Bison|Red meat|North American|Beef-like, slightly sweet|Lean, firm|Burger, stew, roast, chili|Protein; iron; zinc; B12|Cook lean cuts gently; avoid overcooking|Animal product
Venison|Game meat|European; North American|Lean, earthy, gamey|Firm, low fat|Stew, roast, sausage, grill|Protein; iron; B vitamins|Marinade or add fat for lean cuts; cook safely|Animal product
Bacon|Cured pork|European; North American|Smoky, salty, savory|Crisps; rendered fat seasons dishes|Breakfast, soup base, salad, wraps|Protein; fat; sodium varies|Render slowly; use rendered fat in moderation|Animal product; often high sodium
Sausage|Processed meat|European; South Asian; global|Savory, spice-dependent|Juicy, firm; casing may snap|Grill, pasta, stew, breakfast|Protein; fat; sodium varies|Cook through; seasoning and salt vary by style|Animal product; check allergens
Tuna|Fish|Mediterranean; Japanese; global|Savory, meaty|Flaky canned; firm fresh|Salad, sushi-style dishes, pasta, grill|Protein; omega-3; selenium|Species and canning style affect mercury and sodium guidance|Fish allergen
Mackerel|Oily fish|Japanese; Mediterranean; Atlantic|Rich, savory, assertive|Oily, firm|Grill, smoke, curry, rice plate|Protein; omega-3 fats; vitamin D; B12|Cook promptly; strong flavor pairs with acid|Fish allergen
Trout|Freshwater fish|European; North American|Mild, nutty, delicate|Fine flakes, tender|Pan-fry, roast, smoke, stuff|Protein; omega-3; vitamin B12|Check pin bones; avoid overcooking|Fish allergen
Tilapia|White fish|African heritage; global aquaculture|Mild, clean|Lean, flaky|Bake, fry, curry, tacos|Protein; selenium; B12|Season well; cook gently|Fish allergen
Halibut|White fish|North Pacific; North Atlantic|Mild, sweet|Firm, meaty flakes|Roast, grill, stew, fish cakes|Protein; selenium; B vitamins|Lean flesh dries if overcooked|Fish allergen
Anchovy|Small oily fish|Mediterranean; European|Salty, intense umami|Soft cured fillet|Pasta, dressing, pizza, sauce base|Protein; omega-3; calcium varies by form|Melt into warm oil for a savory base|Fish allergen
Crab|Crustacean shellfish|Global coastal|Sweet, briny|Delicate, flaky|Cakes, salad, noodles, soup|Protein; vitamin B12; selenium|Pick shells carefully; cooked crab is perishable|Crustacean shellfish allergen
Scallop|Mollusk|North Atlantic; East Asian|Sweet, delicate, briny|Tender; sears quickly|Pan-sear, pasta, curry, ceviche-style preparations|Protein; B12; selenium|Pat dry; cook briefly for browning|Mollusk allergen
Clam|Mollusk|European; East Asian; North American coastal|Briny, mineral, sweet|Chewy to tender; broth-rich|Pasta, chowder, curry, steam|Protein; iron; B12|Purge grit where needed; discard unopened cooked shells|Mollusk allergen
Oyster|Mollusk|Global coastal|Briny, mineral, sweet|Soft, slippery|Raw, grill, stew, sauce|Protein; zinc; B12|Use reputable fresh product and safe handling|Mollusk allergen
Squid / calamari|Mollusk|Mediterranean; East Asian|Mild, sweet, briny|Tender quickly cooked or slow braised|Grill, fry, stew, noodles|Protein; selenium; B12|Cook very quickly or braise until tender|Mollusk allergen
Octopus|Mollusk|Mediterranean; East Asian|Savory, briny|Needs tenderizing; pleasantly chewy|Grill, salad, stew, rice dishes|Protein; iron; B12|Cook until tender before high-heat charring|Mollusk allergen
Mozzarella|Fresh cheese|Italian|Milky, mild, lightly tangy|Soft, stretchy when heated|Pizza, salad, gratin, sandwiches|Protein; calcium; fat|Fresh and low-moisture styles behave differently|Milk allergen
Parmesan / Parmigiano-style cheese|Hard aged cheese|Italian|Salty, nutty, umami|Hard, crystalline; grates finely|Pasta, risotto, soup, finishing|Protein; calcium; sodium|Rind flavors broth; protected names have specific standards|Milk allergen
Ricotta|Fresh whey cheese|Italian; Mediterranean|Mild, sweet, milky|Soft, grainy-creamy|Pasta filling, dessert, toast, dip|Protein; calcium; fat|Drain for firmer baking fillings|Milk allergen
Blue cheese|Aged cheese|European|Sharp, salty, tangy|Creamy to crumbly|Salad, sauce, dressing, cheese board|Protein; calcium; sodium|Strong flavor; use sparingly in balanced dishes|Milk allergen
Cottage cheese|Fresh cheese|European; North American|Mild, tangy|Curds in creamy whey|Breakfast, filling, dip, baking|Protein; calcium; sodium varies|Curd size and salt vary by product|Milk allergen
Buttermilk|Cultured dairy|European; South Asian|Tangy, lightly sour|Pourable, thick liquid|Pancakes, marinade, dressing, baking|Protein; calcium; B12|Acidic dairy tenderizes; substitutes alter flavor|Milk allergen
Kefir|Cultured dairy drink|Caucasus; Eastern Europe|Tart, lightly yeasty|Thin, pourable|Drink, marinade, smoothie, dressing|Protein; calcium; cultures vary|Live cultures depend on product and handling|Milk allergen
Sour cream|Cultured cream|Eastern European; North American|Tangy, rich|Thick, spoonable|Dip, topping, sauce, baking|Fat; calcium; sodium varies|Stir in off heat to reduce splitting|Milk allergen
Seitan|Wheat gluten protein|East Asian Buddhist cooking; global vegetarian cuisine|Savory, chewy|Dense, meat-like chew|Stir-fry, roast, sandwich, skewers|Protein; iron; sodium varies by product|Season or marinate; texture varies by brand|Contains gluten
Goat meat|Red meat|South Asian; Middle Eastern; Caribbean|Robust, slightly sweet|Lean, firm|Curry, stew, grill, braise|Protein; iron; zinc; B12|Slow-cook tougher cuts; trim excess connective tissue|Animal product
Goose|Poultry|European; Chinese|Rich, gamey|Dark meat; fatty skin|Roast, confit, braise|Protein; iron; B vitamins|Render fat slowly; cook safely|Animal product
Quail|Game bird|East Asian; Mediterranean; European|Mild, delicate|Small, tender|Roast, grill, stew, appetizer|Protein; iron; B12|Cook briefly; small size can dry quickly|Animal product
Rabbit|Game meat|European; Mediterranean|Mild, lean, chicken-like|Lean, fine-grained|Stew, braise, roast, pie|Protein; B12; iron|Braise or add fat to prevent drying|Animal product
Veal|Red meat|European; Italian; French|Mild, delicate|Tender, fine-grained|Scallopini, stew, roast, stock|Protein; B12; zinc|Cook gently; choose cut for method|Animal product
Prosciutto|Cured pork|Italian|Salty, sweet, savory|Thin, delicate slices|Antipasti, pasta, salad, wrapped vegetables|Protein; sodium; fat|Add at finish or serve uncooked as cured|Animal product
Pancetta|Cured pork belly|Italian|Salty, peppery, savory|Fatty, renders crisp|Pasta, soup base, braise, stuffing|Fat; protein; sodium|Render before adding aromatics|Animal product
Ham|Cured pork|European; North American|Salty, smoky or sweet|Tender, sliced or diced|Sandwich, soup, roast, breakfast|Protein; sodium; B vitamins|Curing and smoking styles differ|Animal product
Chorizo|Seasoned sausage|Spanish; Mexican|Garlic, paprika, smoky or chili-hot|Soft fresh or firm cured|Rice dishes, stew, tacos, eggs|Protein; fat; sodium|Spanish cured and Mexican fresh styles are distinct|Animal product; check casing/additives
Salami|Cured sausage|Italian; Central European|Salty, fermented, spiced|Firm, sliceable|Sandwich, antipasti, pizza|Protein; fat; sodium|Ready-to-eat cured types differ from cooked sausages|Animal product
Kielbasa|Smoked sausage|Polish; Central European|Smoky, garlic, savory|Firm, juicy|Soup, skillet, grill, stew|Protein; fat; sodium|Many products are precooked; follow package directions|Animal product
Lobster|Crustacean shellfish|North Atlantic; global coastal|Sweet, briny|Tender, firm tail and claw meat|Boil, grill, roll, pasta, bisque|Protein; selenium; B12|Handle live product safely; cook thoroughly|Crustacean shellfish allergen
Sea bass|White fish|Mediterranean; East Asian|Mild, sweet|Moist, fine flakes|Roast whole, steam, grill, curry|Protein; selenium; B12|Species sold under this name vary by region|Fish allergen
Snapper|White fish|Caribbean; Latin American; Indo-Pacific|Mild, sweet, lightly nutty|Firm, flaky|Whole roast, grill, stew, tacos|Protein; selenium; B12|Check species and local sustainability guidance|Fish allergen
Haddock|White fish|North Atlantic; British|Mild, lightly sweet|Lean flakes|Bake, chowder, smoke, fish cakes|Protein; iodine; B12|Avoid overcooking; smoked and fresh forms differ|Fish allergen
Pollock|White fish|North Pacific; North Atlantic|Mild, clean|Lean, flaky|Batter-fry, stew, fish sticks, curry|Protein; selenium; B12|Frozen fillets are common; thaw safely|Fish allergen
Mahi-mahi|White fish|Tropical Atlantic; Pacific|Mild, sweet|Firm, meaty|Grill, roast, tacos, curry|Protein; selenium; B vitamins|Cook just until opaque|Fish allergen
Sea urchin roe / uni|Echinoderm|Japanese; Mediterranean coastal|Sweet, briny, buttery|Soft, custard-like|Sushi, pasta, toast, sauce|Protein; omega-3; B12|Highly perishable; source and handle carefully|Seafood allergy risk
Fish roe|Fish eggs|Japanese; Scandinavian; global coastal|Salty, briny, savory|Bursting beads or firm eggs|Sushi, garnish, toast, sauce|Protein; omega-3; B12; sodium varies|Roe types differ; keep refrigerated|Fish allergen
Halloumi|Brined cheese|Cypriot; Eastern Mediterranean|Salty, tangy, milky|Firm; grills without melting|Grill, salad, wrap, mezze|Protein; calcium; sodium|Pat dry and sear or grill until golden|Milk allergen
Goat cheese|Fresh or aged cheese|Mediterranean; French|Tangy, earthy|Creamy fresh; firmer aged|Salad, tart, dip, pasta|Protein; calcium; fat|Fresh chèvre and aged styles behave differently|Milk allergen
Brie|Soft-ripened cheese|French|Buttery, mushroomy, mild|Creamy center, edible rind|Cheese board, pastry, sandwich|Protein; calcium; fat|Bring toward room temperature before serving|Milk allergen
Camembert|Soft-ripened cheese|French|Earthy, mushroomy, rich|Soft, creamy center|Bake, board, sandwich, tart|Protein; calcium; fat|Ripeness changes texture and aroma|Milk allergen
Gouda|Aged cheese|Dutch; European|Buttery, nutty, caramel notes when aged|Semi-hard, melts well|Sandwich, gratin, snack, sauce|Protein; calcium; sodium|Young and aged styles have different moisture and flavor|Milk allergen
Gruyère|Hard cheese|Swiss; French|Nutty, savory, sweet|Firm; melts smoothly|Fondue, gratin, soup, quiche|Protein; calcium; sodium|Grate fresh for smooth melting|Milk allergen
Oat milk|Plant-based beverage|Contemporary global pantry|Mild, cereal-sweet|Pourable; barista versions foam|Coffee, cereal, baking, sauce|Carbohydrate; calcium/B12 if fortified|Fortification and sugar vary by brand|Usually vegan; check gluten cross-contact
Soy milk|Plant-based beverage|East Asian; global|Mild, bean-like|Pourable; curdles in acid|Drink, baking, tofu making, sauce|Protein; calcium/B12 if fortified|Choose unsweetened for savory cooking|Soy allergen
Almond milk|Plant-based beverage|Middle Eastern; Mediterranean; global|Mild, nutty|Thin; barista versions foam|Drinks, cereal, baking|Often low protein; calcium if fortified|Unsweetened products vary in body|Tree-nut allergen
Coconut yogurt|Cultured plant-based product|South Asian; Southeast Asian; contemporary global|Tangy, coconut-rich|Creamy|Dips, breakfast, curry finish, desserts|Fat and calcium vary by product|Choose unsweetened for savory dishes|Vegan; check additives and fortification
`),
  'Pantry & Condiments': parse(`
Almond|Tree nut|Middle Eastern; Mediterranean; South Asian|Sweet, buttery|Crunchy; paste is smooth|Snack, baking, sauce, garnish|Unsaturated fat; protein; vitamin E; magnesium|Toast for aroma; almond flour absorbs moisture|Tree-nut allergen
Cashew|Tree nut|South Asian; West African; Southeast Asian|Sweet, creamy|Soft crunch; blends into cream|Curry, sauce, snack, dessert|Unsaturated fat; copper; magnesium|Soak for smooth dairy-free cream|Tree-nut allergen
Sesame seed|Oilseed|West Asian; East Asian; South Asian|Nutty, toasted|Small crunchy seeds; paste is smooth|Tahini, garnish, bread, dressing|Unsaturated fat; calcium; copper|Toast lightly; burns fast|Sesame allergen
Walnut|Tree nut|Persian; Mediterranean; European|Bitter-sweet, earthy|Crunchy, oily|Salad, baking, pesto, garnish|Omega-3 ALA; fiber; copper|Toast gently; rancidity develops with age|Tree-nut allergen
Chia seed|Seed|Central American|Mild, nutty|Forms gel when hydrated|Pudding, thickener, baking|Fiber; omega-3 ALA; calcium|Hydrate before use in large quantities|Vegan; allergy possible
Flaxseed|Seed|West Asian heritage; European|Nutty, earthy|Whole seeds pass through; ground absorbs liquid|Baking, porridge, egg substitute|Fiber; omega-3 ALA; lignans|Grind for better nutrient availability; refrigerate ground|Vegan
Pumpkin seed|Seed|Mesoamerican; global|Nutty, green|Crunchy; toasts well|Salad, granola, sauce, snack|Protein; magnesium; zinc; iron|Toast for crunch; pepitas are hulled|Vegan
Olive oil, extra virgin|Plant oil|Mediterranean|Fruity, peppery, grassy|Liquid fat; carries aroma|Dressing, finishing, saute, dip|Monounsaturated fat; vitamin E|Flavor varies by cultivar and harvest|Vegan
Olive oil, refined|Plant oil|Mediterranean; global|Mild, neutral|Liquid fat, moderate-high heat use|Roast, saute, general cooking|Monounsaturated fat|Less aromatic than extra virgin|Vegan
Sesame oil, toasted|Seed oil|East Asian|Toasty, roasted|Aromatic finishing oil|Noodles, dressing, stir-fry finish|Unsaturated fats; vitamin E|Use as seasoning; high aroma|Sesame allergen
Coconut milk|Plant-based liquid|South Asian; Southeast Asian; Caribbean|Sweet, rich, coconut|Creamy; separates in can|Curry, soup, dessert, drink|Fat varies greatly by product|Shake can; light and full-fat differ|Vegan; coconut labeling varies
Vinegar, rice|Fermented condiment|Chinese; Japanese; Korean|Mildly sweet, gentle acid|Thin liquid|Sushi rice, dipping sauce, dressing|Negligible micronutrients|Seasoned versions contain sugar and salt|Vegan
Vinegar, apple cider|Fermented condiment|European; North American|Fruity, sharp|Thin liquid|Pickle, slaw, marinade, sauce|Negligible micronutrients|Dilute in sauces to balance sharpness|Vegan
Soy sauce|Fermented condiment|Chinese; Japanese; Korean|Salty, umami|Thin, savory liquid|Marinade, dipping sauce, stir-fry|Sodium; small amounts of protein|Regular, light, dark and tamari differ|Soy; often wheat (tamari may be wheat-free)
Tamari|Fermented soy sauce|Japanese|Salty, deep umami|Full-bodied liquid|Dipping, gluten-aware seasoning, marinade|Sodium; small amounts of protein|Check label for wheat-free status|Soy allergen
Fish sauce|Fermented condiment|Thai; Vietnamese; Southeast Asian|Salty, intensely savory|Thin liquid|Curry, dipping sauce, dressing, stir-fry|Sodium; trace protein|Use drops at a time; mellows in cooked dishes|Fish allergen
Oyster sauce|Savory sauce|Chinese; Cantonese|Sweet-salty, umami|Thick, glossy|Stir-fry, glaze, noodle sauce|Sodium; sugars vary|Dilute for even coating|Oyster/shellfish; often soy or wheat
Miso, white|Fermented soybean paste|Japanese|Mild, sweet-salty umami|Thick paste|Soup, glaze, dressing, butter|Protein; sodium; fermentation cultures|Whisk in off heat to retain aroma|Soy; often grains; check label
Tomato paste|Concentrated tomato|Italian; global|Sweet, tangy, umami|Dense paste|Sauce base, curry, stew, braise|Lycopene; potassium; sodium varies|Cook briefly until darkened for depth|Vegan
Canned tomatoes|Preserved fruit vegetable|Italian; global|Acidic, sweet, umami|Juicy, crushed or whole|Sauce, curry, soup, braise|Vitamin C; lycopene; sodium varies|Whole tomatoes often have better texture control|Vegan
Coconut sugar|Sweetener|South Asian; Southeast Asian|Caramel, toffee|Granular|Baking, sauces, drinks|Mostly carbohydrate; minerals in trace amounts|Not nutritionally equivalent to a meaningful mineral source|Vegan
Honey|Sweetener|Global|Floral, sweet|Viscous; browns readily|Glaze, dressing, baking, drinks|Mostly sugars; trace compounds vary|Use lower heat to prevent scorching|Not vegan; not suitable for infants under 1 year
Maple syrup|Tree sap sweetener|North American|Caramel, woody|Thin syrup|Pancakes, glaze, baking, drinks|Mostly sugars; manganese|Grade/color reflect flavor intensity|Vegan
Molasses|Sugar by-product|Caribbean; North American; European baking|Bitter-sweet, dark caramel|Viscous|Baking, barbecue sauce, marinades|Iron, calcium, potassium vary by type|Blackstrap is strongest and bitterest|Vegan
Sugar, granulated|Sweetener|Global|Clean sweetness|Crystalline; tenderizes and browns|Baking, sauce balance, preserving|Carbohydrate; no meaningful micronutrients|Dissolve fully for syrups|Vegan
Cornstarch|Starch thickener|Chinese; global|Neutral|Glossy thickener when heated|Sauce, custard, coating|Mostly carbohydrate|Mix with cold liquid before hot sauce|Usually gluten-free; check cross-contact
Baking powder|Leavening blend|Global baking|Neutral, slight alkaline note|Releases gas with moisture and heat|Cake, biscuit, quick bread|Sodium varies by product|Check freshness and label; single/double acting|May contain starch; check allergen label
Baking soda|Leavening agent|Global baking|Alkaline, slightly salty|Reacts with acid|Cookies, quick bread, tenderizing|Sodium|Requires acidic ingredient for leavening|Vegan
Yeast, dry|Leavening organism|Global baking|Bready, savory|Ferments dough, creates gas|Bread, pizza, rolls|B vitamins vary; small amounts per serving|Activate according to package type|Usually vegan
Stock / broth|Liquid base|French; global|Savory, aromatic|Adds body and moisture|Soup, rice, sauce, braise|Sodium varies widely; nutrients depend on recipe|Choose low-sodium when seasoning control matters|May contain allergens; check label
Tahini|Sesame paste|Levantine; Middle Eastern|Nutty, lightly bitter|Creamy, can seize with water then loosen|Hummus, sauce, dressing, baking|Unsaturated fat; protein; calcium|Whisk with lemon and water to smooth|Sesame allergen
Harissa|Chili paste|Tunisian; North African|Hot, smoky, garlic, spice|Thick paste|Marinade, soup, dip, grain bowls|Sodium varies; chili compounds|Heat level and ingredients vary by brand|Check for caraway, garlic, allergens
Gochujang|Fermented chili paste|Korean|Sweet-hot, fermented umami|Thick, sticky paste|Stew, glaze, marinade, sauce|Sodium; carbohydrate; fermentation|Thin with liquid before tossing|Soy; often wheat
Sambal oelek|Chili condiment|Indonesian; Southeast Asian|Bright chili heat, tart|Coarse chili paste|Noodles, eggs, marinades, sauces|Sodium varies; chili compounds|Add gradually; brand salt differs|Check label for allergens
Mustard, Dijon|Condiment|French|Sharp, tangy, savory|Smooth emulsion|Vinaigrette, sauce, sandwich, marinade|Sodium; negligible micronutrients|Emulsifies oil and vinegar|Mustard allergen
Mayonnaise|Emulsified condiment|French; global|Tangy, rich|Creamy emulsion|Sandwich, dressing, dip, salad|Fat; sodium varies|Refrigerate; egg-free versions available|Usually egg; some contain soy
Vinegar, balsamic|Aged vinegar|Italian|Sweet-tart, woody|Syrupy to thin depending on age|Dressing, glaze, roast vegetables, finishing|Negligible micronutrients; sugars vary|Use aged styles for finishing, lighter styles for cooking|Vegan
Vinegar, red wine|Fermented condiment|French; Mediterranean|Sharp, fruity|Thin liquid|Vinaigrette, marinade, deglazing|Negligible micronutrients|Balance with oil and a little sweetness if needed|Vegan
Vinegar, sherry|Aged vinegar|Spanish|Nutty, sharp, complex|Thin liquid|Finishing, dressing, soup, sauce|Negligible micronutrients|Add late for aroma|Vegan
Rice wine vinegar, seasoned|Seasoned vinegar|Japanese|Mild, sweet-sour|Thin liquid|Sushi rice, quick pickle, dressing|Sugar and sodium vary|Adjust added sugar and salt to taste|Vegan
Mirin|Sweet rice seasoning|Japanese|Sweet, mellow, fermented|Thin syrup-like liquid|Teriyaki-style glaze, simmered dishes, sauces|Sugars; sodium varies|Hon-mirin and mirin-style products differ|May contain alcohol; check label
Shaoxing cooking wine|Rice wine|Chinese|Nutty, aromatic, savory|Thin liquid|Braise, stir-fry, marinades, sauces|Alcohol; sodium varies by product|Cooking versions can be salted|Check label for wheat or other allergens
Sake|Rice wine|Japanese|Clean, lightly sweet, savory|Thin liquid|Broth, marinade, simmered dishes|Alcohol; small amounts of amino acids|Cooking sake may contain added salt|Alcohol-containing
Worcestershire sauce|Fermented savory condiment|British; global|Sour-sweet, savory, spiced|Thin liquid|Marinade, burger, stew, tomato drinks|Sodium; sugars vary|Anchovy is common; recipes differ|Often fish; check label for allergens
Hot sauce|Chili condiment|Americas; global|Hot, acidic, salty|Thin to thick sauce|Eggs, tacos, soups, marinades|Sodium varies; capsaicin|Add after tasting; heat varies by brand|Check label for allergens
Barbecue sauce|Cooked sauce|North American|Sweet, smoky, tangy|Thick glaze|Grill, roast, sandwich, beans|Added sugars and sodium can be high|Apply late on grill to limit burning|Check label for allergens
Pesto|Herb-nut sauce|Italian; Ligurian|Herbal, garlicky, nutty, salty|Oil-bound paste|Pasta, sandwich, roast vegetables, dip|Unsaturated fat; herbs; sodium varies|Loosen with pasta water; classic versions contain cheese and nuts|Often milk and tree nuts
Olives, green|Cured fruit|Mediterranean|Salty, tangy, grassy|Firm, briny|Salad, tapenade, braise, garnish|Monounsaturated fat; sodium|Rinse if overly salty; remove pits as needed|Vegan
Olives, black|Cured fruit|Mediterranean|Savory, mellow, briny|Soft to firm by cure|Salad, pizza, tapenade, stew|Monounsaturated fat; sodium|Curing style changes taste and texture|Vegan
Capers|Pickled flower bud|Mediterranean|Salty, tart, floral|Small, tender burst|Sauce, pasta, salad, fish, dressing|Sodium; negligible energy|Rinse or drain to reduce saltiness|Vegan
Pickled cucumber / gherkin|Pickled vegetable|European; global|Sour, salty, dill or spice notes|Crunchy|Sandwich, burger, salad, relish|Sodium; low energy density|Brine-fermented and vinegar-pickled styles differ|Vegan
Kimchi|Fermented vegetables|Korean|Sour, spicy, savory|Crisp, juicy|Rice, soup, pancake, noodle dishes|Fiber; vitamins; sodium varies|Flavor and heat change with fermentation time|Check fish sauce/shrimp paste in recipe
Sauerkraut|Fermented cabbage|Central/Eastern European|Sour, salty|Crisp to soft|Sausage, sandwich, stew, salad|Fiber; vitamin C; sodium varies|Unpasteurized versions may contain live cultures|Vegan; check added wine or allergens
Coconut aminos|Fermented coconut seasoning|Contemporary alternative condiment|Sweet-salty, mild umami|Thin sauce|Stir-fry, marinade, dipping sauce|Sodium and sugars vary by brand|Usually sweeter and less salty than soy sauce|Check label; often marketed soy-free
Tomato passata|Strained tomato puree|Italian; global|Fresh tomato, sweet-tart|Smooth liquid puree|Sauce, soup, curry, braise|Lycopene; potassium; sodium varies|Season and reduce to desired thickness|Vegan
Tomato sauce|Prepared tomato condiment|Italian-American; global|Sweet, tangy, savory|Smooth or chunky|Pasta, pizza, casserole, dipping|Sodium and added sugar vary|Seasoning differs from passata and canned tomatoes|Check label for allergens
Breadcrumbs|Bread product|European; global|Mild, toasted|Crisp coating or soft binder|Coating, meatballs, gratin, stuffing|Carbohydrate; sodium varies|Fresh and dried crumbs absorb differently|Usually contains wheat/gluten
Panko|Coarse breadcrumb|Japanese|Mild, clean|Large flakes; very crisp when fried|Katsu-style coating, baked topping|Carbohydrate; sodium varies|Press gently for an airy coating|Usually contains wheat/gluten
Nori|Dried seaweed|Japanese; Korean|Briny, roasted, umami|Thin, crisp sheet|Sushi, rice seasoning, garnish|Iodine; protein; minerals; sodium varies|Toast sheets lightly; iodine varies considerably|Vegan; iodine-sensitive diets should consider amount
Kombu|Kelp seaweed|Japanese; East Asian|Deep umami, mineral|Tough dried frond; infuses broth|Dashi, beans, broth, pickle|Iodine; glutamates; minerals|Steep gently; iodine content varies widely|Vegan; iodine-sensitive diets should consider amount
Wakame|Seaweed|Japanese; Korean|Briny, mild, sweet|Rehydrates tenderly|Soup, salad, rice dishes|Iodine; fiber; minerals|Rehydrate briefly; dried form expands substantially|Vegan; iodine content varies widely
Tapioca starch|Cassava starch|South American; Southeast Asian|Neutral|Glossy, elastic thickener|Pudding, baking, sauces, boba|Mostly carbohydrate|Use slurry; too much can turn gummy|Usually gluten-free; check cross-contact
Chocolate, dark|Cacao product|Mesoamerican heritage; global confectionery|Bitter-sweet, roasted|Snaps; melts smoothly|Dessert, baking, sauce, drinks|Fat; carbohydrate; iron and magnesium vary by cacao|Cacao percentage affects sweetness and bitterness|May contain milk, soy or traces of nuts
Cocoa nibs|Roasted cacao pieces|Mesoamerican heritage; global|Bitter, fruity, roasted|Crunchy pieces|Granola, dessert, garnish, savory rub|Fiber; magnesium; iron|Use as crunchy bitter accent|Vegan; check cross-contact
Jam / fruit preserve|Fruit preserve|European; global|Sweet, fruity, tartness varies|Spreadable gel|Toast, pastry, glaze, sauce|Mostly carbohydrate; fruit content varies|Sugar-free and reduced-sugar versions behave differently|Vegan; check gelatin in some products
Marmalade|Citrus preserve|British; Mediterranean|Sweet-bitter, citrus peel|Gel with peel pieces|Toast, glaze, baking, sauce|Mostly carbohydrate; small amounts of vitamin C|Bitterness varies with peel and recipe|Vegan
Agave syrup|Plant sweetener|Mexican; global|Mild, very sweet|Thin syrup|Drinks, baking, glaze|Mostly sugars|Use sparingly; liquid sweetener changes recipe moisture|Vegan
Palm sugar|Palm sap sweetener|Southeast Asian; South Asian|Caramel, smoky, fruity|Solid cakes or granules|Curry, desserts, sauces, drinks|Mostly carbohydrate; trace minerals|Grate or dissolve; sweetness varies|Vegan
Ghee|Clarified butterfat|South Asian; Middle Eastern|Nutty, buttery|Solid or liquid fat|Tempering, roast, baking, finishing|Fat; small amounts of fat-soluble vitamins|Clarification removes most water and milk solids, but traces may remain|Dairy-derived
Canola / rapeseed oil|Seed oil|European; North American|Neutral, mild|Liquid cooking oil|Baking, frying, dressing, saute|Mostly unsaturated fats; vitamin E|Refined oils are neutral; cold-pressed has more flavor|Vegan
Avocado oil|Fruit oil|Mesoamerican; global|Mild, buttery|Liquid oil|Dressing, roast, saute|Mostly monounsaturated fat; vitamin E|Refined and virgin versions differ in flavor and use|Vegan
Grapeseed oil|Seed oil|European; global|Very mild, neutral|Light liquid oil|Dressing, saute, mayonnaise|Mostly polyunsaturated fat; vitamin E|Neutral flavor suits emulsions|Vegan
Peanut oil|Legume seed oil|South Asian; East Asian; West African|Nutty if unrefined; neutral if refined|Liquid frying oil|Stir-fry, deep-fry, dressing|Mostly unsaturated fat|Refined oil allergen status varies by jurisdiction and processing|Peanut-derived; verify allergy safety
MSG|Umami seasoning|East Asian; global food service|Savory, umami|White crystals|Soups, stir-fries, sauces, snacks|Sodium; glutamate|Use a pinch; reduce other salt accordingly|Check blends for added allergens
Doubanjiang|Fermented chili-bean paste|Sichuan, China|Salty, hot, fermented umami|Thick red-brown paste|Mapo tofu-style dishes, noodle sauce, braise|Sodium; protein varies|Fry briefly in oil to deepen aroma|Often soy and wheat; check label
Doenjang|Fermented soybean paste|Korean|Earthy, salty, deep umami|Dense paste|Stew, soup, dip, marinade|Protein; sodium; fermentation|Whisk into hot broth; strength varies by brand|Soy; may contain grains
Hoisin sauce|Sweet savory sauce|Chinese; Cantonese|Sweet, salty, spiced|Thick, glossy|Glaze, dipping sauce, stir-fry, wrap|Added sugar; sodium|Dilute for an even glaze|Often soy and wheat
Teriyaki sauce|Seasoned soy sauce|Japanese-inspired; global|Sweet-salty, glossy umami|Pourable to thick|Glaze, marinade, stir-fry|Added sugar; sodium|Sauce may burn under high heat|Usually soy and wheat; check label
Chili crisp|Chili oil condiment|Chinese; global|Hot, crunchy, savory|Oil with crisp aromatics|Noodles, eggs, rice, dumplings|Fat; sodium varies|Stir jar before spooning; oil carries heat|May contain soy, sesame, peanuts
Preserved lemon|Salt-cured citrus|Moroccan; North African|Salty, tart, floral citrus|Soft rind, tender pulp|Tagine, salad, dressing, fish|Sodium; small amounts of vitamin C|Rinse and use mostly the rind if very salty|Vegan
Pickled jalapeño|Pickled chili|Mexican; Tex-Mex|Hot, sour, salty|Crisp to soft|Nachos, tacos, sandwich, salsa|Sodium; capsaicin|Drain or rinse to moderate brine|Vegan; check preservatives
Fermented black beans|Salted fermented soybean|Chinese|Very salty, intense umami|Whole soft beans|Steamed fish, stir-fry, sauce, tofu|Protein; sodium|Rinse briefly and mince; use sparingly|Soy allergen
Black bean garlic sauce|Savory fermented sauce|Chinese|Salty, garlicky, umami|Thick sauce|Stir-fry, steamed dishes, noodles|Sodium; sugars vary|Add gradually; no extra salt may be needed|Soy; often wheat
Dashi stock|Japanese broth base|Japanese|Clean umami, oceanic|Clear broth|Soup, noodles, simmered vegetables|Sodium varies; glutamates|Kombu-only, bonito and instant types differ|May contain fish; check product
Pomegranate molasses|Fruit reduction|Levantine; Middle Eastern|Tart-sweet, fruity, deep|Thick syrup|Salad, marinade, stew, glaze|Mostly sugars; acids and polyphenols vary|Use in small amounts; brands vary in sweetness|Vegan
Date syrup|Fruit sweetener|Middle Eastern; North African|Caramel, fruity|Thick syrup|Baking, glaze, breakfast, marinade|Mostly sugars; some potassium|Use as liquid sweetener; adjust recipe moisture|Vegan
Brown sugar|Sugar with molasses|Global baking|Caramel, warm|Moist crystals|Baking, barbecue sauce, glaze|Mostly carbohydrate|Dark and light types differ in molasses content|Vegan; processing may vary
Demerara sugar|Partially refined cane sugar|Caribbean; British baking|Light caramel|Large crunchy crystals|Topping, baking, coffee|Mostly carbohydrate|Does not dissolve as readily as fine sugar|Vegan
Powdered sugar|Milled sugar|European; global baking|Clean sweetness|Fine, dissolves quickly|Icing, frosting, dusting, glaze|Mostly carbohydrate; starch may be added|Sift to remove lumps|Often includes cornstarch
Nutritional yeast|Inactive yeast seasoning|Contemporary vegetarian pantry|Savory, cheesy, nutty|Dry flakes or powder|Sauce, popcorn, pasta, vegetables|Protein; B vitamins; fortification varies|Add to sauces or sprinkle at finish|Usually vegan; check fortified product label
Bouillon cube / paste|Concentrated stock seasoning|Global|Salty, savory, aromatic|Concentrated block or paste|Soup, rice, braise, quick broth|Sodium often high; nutrients vary|Dilute before salting; brands vary widely|May contain wheat, soy, dairy or celery
Sunflower oil|Seed oil|Eastern Europe; global|Mild, neutral|Liquid cooking oil|Frying, baking, saute, dressing|Mostly unsaturated fats; vitamin E|Refined and cold-pressed styles differ|Vegan
Safflower oil|Seed oil|South Asian; North American|Very mild|Light liquid oil|Frying, baking, dressing|Mostly unsaturated fats; vitamin E|High-oleic and conventional types differ|Vegan
Walnut oil|Tree-nut oil|European; North American|Nutty, delicate|Liquid finishing oil|Dressing, finishing, baking|Polyunsaturated fats; omega-3 ALA|Use mostly cold; refrigerate after opening|Tree-nut allergen risk; verify processing
Flaxseed oil|Seed oil|European; global|Nutty, grassy|Delicate liquid oil|Cold dressing, porridge, smoothie|Omega-3 ALA|Do not heat; store refrigerated and shielded from light|Vegan
Mushroom powder|Dried mushroom seasoning|East Asian; global|Deep earthy umami|Fine powder|Broth, rub, sauce, vegetarian seasoning|B vitamins; minerals vary by mushroom|Use sparingly; check salt in blends|Vegan; mushroom sensitivity possible
Nori seasoning flakes|Seaweed seasoning|Japanese; Korean|Salty, roasted, umami|Dry flakes|Rice, noodles, eggs, vegetables|Iodine; sodium varies|Use as finishing seasoning; iodine varies|May contain sesame or fish
Vanilla bean paste|Vanilla flavoring|Global pastry; vanilla growing regions|Floral, creamy, sweet|Thick paste with seeds|Custard, cake, ice cream, drinks|Negligible nutrient contribution per use|Use like extract; strength varies by brand|Check label for sugar or alcohol
Gelatin|Animal-derived gelling agent|European; global pastry|Neutral|Sets liquids to a soft gel|Jelly, mousse, confectionery, aspic|Protein; no complete protein role in typical use|Bloom in cold liquid before dissolving|Animal-derived; not vegetarian
Agar agar|Seaweed-derived gelling agent|East Asian; global plant-based cooking|Neutral|Firm plant-based gel|Jelly, dessert, pastry, vegan set dishes|Fiber; small mineral amounts|Boil to activate; sets more firmly than gelatin|Vegan
`)
};

// Second global expansion: regionally distinct cultivars, species, prepared forms and pantry staples.
data['Vegetables'].push(...parse(`
Banana blossom|Flower bud|South Asian; Southeast Asian|Mild, artichoke-like, lightly bitter|Layered, tender when cooked|Curry, salad, fritter, vegan filling|Fiber; potassium; micronutrients vary|Remove tough outer bracts; soak cut pieces to limit browning|Vegan
Ivy gourd / tindora|Gourd|South Asian|Mild, slightly tart|Crisp, small oval fruit|Stir-fry, curry, pickle|Fiber; vitamin C|Choose firm unblemished fruits|Vegan
Ridge gourd / luffa|Gourd|South Asian; East Asian|Mild, sweet, grassy|Softens quickly; sponge-like mature fibers|Curry, soup, stew|Water; fiber; vitamin C|Peel ridges if tough; use young fruit for cooking|Vegan
Snake gourd|Gourd|South Asian|Mild, cucumber-like|Tender, watery|Curry, stir-fry, stew|Water; fiber; vitamin C|Scoop seeds if mature|Vegan
Yardlong bean|Pod vegetable|Southeast Asian; South Asian|Green, sweet|Long, crisp-tender|Stir-fry, curry, pickle|Fiber; folate; vitamin C|Cut into short lengths; cook briefly|Vegan
Water chestnut|Corm|Chinese; Southeast Asian|Sweet, mild|Stays crisp after cooking|Stir-fry, dumpling filling, salad|Carbohydrate; fiber; potassium|Fresh and canned products differ in crunch|Vegan
Jalapeño, red-ripe|Fresh chili variety|Mexican|Fruitier and sweeter than green jalapeño, medium heat|Firm, thick flesh|Salsa, pickle, poppers, sauces|Vitamin C; capsaicin|Heat varies by pepper|Vegan
Shishito pepper|Fresh pepper|Japanese|Mostly mild, grassy; occasional hot fruit|Thin skin, blisters easily|Blistered appetizer, stir-fry, grill|Vitamin C; fiber|Blister over high heat; occasional peppers are hotter|Vegan
Padron pepper|Fresh pepper|Galician, Spain|Green, grassy; occasional hot fruit|Small, thin-walled|Blistered tapas, grill, pickle|Vitamin C; fiber|Cook whole and season after blistering|Vegan
Sweet corn, white|Grain vegetable|Americas; East Asian|Sweet, milky|Juicy kernels|Grill, soup, salad, masa dishes|Carbohydrate; fiber; folate|Cook soon after shucking for sweetness|Vegan
Sweet corn, blue|Grain vegetable|Mesoamerican; Southwestern US|Earthy, sweet, toasted|Firm kernels; blue-purple pigment|Tortillas, atole, roast, salad|Carbohydrate; fiber; anthocyanins|Varieties differ in starch and color|Vegan
Vidalia onion|Sweet onion cultivar|Georgia, United States|Very sweet, mild|Juicy layers|Raw salads, grill, onion rings, sandwich|Vitamin C; fiber|Seasonal, protected geographic designation|Vegan
Walla Walla sweet onion|Sweet onion cultivar|Washington State, United States|Mild, very sweet|Juicy, tender|Raw sandwich, grill, roast|Vitamin C; fiber|Less pungent than storage onions|Vegan
Red kuri squash|Winter squash|Japan; East Asian|Chestnut-like, sweet|Dense, smooth flesh|Soup, roast, curry, puree|Beta-carotene; fiber; potassium|Skin is edible when cooked tender|Vegan
Chayote shoots|Edible leafy shoots|Mexico; Central America; South Asia|Mild, grassy|Tender stems and leaves|Saute, stew, greens, omelet|Folate; vitamin C; fiber|Remove tough fibers from mature shoots|Vegan
Papaya greens|Leafy plant part|South Asia; Southeast Asia|Bitter, peppery|Sturdy leaves|Stew, saute, soup|Folate; vitamins A and C|Cook thoroughly; flavor is distinctly bitter|Vegan
Purslane|Leafy succulent|Mediterranean; West Asian; Latin American|Lemony, lightly salty, green|Juicy, crisp succulent leaves|Salad, stew, stir-fry, pickle|Omega-3 ALA; vitamin C; magnesium|Wash carefully; tender stems are edible|Vegan
Malabar spinach|Leafy vine|South Asian; Southeast Asian|Mild, slightly peppery|Thick, mucilaginous leaves|Dal, stew, stir-fry, soup|Vitamins A and C; iron; folate|Use thick stems in longer cooking|Vegan
Amaranth greens|Leafy plant|West African; South Asian; Caribbean|Earthy, mild, slightly mineral|Tender to sturdy by maturity|Saute, stew, soup, side greens|Folate; vitamins A and C; iron|Cook mature leaves; wash grit well|Vegan
Spider plant leaves|Leafy plant|West African; East African|Mild, spinach-like|Tender cooked leaves|Saute, stew, traditional greens dishes|Vitamins A and C; iron; calcium varies|Cook leaves; local varieties differ|Vegan
Drumstick pods / moringa pods|Legume pod|South Asian; East African|Green, mild, slightly sweet|Fibrous exterior; soft pulp around seeds|Sambar, curry, stew|Vitamin C; fiber; potassium|Simmer until pulp is tender; eat pulp, discard fibers|Vegan
Moringa leaves|Leafy plant|South Asian; East African|Peppery, green|Small leaves soften rapidly|Soup, dal, saute, omelet|Vitamins A and C; calcium; iron|Add late to soups; dry leaves are concentrated|Vegan
Chicory root|Root|European; Mediterranean|Bitter, earthy|Firm root; softens when braised|Roast, braise, coffee substitute|Fiber (inulin); potassium|Roast or braise to mellow bitterness|Vegan
Salsify|Root|European|Oyster-like, mild, earthy|Tender, slightly fibrous|Soup, mash, saute, gratin|Fiber; potassium; iron|Peel and acidulate to limit browning|Vegan
Black radish|Root|Central/Eastern European; Mediterranean|Sharp, peppery|Dense, crisp|Salad, pickle, roast, soup|Vitamin C; fiber|Peel if skin is tough; slice thin for raw use|Vegan
`));

data['Fruit & Citrus'].push(...parse(`
Alphonso mango|Mango cultivar|Western India|Intensely floral, sweet, low-fiber|Silky, juicy flesh|Dessert, puree, lassi, chutney|Vitamin C; folate; beta-carotene|Seasonal; ripeness drives aroma|Vegan
Kesar mango|Mango cultivar|Western India|Sweet, saffron-like floral aroma|Smooth golden flesh|Dessert, puree, drinks, preserve|Vitamin C; folate; beta-carotene|Use ripe fruit for fragrance|Vegan
Ataulfo mango|Mango cultivar|Mexico; Central America|Very sweet, buttery, low fiber|Smooth, soft flesh|Snack, salsa, smoothie, dessert|Vitamin C; folate; beta-carotene|Ripe fruit yields gently to pressure|Vegan
Green mango|Unripe fruit|South Asian; Southeast Asian|Very sour, crisp|Firm, crunchy flesh|Pickle, chutney, salad, curry|Vitamin C; fiber|Use as an acid fruit; ripeness changes acidity|Vegan
Meyer lemon|Citrus cultivar|China; California cultivation|Floral, sweeter, less sharp than lemon|Thin peel, juicy flesh|Dessert, dressing, preserve, roast|Vitamin C; fiber|Use peel and juice together for fragrance|Vegan
Yuzu kosho citrus zest|Citrus-chili condiment base|Japanese|Citrus-forward, salty, hot|Coarse paste|Fish, noodles, soup, grilled foods|Sodium; vitamin C varies|Commercial pastes vary in heat and salt|May contain citrus/chili allergens; check label
Key lime|Citrus cultivar|Southeast Asia heritage; Caribbean cultivation|Sharp, aromatic, floral|Small, seedy, juicy|Pie, ceviche-style acid, drinks, dressing|Vitamin C|Roll before juicing; seeds are common|Vegan
Calamansi|Citrus|Philippines; Southeast Asia|Tart, floral, mandarin-lime|Small, thin-skinned fruit|Dipping sauce, drink, marinade, dessert|Vitamin C|Use fresh juice; frozen concentrate differs|Vegan
Makrut lime fruit|Citrus|Southeast Asia|Very aromatic, bitter-tart|Bumpy thick peel, little juice|Zest, preserve, curry paste, syrup|Vitamin C; small amounts of fiber|Use zest sparingly; juice is intensely aromatic|Vegan
Finger lime|Citrus|Australia|Tart, lime-like|Juice vesicles burst like caviar|Seafood, salad, dessert, garnish|Vitamin C|Cut and squeeze to release vesicles|Vegan
Seville orange|Bitter orange|Mediterranean; Spain|Bitter, tart, aromatic|Juicy, thick peel|Marmalade, sauce, baking, marinade|Vitamin C; flavonoids|Usually cooked or sweetened due to bitterness|Vegan
Ugli fruit|Citrus hybrid|Jamaica|Sweet-tart, mild grapefruit notes|Loose peel, juicy segments|Snack, salad, juice, dessert|Vitamin C; fiber|Remove thick pith for best texture|Vegan
Kaffir plum / marula|Stone fruit|Southern Africa|Sweet-tart, resinous, aromatic|Juicy flesh around stone|Drink, preserve, sauce|Vitamin C; fruit composition varies|Ripeness and processing differ by region|Vegan
Baobab fruit pulp|Tree fruit|Sahel; East and Southern Africa|Tart, citrusy, tangy|Dry powdery pulp|Drink, porridge, sauce, baking|Vitamin C; fiber; minerals vary|Use dried pulp/powder; products differ in concentration|Vegan
Soursop / guanabana|Tropical fruit|Caribbean; Central/South America|Sweet-tart, pineapple-strawberry notes|Soft, fibrous, seeded pulp|Drink, sorbet, dessert, puree|Vitamin C; fiber; potassium|Remove seeds; use ripe pulp|Vegan
Cherimoya|Tropical fruit|Andean; Mediterranean cultivation|Custardy, sweet, banana-pineapple notes|Soft creamy pulp|Spoon dessert, smoothie, ice cream|Fiber; vitamin C; potassium|Remove seeds; do not eat seeds|Vegan
Feijoa / pineapple guava|Fruit|South American; New Zealand|Perfumed, tart-sweet, guava-like|Soft seeded center, grainy|Spoon snack, jam, chutney, baking|Vitamin C; fiber|Ripe fruit is fragrant and slightly soft|Vegan
Loquat|Pome fruit|Chinese; Japanese; Mediterranean|Sweet-tart, apricot-like|Juicy, soft flesh|Fresh snack, jam, sauce, dessert|Vitamin A; vitamin C; fiber|Remove large seeds; varieties differ in sweetness|Vegan
Jujube / Chinese date|Stone fruit|Chinese; South Asian|Fresh apple-like; dried caramel-like|Crisp fresh or chewy dried|Tea, soup, dessert, snack|Vitamin C; fiber|Fresh and dried forms have different taste and uses|Vegan
Salak / snake fruit|Palm fruit|Indonesian; Southeast Asian|Sweet-sour, pineapple-like|Crisp, segmented flesh|Snack, pickle, fruit salad|Carbohydrate; fiber|Peel carefully; remove seed|Vegan
Mangosteen|Tropical fruit|Southeast Asian|Sweet, delicate, lightly tart|Soft white segments|Fresh dessert, fruit plate, preserve|Fiber; vitamin C|Open thick rind without staining flesh|Vegan
`));

data['Grains & Pulses'].push(...parse(`
Lentils, French green / Puy-style|Pulse cultivar|France; Mediterranean|Peppery, earthy|Small, firm; holds shape|Salad, side, stew, braise|Protein; fiber; folate; iron|Simmer gently; some protected-origin labels identify region|Vegan; legume sensitivity possible
Lentils, black beluga|Pulse variety|South Asian; global cultivation|Earthy, rich|Tiny, glossy; stays intact|Salad, bowl, side, curry|Protein; fiber; folate; iron|Rinse and simmer until tender|Vegan; legume sensitivity possible
Lentils, yellow split|Pulse|South Asian; Middle Eastern|Mild, sweet, earthy|Breaks down creamy|Dal, soup, stew|Protein; fiber; folate|Rinse; cook without soaking|Vegan; legume sensitivity possible
Lentils, black urad|Pulse|South Asian|Earthy, rich, savory|Whole bean firm; split form creamy|Dal makhani-style dishes, dosa batter|Protein; fiber; iron; folate|Soak whole beans; pressure cooking is common|Vegan; legume sensitivity possible
Chana dal|Split pulse|South Asian|Nutty, sweet|Keeps slight bite|Dal, snack, tempering, batter|Protein; fiber; folate; iron|Soak for quicker cooking|Vegan; legume sensitivity possible
Toor dal / pigeon pea split|Split pulse|South Asian|Earthy, mild|Softens to thick puree|Sambar, dal, stew|Protein; fiber; folate|Rinse and cook until soft|Vegan; legume sensitivity possible
Arborio rice|Rice variety|Italy|Mild, starchy|Creamy risotto texture, firm center|Risotto, rice pudding|Carbohydrate; small amounts of protein|Add warm broth gradually and stir|Naturally gluten-free
Carnaroli rice|Rice variety|Northern Italy|Mild, starchy, slightly nutty|Creamy outside, resilient center|Risotto, rice dishes|Carbohydrate; small amounts of protein|Higher starch retention suits creamy rice dishes|Naturally gluten-free
Bomba rice|Short-grain rice variety|Valencia, Spain|Mild, absorbent|Distinct, separate grains|Paella, rice casserole|Carbohydrate; small amounts of protein|Absorbs liquid while holding shape|Naturally gluten-free
Japonica rice|Rice subspecies|Japan; Korea; China|Mild, lightly sweet|Short to medium grain; sticky|Rice bowls, sushi, porridge|Carbohydrate; small amounts of B vitamins|Cultivars include many regional types|Naturally gluten-free
Indica rice|Rice subspecies|South Asia; Southeast Asia|Mild, aromatic or neutral by cultivar|Long grain, generally separate|Pilaf, biryani, curry accompaniment|Carbohydrate; small amounts of B vitamins|Varieties span aromatic basmati and non-aromatic types|Naturally gluten-free
Red rice|Whole grain rice variety|Bhutan; South Asia; Southeast Asia|Nutty, earthy|Chewy, red bran layer|Pilaf, salad, bowl, side|Fiber; anthocyanins/phenolics; magnesium|Cooking time varies; rinse before cooking|Naturally gluten-free
Brown basmati rice|Whole grain aromatic rice|South Asian|Fragrant, nutty|Long, chewy bran layer|Pilaf, bowl, curry accompaniment|Fiber; manganese; magnesium|Soak to reduce cooking time|Naturally gluten-free
Wild rice blend|Grain blend|North American; global packaged food|Nutty, earthy|Mixed grains with varied chew|Pilaf, salad, stuffing, soup|Fiber; protein; minerals vary by blend|Check blend label; some include brown rice|Usually gluten-free; check cross-contact
Durum wheat|Hard wheat grain|Mediterranean; West Asian heritage|Nutty, wheaty|Firm protein-rich grain|Pasta, couscous, semolina|Protein; selenium; B vitamins|Durum flour is used for pasta texture|Contains gluten
Spelt|Ancient wheat|Central European|Nutty, sweet|Chewy grain; flour forms gluten|Bread, pasta, grain salad|Protein; fiber; manganese|Whole kernels need longer cooking|Contains gluten
Kamut / khorasan wheat|Wheat variety|West Asian heritage; North American cultivation|Buttery, nutty|Large chewy kernels|Porridge, salad, bread|Protein; selenium; fiber|Soak kernels for faster cooking|Contains gluten
Einkorn wheat|Ancient wheat|West Asian heritage; European|Sweet, nutty|Small kernels; tender chew|Bread, porridge, pasta|Protein; carotenoids; minerals|Flour absorbs differently from modern wheat|Contains gluten
Triticale|Wheat-rye hybrid grain|Europe; North America|Mild, rye-like, nutty|Chewy whole grain|Bread, porridge, grain salad|Fiber; protein; manganese|Use whole berries in pilaf or flour in baking|Contains gluten
Fonio|Small-seeded grain|West Africa|Mild, nutty, delicate|Fine, fluffy|Pilaf, porridge, couscous-style side|Carbohydrate; amino acids; iron varies|Rinse well; cooks quickly|Naturally gluten-free; check cross-contact
Kañiwa / cañihua|Pseudocereal|Andean highlands|Earthy, nutty|Tiny, crunchy when toasted|Porridge, salad, pilaf|Protein; fiber; iron|Rinse; toast to deepen flavor|Naturally gluten-free
`));

data['Spices & Herbs'].push(...parse(`
Mexican oregano|Leaf herb|Mexico; Central America|Citrusy, peppery, floral|Dried leaf|Beans, chili, salsa, marinades|Small culinary quantities|Distinct from Mediterranean oregano; use as a regional herb|Vegan
Epazote, dried|Leaf herb|Mexico|Pungent, resinous, anise-like|Dried leaf|Beans, soups, quesadillas|Small culinary quantities|Use sparingly; dried form is concentrated|Vegan
Lemon verbena|Leaf herb|South American; Mediterranean cultivation|Strong lemon, floral|Narrow aromatic leaf|Tea, dessert, fish, syrup|Small culinary quantities|Infuse and remove; fresh aroma is delicate|Vegan
Lemongrass powder|Dried aromatic herb|Southeast Asian|Citrus, grassy|Fine dried powder|Curry rub, soup, marinade|Small culinary quantities|Use less than fresh stalk; powder texture remains|Vegan
Pandan extract|Flavoring|Southeast Asian|Vanilla-like, grassy|Concentrated liquid|Cake, custard, rice, drinks|Negligible nutrients per use|Strength and color vary by brand|Check ingredients and color additives
Shiso, red|Leaf herb variety|Japanese; East Asian|Cinnamon, mint, clove, herbal|Aromatic purple leaf|Pickle, garnish, drink, rice seasoning|Small culinary quantities|Often used for color and pickling|Vegan
Thai basil|Leaf herb variety|Thai; Southeast Asian|Anise, licorice, peppery|Sturdier leaf than sweet basil|Curry, stir-fry, noodle dishes|Vitamin K; small amounts of vitamin A|Add near the end; distinct from Italian basil|Vegan
Holy basil / tulsi|Leaf herb|South Asian; Thai cooking|Peppery, clove-like, herbal|Tender aromatic leaf|Stir-fry, tea, curry|Small culinary quantities|Culinary varieties vary; use fresh when possible|Vegan
Lemon basil|Leaf herb|Southeast Asian|Citrus, sweet basil|Tender leaf|Salad, fish, curry, tea|Small culinary quantities|Add late to preserve citrus aroma|Vegan
Garlic chives|Allium herb|Chinese; East Asian|Garlicky, onion-like|Flat leaves; blossoms edible|Dumplings, pancake, stir-fry, omelet|Vitamin K; vitamin C|Use raw or quick-cooked; blossoms are edible|Vegan
Mexican cinnamon / canela|Bark spice|Mexico; Central America|Sweet, soft, citrusy|Thin bark rolls|Hot chocolate, rice pudding, mole-style dishes|Small culinary quantities|Softer bark crumbles more readily than cassia|Vegan
Ceylon cinnamon|Bark spice|Sri Lanka|Delicate, floral, mildly sweet|Thin layered quills|Baking, tea, curry, fruit|Small culinary quantities|Flavor is gentler than cassia cinnamon|Vegan
Cassia cinnamon|Bark spice|China; Indonesia; Vietnam|Strong, hot, woody|Thick single bark curl|Baking, chai, braise, spice blends|Small culinary quantities|Commonly sold as cinnamon; varieties differ|Vegan
Tellicherry peppercorn|Black pepper variety|Malabar Coast, India|Citrusy, woody, robust|Large, mature peppercorn|Finishing, steak, sauces, spice blends|Small culinary quantities|Grind fresh; origin grading varies|Vegan
Kampot pepper|Pepper variety|Kampot, Cambodia|Floral, fruity, mild heat|Whole peppercorns, multiple colors|Seafood, salad, fruit, finishing|Small culinary quantities|Geographic indication; black/red/white types differ|Vegan
Long pepper|Pepper spice|South Asia; Southeast Asia|Sweet-hot, cinnamon-like|Long dried catkin|Pickle, braise, spice blends, dessert|Small culinary quantities|Grind or simmer whole; stronger than black pepper|Vegan
Green cardamom, smoked|Seed pod preparation|South Asia; Middle Eastern specialty|Smoky, floral, eucalyptus|Aromatic pod|Rice, meat, desserts, tea|Small culinary quantities|Use as a specialty variation; smoke level varies|Vegan
Kaffir lime leaf powder|Dried citrus leaf|Southeast Asian|Intensely citrus, floral|Fine leaf powder|Curry paste, soup, marinade|Small culinary quantities|Use lightly; fresh leaf is brighter|Vegan
Sumac, whole berry|Berry spice|Levantine; Middle Eastern|Tart, fruity, tannic|Coarse dried berry|Infusion, spice rub, souring agent|Small culinary quantities|Ground sumac is more common; salt may be added|Vegan
Svanetian salt|Seasoned salt blend|Georgia, Caucasus|Salty, garlic, coriander, fenugreek|Coarse blend|Potatoes, beans, grilled meat, vegetables|Sodium; spice nutrients negligible per use|Recipes vary regionally; reduce additional salt|Check blend for allergens
`));

data['Protein & Dairy'].push(...parse(`
King salmon / Chinook|Oily fish species|North Pacific|Rich, buttery, robust|Large, moist flakes|Roast, grill, smoke, rice bowl|Protein; omega-3; vitamin D; B12|Wild and farmed products differ; cook safely|Fish allergen
Sockeye salmon|Oily fish species|North Pacific|Deep, savory, more pronounced than mild salmon|Firm, red flesh|Grill, roast, smoke, salad|Protein; omega-3; vitamin D; B12|Lean firm flesh can dry if overcooked|Fish allergen
Arctic char|Oily fish|Arctic and sub-Arctic regions|Mild, trout-like, lightly rich|Fine, tender flakes|Pan-sear, roast, smoke|Protein; omega-3; B12|Cook gently; skin crisps well|Fish allergen
Yellowtail / hamachi|Fish|Japan; Pacific|Rich, buttery, clean|Tender, fatty flesh|Sashimi-style dishes, grill, rice bowls|Protein; omega-3; B12|Use raw only when handled for raw consumption|Fish allergen
Sablefish / black cod|Oily fish|North Pacific|Very rich, buttery|Moist, silky|Miso glaze, roast, smoke|Protein; omega-3; selenium|High fat suits broiling and roasting|Fish allergen
Monkfish|White fish|North Atlantic; Mediterranean|Mild, sweet, lobster-like|Firm, dense medallions|Braise, grill, stew, skewer|Protein; selenium; B12|Remove membrane before cooking|Fish allergen
Eel / unagi|Fish|East Asian; European|Rich, sweet-savory|Firm, oily flesh|Grill, sushi, rice bowl, stew|Protein; vitamin A; omega-3|Commercial cooked eel is commonly glazed|Fish allergen
Smoked haddock|Smoked white fish|British; Northern European|Smoky, salty, savory|Firm flakes|Chowder, kedgeree-style rice, pie|Protein; iodine; sodium varies|Salt content varies; adjust recipe seasoning|Fish allergen
Kippers|Cured herring|British; Northern European|Smoky, salty, oily|Firm, flaky|Breakfast, toast, potato dishes|Protein; omega-3; vitamin D|Often ready-cooked but check package|Fish allergen
Herring|Oily fish|Northern European; Baltic; Japanese|Rich, savory|Soft, oily|Pickle, smoke, grill, salad|Protein; omega-3; vitamin D; B12|Pickled styles are high in sodium|Fish allergen
Anchovy paste|Fish condiment|Mediterranean; European|Concentrated salty umami|Smooth paste|Dressing, pasta, sauce base, stew|Protein; sodium; omega-3|Use small amount and account for salt|Fish allergen
Razor clam|Mollusk|European; Pacific coastal|Sweet, briny|Long, tender muscle|Grill, saute, pasta, steam|Protein; iron; B12|Purge grit; cook briefly|Mollusk allergen
Abalone|Mollusk|East Asian; Pacific coastal|Sweet, briny, mildly nutty|Firm, chewy if overcooked|Steam, braise, grill, soup|Protein; iron; B12|Tenderize and cook according to species/size|Mollusk allergen
Lamb sweetbreads|Offal|French; Mediterranean; South Asian|Mild, creamy, savory|Tender, delicate|Fry, grill, braise, sauce|Protein; B12; minerals|Requires careful trimming and food-safe handling|Animal product
Beef liver|Offal|Global|Mineral, intense, slightly sweet|Soft, dense|Pate, saute, stew, dumpling filling|Protein; iron; vitamin A; B12|Avoid excess consumption due to high vitamin A; cook safely|Animal product
Chicken liver|Offal|European; South Asian; global|Rich, mineral, savory|Soft, creamy when cooked gently|Pate, curry, saute, stuffing|Protein; iron; vitamin A; B12|Cook through; chill promptly|Animal product
Pork liver|Offal|East Asian; European|Strong, mineral, savory|Dense, tender if brief-cooked|Pate, stir-fry, sausage, curry|Protein; iron; vitamin A; B12|Cook safely; flavor is stronger than chicken liver|Animal product
Paneer, smoked|Fresh cheese variation|South Asian; contemporary|Mild, milky with smoke|Firm, holds shape|Grill, curry, sandwich, salad|Protein; calcium; fat|Smoke level varies by producer|Milk allergen
Fior di latte|Fresh mozzarella style|Southern Italy|Milky, sweet, lightly tangy|Soft, moist, stretchy|Pizza, salad, caprese-style dishes|Protein; calcium; sodium|Drain before use on pizza|Milk allergen
Burrata|Fresh cheese|Apulia, Italy|Creamy, buttery, mild|Outer mozzarella with soft cream center|Salad, toast, pasta, starter|Protein; calcium; fat|Serve fresh; highly perishable|Milk allergen
Pecorino Romano|Hard sheep cheese|Italy; Sardinia|Salty, sharp, sheep-milk tang|Hard, grates finely|Pasta, soup, salad, finishing|Protein; calcium; sodium|Use less salt in the dish|Milk allergen
Manchego|Aged sheep cheese|La Mancha, Spain|Nutty, savory, lightly sweet|Firm, sliceable|Tapas, sandwich, salad, cheese board|Protein; calcium; fat|Aging changes firmness and intensity|Milk allergen
Feta, sheep/goat milk|Brined cheese type|Greece; Eastern Mediterranean|Tangy, salty, more sheep-milk richness|Crumbly, creamy|Salad, pastry, dip, grain bowl|Protein; calcium; sodium|Protected-origin versions have defined standards|Milk allergen
`));

data['Pantry & Condiments'].push(...parse(`
Tamari, wheat-free|Fermented soy condiment|Japan; global|Deep, salty umami|Thin liquid|Dipping, marinade, stir-fry|Sodium; small amounts of protein|Only use if label confirms gluten-free status|Soy allergen
Soy sauce, light|Fermented condiment|Chinese|Salty, bright umami|Thin, pale brown liquid|Stir-fry, marinade, seasoning|Sodium; small amounts of protein|Light refers to color/body, not reduced sodium|Soy; often wheat
Soy sauce, dark|Fermented condiment|Chinese|Mellow, slightly sweet, deep umami|Thicker, darker liquid|Braise, color, sauce base|Sodium; sugars vary|Use for color and depth, not as sole salt|Soy; often wheat
Kecap manis|Sweet soy sauce|Indonesian|Sweet, molasses-like, savory|Thick, glossy|Nasi goreng-style dishes, glaze, marinade|Added sugars; sodium|Very sweet; reduce other sweeteners|Soy; often wheat
Ponzu sauce|Citrus soy condiment|Japanese|Citrus, salty, umami|Thin sauce|Dipping, dressing, grilled fish, salad|Sodium; sugars vary|Yuzu and other citrus versions vary|Soy; may contain wheat or fish
Yuzu kosho|Citrus chili paste|Japanese|Salty, hot, aromatic citrus|Coarse paste|Noodles, fish, soup, grilled meats|Sodium; chili compounds|A little gives strong flavor|Citrus/chili; check label
Chimichurri|Herb sauce|Argentine; Uruguayan|Garlic, parsley, vinegar, chili|Loose herb-oil sauce|Grilled meat, vegetables, bread, beans|Unsaturated fat; sodium varies|Fresh herbs and oil; recipe varies|Vegan
Salsa macha|Chili oil condiment|Mexican|Nutty, smoky, hot|Oil with toasted chili and seeds/nuts|Tacos, eggs, noodles, roasted vegetables|Fat; sodium varies|Chili and nut/seed content vary by maker|May contain peanuts, sesame or tree nuts
Salsa verde, Mexican|Tomatillo sauce|Mexican|Tart, herbal, chili heat|Chunky or smooth|Tacos, eggs, stew, enchiladas|Vitamin C; sodium varies|Roast or boil tomatillos for different flavor|Vegan; check commercial product
Harissa, rose|Chili paste variation|North African|Hot, floral, sweet-spiced|Thick paste|Marinade, dip, couscous, roast vegetables|Sodium varies; chili compounds|Rose and spice additions vary by recipe|Check label for allergens
Chutney, mango|Fruit condiment|South Asian; British Indian|Sweet, sour, spiced|Chunky to smooth|Curry, cheese, sandwich, snack|Added sugar; sodium varies|Fresh, cooked and preserved versions differ|Vegan; check packaged label
Chutney, mint-cilantro|Herb condiment|South Asian|Fresh, sour, herbal, chili-hot|Smooth green paste|Snack, sandwich, grilled foods, chaat|Vitamin C; sodium varies|Add acid to maintain brightness; keep chilled|Vegan; yogurt may be added in some recipes
Relish, corn|Vegetable condiment|North American|Sweet, tangy|Chunky|Hot dog, sandwich, beans, salad|Added sugar; sodium varies|Sweet and dill styles differ|Vegan; check label
Ajvar|Pepper-eggplant relish|Balkan|Sweet roasted pepper, smoky, savory|Thick spread|Bread, grilled meat, vegetables, dip|Vitamin C; sodium varies|Roasted pepper and eggplant ratios vary|Usually vegan; check product
Tkemali|Sour plum sauce|Georgian|Tart, herbal, fruity|Thin to spoonable sauce|Grilled meat, potatoes, beans, dumplings|Sugars and sodium vary|Green and ripe plum versions taste different|Often vegan; check recipe
Amba|Pickled mango condiment|Iraqi; Indian Jewish; Middle Eastern|Sour, salty, spicy, fenugreek|Thick sauce|Sandwich, falafel, fish, eggs|Sodium; sugars vary|Heat and sourness vary by brand|Check mustard/fenugreek and allergens
Pomegranate molasses, sour|Fruit reduction variation|Levantine; Middle Eastern|Very tart, concentrated|Thick syrup|Stew, dressing, glaze, marinade|Mostly sugars; organic acids|Some products are unsweetened; label varies|Vegan
Malt vinegar|Malt condiment|British|Toasty, sharp, malty|Thin liquid|Chips, pickle, dressing, marinade|Negligible micronutrients|Barley malt may contain gluten|Contains barley/gluten
White wine vinegar|Fermented vinegar|French; Mediterranean|Sharp, light, fruity|Thin liquid|Vinaigrette, sauce, pickle|Negligible micronutrients|Use when colorless acidity is desired|Vegan
Champagne vinegar|Wine vinegar|French|Mild, delicate, floral acid|Thin liquid|Vinaigrette, fish, fruit, sauce|Negligible micronutrients|Gentler than many red wine vinegars|Vegan
Red palm oil|Plant oil|West African; Central African|Earthy, rich, carotenoid aroma|Deep orange-red fat|Soups, stews, rice dishes|Fat; carotenoids; vitamin E|Use sustainably sourced product; flavor is strong|Vegan
Mustard oil|Seed oil|South Asian; Bengali; Nepalese|Pungent, sharp, mustardy|Liquid oil|Pickle, fish curry, saute|Unsaturated fat; erucic acid varies|Regulatory food use and cooking practices vary by country|Mustard allergen risk
Ghee, cultured|Clarified butter variation|South Asian|Deep, nutty, cultured tang|Clarified fat|Tempering, rice, roast, baking|Fat; small amounts of fat-soluble vitamins|Flavor differs from sweet-cream ghee|Dairy-derived
Palm shortening|Plant fat|Global baking; tropical regions|Neutral|Solid fat|Pastry, frying, baking|Fat; composition varies|Palm source and blend vary; check label|Vegan
Rice bran oil|Rice by-product oil|East Asian; South Asian|Mild, nutty|Liquid cooking oil|Stir-fry, frying, baking|Mostly unsaturated fats; vitamin E|Refined products are neutral|Vegan
Grapeseed oil, cold-pressed|Seed oil variation|European|Fruity, green, light|Liquid finishing oil|Dressing, finishing, mayonnaise|Polyunsaturated fat; vitamin E|More aromatic than refined grapeseed oil|Vegan
Hemp seed oil|Seed oil|Central Asian heritage; global|Nutty, grassy|Delicate liquid oil|Cold dressing, finishing, dips|Omega-3 ALA; omega-6 fats|Do not heat strongly; refrigerate after opening|Vegan
Cashew cream|Plant-based preparation|South Asian; contemporary plant-based cooking|Mild, creamy, nutty|Smooth thick cream|Curry, soup, sauce, dessert|Unsaturated fat; protein; copper|Blend soaked cashews with water; thin to desired body|Tree-nut allergen
`));

data['Vegetables'].push(...parse(`
Jute mallow / saluyot|Leafy plant|West Africa; Egypt; South Asia|Mild, green, mucilaginous|Silky leaves when cooked|Soup, stew, greens dishes|Fiber; vitamin A precursors; iron|Cook leaves thoroughly; texture thickens broth|Vegan
Callaloo leaves|Leafy greens blend|Caribbean|Earthy, spinach-like|Tender cooked greens|Stew, soup, saute, rice dishes|Folate; vitamins A and C; iron|Name can refer to different local greens by island|Vegan
Water spinach / kangkong|Aquatic leafy vegetable|Southeast Asia; South China|Mild, grassy|Hollow stems, tender leaves|Stir-fry, soup, curry|Vitamin A; vitamin C; iron|Wash thoroughly; cook quickly|Vegan
Ong choy|Aquatic leafy vegetable|Cantonese; Southeast Asian|Fresh, mild, grassy|Hollow crisp stems|Stir-fry, garlic saute, soup|Vitamin A; vitamin C; folate|Stems need slightly more cooking than leaves|Vegan
Tatsoi|Leafy brassica|East Asian|Mild mustard, sweet|Tender spoon-shaped leaves|Salad, stir-fry, soup|Vitamins C and K; folate|Use raw when young or quick-cook|Vegan
Mizuna|Leafy brassica|Japanese; East Asian|Peppery, mild mustard|Feathery crisp leaves|Salad, hot pot, stir-fry|Vitamins A, C and K|Add late to cooked dishes|Vegan
Komatsuna|Leafy brassica|Japanese|Mild mustard, sweet|Sturdy stems and leaves|Soup, pickle, saute, hot pot|Vitamins C and K; calcium|Cook stems briefly before leaves|Vegan
Gai choy / mustard cabbage|Brassica|Chinese; Southeast Asian|Peppery, mustardy|Sturdy leaves and stems|Pickle, stir-fry, soup|Vitamins A, C and K; folate|Salt or blanch to moderate bitterness|Vegan
Yu choy|Leafy brassica|Chinese; Cantonese|Sweet, mild mustard|Tender stems and leaves; edible flowers|Blanch, stir-fry, soup|Vitamin C; folate; calcium|Cook until stems are just tender|Vegan
Tindora leaves|Leafy plant part|South Asian|Mild, green|Tender leaves|Saute, curry, greens dishes|Fiber; micronutrients vary|Use young leaves; wash well|Vegan
Fiddlehead ferns|Young fern frond|Northeastern North America; East Asia|Green, grassy, asparagus-like|Tender, curled frond|Saute, pickle, pasta, side|Fiber; vitamin A; vitamin C|Cook thoroughly; never eat raw|Vegan
Nopales / cactus pads|Cactus stem vegetable|Mexico; Southwestern US|Tart, green, slightly citrusy|Crisp-tender; mucilage when cut|Tacos, salad, stew, grill|Fiber; vitamin C; magnesium|Remove spines; rinse or grill to reduce mucilage|Vegan
Huitlacoche / corn smut|Edible fungus on corn|Mexico|Earthy, smoky, mushroom-like|Soft, juicy|Taco, quesadilla, soup, sauce|Protein; fiber; minerals vary|Fresh and canned forms differ|Vegan; fungus on corn
Sambal leaf / daun kesum|Aromatic leafy herb|Malaysia; Vietnam; Southeast Asia|Citrusy, minty, pungent|Narrow aromatic leaf|Laksa-style soup, salad, herb plate|Small culinary quantities|Use fresh; names and species vary regionally|Vegan
Tamarind leaves|Leafy plant part|South Asia; Southeast Asia|Tart, sour, green|Tender leaves|Soup, curry, sour stew|Vitamin C; micronutrients vary|Use young leaves; cook briefly|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Ackee|Tropical fruit|Jamaica; West Africa heritage|Mild, buttery, savory|Soft yellow arils|Ackee and saltfish, stew|Fat; vitamin C; protein varies|Only eat fully ripe arils; never eat unopened fruit|Food safety depends on ripeness and preparation
Breadfruit|Multiple fruit|Pacific Islands; Caribbean|Starchy, mild; sweet when ripe|Potato-like when cooked|Roast, fry, curry, chips, flour|Carbohydrate; fiber; potassium|Maturity determines whether used as vegetable or fruit|Vegan
Sweetsop / sugar apple|Tropical fruit|Caribbean; South Asia|Very sweet, custardy|Soft segmented flesh with seeds|Fresh dessert, drink, puree|Vitamin C; fiber; carbohydrate|Remove seeds; eat ripe flesh only|Vegan
Jaboticaba|Tree fruit|Brazil; South America|Sweet-tart, grape-like|Juicy dark skin and pulp|Jelly, wine, juice, fresh snack|Vitamin C; anthocyanins|Highly perishable after harvest|Vegan
Cupuaçu|Amazonian fruit|Brazil; Amazon basin|Tart, chocolate-pineapple aroma|Creamy pulp around seeds|Juice, confectionery, dessert|Fiber; vitamin C; composition varies|Usually sold as pulp or frozen puree|Vegan
Tucuma|Palm fruit|Amazon basin|Nutty, earthy, mildly sweet|Oily orange flesh|Savory filling, snack, regional dishes|Fat; carotenoids|Local varieties and culinary methods differ|Vegan
Nance / nanche|Tropical fruit|Central America; Caribbean|Sweet-tart, pungent, cheesy-fruity|Small soft drupe|Drink, preserve, dessert, fermented beverage|Vitamin C; carotenoids|Ripe fruit is aromatic; often preserved|Vegan
Mamey sapote|Tropical fruit|Mexico; Caribbean|Sweet, pumpkin-pie, almond notes|Dense, creamy orange flesh|Shake, dessert, custard, ice cream|Vitamin A; fiber; potassium|Allow to ripen fully; remove large seed|Vegan
Black sapote|Tropical fruit|Mexico; Central America|Sweet, mild, cocoa-like when ripe|Soft dark custard flesh|Dessert, drink, puree|Vitamin C; fiber|Eat fully ripe; immature fruit is astringent|Vegan
White sapote|Tropical fruit|Mexico; Central America|Sweet, pear-banana, floral|Soft custard-like flesh|Fresh dessert, smoothie|Vitamin C; carbohydrate|Remove seeds; ripe fruit bruises easily|Vegan
Canistel / eggfruit|Tropical fruit|Central America; Caribbean|Sweet, pumpkin, egg-yolk-like|Dry, dense, creamy flesh|Dessert, shake, custard|Carotenoids; fiber|Eat fully ripe; texture is naturally mealy|Vegan
Ambarella / June plum|Tropical fruit|Pacific; Caribbean; South Asia|Tart unripe; sweet-sour ripe|Crisp to juicy|Pickle, chutney, juice, snack|Vitamin C; fiber; carbohydrate|Ripeness changes acidity and texture|Vegan
Rose apple|Tropical fruit|South/Southeast Asia|Light, floral, watery|Crisp, airy flesh|Fresh snack, salad, pickle|Vitamin C; water|Several Syzygium species share common names|Vegan
Wax apple|Tropical fruit|Southeast Asia|Mild, lightly sweet|Crisp, juicy, bell-shaped|Fresh snack, salad, pickle|Water; vitamin C|Cultivars vary in sweetness and color|Vegan
Langsat / lanzones|Tropical fruit|Southeast Asia|Sweet-tart, citrusy|Juicy segmented flesh|Fresh snack, dessert|Vitamin C; carbohydrate|Peel and remove seeds; bitter latex may occur|Vegan
Santol|Tropical fruit|Southeast Asia|Sweet-sour, cottony pulp|Soft segments around seeds|Sour curry, preserve, snack|Fiber; vitamin C|Seeds are not edible; texture varies by cultivar|Vegan
Tamarillo / tree tomato|Fruit|Andean; New Zealand cultivation|Tart, tangy, savory-fruity|Juicy seed-filled flesh|Chutney, sauce, drink, salad|Vitamin C; provitamin A|Peel can be bitter; use ripe fruit|Vegan
Cape gooseberry / goldenberry|Berry|Andean; global cultivation|Sweet-tart, tropical|Juicy small berry in husk|Dessert, salad, preserve, garnish|Vitamin C; fiber; carotenoids|Remove papery husk; rinse sticky surface|Vegan
Sea buckthorn berry|Berry|Northern Europe; Central Asia|Very tart, resinous, citrusy|Small, oily berry|Juice, syrup, puree, sauce|Vitamin C; carotenoids; omega-7 fats|Often processed rather than eaten plain|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Pearl millet / bajra|Whole grain|Sahel; South Asia|Earthy, nutty|Dense grain or flour|Flatbread, porridge, fermented batter|Fiber; iron; magnesium|Toast or ferment flour for flavor|Naturally gluten-free; cross-contact possible
Finger millet / ragi|Whole grain|East Africa; South India|Earthy, malty|Fine flour; porridge-like|Porridge, flatbread, malted drinks|Calcium; fiber; iron|Roast flour lightly before cooking|Naturally gluten-free
Foxtail millet|Whole grain|China; South Asia|Mild, nutty|Small, fluffy grains|Porridge, pilaf, fermented foods|Fiber; protein; minerals|Rinse; use water ratio per package|Naturally gluten-free
Barnyard millet|Whole grain|South Asia; East Asia|Mild, grassy|Quick-cooking small grain|Porridge, pilaf, steamed side|Fiber; iron; phosphorus|Rinse and cook until fluffy|Naturally gluten-free
Kodo millet|Whole grain|South Asia|Earthy, nutty|Small chewy grain|Pilaf, porridge, fermented batter|Fiber; protein; minerals|Sort and rinse well before cooking|Naturally gluten-free
Little millet|Whole grain|South Asia|Mild, nutty|Small, tender grains|Pilaf, porridge, side dish|Fiber; iron; B vitamins|Rinse; cook like small-grain rice|Naturally gluten-free
Proso millet|Whole grain|Eurasia; North America|Mild, sweet|Fluffy when cooked|Porridge, pilaf, birdseed and flour uses|Protein; magnesium; phosphorus|Toast for nuttiness|Naturally gluten-free
Tepary beans|Pulse|Sonoran desert, Mexico and Arizona|Nutty, earthy, slightly sweet|Firm, creamy when cooked|Stew, salad, beans and grains|Protein; fiber; iron|Drought-adapted traditional crop; soak and simmer|Vegan; legume sensitivity possible
Bambara groundnut|Pulse|West Africa|Earthy, nutty|Creamy cooked seed|Stew, flour, snack, porridge|Protein; carbohydrate; minerals|Cook thoroughly; local processing varies|Vegan; legume sensitivity possible
Lablab / hyacinth bean|Pulse|South Asia; East Africa|Mild, nutty|Firm pods or dried beans|Curry, stew, flour, vegetable dish|Protein; fiber; iron|Mature dried beans require thorough cooking|Vegan; legume sensitivity possible
Winged bean|Pod legume|Southeast Asia; Pacific|Nutty, green|Crisp pods; edible leaves and tubers|Stir-fry, salad, curry, soup|Protein; fiber; vitamins|Cook pods; use other plant parts only as locally prepared|Vegan; legume sensitivity possible
Velvet bean / mucuna|Pulse|South Asia; Africa|Earthy, bean-like|Dried seed|Traditional processed foods|Protein; carbohydrate|Requires detoxifying processing; never eat raw or casually prepare|Use only food-grade, properly processed products
Lupini beans|Pulse|Mediterranean; Andean|Firm, tangy, nutty|Chewy, brined bean|Snack, salad, antipasti|Protein; fiber; low starch|Must be debittered and brined before eating|Legume allergen; lupin is a regulated allergen in some regions
Pigeon pea, green|Fresh pulse|South Asia; Caribbean; East Africa|Sweet, grassy, nutty|Tender fresh bean|Rice dishes, stew, curry, soup|Protein; fiber; folate|Fresh/frozen beans cook faster than dried|Vegan; legume sensitivity possible
Black gram split|Pulse|South Asian|Earthy, rich|Creamy split bean|Dal, batter, papad, tempering|Protein; fiber; iron|Soak where needed; recipes vary by split/skin-on form|Vegan; legume sensitivity possible
Red kidney bean, dark|Bean cultivar|Central America; Caribbean|Earthy, robust|Firm, holds shape|Chili, stew, rice dishes|Protein; fiber; folate|Boil raw soaked beans thoroughly before slow cooking|Vegan; legume sensitivity possible
White hominy|Nixtamalized corn|Mesoamerican; Mexican; Southwestern US|Sweet corn, alkaline notes|Chewy kernels|Pozole, stew, grits|Carbohydrate; calcium varies|Prepared dried hominy needs long cooking; canned is ready-to-use|Naturally gluten-free
Hominy grits|Milled nixtamalized corn|Southern US|Toasty corn, mild|Creamy porridge|Breakfast, shrimp and grits, side dish|Carbohydrate; fiber; calcium varies|Stir while simmering to prevent lumps|Naturally gluten-free
Purple corn flour|Milled corn|Andean; Peru|Earthy, mildly sweet|Fine flour, purple pigments|Drink, porridge, baked goods|Anthocyanins; carbohydrate|Color and flavor vary by corn variety|Naturally gluten-free; check cross-contact
`));
data['Spices & Herbs'].push(...parse(`
Sansho pepper|Citrus pepper spice|Japan|Citrus, floral, mildly numbing|Fine powder or dried husk|Unagi, noodle, fish, finishing|Small culinary quantities|Japanese relative of Sichuan pepper; aroma is more citrus-forward|Vegan
Grains of paradise|Seed spice|West Africa|Peppery, ginger, citrus|Small seeds|Stew, meat, gin-style drinks, spice blends|Small culinary quantities|Grind fresh for aromatic heat|Vegan
Negro pepper / uda|Seed spice|West Africa|Woody, smoky, peppery|Dark pods or seeds|Soup, stew, pepper soup|Small culinary quantities|Use sparingly; names vary by region|Vegan
Selim pepper|Seed spice|West Africa|Smoky, resinous, peppery|Long pods|Stew, soup, spice blend|Small culinary quantities|Often used whole in broths|Vegan
Kororima / Ethiopian cardamom|Seed spice|Ethiopia; Eritrea|Smoky, citrus, cardamom-like|Dark seed pods|Berbere, coffee, stew, spice blends|Small culinary quantities|Distinct from green cardamom|Vegan
Lesser galangal|Rhizome|Southeast Asia|Camphor, pine, citrus|Hard, fibrous rhizome|Herbal pastes, soup, tonic preparations|Small culinary quantities|Use thin slices or pound into paste|Vegan
Kencur / aromatic ginger|Rhizome|Indonesia; Southeast Asia|Peppery, camphor, earthy|Small rhizome|Javanese spice pastes, rice dishes, herbal drinks|Small culinary quantities|Distinct from common ginger and galangal|Vegan
Fingerroot / krachai|Rhizome|Thailand; Southeast Asia|Mild ginger, earthy|Finger-like roots|Thai curry paste, fish dishes, stir-fry|Small culinary quantities|Fresh, frozen and pickled forms differ|Vegan
Pandan seed spice blend|Aromatic leaf blend|Southeast Asia|Grassy, vanilla-like, toasted|Dried leaf or mixed powder|Rice, dessert, drink|Small culinary quantities|Blends are commercial preparations; check label|Check additives
Mastic resin|Tree resin spice|Greek islands; Eastern Mediterranean|Piney, resinous, lightly sweet|Hard resin crystals|Bread, ice cream, custard, liqueur|Small culinary quantities|Freeze briefly before grinding to reduce stickiness|Vegan
Mahleb / mahlab kernel|Stone-fruit kernel spice|Eastern Mediterranean|Almond, cherry, floral|Ground kernel|Sweet breads, cookies, brioche|Small culinary quantities|Use sparingly; strong aroma|Possible tree-nut cross-reactivity
Amchoor / dried mango powder|Fruit spice|South Asian|Tart, fruity, dry sourness|Fine tan powder|Chaat, curry, lentils, marinade|Small amounts of vitamin C; acids|Add late for bright acidity|Vegan
Kokum|Dried fruit spice|Western India; Konkan|Tart, fruity, lightly smoky|Dried dark rind|Curry, sol kadhi, sour broth|Organic acids; small culinary quantities|Soak in warm water before adding|Vegan
Asam keping / dried sour fruit|Dried fruit spice|Malaysia; Southeast Asia|Sharp sour, fruity|Dried slices|Fish curry, soup, tamarind-style souring|Small culinary quantities|Use to sour broths; potency varies|Vegan
Pandanus / kewra water|Floral extract|South Asia; Southeast Asia|Floral, sweet, perfumed|Liquid essence|Rice, dessert, drinks, biryani-style dishes|Negligible nutrition per use|Use a few drops; concentrated flavor|Check label for artificial flavor or additives
Rose water|Floral extract|Middle East; South Asia|Floral, sweet, perfumed|Liquid essence|Dessert, drinks, syrup, rice pudding|Negligible nutrition per use|Add lightly to avoid soapy notes|Vegan; check product ingredients
Orange blossom water|Floral extract|North Africa; Middle East; Mediterranean|Citrus-floral, delicate|Liquid essence|Pastry, drink, syrup, fruit salad|Negligible nutrition per use|Strength differs by brand; use sparingly|Vegan
Vanilla, Mexican style|Vanilla type|Mexico; global production|Woody, spicy, creamy floral|Extract, paste or pod|Chocolate, custard, baked goods|Negligible nutrients per use|Pure extracts vary; avoid confusing imitation products|Check alcohol/additives
Tajín-style chili-lime seasoning|Seasoned spice blend|Mexico|Salty, tart, chili, citrus|Dry granular blend|Fruit, vegetables, snacks, rim seasoning|Sodium; chili compounds|Use as finishing seasoning; blend ratios vary|Check label for allergens
Berbere, mitmita-style heat blend|Spice blend|Ethiopia|Very hot, warm, aromatic|Fine red powder|Stew, meat, legumes|Small culinary quantities|Mitmita is generally hotter and distinct from berbere|Check blend composition
`));
data['Protein & Dairy'].push(...parse(`
Tilapia, Nile type|Freshwater fish species|Nile basin; global aquaculture|Mild, clean|Lean, flaky|Bake, fry, curry, stew|Protein; selenium; B12|Species and farmed product labeling vary|Fish allergen
Catfish|Freshwater fish|West Africa; Southern US; Southeast Asia|Mild, slightly sweet|Firm, moist|Fry, grill, stew, curry|Protein; B12; selenium|Species and flavor vary by region|Fish allergen
Carp|Freshwater fish|East/Central Europe; East Asia|Earthy, mild to rich by species|Firm, bony|Soup, braise, roast, dumpling|Protein; B12; phosphorus|Check bones carefully; species vary widely|Fish allergen
Tilapia, red hybrid|Freshwater fish|Southeast Asian aquaculture|Mild, clean|Firm, flaky|Roast, steam, curry, grill|Protein; selenium; B12|Commercial hybrid; quality depends on farming|Fish allergen
Barramundi / Asian sea bass|Fish|Australia; Southeast Asia|Mild, buttery|Moist, medium flakes|Pan-fry, grill, curry, roast|Protein; omega-3 varies; B12|Cook just until opaque|Fish allergen
Walleye|Freshwater fish|North American Great Lakes|Sweet, mild|Tender, fine flakes|Pan-fry, bake, fish cakes|Protein; selenium; B12|Check for bones; cook gently|Fish allergen
Perch, yellow|Freshwater fish|North American Great Lakes|Mild, sweet|Small, delicate flakes|Fry, bake, stew|Protein; B12; selenium|Small fillets cook quickly|Fish allergen
Sturgeon|Fish|Caspian basin; Eurasia; aquaculture worldwide|Rich, savory|Firm, meaty|Roast, smoke, grill, caviar|Protein; selenium; omega-3 varies|Species and sourcing are regulated; check provenance|Fish allergen
Caviar|Cured roe|Caspian; global aquaculture|Briny, buttery, complex|Small bursting eggs|Garnish, toast, pasta, potato|Protein; B12; sodium|Keep cold; product grade and species vary|Fish allergen
Bottarga|Cured fish roe|Sardinia; Sicily; Mediterranean|Salty, savory, oceanic|Firm cured roe|Grate over pasta, salad, seafood|Protein; sodium; omega-3|Use as a concentrated seasoning|Fish allergen
Surimi|Processed seafood paste|Japan; global seafood industry|Mild, sweet, crab-like|Springy, molded|Sushi rolls, salad, soup|Protein; sodium; starch varies|Product ingredients differ; not whole crab meat|Fish; often wheat, egg or shellfish traces
Fish balls|Prepared seafood product|East Asia; Southeast Asia|Mild, savory|Springy, tender|Soup, hot pot, noodle dishes|Protein; sodium varies|Species and starch binders vary|Fish; check shellfish/egg/wheat additives
Sea cucumber|Marine invertebrate|East Asia; Pacific|Mild, absorbs sauce|Gelatinous, tender|Soup, braise, banquet dishes|Protein; collagen; sodium varies|Often sold dried; must be rehydrated and prepared correctly|Seafood allergy risk
Jellyfish, prepared|Marine invertebrate|East Asia|Mild, salty, sesame-ready|Crunchy, elastic|Cold salad, appetizer|Protein; sodium varies|Use commercially prepared product; rinse as directed|Seafood allergy risk
Escargot|Land snail|French; Mediterranean|Mild, earthy|Tender, chewy|Garlic butter, stew, pastry|Protein; iron; selenium|Use food-grade farmed product and cook thoroughly|Mollusk allergy risk
Frog legs|Amphibian meat|French; Southeast Asian; South Asian|Mild, chicken-like|Tender, fine-grained|Fry, grill, curry, stew|Protein; selenium; B12|Cook thoroughly; source legally and safely|Animal product; allergy cross-reactions possible
Camel meat|Red meat|North/East Africa; Middle East|Lean, savory, mildly sweet|Firm, low fat|Stew, roast, curry, grill|Protein; iron; zinc|Cook tougher cuts slowly; food traditions vary|Animal product
Yak meat|Red meat|Himalayan plateau; Central Asia|Lean, rich, gamey|Dense, low-fat|Stew, dried meat, grill|Protein; iron; B12|Avoid overcooking lean cuts|Animal product
Reindeer / caribou|Game meat|Arctic and sub-Arctic regions|Lean, earthy, mineral|Fine-grained, low-fat|Stew, roast, sausage, dried meat|Protein; iron; B12|Wild and farmed sourcing differs by region|Animal product
Llama / alpaca meat|Red meat|Andean region|Lean, mild game flavor|Fine-grained|Stew, grill, dried meat|Protein; iron; B12|Traditional regional use; source food-grade meat|Animal product
Insects, cricket flour|Edible insect product|Global; traditional use in many regions|Nutty, toasted|Fine protein-rich powder|Baking, snack, protein blend|Protein; iron; fiber/chitin|Use only regulated food-grade product; allergen cross-reaction possible|Insect allergen; possible crustacean cross-reactivity
Mealworms, edible|Edible insect|Europe; East Asia; traditional global use|Nutty, toasted|Crunchy dried or soft cooked|Snack, protein powder, garnish|Protein; fat; minerals|Only food-grade farmed product; cook according to supplier|Insect allergen; possible shellfish cross-reactivity
Buffalo mozzarella|Fresh cheese|Campania, Italy|Milky, tangy, rich|Soft, juicy|Pizza, salad, baked dishes|Protein; calcium; fat|Drain before cooking; buffalo milk has richer flavor|Milk allergen
Queso Oaxaca|Pasta-filata cheese|Oaxaca, Mexico|Mild, buttery|Stringy, melts smoothly|Quesadilla, tlayuda, baked dishes|Protein; calcium; sodium|Shred by hand for even melting|Milk allergen
Queso fresco|Fresh cheese|Mexico; Central America|Mild, milky, lightly salty|Crumbly, moist|Tacos, beans, salad, garnish|Protein; calcium; sodium|Does not melt like mozzarella|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Vinegar, coconut|Fermented vinegar|Southeast Asia|Mild, fruity, tart|Thin liquid|Pickle, dipping sauce, dressing|Negligible micronutrients|Acidity varies by producer|Vegan
Vinegar, cane|Fermented condiment|Philippines; Southeast Asia|Mild, slightly sweet, clean acid|Thin liquid|Adobo-style braise, pickle, sauce|Negligible micronutrients|Use for gentle acidity|Vegan
Vinegar, palm|Fermented condiment|Southeast Asia|Fruity, tart, lightly smoky|Thin liquid|Sour curry, pickle, dressing|Negligible micronutrients|Local varieties differ|Vegan
Vinegar, ume plum|Seasoned fruit vinegar|Japan|Salty, sour, fruity|Thin pink liquid|Dressing, pickle, rice, drink|Sodium varies|Often made from ume brine rather than vinegar fermentation|Vegan
Fermented tofu / doufu ru|Fermented soybean curd|China; East Asia|Salty, creamy, pungent umami|Soft cubes in brine|Congee, stir-fry, sauce, dipping|Protein; sodium|Use small amounts; styles vary greatly|Soy allergen
Stinky tofu, fermented|Fermented soybean curd|Taiwan; China|Pungent aroma, savory fermented taste|Soft center, crisp fried exterior|Street food, stew, hot pot|Protein; sodium; fat depends on cooking|Fermentation and frying styles vary|Soy allergen
Natto|Fermented soybean|Japan|Strong fermented aroma, savory|Sticky, stringy beans|Rice, sushi, toast, soup|Protein; vitamin K2; fiber|Stirring increases stringiness; product flavor varies|Soy allergen
Tempeh starter culture|Fermentation culture|Indonesia; global fermentation practice|Mild mushroomy fermentation|Culture for soybean cake|Tempeh production|No direct nutrition as culture ingredient|Use food-grade Rhizopus starter and controlled conditions|Soy present in finished product
Kombucha starter / SCOBY|Fermentation culture|East Asian tea fermentation; global craft practice|Tart, yeasty, lightly fruity|Gelatinous culture|Fermented tea beverage|Nutrition depends on final drink|Sanitary brewing and acidity control are essential|May contain trace alcohol
Shrimp paste / belacan|Fermented seafood paste|Malaysia; Indonesia; Southeast Asia|Intense briny umami|Dense block or paste|Sambal, curry, stir-fry|Protein; sodium|Toast or fry briefly to mellow aroma|Crustacean shellfish allergen
Bagoong alamang|Fermented shrimp condiment|Philippines|Salty, strong, savory|Wet paste|Stew, vegetable dip, sauce|Protein; sodium|Fermentation and salt levels vary|Crustacean shellfish allergen
Bagoong isda|Fermented fish condiment|Philippines|Salty, pungent, savory|Fish paste or sauce|Vegetable dishes, rice, dipping sauce|Protein; sodium|Use as seasoning in small amounts|Fish allergen
Tobiko|Flying fish roe|Japan|Salty, briny, lightly smoky|Small crunchy eggs|Sushi, garnish, rice|Protein; sodium|Colored products may contain additives|Fish allergen
Ikura|Salmon roe|Japan; North Pacific|Salty, rich, oceanic|Large bursting eggs|Sushi, rice, garnish|Protein; omega-3; sodium|Keep refrigerated and use fresh|Fish allergen
Nduja|Spreadable cured pork sausage|Calabria, Italy|Hot, smoky, fatty|Soft, spreadable|Pasta, toast, sauce, pizza|Fat; protein; sodium|Render into sauce; chili level varies|Animal product
Marmite / yeast extract|Yeast spread|British; global|Salty, bitter, intense umami|Thick sticky paste|Toast, broth, gravy, marinade|B vitamins may be fortified; sodium|Use sparingly as a savory concentrate|May contain gluten; check label
Vegemite|Yeast spread|Australia|Salty, bitter, savory|Thick paste|Toast, gravy, broth, savory baking|B vitamins may be fortified; sodium|Very concentrated; use a thin spread|May contain gluten
Black garlic|Aged garlic|East Asia; global specialty pantry|Sweet, balsamic, molasses-like|Soft, sticky cloves|Sauce, dressing, puree, finishing|Small amounts of minerals|Aged under controlled heat/humidity; not fermented in the usual sense|Vegan
Fermented chili bean paste, broad bean|Fermented pulse paste|China; East Asia|Salty, earthy, chili umami|Coarse paste|Noodle sauce, stir-fry, stew|Protein; sodium|Use sparingly and fry in oil for aroma|May contain soy/wheat; label varies
Palm sugar, liquid|Palm sap sweetener|Thailand; Southeast Asia|Caramel, smoky, fruity|Soft paste or syrup|Curry, dessert, sauce, drink|Mostly carbohydrate|Texture and sweetness vary by palm species and producer|Vegan
Sorghum syrup|Grain syrup|Southern United States; Africa|Earthy, mild caramel|Thick syrup|Baking, glaze, breakfast, sauce|Mostly sugars; minerals vary|Not the same as molasses or maple syrup|Vegan
Golden syrup|Sugar syrup|British; global baking|Buttery caramel, mild|Thick invert syrup|Baking, pudding, glaze|Mostly sugars|Corn syrup or treacle are different substitutes|Vegan
Treacle, black|Cane sugar syrup|British; Caribbean|Dark, bitter-sweet, robust|Thick syrup|Gingerbread, pudding, sauce|Mostly sugars; minerals vary|Flavor is stronger than golden syrup|Vegan
Arrowroot starch|Tuber starch|Caribbean; South American; East Asian|Neutral|Clear, glossy thickener|Fruit glaze, sauce, baking|Mostly carbohydrate|Add slurry near end; prolonged boiling can thin|Usually gluten-free
Kudzu starch|Root starch|Japan; East Asia|Neutral, delicate|Smooth translucent gel|Sauce, dessert, drink|Mostly carbohydrate|Dissolve in cold liquid and heat gently|Usually gluten-free
Konjac flour / glucomannan|Tuber fiber ingredient|East Asia|Neutral|Forms elastic gel or noodle|Shirataki noodles, jelly, thickener|Soluble fiber|Hydrate and prepare according to food product instructions|Check choking warnings for jelly products
`));

// Additional regional and less commonly catalogued foods for this expansion pass.
data['Vegetables'].push(...parse(`
Oca|Andean tuber|Andes, especially Peru and Bolivia|Tart, lemony to sweet|Waxy, colorful tuber|Roast, stew, pickle, side dish|Carbohydrate; vitamin C; fiber|Cook to soften acidity; cultivar varies|Vegan
Ulluco|Andean tuber|Andes, especially Peru and Bolivia|Mild, earthy|Moist, waxy flesh|Soup, stew, saute, pickle|Carbohydrate; vitamin C|Small tubers may be cooked whole|Vegan
Mashua|Andean tuber|Andes|Peppery, pungent, sweetens on cooking|Dense tuber|Roast, stew, preserve|Carbohydrate; vitamin C|Cook thoroughly; flavors mellow with heat|Vegan
Yacon|Root tuber|Andes; cultivated elsewhere|Juicy, crisp, mildly sweet|Watery, pear-like crunch|Fresh salad, pickle, syrup|Water; fiber; fructooligosaccharides|Store cool; sweetness changes after harvest|Vegan
Burdock root / gobo|Root|Japan, Korea, China|Earthy, sweet, mildly bitter|Fibrous, crisp when thin-sliced|Kinpira, soup, pickle, braise|Fiber; potassium|Soak cut pieces briefly to limit browning|Vegan
Water chestnut, fresh|Corm|China; Southeast Asia|Sweet, clean, nutty|Very crisp even after cooking|Stir-fry, dumplings, salad, dessert|Carbohydrate; fiber; potassium|Peel brown skin; canned has softer texture|Vegan
Arrowhead tuber / cigu|Corm|China; Japan|Mild, chestnut-like, earthy|Starchy, crumbly|Stew, simmered dishes, roast|Carbohydrate; fiber|Cook fully; texture varies by cultivar|Vegan
Taro leaves|Leaf|Pacific Islands; South Asia; Caribbean|Earthy, mild|Softens and thickens when cooked|Stew, wrapped foods, greens dishes|Fiber; micronutrients vary|Never eat raw; cook thoroughly to remove irritants|Vegan
Chaya leaves|Leaf|Yucatan, Mexico; Central America|Spinach-like, green|Tender cooked leaf|Soup, saute, stew, tamales|Protein; calcium; iron|Must be thoroughly cooked; raw leaves are unsafe|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Bael fruit|Tree fruit|South Asia|Sweet, aromatic, resinous|Hard shell; fragrant fibrous pulp|Drink, preserve, dessert|Fiber; carbohydrate|Break hard shell carefully; ripe and unripe uses differ|Vegan
Wood apple / elephant apple|Tree fruit|South Asia; Sri Lanka|Tart, fermented, tangy|Soft pulp in hard shell|Chutney, drink, preserve|Fiber; carbohydrate|Ripe fruit aroma is strong; discard shell|Vegan
Phalsa|Berry|South Asia|Sweet-tart, refreshing|Small soft berry|Sherbet, drink, preserve|Vitamin C; anthocyanins|Highly perishable; often processed quickly|Vegan
Karonda|Berry|South Asia|Tart, astringent, resinous|Firm, small fruit|Pickle, preserve, chutney|Vitamin C; fiber|Usually cooked or preserved before fully ripe|Vegan
Marang|Multiple fruit|Borneo; Philippines|Sweet, creamy, fragrant|Soft segments around seeds|Fresh eating, dessert|Carbohydrate; fiber|Perishable; handle gently|Vegan
Soursop|Tropical fruit|Caribbean; Central and South America|Sweet-tart, citrusy, creamy|Soft fibrous pulp|Drink, dessert, puree|Vitamin C; fiber|Remove seeds; use ripe pulp|Vegan
Grumichama|Berry|Brazil; South America|Sweet, cherry-like, lightly tart|Juicy, thin-skinned|Fresh eating, jelly, dessert|Vitamin C; anthocyanins|Ripe fruit bruises easily|Vegan
Lucuma|Tropical fruit|Andes, especially Peru|Maple-like, sweet, mildly dry|Dense, yellow, creamy-mealy|Ice cream, pastry, drink, puree|Carbohydrate; fiber; carotenoids|Often sold as frozen pulp or powder|Vegan
Muscadine grape|Grape|Southeastern United States|Musky, sweet-tart|Thick skin, large seeds|Juice, wine, jelly, dessert|Polyphenols; carbohydrate|Skin and seed texture differs from bunch grapes|Vegan
Saskatoon berry|Berry|North American prairie and boreal regions|Sweet, almond-like, mild|Small berry with soft seeds|Pie, preserve, fresh eating|Fiber; manganese; anthocyanins|Fresh and dried forms differ in concentration|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Job's tears / coix|Grain|East and Southeast Asia|Mild, nutty|Chewy, pearl-like grain|Soup, porridge, tea, grain bowl|Protein; fiber; minerals|Soak and cook until tender|Vegan
Khorasan wheat / Kamut-type|Ancient wheat|West Asia; grown elsewhere|Buttery, nutty|Large chewy kernels|Pilaf, salad, bread, porridge|Protein; fiber; selenium|Contains gluten; cook whole kernels until tender|Wheat/gluten allergen
Emmer / farro|Ancient wheat|Fertile Crescent; Mediterranean|Earthy, nutty|Chewy whole kernels|Soup, salad, pilaf, bread|Fiber; protein; magnesium|Farro labels can refer to different wheat species|Wheat/gluten allergen
Black chickpea / kala chana|Pulse|South Asia|Earthy, nutty, robust|Firm skin, creamy center|Chana masala, sprout, salad|Protein; fiber; iron|Soak dried beans; cook thoroughly|Legume sensitivity possible
Moth bean / matki|Pulse|South Asia|Mild, earthy|Small firm bean; sprouts crisp|Usal, curry, sprout salad|Protein; fiber; folate|Soaking and sprouting change cook time|Legume sensitivity possible
Grass pea / khesari|Pulse|South Asia; Mediterranean|Earthy, nutty|Small, firm legume|Flour, stew, pulse dishes|Protein; fiber|Use as part of varied diet; avoid chronic heavy reliance on improperly processed grass pea|Legume sensitivity possible
Adzuki bean, black|Pulse cultivar|Japan; East Asia|Sweet, earthy|Small firm bean|Dessert, rice, paste, soup|Protein; fiber; iron|Cook until tender; naming varies by market|Legume sensitivity possible
Mung bean starch noodles|Pulse-derived starch|China; East and Southeast Asia|Neutral, absorbs sauce|Slippery, springy strands|Hot pot, salad, stir-fry, soup|Mostly carbohydrate|Soak or cook per package directions|Vegan
`));
data['Spices & Herbs'].push(...parse(`
Uziza leaf / West African pepper leaf|Leaf spice|Nigeria; Gulf of Guinea|Peppery, aromatic, warm|Tender leaf|Soup, stew, fish dishes|Small culinary quantities|Add late for fresh aroma|Vegan
Alligator pepper|Spice seed|West Africa|Pungent, ginger-like, smoky|Hard aromatic seeds|Soup, spice mix, ceremonial foods|Used in small quantities|Flavor strength varies; grind as needed|Vegan
Cubeb pepper|Spice berry|Indonesia; South Asia|Camphor, pepper, eucalyptus|Dried tailed berries|Pickle, spice blend, meat, drink|Used in small quantities|Use sparingly; distinctive aroma|Vegan
Mahleb|Stone-fruit kernel spice|Eastern Mediterranean; West Asia|Almond, cherry, floral|Fine powder|Bread, pastry, cheese dishes|Used in small quantities|Grind shortly before use; bitter if overdosed|Tree nut/stone-fruit cross-contact varies
Dried lime, black loomi|Dried citrus|Iran; Gulf region|Tart, fermented, smoky citrus|Hard dried lime|Stew, rice, broth, tea|Small amounts of vitamin C remain|Pierce or crush; simmer whole|Vegan
Amchur|Dried green mango powder|South Asia|Tart, fruity, dry acidity|Fine powder|Chaat, curry, legumes, marinades|Small amounts of micronutrients|Add toward end to preserve brightness|Vegan
Kokum rind|Dried fruit rind|Western India; Konkan|Tart, fruity, gently smoky|Dried leathery rind|Curry, cooling drink, sour broth|Anthocyanins; organic acids|Soak or simmer; intensity varies|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Mutton, goat|Red meat|South Asia; Middle East; Africa|Savory, earthy, rich|Varies by age and cut|Curry, stew, roast, grill|Protein; iron; B12|Slow-cook tougher cuts; trim to preference|Animal product
Guinea fowl|Poultry|West Africa; Europe|Gamey, rich, poultry-like|Lean, firm meat|Roast, stew, braise|Protein; iron; B vitamins|Avoid overcooking lean breast|Animal product
Pigeon / squab|Poultry|Europe; North Africa; East Asia|Rich, gamey, deep savory|Small, tender, dark meat|Roast, braise, grill|Protein; iron; B12|Use food-grade farmed product; cook safely|Animal product
Quail eggs|Egg|East Asia; Europe; global|Mild, rich, egg-like|Small, delicate shell|Pickle, garnish, roast, snack|Protein; choline; B12|Cook thoroughly; handle like other eggs|Egg allergen
Donkey milk cheese|Cheese|Mediterranean; Balkans|Mild, tangy, distinctive|Soft to semi-firm|Cheese board, cooking|Protein; calcium; fat varies|Rare specialty; verify source and milk declaration|Milk allergen
Clotted cream|Cultured cream|Southwest England; South Asia variants|Rich, sweet, buttery|Thick, spoonable|Scone, dessert, curry enrichment|Fat; saturated fat|Keep refrigerated; use in modest portions|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Kashk|Dried fermented dairy|Iran; Central Asia|Tangy, salty, savory|Paste, powder or dried pieces|Soup, dip, stew, grain dishes|Protein; calcium; sodium|Dilute and cook; product forms vary|Milk allergen
Qurut / kurut|Dried cultured dairy|Central Asia; Afghanistan|Sour, salty, concentrated|Hard dried balls or powder|Soup, sauce, snack|Protein; calcium; sodium|Rehydrate; salt varies by maker|Milk allergen
Shio koji|Fermented rice seasoning|Japan|Sweet, salty, fermented umami|Wet grain paste|Marinade, pickle, sauce|Carbohydrate; sodium|Refrigerate after opening; fermentation continues|May contain rice/barley; check label
Amazake, rice|Fermented rice drink/base|Japan|Sweet, mellow, rice aroma|Thick drink or porridge|Drink, dessert, marinade|Carbohydrate; small protein|Sweetness varies by fermentation and added alcohol|Check product label
Saba banana vinegar|Fruit vinegar|Philippines; Southeast Asia|Fruity, sweet-sour|Thin liquid|Pickle, dipping sauce, adobo-style cooking|Negligible micronutrients|Acidity varies by producer|Vegan
Tamarind concentrate, seedless|Fruit concentrate|South Asia; Southeast Asia; Africa|Sour, fruity, caramel-like|Thick paste|Sauce, soup, curry, drink|Carbohydrate; organic acids|Dilute to taste; products vary in sugar and salt|Vegan
`));

data['Vegetables'].push(...parse(`
Malabar spinach / basella|Leafy vine|South Asia; Southeast Asia; Africa|Mild, earthy, slightly tart|Succulent leaves; mucilaginous when cooked|Soup, curry, stir-fry, stew|Vitamin A; vitamin C; iron|Add near the end; texture thickens liquid|Vegan
Pointed gourd / parwal|Gourd|South Asia|Mild, lightly sweet|Firm, holds shape|Curry, stuffed vegetable, braise|Fiber; vitamin C|Scrape or peel skin as desired|Vegan
Tindora / scarlet gourd|Gourd|South Asia|Mild, grassy, lightly sweet|Crisp when young|Stir-fry, pickle, curry|Fiber; vitamin C|Avoid overripe red fruit if a firm vegetable texture is wanted|Vegan
Drumstick leaves|Leafy greens|South Asia; East Africa|Green, peppery, gently bitter|Small tender leaflets|Dal, soup, stir-fry, flatbread|Vitamin A; vitamin C; calcium; iron|Strip leaflets from stems; cook briefly|Vegan
Jute mallow, Egyptian molokhia|Leafy greens|Egypt; Levant; North Africa|Earthy, savory, gently grassy|Silky, mucilaginous when cooked|Soup, stew, rice accompaniment|Fiber; vitamin A; iron|Chop leaves and cook thoroughly|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Langsat, duku type|Tropical fruit|Malaysia; Indonesia; Southeast Asia|Sweet, tart, slightly resinous|Juicy segments|Fresh eating, preserve|Vitamin C; carbohydrate|Peel and remove seeds; latex may be bitter|Vegan
Santol, red type|Tree fruit|Southeast Asia|Sweet-sour, aromatic|Soft cottony pulp around seeds|Curry, preserve, relish|Fiber; vitamin C|Do not eat seeds; cultivar changes sourness|Vegan
Pulasan|Tropical fruit|Malaysia; Indonesia|Sweet, floral, mild|Juicy white flesh|Fresh eating, fruit salad|Vitamin C; carbohydrate|Remove seed; handle gently|Vegan
Burmese grape / baccaurea|Tropical fruit|Southeast Asia|Tart, sweet, lightly astringent|Small segmented fruit|Fresh eating, preserve, pickle|Vitamin C; fiber|Ripeness strongly changes acidity|Vegan
Tamarillo, purple cultivar|Fruit|Andes; New Zealand|Tart, aromatic, tomato-like|Juicy, seedy flesh|Chutney, sauce, drink, salad|Vitamin C; provitamin A|Peel may taste bitter; use ripe fruit|Vegan
Fingered citron / Buddha's hand|Citrus|East Asia; Mediterranean cultivation|Intensely fragrant, floral, mildly bitter|Mostly aromatic rind; little pulp|Candied peel, zest, infusion, garnish|Small amounts of vitamin C|Wash peel well; generally used for zest|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Black-eyed pea, cream type|Cowpea cultivar|West Africa; Americas; South Asia|Mild, earthy|Tender, creamy bean|Rice dishes, stew, salad, fritter|Protein; fiber; folate|Soak optional; cook until tender|Legume sensitivity possible
Cowpea, yardlong bean|Pod legume|South Asia; Southeast Asia; Africa|Green, bean-like, mildly sweet|Long crisp pods|Stir-fry, curry, stew|Fiber; folate; vitamin C|Trim ends and cook thoroughly|Vegan; legume sensitivity possible
Winged bean seed|Pulse|Southeast Asia; Pacific|Nutty, rich|Firm seed when cooked|Stew, flour, snack|Protein; fat; minerals|Use food-grade mature seed and cook thoroughly|Legume sensitivity possible
Horse gram / kulthi|Pulse|South Asia|Earthy, robust, slightly smoky|Small firm seed; creamy when cooked|Rasam, soup, curry, sprout|Protein; fiber; iron|Soaking and pressure cooking improve tenderness|Legume sensitivity possible
Rice bean / moth bean relative|Pulse|Himalayan foothills; South Asia|Mild, nutty|Small, firm bean|Stew, curry, sprout dishes|Protein; fiber; minerals|Cook thoroughly; local names vary|Legume sensitivity possible
Yam bean / jicama seed pod|Pulse seed|Mexico; Central America|Bean-like, mild|Hard mature seed|Not a routine culinary ingredient|Protein; plant compounds vary|Do not eat seed casually; seeds contain rotenone and require expert, regulated preparation|Potentially toxic; avoid unsupervised use
Red lentil, whole brown-skinned|Pulse cultivar|South Asia; Middle East|Earthy, mildly sweet|Holds shape better than split red lentil|Soup, stew, salad, dal|Protein; fiber; folate|Rinse and simmer; cooking time varies by age|Legume sensitivity possible
Puy lentil / French green lentil|Pulse cultivar|Auvergne, France|Peppery, earthy, mineral|Firm, holds shape|Salad, braise, side dish|Protein; fiber; iron|Do not overcook if shape is desired|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Ajwain leaf / Indian borage|Aromatic herb|South Asia; Caribbean|Oregano-like, thyme, camphor|Fleshy aromatic leaves|Fritter, chutney, tea, stew|Small culinary quantities|Use fresh leaves sparingly; flavor is potent|Vegan
Pandanus amaryllifolius, wild form|Aromatic leaf|Southeast Asia|Grassy, vanilla-like, toasted rice|Long fragrant leaves|Rice, sweets, custard, infusion|Negligible per serving|Use correctly identified culinary pandan species|Vegan
Kaffir lime leaf / makrut lime leaf|Citrus leaf|Thailand; Laos; Cambodia; Indonesia|Intense citrus, floral, green|Tough glossy double leaves|Curry, soup, paste, marinade|Small culinary quantities|Bruise, tear or finely slice; remove tough ribs|Vegan
Vietnamese coriander / rau ram|Aromatic herb|Vietnam; Southeast Asia|Peppery, citrus, coriander-like|Tender narrow leaf|Salad, noodle soup, egg dishes|Small culinary quantities|Add fresh near service|Vegan
Sawtooth coriander / culantro|Aromatic herb|Caribbean; Central America; Southeast Asia|Strong cilantro, pungent|Long serrated leaves|Soup, sofrito, curry, seasoning paste|Small culinary quantities|More robust than cilantro; use less initially|Vegan
Lemon myrtle|Aromatic leaf|Australia|Strong lemon, floral|Dried leaf or powder|Tea, dessert, seafood, seasoning|Small culinary quantities|Use sparingly; can become bitter with long cooking|Vegan
Mountain pepper leaf|Aromatic leaf|Australia|Peppery, herbal, berry-like|Tender leaf|Seasoning, sauce, seafood|Small culinary quantities|Use food-grade cultivated product|Vegan
Myrtle, bog myrtle leaf|Aromatic leaf|Northern Europe|Balsamic, herbal, resinous|Dried leaf|Infusion, beer, seasoning|Small culinary quantities|Use culinary-grade product; species identification matters|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Carp, grass carp|Freshwater fish|East Asia; Central Europe|Mild, earthy, sweet|Firm, bony flesh|Stew, braise, grill, smoked|Protein; B12; omega-3 varies|Remove pin bones; cook fully|Fish allergen
Catfish, African|Freshwater fish|West Africa; Nile basin|Mild, earthy, savory|Firm, moist flesh|Grill, stew, fry, smoke|Protein; B12; selenium|Source from inspected aquaculture; cook safely|Fish allergen
Tilapia, Nile|Freshwater fish|Nile basin; aquaculture worldwide|Mild, clean, slightly sweet|Lean, flaky|Grill, curry, bake, fry|Protein; selenium; B12|Cook to safe internal temperature|Fish allergen
Milkfish / bangus|Freshwater and coastal fish|Philippines; Indonesia|Rich, savory, mildly sweet|Tender, many fine bones|Grill, pickle, stew, fry|Protein; omega-3; B12|Debone carefully or use prepared product|Fish allergen
Snakehead fish|Freshwater fish|South and Southeast Asia|Mild, savory|Firm, white flesh|Soup, curry, braise, grill|Protein; minerals|Use legally sourced product; cook thoroughly|Fish allergen
Buffalo curd / buffalo yogurt|Cultured dairy|South Asia; Balkans|Tangy, rich, creamy|Thick, high-fat cultured dairy|Dessert, sauce, marinade, side|Protein; calcium; fat|Keep chilled; culture and fat content vary|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Mackerel paste, fermented|Fermented fish condiment|Scandinavia; Southeast Asia|Intense, salty, fishy umami|Thick paste|Broth, sauce, seasoning|Protein; sodium|Small doses; aroma and sodium vary widely|Fish allergen
Fermented tea leaf / laphet|Fermented tea leaf|Myanmar|Bitter, tangy, earthy, savory|Chewy leaf paste or whole leaf|Tea leaf salad, rice accompaniment|Polyphenols; caffeine varies|Fermentation and added oil/salt vary by product|Vegan; check additions
Fermented bamboo shoot, sour|Fermented vegetable|South Asia; Southeast Asia|Sour, grassy, fermented|Tender strips or paste|Curry, soup, stew, condiment|Fiber; sodium varies|Rinse if very salty; cook as customary|Vegan
Fermented mustard greens, suan cai|Fermented leafy vegetable|China; East Asia|Sour, salty, savory|Crunchy to tender|Noodle soup, fish, stir-fry|Fiber; sodium|Rinse or soak to adjust salt|Vegan
Fermented mustard greens, kiam chye|Fermented leafy vegetable|Southeast Asia|Sour, salty, mustardy|Tender, juicy|Soup, braise, fish, stir-fry|Fiber; sodium|Product salt and acidity vary|Vegan
Fermented radish, takuan|Pickled vegetable|Japan|Sweet-sour, salty, turmeric-like|Crunchy|Rice, bento, sandwich, pickle plate|Fiber; sodium|Sweetness and coloring vary by product|Vegan; check additives
Fermented cabbage, curtido|Fermented vegetable relish|El Salvador; Central America|Sour, gently spicy, aromatic|Crisp shredded vegetables|Pupusa, taco, sandwich|Fiber; sodium|Fermentation time changes acidity|Vegan
Fermented cassava, puba|Fermented starch|Brazil|Mildly sour, earthy|Soft fermented mash|Porridge, cake, regional dishes|Carbohydrate; composition varies|Use properly processed food-grade product|Vegan
`));

data['Vegetables'].push(...parse(`
Salsify / oyster plant|Root|Europe; Mediterranean|Mild, sweet, oyster-like|Tender, creamy when cooked|Roast, soup, gratin, puree|Fiber; potassium; vitamin C|Peel and acidulate cut roots to limit browning|Vegan
Skirret|Root|Europe; East Asia|Sweet, parsnip-like|Clustered, tender roots|Roast, puree, soup|Carbohydrate; fiber|Trim fine roots; cook until tender|Vegan
Sea kale shoots|Brassica shoot|Atlantic Europe|Mild, sweet, asparagus-like|Tender blanched shoots|Steam, butter, salad, side|Vitamin C; fiber|Blanch pale shoots before finishing|Vegan
Cardoon|Stem vegetable|Mediterranean; Italy; North Africa|Bitter, artichoke-like|Fibrous ribs soften with cooking|Braise, gratin, soup|Fiber; potassium|Trim strings and blanch to reduce bitterness|Vegan
Rampion / rapunzel root|Root|Europe; West Asia|Mild, nutty, slightly sweet|Small crisp roots|Salad, pickle, soup|Fiber; micronutrients vary|Use cultivated food-grade roots|Vegan
Good King Henry|Leafy perennial|Europe|Mineral, spinach-like|Tender shoots and leaves|Greens, soup, side dish|Iron; vitamin C; fiber|Use only correctly identified cultivated plants|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Imbe / African mangosteen|Tropical fruit|East Africa|Sweet-tart, resinous, tropical|Juicy orange segments|Fresh eating, juice, preserve|Vitamin C; carbohydrate|Remove seed; ripe fruit is delicate|Vegan
Marula fruit|Tropical fruit|Southern Africa|Tart, aromatic, citrus-like|Juicy pulp around stone|Drink, preserve, sauce|Vitamin C; carbohydrate|Use ripe food-grade fruit; fermentation begins readily|Vegan
Baobab pulp|Tree fruit pulp|Sahel; East Africa; Southern Africa|Tart, citrusy, creamy powder|Dry powder or fresh pulp|Drink, porridge, sauce, baking|Vitamin C; fiber; minerals|Powder concentration varies by processing|Vegan
Monkey orange|Tree fruit|Southern Africa|Sweet-sour, aromatic|Firm shell; juicy pulp|Fresh eating, drink, preserve|Vitamin C; carbohydrate|Open shell carefully; species and maturity vary|Vegan
Kei apple|Tree fruit|Southern Africa|Tart, floral, apricot-like|Small juicy fruit|Preserve, sauce, drink|Vitamin C; organic acids|Often cooked or sweetened due to acidity|Vegan
Naranjilla / lulo|Andean fruit|Ecuador; Colombia; Andes|Tart, citrusy, pineapple-like|Juicy pulp with seeds|Drink, sauce, dessert|Vitamin C; fiber|Peel and strain if desired; do not confuse with unrelated nightshades|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Black fonio|Small grain|West Africa; Sahel|Nutty, delicate, earthy|Tiny fluffy grain|Pilaf, porridge, couscous-style dishes|Protein; fiber; minerals|Rinse well; cook with measured liquid|Vegan; naturally gluten-free
African rice / Oryza glaberrima|Grain|West Africa|Nutty, earthy, robust|Chewy grain, often red-brown|Rice dishes, porridge, traditional meals|Carbohydrate; fiber; minerals|Landraces and milling change cooking time|Vegan; naturally gluten-free
Job's tears, hulled Chinese type|Grain|China; East Asia|Mild, nutty|Chewy pearls|Soup, porridge, tea|Protein; fiber; minerals|Soak if needed; cook until tender|Vegan; naturally gluten-free
Chickling vetch|Pulse|Mediterranean; West Asia|Earthy, nutty|Firm pulse|Traditional flour and pulse dishes|Protein; fiber|Use food-grade varieties and traditional processing; avoid chronic reliance|Legume sensitivity possible
Kersting's groundnut|Pulse|West Africa|Nutty, earthy|Small firm seeds|Stew, flour, snack|Protein; fiber; minerals|Cook thoroughly; local names vary|Legume sensitivity possible
`));
data['Protein & Dairy'].push(...parse(`
Kefir grains culture|Fermentation culture|Caucasus; global use|Tangy, yeasty, lightly effervescent|Culture for fermented milk|Kefir, cultured sauces, baking|Nutrition depends on final milk product|Use clean equipment and food-grade culture|Milk allergen in finished dairy kefir
Filmjölk|Cultured milk|Sweden; Nordic countries|Mildly sour, buttery|Pourable, lightly viscous|Breakfast, sauce, baking|Protein; calcium; cultures vary|Keep refrigerated; cultures are product-specific|Milk allergen
Skyr|Cultured dairy|Iceland|Tangy, clean, high-protein|Thick, strained|Breakfast, dessert, sauce|Protein; calcium; fat varies|Straining and fat content vary by product|Milk allergen
Quesillo, Nicaraguan style|Fresh cheese|Nicaragua|Mild, milky, salty|Soft, stringy|Tortilla, snack, street food|Protein; calcium; sodium|Keep chilled; regional versions differ|Milk allergen
Chhurpi, soft|Cultured cheese|Himalayan regions|Tangy, smoky, savory|Soft or semi-firm|Soup, stew, snack|Protein; calcium; fat varies|Traditional milk source and aging vary|Milk allergen
Kumis / fermented mare milk|Fermented dairy drink|Central Asia; Mongolia|Tart, yeasty, lightly fizzy|Thin drink|Beverage, sauces, regional foods|Protein; lactose and alcohol vary|Fermentation varies; use reputable food-grade products|Milk allergen; may contain alcohol
`));
data['Pantry & Condiments'].push(...parse(`
Fermented black beans / douchi|Fermented soybean condiment|China|Salty, earthy, concentrated umami|Dry or moist beans|Steamed fish, stir-fry, sauce|Protein; sodium|Rinse or soak; use sparingly|Soy allergen
Fermented rice bran pickle bed / nukadoko|Fermentation medium|Japan|Sour, nutty, savory|Bran paste for pickling|Pickled vegetables|Nutrition depends on pickled food|Maintain food-safe salt, acidity and refrigeration practices|Vegan; check additions
Fermented locust bean condiment / iru|Fermented seed condiment|West Africa|Pungent, savory, umami|Sticky fermented seeds|Soup, stew, seasoning|Protein; sodium varies|Use small amounts; aroma and salt vary by maker|Legume sensitivity possible
Ogiri, fermented melon seed|Fermented seed condiment|Nigeria; West Africa|Strong, nutty, fermented|Soft paste or crumbled cake|Soup, stew, vegetable dishes|Protein; fat; sodium|Fermentation and seed type vary by region|Seed allergen possible
Shito, Ghanaian pepper sauce|Chili condiment|Ghana|Hot, smoky, fishy, savory|Thick oil-based sauce|Rice, fish, plantain, stew|Fat; sodium varies|Contains seafood in many recipes; check label|May contain fish or crustacean shellfish
Green seasoning, Caribbean|Fresh herb paste|Trinidad; Guyana; Caribbean|Herbal, citrusy, garlicky|Wet chopped paste|Marinade, stew, rice, roast|Small amounts of micronutrients|Fresh formulations vary; refrigerate promptly|Check ingredients for allergens
`));

data['Vegetables'].push(...parse(`
Chayote shoot tips|Leafy shoots|Mexico; Central America|Mild, green, slightly sweet|Tender curling shoots|Soup, saute, stew|Fiber; micronutrients vary|Use young food-grade shoots; cook briefly|Vegan
Jicama leaves, cultivated food use|Leaf|Mexico; Central America|Green, mild|Tender when cooked|Traditional cooked greens in some local uses|Micronutrients vary|Use only verified food-grade leaves; seeds are toxic and excluded|Vegan; leaf safety requires correct plant part
Sea beet leaves|Leafy coastal vegetable|Europe; Mediterranean|Mineral, salty, spinach-like|Tender leaf, firm stems|Saute, soup, side dish|Fiber; minerals|Use cultivated or correctly identified food-grade plants|Vegan
Amaranth stem shoots|Stem vegetable|West Africa; South Asia; Caribbean|Earthy, green|Tender to fibrous|Stew, saute, soup|Fiber; micronutrients vary|Peel fibrous mature stems and cook thoroughly|Vegan
Water caltrop / singhara|Aquatic nut|South Asia; East Asia|Mild, sweet, chestnut-like|Crisp raw; floury cooked|Snack, flour, curry, dessert|Carbohydrate; potassium|Cook or peel carefully; regional forms vary|Vegan
Lotus seed, fresh|Aquatic seed|South Asia; China|Mild, sweet, nutty|Crisp-tender, starchy|Soup, curry, dessert, snack|Protein; carbohydrate; minerals|Remove bitter green embryo if present|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Jujube, fresh|Tree fruit|China; Central Asia|Sweet, crisp, apple-like|Crunchy skin and flesh|Fresh eating, tea, dessert|Vitamin C; carbohydrate|Fresh and dried jujubes differ greatly|Vegan
Jujube, dried red date|Dried fruit|China; East Asia|Caramel-like, sweet, mildly fruity|Chewy, dense|Tea, soup, dessert, stuffing|Carbohydrate; fiber|Pit before chopping; sugar is concentrated|Vegan
Chinese bayberry / yangmei|Tree fruit|China; East Asia|Sweet-tart, berry-like, floral|Juicy, textured surface|Fresh eating, drink, preserve|Vitamin C; polyphenols|Perishable; use ripe fruit|Vegan
Sea buckthorn, orange cultivar|Berry|Central Asia; Northern Europe|Sharp tart, citrus, resinous|Small oily berry|Juice, sauce, syrup, puree|Vitamin C; carotenoids; oils|Often processed; intense acidity|Vegan
Salmonberry|Berry|Pacific Northwest, North America|Mild sweet, floral|Soft, juicy|Fresh eating, preserve, dessert|Vitamin C; fiber|Delicate; harvest and use promptly|Vegan
Cloudberry|Berry|Arctic and boreal regions|Tart, honeyed, aromatic|Soft, seedy|Jam, sauce, dessert, fresh eating|Vitamin C; fiber|Seasonal and delicate; frozen fruit is common|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Buckwheat groats, roasted kasha|Pseudocereal|Eastern Europe; Russia; Japan|Toasty, earthy, nutty|Separate, chewy grains|Kasha, porridge, pilaf, stuffing|Fiber; magnesium; rutin|Rinse; roast level affects cooking time|Vegan; naturally gluten-free
Buckwheat groats, green|Pseudocereal|Eastern Europe; Asia|Mild, grassy, nutty|Soft, creamy if overcooked|Porridge, salad, flour, sprouting|Fiber; magnesium; protein|Rinse and simmer gently|Vegan; naturally gluten-free
Cornmeal, blue maize|Milled grain|Southwestern United States; Mexico|Sweet, earthy, corn-forward|Fine to coarse meal|Masa, porridge, bread, tortillas|Carbohydrate; fiber; anthocyanins|Stone-ground texture varies; not all blue corn is nixtamalized|Vegan; naturally gluten-free
Masa harina, yellow corn|Nixtamalized corn flour|Mexico; Central America|Sweet corn, alkaline, warm|Fine flour; cohesive dough|Tortilla, tamale, pupusa|Carbohydrate; calcium varies|Hydration differs by brand and grind|Vegan; naturally gluten-free
Chestnut flour|Nut flour|Corsica; Italy; France|Sweet, earthy, roasted|Fine, slightly grainy flour|Pasta, bread, cake, crepe|Carbohydrate; fiber; minerals|Naturally sweet; combine with other flours as needed|Tree nut allergen
Tiger nut flour|Tuber flour|West Africa; Spain|Sweet, nutty, coconut-like|Fine, fibrous powder|Baking, horchata, porridge|Fiber; carbohydrate; minerals|Absorbs liquid; product grind varies|Check cross-contact and label
`));
data['Spices & Herbs'].push(...parse(`
Wattle seed, roasted|Native Australian seed spice|Australia|Coffee, chocolate, hazelnut-like|Fine roasted meal|Bread, dessert, sauce, rub|Protein; fiber; minerals|Use food-grade roasted seed; grind as needed|Seed allergen cross-contact possible
Lemon aspen|Native Australian fruit spice|Australia|Tart, citrus, floral|Dried powder or puree|Sauce, dessert, seafood, drink|Vitamin C; organic acids|Use food-grade species and product|Vegan
Davidson plum|Native Australian fruit spice|Australia|Tart, berry, earthy|Juicy dark fruit; puree|Sauce, dessert, preserve|Anthocyanins; organic acids|Strong acidity; often sweetened|Vegan
Mountain pepper berry|Native Australian spice|Australia|Peppery, berry, warm numbing|Dried berries|Rub, sauce, chocolate, game|Used in small quantities|Grind freshly; flavor can be intense|Vegan
Tasmanian pepper leaf|Native Australian herb|Australia|Peppery, fruity, herbal|Dried leaf|Seasoning, sauce, seafood|Small culinary amounts|Use cultivated food-grade product|Vegan
Finger lime leaf|Citrus leaf|Australia|Citrusy, floral, green|Aromatic leaf|Seafood, tea, seasoning|Small culinary amounts|Use food-grade leaves; add late|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Paneer, buffalo milk|Fresh cheese|South Asia|Mild, milky, rich|Firm, holds shape when cooked|Curry, grill, saute, stuffing|Protein; calcium; fat|Keep refrigerated; homemade moisture varies|Milk allergen
Ricotta salata|Salted aged whey cheese|Sicily, Italy|Salty, milky, gently tangy|Firm, crumbly|Grate over pasta, salad, vegetables|Protein; calcium; sodium|Salt and aging vary by producer|Milk allergen
Labneh, dried balls|Strained cultured dairy|Levant|Tangy, salty, herbal if preserved in oil|Dense, chewy|Mezze, bread, salad, garnish|Protein; calcium; sodium|Keep chilled unless commercially shelf-stable|Milk allergen
Kashkaval cheese|Pasta-filata cheese|Balkans; Turkey; Eastern Mediterranean|Salty, tangy, buttery|Semi-hard, melts well|Toast, pastry, bake, grill|Protein; calcium; sodium|Milk blend and aging differ by country|Milk allergen
Gariss fermented camel milk|Fermented dairy|Sudan; East Africa|Sharp, sour, yeasty|Thin cultured drink|Drink, porridge, regional dishes|Protein; calcium; composition varies|Use regulated pasteurized product where available|Milk allergen
Zebu milk curd|Cultured dairy|Madagascar; East Africa|Tangy, rich, mildly grassy|Soft curd|Side dish, sauce, snack|Protein; calcium; fat varies|Milk source and cultures affect acidity|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Gochugaru, coarse flakes|Dried chili product|Korea|Fruity, smoky, mild to hot|Coarse red flakes|Kimchi, stew, seasoning|Capsaicin; sodium varies by blend|Heat and color vary by pepper and grind|Vegan; check blends
Gochugaru, fine powder|Dried chili product|Korea|Fruity, warm, concentrated|Fine red powder|Paste, marinade, stew|Capsaicin; sodium varies by blend|Use less than coarse flakes by volume|Vegan; check blends
Mole negro paste|Regional sauce base|Oaxaca, Mexico|Roasted, bitter-sweet, smoky, spiced|Thick complex paste|Sauce, braise, enchiladas|Fat; sodium; sugar varies|Recipes and ingredient lists vary greatly|May contain sesame, nuts, wheat; check label
Mole coloradito paste|Regional sauce base|Oaxaca, Mexico|Sweet, chili, warm spice, mild cocoa|Thick smooth paste|Chicken, vegetables, enchiladas|Fat; sodium; sugar varies|Regional and household recipes differ|Allergen ingredients vary
Shatta red chili paste|Chili condiment|Levant; Palestine|Fresh chili, garlic, tangy heat|Coarse wet paste|Sandwich, falafel, stew, eggs|Vitamin C; sodium varies|Fresh and fermented types differ|Vegan; check additions
`));

data['Vegetables'].push(...parse(`
Ogonori / gracilaria vegetable|Edible sea vegetable|Japan; Southeast Asia; Caribbean|Mild, briny, slightly sweet|Crisp to tender strands|Salad, soup, agar dishes|Fiber; minerals; iodine varies|Use food-grade species; iodine levels vary by source|Vegan
Chaya, cooked leaf clusters|Leafy green|Yucatan; Belize; Guatemala|Earthy, spinach-like|Tender after thorough cooking|Stew, soup, tamale, side|Protein; calcium; iron|Cook leaves fully; uncooked leaves contain cyanogenic compounds|Vegan
Cassava leaves, tender|Leafy plant part|West Africa; Central Africa; Indonesia|Earthy, savory, mild bitterness|Soft, finely chopped when cooked|Stew, soup, rice accompaniment|Protein; fiber; minerals|Cook thoroughly using established local preparation; never eat raw|Vegan
Sweet potato leaves|Leafy green|East Asia; Pacific; Africa|Mild, green, slightly sweet|Tender leaves and stems|Stir-fry, soup, stew|Vitamin A; vitamin C; fiber|Separate stems and leaves; cook briefly|Vegan
Chrysanthemum greens / shungiku|Leafy herb vegetable|Japan; Korea; China|Herbal, bitter, floral|Tender leaves, crisp stems|Hot pot, salad, soup, saute|Vitamin A; folate; potassium|Use young leaves raw or cook briefly|Vegan
Mitsuba|Leafy herb vegetable|Japan|Celery-like, parsley, anise|Tender leaves and stems|Soup, rice bowl, garnish|Vitamin C; small culinary amounts|Add late; aroma fades with heat|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Kepel fruit|Tropical fruit|Java, Indonesia|Sweet, fragrant, mango-like|Soft juicy pulp|Fresh eating, regional dessert|Carbohydrate; fiber|Ripe fruit is seasonal; local usage varies|Vegan
Safou / African pear|Tree fruit|Central and West Africa|Buttery, savory, mildly acidic|Firm raw; softens when warmed|Roast, boil, snack, stew|Fat; fiber; minerals|Often warmed or cooked; seed is discarded|Vegan
African star apple / agbalumo|Tree fruit|West Africa|Sweet-tart, resinous, astringent|Juicy segments around seeds|Fresh eating, drink, preserve|Vitamin C; fiber|Ripeness strongly changes astringency|Vegan
African mangosteen / bush mango|Tree fruit|West and Central Africa|Sweet-tart, aromatic|Soft pulp; seed kernel used in some cuisines|Fresh fruit, soup thickener, seed product|Fiber; fats in kernel|Seed products and fruit pulp are distinct; use food-grade source|Vegan
Miracle berry|Berry|West Africa|Mild itself; temporarily changes sour perception|Small red berry|Fresh tasting fruit, specialty culinary use|Small culinary serving|Effect is temporary; not a sweetener substitute for dietary needs|Vegan
Gac fruit|Tropical fruit|Vietnam; Southeast Asia|Mild, vegetal, lightly sweet|Bright red oily arils|Rice, juice, dessert, coloring|Carotenoids; fat; vitamin E|Use ripe arils; seeds and outer rind are not eaten|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Fonio, white Digitaria exilis|Ancient grain|West Africa; Sahel|Delicate, nutty, mild|Fine, fluffy grain|Porridge, couscous-style dishes, pilaf|Protein; fiber; minerals|Rinse thoroughly; grains cook quickly|Vegan; naturally gluten-free
Fonio, black Digitaria iburua|Ancient grain|West Africa; Nigeria; Benin|Earthy, aromatic, nutty|Small, slightly firmer grain|Porridge, couscous, festive dishes|Protein; fiber; minerals|Rinse and inspect for debris|Vegan; naturally gluten-free
Tef, ivory variety|Ancient grain|Ethiopia; Eritrea|Mild, sweet, earthy|Fine flour or tiny grains|Injera, porridge, baked goods|Iron; fiber; protein|Fermentation and flour shade vary|Vegan; naturally gluten-free
Tef, red-brown variety|Ancient grain|Ethiopia; Eritrea|Earthier, robust, nutty|Fine flour or tiny grains|Injera, porridge, grain blend|Iron; fiber; protein|Color and flavor depend on cultivar and milling|Vegan; naturally gluten-free
Amaranth, grain type Kiwicha|Pseudocereal|Andes; South America|Nutty, earthy|Tiny seed; crisp when popped|Porridge, flour, popped snack|Protein; iron; fiber|Rinse and cook; pop in a covered pan|Vegan; naturally gluten-free
Andean lupin / tarwi|Pulse|Peru; Bolivia; Andes|Nutty, rich, mildly bitter before processing|Firm bean|Salad, stew, snack, flour|Protein; fiber; fat|Must be debittered by proper soaking and washing; never eat raw bitter seed|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Korarima / Ethiopian cardamom|Spice seed|Ethiopia; Eritrea|Smoky, resinous, cardamom-like|Large dried pods|Berbere, stew, coffee, broth|Small culinary amounts|Not interchangeable with green cardamom; toast lightly|Vegan
Besobela / Ethiopian sacred basil|Aromatic herb|Ethiopia|Clove, basil, mint, warm spice|Dried leaf|Stew, spice blends, lentils|Small culinary amounts|Use dried leaf sparingly; aroma varies|Vegan
Ethiopian rue / tena adam|Aromatic herb and seed|Ethiopia; Eritrea|Bitter, citrus, warm resin|Leaves and seeds|Coffee, spice blends, stew|Small culinary quantities|Culinary rue differs from common rue; verify identity|Vegan
Melegueta pepper|Spice seed|West Africa; Gulf of Guinea|Peppery, citrus, ginger warmth|Small aromatic seeds|Soup, stew, spice blend, drink|Small culinary amounts|Grind as needed; flavor intensity varies|Vegan
Korarima leaf|Aromatic leaf|Ethiopian highlands|Camphor, citrus, resinous|Dried leaf|Stew, broth, seasoning|Small culinary amounts|Use food-grade product; potency varies|Vegan
Roselle leaf|Leafy herb|West Africa; South Asia; Caribbean|Tart, sour, green|Tender leaf|Soup, stew, greens dishes|Vitamin C; fiber|Leaf flavor differs from roselle calyx|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Camel milk, fermented shubat|Fermented dairy|Central Asia; Mongolia|Tangy, yeasty, lightly sour|Pourable cultured drink|Drink, porridge, regional foods|Protein; minerals; composition varies|Use safe, pasteurized product where available|Milk allergen
Mare milk, fresh|Milk|Central Asia; Mongolia|Sweet, mild, thin|Light liquid|Fermentation, drink, regional foods|Lactose; protein; composition varies|Use pasteurized food-grade milk where possible|Milk allergen
Yak butter|Dairy fat|Himalayan plateau|Rich, grassy, slightly smoky|Dense, firm fat|Tea, cooking, dough, finishing|Fat; fat-soluble vitamins vary|Flavor depends on pasture and churning|Milk allergen
Yak cheese, chhurpi hard|Hard dried cheese|Himalayan regions|Smoky, tangy, salty|Very hard, chewy|Snack, soup, tea accompaniment|Protein; calcium; sodium|Traditional long-dried form; chew slowly|Milk allergen
Queso de mano|Fresh stretched-curd cheese|Venezuela; Caribbean coast|Mild, milky, lightly salty|Soft, layered, elastic|Arepa, snack, sandwich|Protein; calcium; sodium|Keep chilled; milk source and salt vary|Milk allergen
Wara cheese / wara|Fresh curd cheese|Nigeria; West Africa|Mild, milky, lightly tangy|Soft, squeaky curd|Fry, stew, snack|Protein; calcium; fat varies|Use pasteurized milk product; refrigeration is important|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Ogi / akamu, fermented maize paste|Fermented cereal|Nigeria; West Africa|Sour, mild, cereal aroma|Smooth wet paste or powder|Breakfast porridge, drink, baking|Carbohydrate; protein varies|Cook thoroughly; fermentation and drying vary|Vegan; naturally gluten-free if uncontaminated
Kenkey dough, fermented maize|Fermented cereal|Ghana|Sour, yeasty, corn-forward|Dense fermented dough|Steamed dumpling with fish and pepper sauce|Carbohydrate; fiber|Fermentation time affects sourness; cook thoroughly|Vegan; naturally gluten-free if uncontaminated
Banku dough, fermented maize-cassava|Fermented cereal and root|Ghana|Sour, mild, tangy|Smooth cohesive dough|Steamed or simmered staple|Carbohydrate; fiber|Cook while stirring to a safe, fully cooked texture|Vegan
Injera starter / ersho|Fermentation starter|Ethiopia; Eritrea|Sour, yeasty, cereal aroma|Live batter starter|Fermented flatbread|Nutrition depends on flour and fermentation|Maintain clean utensils and controlled fermentation|Vegan; naturally gluten-free only with pure teff flour
Tamarind seed powder|Seed flour|South Asia; Africa|Mild, earthy, tart edge|Fine flour|Thickener, roasted flour, traditional foods|Protein; fiber; minerals|Use food-grade processed seed product|Vegan
Baobab seed oil|Seed oil|West Africa; Southern Africa|Mild, nutty, slightly fruity|Light oil|Dressing, finishing, specialty cooking|Unsaturated fats; tocopherols|Food-grade oils vary by extraction and storage|Vegan
`));

data['Vegetables'].push(...parse(`
Nalta jute stem|Leafy plant part|South Asia; Bangladesh|Earthy, green, lightly tart|Soft, silky when cooked|Soup, stew, greens dishes|Fiber; iron; micronutrients vary|Cook thoroughly; leaves and stems have different texture|Vegan
Sesbania flowers|Edible flower|South Asia; Southeast Asia|Bitter, green, floral|Delicate petals|Curry, fritter, salad, soup|Fiber; micronutrients vary|Use correctly identified food-grade flowers|Vegan
Moringa flowers|Edible flower|South Asia; East Africa|Mild, herbal, slightly sweet|Tender clusters|Soup, fritter, curry|Small amounts of micronutrients|Cook gently; use food-grade flowers|Vegan
Bamboo shoot, fermented sour|Fermented vegetable|South Asia; Southeast Asia|Sour, grassy, fermented|Crisp-tender strips|Stew, soup, curry, pickle|Fiber; sodium varies|Rinse or cook per traditional preparation; source food-grade product|Vegan
Taro stem / dasheen stalk|Stem vegetable|Caribbean; South Asia; Pacific|Mild, earthy|Fibrous; softens when cooked|Callaloo, stew, soup|Fiber; minerals|Peel and cook thoroughly; never eat raw|Vegan
Suran / elephant foot yam|Tuber|South Asia; Southeast Asia|Earthy, starchy, mildly sweet|Dense, firm, creamy when cooked|Curry, roast, stew, fry|Carbohydrate; fiber; potassium|Cook thoroughly; raw corm can irritate mouth|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
African medlar / Vangueria|Tree fruit|East Africa; Southern Africa|Sweet-tart, aromatic, slightly astringent|Soft, pulpy flesh|Fresh eating, drink, preserve|Vitamin C; fiber|Ripe fruit texture and flavor vary by species|Vegan
Uapaca fruit|Tree fruit|Central and Southern Africa|Sweet, creamy, mild tartness|Soft pulp around seed|Fresh eating, fermented drink|Carbohydrate; fiber|Use ripe fruit; local species names vary|Vegan
Monkey orange, Strychnos cocculoides|Tree fruit|Southern Africa|Sweet-sour, aromatic|Pulpy orange flesh in hard shell|Fresh eating, juice, preserve|Vitamin C; carbohydrate|Confirm species and ripeness; common name covers different fruits|Vegan
Pawpaw / Asimina triloba|Temperate fruit|Eastern North America|Custard-like, banana-mango aroma|Soft creamy flesh|Fresh eating, ice cream, puree|Vitamin C; manganese; carbohydrate|Perishable; remove large seeds|Vegan
Maypop / passionflower fruit|Temperate fruit|Eastern North America|Tart, tropical, aromatic|Juicy pulp and seeds|Drink, preserve, fresh eating|Vitamin C; fiber|Use ripe fruit; species identity matters|Vegan
Rose hip|Accessory fruit|Europe; West Asia; North America|Tart, floral, apple-like|Fleshy shell with irritating hairs and seeds|Tea, syrup, jam, sauce|Vitamin C; carotenoids|Remove seeds and irritating hairs; cook or strain|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Fonio flour, pre-cooked|Grain flour|West Africa|Mild, nutty, lightly sweet|Fine, fast-hydrating flour|Porridge, baking, thickener|Carbohydrate; protein; fiber|Hydration and particle size vary by product|Vegan; naturally gluten-free if uncontaminated
Sorghum malt|Malted grain|Africa; East Asia; North America|Sweet, malty, cereal-like|Sprouted and dried grain|Beverage, porridge, baking|Carbohydrate; enzymes|Malt processing and roast level vary|Vegan; naturally gluten-free only if uncontaminated
Rice, wild Zizania|Aquatic grass seed|Great Lakes, North America|Nutty, earthy, smoky|Long, chewy grains|Pilaf, soup, grain salad|Protein; fiber; minerals|Rinse and simmer; harvest source matters|Vegan; naturally gluten-free
Quinoa, royal black variety|Pseudocereal|Bolivia; Andes|Earthy, nutty, slightly bitter|Firm tiny grains|Pilaf, salad, porridge|Protein; fiber; minerals|Rinse well to remove saponins|Vegan; naturally gluten-free
Lima bean, Sieva type|Pulse|Caribbean; Central America|Mild, buttery, starchy|Creamy, medium-sized bean|Stew, succotash, puree|Protein; fiber; potassium|Cook dried beans thoroughly; raw beans contain cyanogenic compounds|Legume sensitivity possible
Pigeon pea, split dhal|Pulse product|South Asia; Caribbean; East Africa|Earthy, mild, nutty|Softens into creamy dal|Dal, soup, stew, rice dish|Protein; fiber; folate|Rinse and simmer; split and whole forms differ in cooking|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Black cardamom, Amomum subulatum|Spice pod|Nepal; India; Himalayas|Smoky, camphoraceous, resinous|Large ridged dried pod|Stew, rice, spice blend|Small culinary amounts|Crush pod; remove husk after infusion if desired|Vegan
Pepper elder / hoja santa|Aromatic leaf|Mexico; Central America|Anise, sassafras, peppery|Large tender leaf|Tamale, fish, bean dishes, sauce|Small culinary quantities|Use culinary-grade leaves; flavor softens with cooking|Vegan
Epazote, Mexican herb|Aromatic herb|Mexico; Central America|Pungent, resinous, citrus-herbal|Soft leaves and stems|Beans, quesadilla, stew|Small culinary amounts|Use sparingly; avoid medicinal dosing|Vegan
Papalo|Aromatic herb|Mexico; Central America|Cilantro-like, citrus, peppery|Broad tender leaves|Taco, salsa, sandwich, salad|Small culinary quantities|Use fresh; heat reduces aroma|Vegan
Huacatay / black mint|Aromatic herb|Andes, especially Peru|Mint, basil, citrus, pungent|Soft leaves|Sauce, potato dish, stew|Small culinary amounts|Fresh and paste forms vary in intensity|Vegan
Shallot, banana type|Allium bulb|Europe; Mediterranean|Sweet, mild onion-garlic|Long layered bulb|Sauce, roast, pickle, vinaigrette|Fiber; vitamin C|Peel thick outer skin; cook gently|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Lamb, Awassi type|Sheep meat|Levant; West Asia|Rich, grassy, savory|Tender to firm by cut|Roast, grill, mansaf-style dishes|Protein; iron; B12|Trim to preference; cook whole cuts safely|Animal product
Goat milk, cultured laban|Cultured dairy|Middle East; South Asia; Africa|Tangy, goaty, refreshing|Pourable fermented milk|Drink, sauce, marinade|Protein; calcium; fat varies|Use pasteurized product; cultures vary|Milk allergen
Buffalo ghee|Clarified dairy fat|South Asia; Middle East|Nutty, rich, buttery|Clear solid/semi-solid fat|Tempering, frying, finishing|Fat; energy dense|Clarification and storage change aroma|Milk allergen trace possible
Smoked sheep cheese, bryndza type|Sheep cheese|Carpathian regions|Tangy, salty, smoky|Soft to crumbly|Bread, potato, soup, dumpling|Protein; calcium; sodium|Aging and smoke levels vary by producer|Milk allergen
Chalap / fermented cheese curd|Fermented dairy|Central Asia|Sour, salty, cultured|Thin drink or soft curd|Drink, soup, regional foods|Protein; calcium; sodium|Product naming and milk source vary regionally|Milk allergen
Fish, mola carplet|Small freshwater fish|Bangladesh; Northeast India|Mild, sweet, delicate|Small bony fish|Stew, fry, greens dishes|Protein; calcium; vitamin A|Cook thoroughly; bones are fine and often eaten|Fish allergen
`));
data['Pantry & Condiments'].push(...parse(`
Sukuti, dried meat|Dried meat product|Nepal; Himalayan regions|Smoky, savory, concentrated|Chewy strips|Curry, snack, soup|Protein; sodium|Rehydrate or cook fully; storage method varies|Animal product
Ngari, fermented fish|Fermented fish condiment|Manipur, India|Pungent, salty, savory|Dense fermented fish|Eromba, chutney, curry|Protein; sodium|Use small amounts; fermentation and salt vary|Fish allergen
Hentak, fermented fish paste|Fermented fish condiment|Manipur, India|Strong, salty, earthy umami|Thick paste|Stew, chutney, curry|Protein; sodium|Use food-grade product; cook as customary|Fish allergen
Bamboo shoot pickle, Assamese|Fermented vegetable condiment|Assam, India|Sour, smoky, grassy|Crisp strips|Fish curry, pork, rice dishes|Fiber; sodium|Salt and fermentation vary by household|Vegan
Kapi / Thai shrimp paste|Fermented seafood condiment|Thailand|Salty, pungent, marine umami|Dense purple-brown paste|Curry paste, nam prik, stir-fry|Protein; sodium|Toast or fry briefly; use small amounts|Crustacean shellfish allergen
Yam bean paste, douchi style|Fermented soybean condiment|China; East Asia|Salty, earthy, savory|Soft fermented beans|Steamed fish, stir-fry, sauce|Protein; sodium|Rinse or soak if very salty|Soy allergen
`));

data['Vegetables'].push(...parse(`
Chenopodium greens / goosefoot|Leafy green|Europe; South Asia; Americas|Mineral, spinach-like, earthy|Tender to firm leaves|Saute, soup, stew, flatbread|Fiber; iron; vitamins vary by species|Use correctly identified food-grade leaves; wash well|Vegan
Gynura / Okinawan spinach|Leafy green|Southeast Asia; Okinawa|Mild, nutty, slightly mucilaginous|Succulent leaves|Stir-fry, soup, salad|Fiber; minerals|Cook lightly; fresh and cooked textures differ|Vegan
Fuki / butterbur stalk|Stem vegetable|Japan; Korea|Bitter, green, celery-like|Firm, fibrous stems|Simmered side dish, miso soup|Fiber; micronutrients vary|Use edible petioles only after proper parboiling and peeling; raw plant contains toxins|Vegan; careful preparation required
Fiddlehead, ostrich fern|Young fern frond|Northeastern North America; Japan|Green, asparagus-like, earthy|Tender curled frond|Saute, soup, pickle|Fiber; vitamin C; potassium|Cook thoroughly; do not eat raw or misidentified ferns|Vegan
Kelp stem, salted|Sea vegetable part|Korea; Japan|Briny, mineral, umami|Crunchy, chewy strips|Salad, soup, braise|Fiber; iodine|Rinse and soak; iodine and salt vary|Vegan
Sea grapes, Okinawan green caviar|Sea vegetable|Okinawa, Japan|Briny, fresh, mild umami|Small popping beads|Salad, rice, seafood|Minerals; iodine varies|Use food-grade cultivated product; avoid prolonged soaking|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Pawpaw, Shenandoah cultivar|Temperate fruit|Eastern United States|Custard-like, banana-mango, less seedy|Soft creamy flesh|Fresh eating, ice cream, puree|Vitamin C; manganese; carbohydrate|Perishable; remove large seeds|Vegan
Pawpaw, Susquehanna cultivar|Temperate fruit|Eastern United States|Sweet, rich, tropical aroma|Creamy flesh|Fresh eating, dessert, beverage|Vitamin C; carbohydrate|Keep chilled when ripe; remove seeds|Vegan
Hackberry fruit|Tree fruit|North America; Eurasia|Sweet, date-like, mildly nutty|Small fruit with hard seed|Fresh eating, flour, porridge|Carbohydrate; fiber|Use correctly identified ripe fruit; remove hard seeds|Vegan
Sea grape, Caribbean|Coastal fruit|Caribbean; Florida|Sweet-tart, grape-like|Juicy clusters|Jelly, wine, fresh eating|Vitamin C; antioxidants|Seasonal coastal fruit; observe local harvest rules|Vegan
Guavaberry / rumberry|Tree fruit|Caribbean|Spicy, tart, aromatic|Small juicy fruit|Liqueur, preserve, drink|Vitamin C; polyphenols|Often processed; ripe fruit is seasonal|Vegan
Kakadu plum|Native Australian fruit|Northern Australia|Very tart, green, aromatic|Small firm fruit|Sauce, preserve, powder|Vitamin C; polyphenols|Use food-grade processed products; acidity is strong|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Teff injera flour blend|Fermented grain flour|Ethiopia; Eritrea|Sour, earthy, lightly nutty|Fine flour, cohesive batter|Injera and fermented flatbread|Carbohydrate; fiber; iron varies|Commercial blends may include wheat or other grains|Vegan; check gluten cross-contact
Sorghum meal, white kafir|Milled grain|Southern Africa|Mild, sweet, earthy|Fine meal or coarse porridge|Porridge, flatbread, fermented foods|Fiber; protein; minerals|Hydration varies by milling|Vegan; naturally gluten-free if uncontaminated
Millet, fonio couscous product|Prepared grain|West Africa|Delicate, nutty|Light, separate grains|Side dish, salad, porridge|Carbohydrate; protein; fiber|Ready-to-cook products vary in pre-steaming|Vegan; naturally gluten-free if pure
Black gram, whole urad|Pulse|South Asia|Earthy, rich, nutty|Creamy but holds skin|Dal makhani, stew, fermented batter|Protein; fiber; iron|Soak and cook thoroughly; cooking time varies|Legume sensitivity possible
Bambara groundnut, red seed|Pulse|West Africa|Nutty, earthy, slightly sweet|Firm bean; creamy when cooked|Stew, snack, flour|Protein; carbohydrate; minerals|Soak and cook fully|Legume sensitivity possible
Jack bean|Pulse|South Asia; Southeast Asia|Mild, earthy|Large firm seed|Traditional processed food, stew|Protein; fiber|Requires proper detoxification and cooking; do not eat raw|Use only professionally processed food-grade products
`));
data['Spices & Herbs'].push(...parse(`
Ethiopian basil, basilicum type|Aromatic herb|Ethiopia; Eritrea|Sweet, clove-like, basil|Tender leaf|Spice blends, stew, tea|Small culinary quantities|Fresh and dried forms differ|Vegan
Koseret / Lippia adoensis|Aromatic herb|Ethiopia|Minty, resinous, oregano-like|Dried leaf|Spiced butter, stews, lentils|Small culinary quantities|Use food-grade product; species and names vary|Vegan
Ras el hanout, Taliouine style|Spice blend|Morocco|Floral, warm, savory, layered|Dry blend|Tagine, couscous, roasted foods|Sodium varies by maker|Recipes are not standardized; check allergens|May contain sesame or nuts; inspect label
Pepper leaf, Piper sarmentosum|Aromatic leaf|Southeast Asia|Peppery, citrus, herbal|Broad soft leaf|Wraps, curry, salad, grilled foods|Small culinary quantities|Use the culinary species, not unrelated pepper leaves|Vegan
Wild lime leaf, Citrus hystrix young leaf|Citrus leaf|Southeast Asia|Strong citrus, floral, pungent|Aromatic tender leaf|Paste, soup, curry|Small culinary amounts|Add late; leaf aroma fades with long boiling|Vegan
Mango ginger / Curcuma amada|Rhizome spice|South Asia|Raw mango, ginger, citrus|Firm rhizome|Pickle, chutney, curry, salad|Small culinary amounts|Use fresh or pickled; do not confuse with mango fruit|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Goat milk, raw-milk artisanal cheese|Cheese|Mediterranean; West Asia|Tangy, grassy, lactic|Soft to semi-firm|Cheese board, salad, bake|Protein; calcium; fat varies|Raw-milk status and aging rules vary; use safe regulated products|Milk allergen
Buffalo milk, cultured buttermilk|Cultured dairy|South Asia; Balkans|Tangy, rich, buttery|Pourable cultured drink|Drink, marinade, baking|Protein; calcium; fat varies|Use pasteurized product; fermentation differs by culture|Milk allergen
Yak yogurt|Cultured dairy|Himalayan plateau|Tangy, rich, grassy|Thick cultured dairy|Sauce, side dish, drink|Protein; calcium; fat|Milk source, altitude and culture affect taste|Milk allergen
Ewe's milk curd|Fresh dairy curd|Mediterranean; West Asia|Sweet, grassy, mild tang|Soft curds|Dessert, cheese, savory dishes|Protein; calcium; fat|Use pasteurized milk products; fresh curd is perishable|Milk allergen
Horse milk cheese, artisanal|Cheese|Central Asia; Eastern Europe|Tangy, sour, distinct|Firm to crumbly|Regional cheese, soup, snack|Protein; minerals|Rare product; use regulated producer and check fermentation|Milk allergen
Smoked goat cheese, Balkan style|Smoked cheese|Balkans|Tangy, smoky, grassy|Firm, sliceable|Cheese board, sandwich, grill|Protein; calcium; sodium|Smoke and aging vary by maker|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Shio koji, dried powder|Fermented rice seasoning|Japan|Sweet, salty, mellow umami|Fine powder|Marinade, seasoning, pickles|Carbohydrate; sodium varies|Powdered products vary in enzyme activity and salt|May contain barley; check label
Malt vinegar, unfiltered|Fermented vinegar|United Kingdom|Malty, sharp, toasted|Thin liquid|Pickle, chips, sauces|Negligible micronutrients|Acidity varies by producer|Vegan
Cane vinegar, Ilocos sukang iloko|Fermented vinegar|Philippines|Sharp, fruity, cane aroma|Thin liquid|Adobo, dipping sauce, pickle|Negligible micronutrients|Traditional ferment and acidity vary|Vegan
Date vinegar|Fermented fruit vinegar|Middle East; North Africa|Sweet-tart, fruity, caramel-like|Thin dark liquid|Salad, marinade, sauce|Sugars and acidity vary|Check whether sweetener is added|Vegan
Chili crisp, Sichuan style|Chili oil condiment|China|Toasted chili, numbing spice, savory|Crunchy solids in oil|Noodles, dumplings, eggs, vegetables|Fat; sodium varies|Oil and allergen ingredients vary by brand|May contain soy, sesame, peanut; check label
Toum, garlic emulsion|Garlic condiment|Levant|Sharp garlic, lemon, creamy|Fluffy emulsified paste|Grill, sandwich, dip|Fat; sodium varies|Refrigerate; homemade product is perishable|May contain egg depending on recipe
`));

// Back-to-back refresh pass 1 of 10: regional vegetables and sour fruits.
data['Vegetables'].push(...parse(`
Lamb's quarters / chenopodium album|Leafy green|Europe; South Asia; North America|Mineral, spinach-like, earthy|Tender leaf, slight chew|Saute, soup, stew, flatbread|Fiber; iron; vitamins vary|Use correctly identified food-grade leaves; wash well|Vegan
Jute plant stem|Stem vegetable|South Asia; Bangladesh|Mild, green, lightly tart|Tender, fibrous when mature|Stew, soup, greens dishes|Fiber; micronutrients vary|Peel mature stems and cook thoroughly|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Batuan / Garcinia binucao|Souring fruit|Philippines|Strong sour, fruity, lightly resinous|Firm green fruit|Sinigang-style sour soup, pickle, preserve|Organic acids; fiber|Used as a souring agent; maturity changes acidity|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Rye berries, whole|Whole grain|Northern Europe; Eastern Europe|Earthy, malty, pleasantly sour|Chewy kernels|Porridge, salad, soup, bread|Fiber; manganese; B vitamins|Soak and simmer until tender; contains gluten|Wheat/rye gluten allergen
`));
data['Spices & Herbs'].push(...parse(`
Kachri powder|Dried wild melon spice|Rajasthan, India|Sour, fruity, earthy|Fine powder|Meat tenderizer, chutney, curry|Small culinary amounts|Acidity and aroma vary by fruit and drying|Vegan
`));
data['Protein & Dairy'].push(...parse(`
`));
data['Pantry & Condiments'].push(...parse(`
Enset kocho|Fermented starchy staple|Ethiopia, especially southern highlands|Mildly sour, earthy|Dense, elastic fermented bread|Flatbread, porridge, stew accompaniment|Carbohydrate; fiber|Traditional processing is essential; use prepared food-grade product|Vegan
`));

// Back-to-back refresh pass 2 of 10: mountain crops and aromatic masala ingredients.
data['Vegetables'].push(...parse(`
Chayote, spiny cultivar|Gourd|Central America; South Asia|Mild, crisp, lightly sweet|Firm skin; juicy flesh|Curry, stew, pickle, saute|Water; fiber; vitamin C|Peel if spines are tough; cook until tender|Vegan
Pointed cabbage / Hispi|Brassica|Northern Europe|Sweet, mild mustard|Tender conical leaves|Slaw, roast, braise|Vitamin C; vitamin K; fiber|Cooks faster than dense storage cabbage|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Kaffir plum / Dovyalis caffra|Tree fruit|Southern Africa|Tart, citrusy, aromatic|Juicy, seedy fruit|Jam, chutney, drink|Vitamin C; organic acids|Usually cooked or sweetened due to sharp acidity|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Lentil, Spanish pardina|Pulse cultivar|Castile and León, Spain|Earthy, peppery, nutty|Small, firm, holds shape|Soup, salad, stew|Protein; fiber; folate|Rinse and simmer; age affects cooking time|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Chitrak root, culinary processed|Root spice|South Asia|Pungent, earthy, bitter|Dried root powder|Traditional spice blends, pickles|Small culinary amounts|Use only food-grade prepared product; raw plant parts can be irritating|Vegan; use measured culinary amounts
`));
data['Protein & Dairy'].push(...parse(`
Lamb kidney|Offal|Global; especially South Asia and Europe|Mineral, savory, gently sweet|Firm, tender when cooked briefly|Grill, curry, pie, saute|Protein; iron; B12|Trim core and cook safely; avoid overcooking|Animal product
`));
data['Pantry & Condiments'].push(...parse(`
Sour plum paste, umeboshi paste|Fruit condiment|Japan|Salty, sour, fruity, fermented|Thick paste|Rice, dressing, sauce, pickle|Sodium; organic acids|Very salty; use small amounts|Vegan; check additions
`));

// Back-to-back refresh pass 3 of 10: tropical crops and regional grain products.
data['Vegetables'].push(...parse(`
Kangkong flower shoots|Aquatic vegetable part|Southeast Asia|Mild, green, slightly sweet|Tender hollow stems and buds|Stir-fry, soup, curry|Vitamin A; vitamin C; folate|Cook promptly; separate tough stems|Vegan
Fuki-no-tou / butterbur flower bud|Edible flower bud|Japan|Bitter, herbal, earthy|Compact tender bud|Tempura, miso, saute|Fiber; micronutrients vary|Use correctly identified food-grade buds; blanch or cook|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Santol, yellow cultivar|Tree fruit|Southeast Asia|Tart-sweet, floral|Soft cottony pulp|Curry, preserve, relish|Fiber; vitamin C|Do not eat seeds; variety changes acidity|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Barley, pearl pot type|Whole grain|Europe; West Asia|Nutty, malty, mild|Chewy, plump grains|Soup, stew, pilaf|Fiber; beta-glucan; selenium|Rinse; contains gluten|Barley gluten allergen
`));
data['Spices & Herbs'].push(...parse(`
Anardana, pomegranate seed powder|Sour spice|South Asia|Tart, fruity, lightly tannic|Coarse or fine dried seed|Chaat, curry, chutney, marinade|Fiber; organic acids|Whole seeds and powder vary in sharpness|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Lamb liver|Offal|Middle East; South Asia; Mediterranean|Rich, mineral, slightly sweet|Soft, delicate|Grill, saute, curry, pate|Protein; iron; vitamin A; B12|Cook safely; limit portion frequency due to high vitamin A|Animal product
`));
data['Pantry & Condiments'].push(...parse(`
Fermented black bean sauce, Cantonese|Fermented condiment|Guangdong, China|Salty, earthy, deep umami|Loose sauce with fermented beans|Steamed fish, tofu, stir-fry|Sodium; protein varies|Use sparingly; brands differ in sugar and salt|Soy allergen
`));

// Back-to-back refresh pass 4 of 10: roots, berries, and dairy styles.
data['Vegetables'].push(...parse(`
Taro, cormel / dasheen small tubers|Corm|Caribbean; South Asia|Nutty, starchy, mild|Small dense tubers|Boil, roast, stew, chips|Carbohydrate; fiber; potassium|Cook fully; peel carefully|Vegan
Radish, watermelon type|Root|China; East Asia|Peppery skin, sweet juicy center|Crisp white-green exterior, pink center|Salad, pickle, slaw|Vitamin C; fiber|Slice thin; storage changes crispness|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Bilberry|Berry|Northern Europe; boreal regions|Tart, deep berry, tannic|Small juicy dark berry|Jam, pastry, drink, fresh eating|Vitamin C; anthocyanins; fiber|Delicate; wild and cultivated fruit differ|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Split fava, broad bean|Pulse product|Mediterranean; North Africa; West Asia|Earthy, buttery, mild|Creamy split bean|Ful, soup, puree, falafel-style dishes|Protein; fiber; folate|Cook thoroughly; fava sensitivity can occur in G6PD deficiency|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Grains of paradise, smoked type|Spice seed|West Africa|Citrus, pepper, warm smoke|Hard aromatic seeds|Stew, rub, spice blend|Small culinary amounts|Smoke-dried forms are intense; grind sparingly|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Kumis cheese curd|Cultured dairy|Central Asia|Sour, yeasty, lightly salty|Soft to firm curds|Regional snacks, soup, dough|Protein; calcium; composition varies|Milk source and fermentation vary|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Preserved mustard seed, rai achar|Pickled condiment|Nepal; Himalayan South Asia|Sharp, sour, pungent|Whole mustard seeds in oil|Rice, dal, pickle plate|Fat; sodium varies|Mustard intensity and fermentation vary|Mustard allergen
`));

// Back-to-back refresh pass 5 of 10: rainforest foods and culinary oils.
data['Vegetables'].push(...parse(`
Amazonian heart of palm, pupunha|Palm shoot|Amazon Basin|Mild, nutty, sweet|Tender cylindrical core|Salad, soup, grill, stew|Fiber; potassium|Use sustainably harvested or cultivated product|Vegan
Chicory root, Belgian forcing type|Root|Belgium; Northern Europe|Bitter, earthy, mildly sweet|Firm root, tender blanched shoots|Salad, braise, coffee substitute|Fiber; inulin|Bitter flavor varies by blanching and roast|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Araza / Amazonian guava|Amazonian fruit|Brazil; Peru; Colombia|Very tart, citrusy, aromatic|Soft yellow pulp|Drink, sorbet, preserve|Vitamin C; organic acids|Usually processed due to high acidity|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Canihua, Kañiwa, yellow-brown landrace|Pseudocereal|Andes, Peru and Bolivia|Nutty, earthy, slightly sweet|Tiny seed|Porridge, flour, grain bowl|Protein; iron; fiber|Rinse and simmer; landraces vary|Vegan; naturally gluten-free
`));
data['Spices & Herbs'].push(...parse(`
Pili nut oil|Nut oil|Philippines|Mild, buttery, nutty|Rich edible oil|Dressing, baking, finishing|Monounsaturated fat; vitamin E|Store sealed and cool to protect delicate oil|Tree nut allergen
`));
data['Protein & Dairy'].push(...parse(`
Pili nut, roasted|Tree nut|Philippines|Buttery, sweet, rich|Tender crisp kernel|Snack, dessert, sauce|Unsaturated fats; magnesium; vitamin E|Roast gently; rancidity develops with age|Tree nut allergen
`));
data['Pantry & Condiments'].push(...parse(`
Palm kernel oil, red unrefined|Edible oil|West Africa; Central Africa|Earthy, nutty, palm-fruit aroma|Rich red-orange oil|Stew, sauce, frying|Fat; carotenoids vary|Processing and sustainability vary by source|Vegan
`));

// Back-to-back refresh pass 6 of 10: central Asian and steppe staples.
data['Vegetables'].push(...parse(`
Turnip greens, Hakurei type|Leafy root crop tops|Japan; East Asia|Mild, peppery, sweet|Tender leaves and stems|Soup, pickle, saute|Vitamin C; vitamin K; fiber|Separate stems from leaves for even cooking|Vegan
Winter melon, fuzzy type|Gourd|China; Southeast Asia|Mild, cucumber-like|Dense white flesh|Soup, braise, curry|Water; fiber|Peel waxy or fuzzy skin; remove seeds|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Goumi berry / Elaeagnus multiflora|Berry|East Asia; cultivated in Europe|Tart, sweet, mildly astringent|Small juicy red fruit|Jam, sauce, fresh eating|Vitamin C; lycopene|Use ripe fruit; seeds are edible but firm|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Pea, grass pea, yellow seed|Pulse|Mediterranean; West Asia|Earthy, nutty|Small dry seed|Flour, stew, pulse dishes|Protein; fiber|Properly process; avoid heavy reliance due to neurotoxin risk in inadequately processed seed|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Safflower petal|Edible flower spice|West Asia; South Asia|Mild, hay-like, floral|Dried red-yellow petals|Rice, tea, coloring, stew|Small culinary amounts|Not equivalent to saffron in flavor; use food-grade product|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Mare milk, dried curd|Dried dairy|Central Asia; Mongolia|Tangy, sour, concentrated|Hard curd or powder|Soup, tea, snack|Protein; calcium; sodium varies|Traditional fermentation and drying vary|Milk allergen
`));
data['Pantry & Condiments'].push(...parse(`
Kumis, distilled spirit|Fermented dairy spirit|Central Asia|Sour, yeasty, alcoholic|Thin beverage|Drink, culinary specialty|Alcohol; nutrients vary|Alcohol content varies; not a dairy substitute|Contains alcohol; milk origin
`));

// Back-to-back refresh pass 7 of 10: island produce and preserved fish.
data['Vegetables'].push(...parse(`
Kombu stalk, su-kombu|Seaweed stem|Japan|Sweet-sour, mineral, umami|Chewy strips|Pickle, rice, snack, salad|Iodine; fiber|Commercially prepared product; iodine varies|Vegan
Matsumae greens mix|Prepared vegetable and seafood relish|Hokkaido, Japan|Salty, briny, savory|Crunchy chopped vegetables and kelp|Rice topping, side dish|Fiber; iodine; sodium|Contains seafood in many recipes; check label|May contain mollusk or fish allergens
`));
data['Fruit & Citrus'].push(...parse(`
Pandanus tectorius fruit, Karuka|Pacific fruit|Papua New Guinea; Pacific Islands|Sweet, nutty, mildly resinous|Fibrous keys around central core|Roast, chew, paste, traditional food|Carbohydrate; fiber|Use local edible species and preparation; fruit varies widely|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Sago starch, wet pearls|Palm starch product|Papua New Guinea; Indonesia|Neutral, mildly sweet|Translucent chewy pearls|Porridge, dessert, soup|Mostly carbohydrate|Cook thoroughly; product size changes cooking time|Vegan; naturally gluten-free
`));
data['Spices & Herbs'].push(...parse(`
Lemon ironbark honeybush herb|Australian herbal infusion|Australia|Lemon, floral, woody|Dried leaf|Tea, dessert infusion|Small culinary amounts|Use food-grade dried leaves; not a true tea plant|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Mullet roe, bottarga|Cured fish roe|Mediterranean; Japan|Salty, marine, concentrated umami|Firm, shaveable roe|Pasta, rice, eggs, salad|Protein; omega-3; sodium|Use sparingly; curing and species vary|Fish allergen
`));
data['Pantry & Condiments'].push(...parse(`
Sambal roa, smoked fish chili condiment|Chili-fish condiment|North Sulawesi, Indonesia|Smoky, hot, fishy umami|Oily coarse relish|Rice, noodles, grilled foods|Fat; sodium|Recipes vary; use small portions|Fish allergen
`));

// Back-to-back refresh pass 8 of 10: regional dairy and meat cuts.
data['Vegetables'].push(...parse(`
Pickled garlic scape|Pickled allium stem|East Asia; North America|Sour, garlicky, gently sweet|Crisp tender stems|Condiment, salad, rice, sandwich|Fiber; sodium varies|Brine and sugar vary by producer|Vegan
Mizuna red-stem type|Leafy brassica|Japan|Peppery, mild mustard, earthy|Feathery leaves, crisp stems|Salad, hot pot, stir-fry|Vitamins A, C and K|Add late to cooking|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Gandaria / plum mango|Tree fruit|Indonesia; Malaysia|Sweet-sour, mango-plum aroma|Firm juicy flesh|Rujak, pickle, sauce, fresh eating|Vitamin C; fiber|Unripe and ripe fruit uses differ|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Chickpea flour, roasted gram flour / sattu|Pulse flour|North India; Pakistan|Nutty, roasted, earthy|Fine dry flour|Drink, litti, stuffing, dough|Protein; fiber; iron|Roasted flour absorbs liquid; check grind|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Mahlab, Syrian kernel spice|Stone-fruit kernel spice|Levant|Almond, cherry, floral|Fine powder|Bread, pastry, cheese dishes|Small culinary amounts|Use food-grade product; avoid bitter overuse|Kernel allergy cross-contact possible
`));
data['Protein & Dairy'].push(...parse(`
Beef cheek|Beef cut|Global; braising traditions|Rich, beefy, gelatinous|Dense connective tissue|Braise, stew, taco, ragu|Protein; iron; B12|Slow-cook until tender; cook safely|Animal product
`));
data['Pantry & Condiments'].push(...parse(`
Liver sausage, leberwurst|Cured offal spread|Germany; Central Europe|Savory, mineral, spiced|Soft spread|Bread, snack, sauce|Protein; iron; sodium varies|Refrigerate according to label|Animal product
`));

// Back-to-back refresh pass 9 of 10: fruit cultivars and fermented staples.
data['Vegetables'].push(...parse(`
Rakkyo, pickled scallion bulb|Allium bulb|Japan; China|Sweet-sour, pungent, crisp|Small crunchy bulbs|Curry, rice, pickle plate|Fiber; sodium varies|Pickled products vary in sugar and salt|Vegan
Kohlrabi, purple type|Brassica stem|Europe; South Asia|Sweet, mild, cabbage-like|Crisp bulb; tender leaves|Slaw, roast, soup, pickle|Vitamin C; fiber|Peel thick outer skin; leaves are edible|Vegan
`));
data['Fruit & Citrus'].push(...parse(`
Mangosteen, Mesta type|Tropical fruit|Southeast Asia|Sweet, tart, floral|Soft white segments|Fresh eating, dessert, preserve|Vitamin C; fiber|Remove rind; avoid staining latex|Vegan
`));
data['Grains & Pulses'].push(...parse(`
Rice, Bhutanese red rice|Whole grain|Bhutan|Nutty, earthy, aromatic|Soft but distinct grains|Steamed rice, pilaf|Fiber; minerals; carbohydrate|Rinse and simmer; parboiled styles cook faster|Vegan; naturally gluten-free
`));
data['Spices & Herbs'].push(...parse(`
Black cumin, Bunium persicum|Spice seed|Iran; Afghanistan; Kashmir|Earthy, caraway, resinous|Small dark seeds|Rice, bread, meat, spice blend|Small culinary quantities|Distinct from Nigella sativa; grind or toast lightly|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Pork jowl|Pork cut|East Asia; Italy; Southern US|Rich, savory, fatty|High fat, tender when braised|Guanciale, stew, cured meat|Protein; fat; sodium varies|Cured and fresh forms differ; cook safely|Animal product
`));
data['Pantry & Condiments'].push(...parse(`
Fermented soybean paste, doenjang regional house style|Fermented soybean condiment|Korea|Earthy, salty, fermented umami|Coarse paste|Soup, stew, dip, marinade|Protein; sodium|Salt, aging and grain content vary by maker|Soy allergen; check wheat/barley
`));

// Back-to-back refresh pass 10 of 10: less common seeds, herbs, and animal ingredients.
data['Vegetables'].push(...parse(`
Garlic chives, yellow blanched type|Allium leaf|China; East Asia|Mild garlic, sweet, onion-like|Soft pale leaves|Dumpling, stir-fry, pancake|Vitamin K; vitamin C|Blanching changes flavor and texture|Vegan
Yam bean leaves, food-grade cultivated|Leafy plant part|Mexico; Central America|Mild, green|Tender cooked leaf|Traditional cooked greens in some local preparations|Micronutrients vary|Only correctly identified cultivated leaves; seeds and pods are toxic|Vegan; plant-part safety critical
`));
data['Fruit & Citrus'].push(...parse(`
Chokecherry|Stone fruit|North America|Tart, tannic, almond-like aroma|Small dark fruit around pit|Jelly, sauce, syrup|Anthocyanins; fiber|Remove pits; leaves and seeds contain cyanogenic compounds|Vegan; use fruit flesh only
`));
data['Grains & Pulses'].push(...parse(`
Lentil, Eston green|Pulse cultivar|Canada; North America|Mild, earthy, peppery|Small firm lentil|Soup, salad, stew|Protein; fiber; folate|Rinse and simmer; do not overcook for salads|Legume sensitivity possible
`));
data['Spices & Herbs'].push(...parse(`
Myrtle berry, dried|Spice berry|Mediterranean; Corsica; Sardinia|Juniper-like, herbal, balsamic|Small dried berry|Game, stew, sausage, sauce|Small culinary amounts|Use culinary-grade berries; grind as needed|Vegan
`));
data['Protein & Dairy'].push(...parse(`
Beef heart|Offal|Global; South America; Europe|Lean, mineral, beefy|Dense, fine-grained|Grill, stew, anticucho-style skewer|Protein; iron; B12|Trim connective tissue and cook safely|Animal product
`));
data['Pantry & Condiments'].push(...parse(`
Fermented locust bean, Dawadawa type|Fermented seed condiment|Ghana; Nigeria; West Africa|Pungent, savory, earthy umami|Soft fermented cakes or paste|Soup, stew, seasoning|Protein; sodium varies|Fermentation and seed species differ regionally|Legume sensitivity possible
`));

// Ten consecutive expansion rounds requested on 2026-10-23. Each round adds
// region-specific ingredient forms across the six requested catalog families.
const refreshRounds = [
  {
    'Vegetables': `African eggplant, bitter|Solanum aethiopicum|West and Central African; Ghanaian|Bitter, savory|Firm fruit softens into a creamy texture|Stew, soup, relish|Fiber; potassium; carotenoids|Cook thoroughly; bitterness varies by cultivar|Vegan\nChayote shoot tips|Cucurbit shoot|Mexican; Central American; Filipino|Green, mild, lightly sweet|Tender tips with crisp stems|Stir-fry, soup, sauté|Fiber; folate; vitamin C|Pick young shoots; rinse well|Vegan`,
    'Fruit & Citrus': `African star apple (agbalumo/udara)|Chrysophyllum albidum|West African; Nigerian; Ghanaian|Sweet-tart, milky, resinous|Juicy pulp around seeds|Fresh snack, drink, fruit salad|Vitamin C; fiber|Eat ripe pulp; discard seeds|Vegan\nMarula fruit|Sclerocarya birrea|Southern African; Bantu food traditions|Tart, aromatic, tropical|Juicy, soft when ripe|Juice, fermented drink, preserves|Vitamin C; organic acids|Use ripe fruit; flavor varies with ripeness|Vegan`,
    'Grains & Pulses': `Bambara groundnut, fresh|Vigna subterranea|West African; Sahelian|Nutty, earthy|Dense, chestnut-like when cooked|Boil, stew, flour, snack|Protein; starch; fiber; minerals|Cook until fully tender|Legume sensitivity possible\nKerala red matta rice|Oryza sativa, Kerala landrace|Kerala, India|Earthy, nutty|Robust grain with bran|Meals, kanji, idli blends|Carbohydrate; fiber; B vitamins|Rinse and allow longer cooking than polished rice|Gluten-free`,
    'Spices & Herbs': `Grains of Selim (uda)|Xylopia aethiopica|West African; Ghanaian; Nigerian|Smoky, peppery, resinous|Hard aromatic pods|Soups, stews, spice blends|Aromatic compounds; used in small amounts|Crush or simmer whole; remove woody pods|Vegan\nLippia javanica (lemon bush)|Lippia javanica|Southern African|Lemon-like, herbal, camphoraceous|Aromatic leaves|Herbal infusion, seasoning, smoke flavor|Plant compounds; culinary quantities are small|Use sparingly for strong aroma|Vegan`,
    'Protein & Dairy': `Kashkaval cheese, sheep milk|Aged pasta filata cheese|Balkans; Romania; Turkey|Salty, buttery, tangy|Firm, sliceable; melts well|Grill, bake, sandwiches|Protein; calcium; fat|Salt and aging vary by producer|Dairy; may contain rennet\nKurut (dried yogurt balls)|Dried fermented dairy|Afghanistan; Tajikistan; Central Asia|Sharp, salty, sour|Hard, dry; dissolves when soaked|Soups, sauces, snack|Protein; calcium; sodium|Rehydrate and adjust salt|Dairy`,
    'Pantry & Condiments': `Ogiri (fermented melon-seed condiment)|Fermented egusi/ogbono seed paste|Nigeria; Igbo cuisine|Pungent, savory, deeply umami|Soft paste|Soups, stews, sauces|Protein and fat from seeds; sodium varies|Use a small amount; cook into dish|May contain seed allergens\nDawadawa, fermented roselle seeds|Fermented roselle seed condiment|Ghana; West Africa|Pungent, savory, earthy|Dense fermented paste|Soups, stews, sauces|Protein; minerals; sodium varies|Cook into hot dishes|Seed source varies`,
  },
  {
    'Vegetables': `Nalta jute leaves (molokhia)|Corchorus olitorius|Egyptian; Levantine; South Asian|Green, mild, faintly bitter|Leaves become silky and mucilaginous|Stew, soup, rice accompaniment|Folate; vitamins A and C; fiber|Chop and simmer briefly|Vegan\nTindora (ivy gourd)|Coccinia grandis|South Asian; Sri Lankan|Mild, grassy, slightly tart|Crisp flesh; tender when cooked|Fry, curry, pickle|Fiber; vitamin C|Cook young fruit; mature fruit can be seedy|Vegan`,
    'Fruit & Citrus': `Kei apple|Dovyalis caffra|Southern African|Brightly tart, fruity|Juicy with many small seeds|Jam, jelly, sauce, fresh when fully ripe|Vitamin C; fiber|Often sweetened or cooked due to acidity|Vegan\nBael fruit (wood apple)|Aegle marmelos|South Asian; Indian|Sweet, aromatic, resinous|Hard shell; fragrant pulp|Drink, preserve, chutney|Fiber; carbohydrates|Crack shell; strain fibrous pulp|Vegan`,
    'Grains & Pulses': `Adzuki bean, Hokkaido type|Vigna angularis cultivar|Japan; Hokkaido|Sweet, nutty|Holds shape; soft interior|Anko, rice, soup, savory dishes|Protein; fiber; folate|Soak optional; simmer until tender|Legume sensitivity possible\nFonio, white|Digitaria exilis|West African; Guinea; Mali; Senegal|Mild, nutty|Tiny grains cook fluffy|Couscous-style side, porridge, pilaf|Carbohydrate; amino acids; minerals|Rinse carefully; cooks quickly|Naturally gluten-free`,
    'Spices & Herbs': `Calabash nutmeg (ehuru)|Monodora myristica|West African; Nigerian|Warm, nutmeg-like, woody|Hard seed, finely ground|Soups, pepper soup, spice blends|Aromatic compounds|Grate or grind shortly before use|Vegan\nKorarima (Ethiopian cardamom)|Aframomum corrorima|Ethiopian; Eritrean|Floral, smoky, resinous|Aromatic pods|Berbere, stews, coffee blends|Aromatic compounds|Toast pods lightly and grind seeds|Vegan`,
    'Protein & Dairy': `Smen (aged clarified butter)|Fermented/aged butterfat|Moroccan; Maghrebi|Intensely savory, cheesy, nutty|Solid fat, melts readily|Couscous, tagine, finishing|Fat-soluble vitamins; predominantly fat|Use sparingly; salt and aging differ|Dairy\nKashk (dried whey/yogurt product)|Fermented dairy concentrate|Iranian; Persian|Sour, salty, umami|Powder, paste, or dried pieces|Ash reshteh, dips, soups|Protein; calcium; sodium|Reconstitute; dilute to taste|Dairy`,
    'Pantry & Condiments': `Tamarind concentrate, Ghanaian style|Concentrated tamarind pulp|Ghana; West Africa|Sour, fruity, lightly sweet|Thick paste|Soups, drinks, sauces, marinades|Carbohydrate; organic acids|Dilute and strain if fibrous|Vegan\nFermented locust bean seasoning, iru woro|Fermented Parkia seeds|Yoruba cuisine; Nigeria|Pungent, savory, umami|Whole fermented seeds|Soups, stews, sauces|Protein; fat; sodium varies|Add during cooking|Legume/seed source`,
  },
  {
    'Vegetables': `Celtuce stem|Lactuca sativa var. angustana|Chinese; East Asian|Mild, nutty, celery-like|Crisp peeled stem; tender leaves|Stir-fry, pickle, soup|Fiber; potassium; folate|Peel fibrous exterior|Vegan\nLotus root, fresh rhizome|Nelumbo nucifera|East Asian; South Asian|Mild, sweet, starchy|Crisp with characteristic holes|Stir-fry, soup, pickle, lotus chips|Fiber; vitamin C; potassium|Acidulate after slicing to limit browning|Vegan`,
    'Fruit & Citrus': `Finger lime, green caviar type|Citrus australasica|Australian native; Queensland|Tart, lime-like, aromatic|Juice vesicles pop like pearls|Seafood, desserts, drinks, garnish|Vitamin C; fiber in whole fruit|Cut crosswise and press pearls out|Vegan\nSantol, red type|Sandoricum koetjape|Southeast Asian; Thai; Filipino|Sweet-sour, resinous|Soft segmented pulp|Fresh, preserves, curries|Fiber; vitamin C|Remove large seeds; peel is not eaten|Vegan`,
    'Grains & Pulses': `Black gram, split dehusked (urad dal)|Vigna mungo|South Asian; Indian|Earthy, creamy|Breaks down and foams when ground|Dal, idli, dosa, papad|Protein; fiber; iron; folate|Rinse; soak for batter applications|Legume sensitivity possible\nHorse gram, whole kulthi|Macrotyloma uniflorum|South Asian; Indian; Himalayan|Earthy, robust|Firm; thickens after long cooking|Rasam, soup, dal, sprouted salads|Protein; fiber; iron|Soak and pressure-cook for tenderness|Legume sensitivity possible`,
    'Spices & Herbs': `Ajwain leaf (Indian borage)|Coleus amboinicus|South Asian; Caribbean|Oregano-like, thymol-rich|Fleshy aromatic leaves|Fritters, chutney, tea, stews|Aromatic compounds; micronutrients in leaves|Use sparingly; aroma is potent|Vegan\nCubeb pepper|Piper cubeba|Indonesian; Javanese|Peppery, camphoraceous, citrusy|Small dried tailed berries|Spice blends, sauces, pickles|Aromatic compounds|Grind fresh; flavor is distinctive|Vegan`,
    'Protein & Dairy': `Akkawi cheese|Brined white cheese|Palestinian; Levantine|Mild, salty, milky|Elastic, springy; melts|Pastries, sandwiches, desserts|Protein; calcium; sodium|Soak briefly if very salty|Dairy\nLabneh balls in oil|Strained fermented yogurt|Levantine|Tangy, creamy, salty|Dense and spreadable|Mezze, sandwiches, salads|Protein; calcium; fat varies|Keep chilled; oil is not a substitute for refrigeration|Dairy`,
    'Pantry & Condiments': `Kokum rind, dried|Garcinia indica rind|Western India; Konkan|Tart, fruity, mildly astringent|Leathery dried rind colors liquid|Sol kadhi, curries, dals|Organic acids; anthocyanins|Soak or simmer; add gradually|Vegan\nBagoong alamang, fermented shrimp paste|Fermented small shrimp|Philippines|Salty, pungent, savory|Coarse paste|Kare-kare, sautéed greens, sauces|Protein; sodium|Cook before serving; salt varies|Crustacean allergen`,
  },
  {
    'Vegetables': `Bitter melon, white cultivar|Momordica charantia cultivar|South Asian; East Asian|Bitter, vegetal|Firm flesh softens but holds shape|Stuffed, stir-fried, curry, soup|Vitamin C; folate; fiber|Salt or blanch if reducing bitterness|Vegan\nGai lan flower stems|Brassica oleracea var. alboglabra|Chinese; Cantonese|Sweet, mustardy, lightly bitter|Crisp stems and tender florets|Blanch, stir-fry, oyster-style sauce|Vitamins C and K; folate|Trim woody bases; cook stems first|Vegan`,
    'Fruit & Citrus': `Jabuticaba|Plinia cauliflora|Brazilian; Atlantic Forest|Sweet, grapey, floral|Juicy translucent pulp; thick skin|Fresh, wine, jelly, syrup|Vitamin C; anthocyanins; fiber|Use soon after harvest; skin is edible|Vegan\nMamey sapote|Pouteria sapota|Mesoamerican; Caribbean|Sweet, pumpkin-like, almondy|Dense, creamy orange flesh|Milkshakes, desserts, fresh eating|Carotenoids; fiber; potassium|Eat only when fully ripe; remove large seed|Vegan`,
    'Grains & Pulses': `Black-eyed pea, crowder type|Vigna unguiculata cultivar|Southern United States; West African diaspora|Earthy, creamy|Creamy center; tender skin|Stew, fritters, rice dishes|Protein; fiber; folate|Rinse; simmer until tender|Legume sensitivity possible\nLupini beans, brined|Lupinus albus and related species|Mediterranean; Portuguese; Italian|Nutty, pleasantly bitter|Firm, chewy|Snack, salad, antipasto|Protein; fiber; low starch|Commercial brining required; raw lupin needs debittering|Legume allergen; lupin is a regulated allergen in some regions`,
    'Spices & Herbs': `Long pepper (pippali)|Piper longum|South Asian; Indonesian|Warm, sweet, peppery|Tiny clustered fruit spikes|Pickles, spice mixes, braises|Aromatic compounds|Grind or infuse whole|Vegan\nGrains of paradise, West African type|Aframomum melegueta|Ghana; Liberia; West Africa|Peppery, citrus, ginger-like|Small seeds|Spice blends, fish, stews|Aromatic compounds|Grind just before use|Vegan`,
    'Protein & Dairy': `Kumis (fermented mare milk)|Fermented equine milk|Kazakh; Kyrgyz; Central Asian|Tart, yeasty, lightly sparkling|Thin, drinkable|Beverage, cold soups|Protein; calcium; fermentation-dependent cultures|Fermentation and alcohol levels vary|Dairy; equine milk\nSulguni cheese|Brined stretched-curd cheese|Georgian; Caucasus|Salty, tangy, milky|Elastic, layered; melts|Khachapuri, grilled, salads|Protein; calcium; sodium|Fresh and smoked styles differ|Dairy`,
    'Pantry & Condiments': `Bissap (hibiscus calyx), dried|Hibiscus sabdariffa calyx|West African; Senegalese; Sahelian|Tart, cranberry-like, floral|Dried fleshy calyces|Infusion, syrup, sauces|Organic acids; anthocyanins|Steep; sweeten to taste|Vegan\nPomegranate molasses, Lebanese style|Reduced pomegranate juice|Levantine|Sweet-sour, fruity, tannic|Thick syrup|Dressings, marinades, stews|Carbohydrate; organic acids|Check for added sugar|Vegan`,
  },
  {
    'Vegetables': `Samphire (sea asparagus)|Salicornia europaea and related species|North Atlantic; Mediterranean|Briny, mineral, green|Crisp succulent stems|Blanch, pickle, seafood side|Fiber; minerals; sodium can be high|Rinse and taste before salting|Vegan\nCardoon stalk|Cynara cardunculus|North African; Italian; Mediterranean|Bitter, artichoke-like|Fibrous stalk becomes tender when braised|Gratin, stew, soup|Fiber; folate|Peel strings; blanch to temper bitterness|Vegan`,
    'Fruit & Citrus': `Naranjilla (lulo)|Solanum quitoense|Andean; Ecuadorian; Colombian|Tart, citrusy, tropical|Juicy seed-filled pulp|Juice, sauces, desserts|Vitamin C; fiber|Strain for juice; fruit is commonly cooked or blended|Vegan\nArazá (Amazonian guava)|Eugenia stipitata|Amazonian; Brazilian; Peruvian|Intensely tart, aromatic|Soft yellow pulp|Nectar, jam, sauce, ice cream|Vitamin C; carotenoids|Usually sweetened or cooked due to acidity|Vegan`,
    'Grains & Pulses': `Bambara groundnut, roasted flour|Vigna subterranea processed flour|West African; Nigerian; Ghanaian|Toasty, nutty|Fine, dense flour|Porridge, infant foods, baking|Protein; starch; fiber|Roast and mill; formulations vary|Legume sensitivity possible\nJob's tears (coix)|Coix lacryma-jobi var. ma-yuen|East Asian; Southeast Asian|Mild, nutty|Chewy, pearl-like when cooked|Congee, soup, tea, grain bowls|Carbohydrate; protein; minerals|Soak before simmering|Naturally gluten-free; cross-contact possible`,
    'Spices & Herbs': `Mastic resin|Pistacia lentiscus resin|Greek; Eastern Mediterranean|Piney, resinous, subtly sweet|Hard tears soften when ground|Breads, ice cream, puddings|Aromatic resin; used in small amounts|Freeze briefly and grind with sugar|Tree nut family cross-reactivity possible\nBlack lime (loomi)|Dried lime, Citrus × latifolia or related|Iranian; Iraqi; Gulf|Tart, smoky, fermented-citrus|Hard dried fruit|Stews, rice, fish, spice blends|Organic acids; negligible macro nutrients per use|Pierce and simmer whole or grind|Vegan`,
    'Protein & Dairy': `Caciocavallo cheese|Stretched-curd cow milk cheese|Southern Italian; Sicilian|Savory, buttery, sharper with age|Firm pear shape; melts|Grill, pasta, sandwiches|Protein; calcium; fat|Aging and salt differ by producer|Dairy; rennet varies\nQurut (Central Asian dried yogurt)|Dried cultured dairy|Tajik; Uzbek; Pamiri|Very sour, salty|Hard balls or cakes|Soups, sauces, snack|Protein; calcium; sodium|Dissolve in warm water; adjust salt|Dairy`,
    'Pantry & Condiments': `Berbere, Ethiopian household blend|Chili-based spice blend|Ethiopian; Eritrean|Hot, warm, floral, earthy|Fine powder or paste|Wot stews, lentils, marinades|Plant compounds; sodium low unless salted|Blend composition varies by household|May contain allergens from shared equipment\nMole negro paste|Chili, spice, nut and seed paste|Oaxacan; Mexican|Sweet, smoky, bitter, spicy|Thick, complex paste|Mole, braises, enchiladas|Energy and sodium vary widely|Dilute and simmer; check ingredients|May contain nuts, seeds, wheat, sesame`,
  },
  {
    'Vegetables': `Taro leaves (gabi leaves)|Colocasia esculenta leaves|Filipino; Pacific; South Asian|Earthy, mild|Silky after thorough cooking|Laing, stews, coconut dishes|Fiber; provitamin A; minerals|Never eat raw; cook thoroughly to destroy irritating calcium oxalate|Vegan when fully cooked\nBamboo shoots, fermented (suan sun)|Fermented bamboo shoot|Chinese; Southeast Asian|Sour, vegetal, funky|Crisp-tender|Noodle soups, hot pot, stir-fry|Fiber; low energy|Rinse if very sour; cook as directed|Vegan`,
    'Fruit & Citrus': `Kaffir plum (African mangosteen)|Dovyalis caffra|Southern African|Tart, aromatic, astringent|Juicy pulp, seeds|Jams, preserves, sauces|Vitamin C; fiber|Usually cooked or sweetened|Vegan\nPawpaw, North American|Asimina triloba|Eastern North American|Custardy, banana-mango, floral|Soft creamy flesh with large seeds|Fresh, custard, ice cream|Vitamin C; minerals; carbohydrates|Eat ripe; remove seeds and skin; highly perishable|Vegan`,
    'Grains & Pulses': `Rice, black glutinous (forbidden rice type)|Oryza sativa pigmented glutinous cultivar|China; Southeast Asia|Sweet, nutty|Sticky, chewy; purple-black when cooked|Desserts, porridge, rice cakes|Anthocyanins; fiber; carbohydrate|Soak for more even cooking|Naturally gluten-free\nLablab bean, hyacinth bean|Lablab purpureus|South Asian; African; Caribbean|Nutty, green|Firm bean; tender when cooked|Curry, stew, fresh pods as vegetable|Protein; fiber; folate|Cook mature beans thoroughly; some require soaking and boiling|Legume sensitivity; preparation matters`,
    'Spices & Herbs': `Ethiopian korarima, smoked-dried pods|Aframomum corrorima|Ethiopian; Eritrean|Smoky, citrus, cardamom-like|Aromatic seeds|Coffee, berbere, stews|Aromatic compounds|Crush pod; use seeds|Vegan\nVietnamese coriander (rau răm)|Persicaria odorata|Vietnamese; Southeast Asian|Peppery, citrusy, coriander-like|Tender aromatic leaves|Salads, soups, noodle dishes|Vitamin C; aromatic compounds|Add fresh near serving|Vegan`,
    'Protein & Dairy': `Jameed, dried sheep yogurt|Dried fermented sheep milk|Jordanian; Palestinian|Salty, tangy, deeply savory|Hard ball or block|Mansaf sauce, soups|Protein; calcium; sodium|Soak and blend before simmering|Dairy`,
    'Pantry & Condiments': `Tamarind candy, Mexican pulparindo style|Tamarind pulp confection|Mexican|Sour, sweet, salty, chili heat|Dense, chewy|Snack, garnish, sauce ingredient|Carbohydrate; sodium varies|Check chili and salt level|May contain added sugar\nFermented mustard greens (suan cai)|Brined fermented brassica greens|Chinese; Taiwanese|Sour, salty, savory|Tender, crisp stems|Noodle soups, fish, stir-fry|Fiber; sodium|Rinse if desired; salt varies|Vegan`,
  },
  {
    'Vegetables': `Chayote, spiny Costa Rican type|Sechium edule cultivar|Costa Rican; Central American|Mild, sweet, cucumber-like|Crisp, watery|Stuffed, sautéed, soup|Fiber; vitamin C; folate|Peel if skin is tough; remove seed if desired|Vegan\nAfrican nightshade leaves (managu)|Solanum scabrum and related|East African; Kenyan; Tanzanian|Bitter, green|Leaves soften quickly|Stew, sauté, maize meal side|Folate; iron; vitamins A and C|Use correctly identified edible species; cook thoroughly|Vegan`,
    'Fruit & Citrus': `Monstera deliciosa fruit (ripe only)|Monstera deliciosa|Central American; cultivated tropics|Pineapple-banana, aromatic|Segmented, juicy flesh|Fresh fruit, dessert|Carbohydrate; fiber|Eat only fully ripe segments; unripe fruit contains irritating calcium oxalate crystals|Vegan when ripe\nAfrican baobab pulp, dried|Adansonia digitata fruit pulp|Sahelian; East African|Tangy, citrusy, powdery|Dry powder dissolves in liquid|Drinks, porridge, sauces|Vitamin C; fiber; minerals|Reconstitute; natural powderiness is normal|Vegan`,
    'Grains & Pulses': `Rice, red cargo (Thai type)|Oryza sativa red bran cultivar|Thailand; Southeast Asia|Nutty, earthy|Chewy with intact bran|Rice bowls, salads, sides|Fiber; manganese; anthocyanins|Rinse; longer cooking than white rice|Naturally gluten-free\nMung bean, split dehulled|Vigna radiata processed pulse|South Asian; Chinese|Mild, sweet|Soft, creamy when cooked|Dal, congee, sprouts, desserts|Protein; fiber; folate|Rinse; cooks without long soaking|Legume sensitivity possible`,
    'Spices & Herbs': `Vanilla, Tahitian type|Vanilla × tahitensis|French Polynesia; Papua New Guinea|Floral, fruity, anise-like|Soft fragrant cured pod|Desserts, drinks, sauces|Aromatic compounds|Use pod, seeds, or extract; authenticity varies|Vegan\nGrains of Selim, ground seed|Xylopia aethiopica processed spice|West African|Smoky, peppery, resinous|Coarse aromatic powder|Pepper soup, stews, seasoning|Aromatic compounds|Use sparingly|Vegan`,
    'Protein & Dairy': `Skyr, Icelandic style|Cultured strained skim-milk product|Icelandic; Nordic|Tangy, clean, mildly sweet|Very thick, high-protein style|Breakfast, dips, desserts|Protein; calcium; fat varies|Sweetened products may contain added sugar|Dairy\nKefir, water kefir grains fermented drink|Fermented sugar-water culture|Caucasus-origin dairy tradition; globally adapted|Tart, lightly yeasty|Thin, sometimes fizzy|Drink, dressing, marinades|Organic acids; live cultures vary by product|Culture source and fermentation determine alcohol and microbes|May be dairy-free; check culture and ingredients`,
    'Pantry & Condiments': `Sambal tempoyak|Fermented durian chili condiment|Malay; Malaysian; Indonesian|Funky, sour, hot, savory|Thick fermented paste|Fish curries, sambal, rice dishes|Fat and carbohydrate from durian; sodium varies|Cook into dishes; fermentation is variable|Vegan versions exist\nMango achar, South African atchar|Oil-preserved green mango pickle|South African Indian diaspora|Sour, hot, salty, pungent|Firm chunks in spiced oil|Curries, sandwiches, grilled meats|Sodium; fat from oil|Oil and salt vary by recipe|Vegan; mustard may be present`,
  },
  {
    'Vegetables': `Sea kale shoots|Crambe maritima|Northern European; Atlantic coast|Sweet, cabbage-like, slightly briny|Tender blanched shoots|Steam, sauté, gratin|Fiber; vitamin C; minerals|Traditionally blanched; source only cultivated or legally harvested shoots|Vegan\nNew Zealand spinach|Tetragonia tetragonioides|Australasian; Pacific; cultivated globally|Mild, salty-green|Succulent leaves soften quickly|Sauté, soup, pie filling|Vitamins A and C; fiber|Blanch if sensitive to oxalates; rinse|Vegan`,
    'Fruit & Citrus': `Illawarra plum|Podocarpus elatus|Australian native; New South Wales|Sweet, berry-like, resinous|Dark fleshy receptacle around seed|Jam, sauce, desserts|Anthocyanins; fiber|Use ripe cultivated fruit; remove seed|Vegan\nQuandong (desert peach)|Santalum acuminatum|Australian native; arid regions|Tart, peachy, aromatic|Firm flesh, large stone|Stewed fruit, chutney, pies|Vitamin C; fiber|Often cooked or sweetened; source sustainably|Vegan`,
    'Grains & Pulses': `Rice, Koshihikari cultivar|Oryza sativa japonica cultivar|Japan; Niigata and other regions|Sweet, clean, fragrant|Glossy, tender, cohesive|Sushi, rice bowls, set meals|Carbohydrate; small amounts of B vitamins|Rinse gently; water ratio matters|Naturally gluten-free\nPigeon pea, split toor dal|Cajanus cajan processed pulse|South Asian; Caribbean; East African|Nutty, earthy|Creamy but can retain shape|Sambar, dal, rice dishes|Protein; fiber; folate; potassium|Rinse; soak older dal for faster cooking|Legume sensitivity possible`,
    'Spices & Herbs': `Tasmanian pepperberry|Tasmannia lanceolata|Australian native; Tasmania|Fruity, floral, building heat|Small dried berries|Seasoning, chocolate, game, sauces|Aromatic compounds|Use sparingly; heat intensifies after tasting|Vegan\nLemon myrtle, dried leaf|Backhousia citriodora|Australian native; Queensland|Intense lemon, sweet herbal|Aromatic leaf flakes|Tea, baking, seafood, seasoning|Citral-rich aroma|Use modest amount; can dominate|Vegan`,
    'Protein & Dairy': `Camel milk, fermented shubat|Fermented camel milk|Kazakh; Central Asian|Tart, yeasty, lightly sparkling|Thin drink|Beverage, regional fermented food|Protein; minerals; fat varies|Fermentation and alcohol levels vary|Camel milk; dairy cross-contact possible\nBison, ground meat|Bison bison|North American|Lean, mineral, beef-like|Lean mince|Burgers, chili, meatballs|Protein; iron; zinc|Cook ground meat to safe internal temperature|Mammalian meat`,
    'Pantry & Condiments': `Amba pickle, Iraqi style|Fermented mango pickle|Iraqi; Jewish Iraqi diaspora|Tangy, hot, fenugreek-rich|Soft spicy relish|Falafel, sabich, grilled fish|Sodium; carbohydrate varies|Check fermentation and salt level|Vegan\nTkemali, green plum sauce|Prunus cerasifera sauce|Georgian; Caucasus|Sour, herbal, tart|Pourable sauce|Grilled meats, potatoes, beans|Carbohydrate; sodium varies|Plum variety and herbs change flavor|Often vegan`,
  },
  {
    'Vegetables': `Water spinach, red-stem cultivar|Ipomoea aquatica cultivar|Southeast Asian; Vietnamese; Thai|Green, mild, lightly sweet|Hollow crisp stems and tender leaves|Stir-fry, soup, fermented greens|Fiber; folate; vitamins A and C|Cook thoroughly; source from clean water|Vegan\nYam bean, jicama seed pods excluded|Pachyrhizus erosus root|Mesoamerican; Filipino|Mild, juicy, apple-like|Crunchy raw or cooked|Salad, salsa, stir-fry|Fiber; vitamin C|Eat root only; seeds and other plant parts are toxic|Vegan`,
    'Fruit & Citrus': `Langsat, duku type|Lansium parasiticum cultivar|Indonesian; Malaysian|Sweet, tart, floral|Translucent segmented flesh|Fresh, syrup, preserves|Vitamin C; fiber|Peel and remove seeds; latex can be bitter|Vegan\nRose apple (jambu air)|Syzygium samarangense|Southeast Asian; Indonesian; Malaysian|Mild, floral, watery|Crisp bell-shaped flesh|Fresh, pickle, salad|Vitamin C; fiber|Flavor is delicate; serve chilled|Vegan`,
    'Grains & Pulses': `Rice, Jeerakasala (Wayanadan kaima)|Oryza sativa aromatic landrace|Kerala, India|Fragrant, buttery, mildly sweet|Short slender grains, cohesive|Malabar biryani, ghee rice|Carbohydrate; small amounts of B vitamins|Rinse gently; fragrance is easily masked|Naturally gluten-free\nPea, yellow split|Pisum sativum dried pulse|Northern Europe; South Asia|Mild, sweet, earthy|Breaks down to thick puree|Pease pudding, soup, dal|Protein; fiber; folate|Rinse; simmer until fully tender|Legume sensitivity possible`,
    'Spices & Herbs': `Mahlab (mahlepi)|Prunus mahaleb seed kernel spice|Eastern Mediterranean; Turkish; Greek|Almond-like, cherry, floral|Fine aromatic powder|Breads, sweet pastries, cheese desserts|Aromatic compounds; used sparingly|Grind fresh; store airtight|Stone fruit seed; avoid excessive use\nGrains of paradise, whole pod|Aframomum melegueta|Ghanaian; Liberian; West African|Citrus, ginger, pepper|Aromatic seeds in pod|Soup, fish, spice blends|Aromatic compounds|Toast gently and crush|Vegan`,
    'Protein & Dairy': `Yak milk churpi, soft type|Yak milk cheese|Himalayan; Nepalese; Bhutanese|Tangy, milky, savory|Chewy fresh cheese|Soup, curry, snack|Protein; calcium; fat|Fresh and dried forms differ greatly|Dairy\nPork jowl, cured guanciale style|Cured pork cheek/jowl|Italian; Roman|Rich, salty, porky|Fatty, renders crisp|Pasta, beans, greens|Protein; fat; sodium|Render gently; cure and salt vary|Pork`,
    'Pantry & Condiments': `Shito, Ghanaian black pepper sauce|Chili, fish, oil condiment|Ghanaian; Ga|Hot, smoky, salty, umami|Dark oily paste|Rice, fish, eggs, stews|Fat; protein if fish included; sodium|Recipe may include dried fish or shrimp|May contain fish or crustacean allergens\nChinchalok (fermented tiny shrimp)|Fermented small shrimp condiment|Malaysian; Peranakan|Salty, pungent, sour|Wet coarse relish|Sambal, lime condiment, stir-fry|Protein; sodium|Fermentation and salt vary; cook when appropriate|Crustacean allergen`,
  },
  {
    'Vegetables': `Fiddlehead fern (ostrich fern only)|Matteuccia struthiopteris young frond|Northeastern North American; Japanese|Green, grassy, lightly nutty|Crisp coils|Blanch, sauté, pickle|Fiber; vitamins A and C|Identify species correctly; boil at least 10–15 minutes before other cooking|Vegan; never eat raw\nTurmeric leaf|Curcuma longa leaf|South Asian; Indonesian|Green, ginger-like, warm|Flexible aromatic leaf|Wrap fish, steam rice, sweets|Aromatic compounds; limited nutrient contribution per use|Use as wrap or infusion; do not confuse with root|Vegan`,
    'Fruit & Citrus': `Kepel fruit|Stelechocarpus burahol|Javanese; Indonesian|Sweet, floral, mango-like|Soft aromatic pulp|Fresh, preserves|Fiber; carbohydrates|Eat ripe pulp; remove seeds|Vegan\nBignay (bugnay)|Antidesma bunius|Filipino; Southeast Asian|Tart, berry-like|Small clustered berries|Wine, vinegar, jam, sauce|Anthocyanins; vitamin C|Ripe fruit darkens; unripe is strongly tart|Vegan`,
    'Grains & Pulses': `Rice, red cargo (Bhutanese heirloom)|Oryza sativa red landrace|Bhutanese Himalaya|Nutty, earthy, mildly sweet|Chewy, colored bran retained|Rice meals, porridge, bowls|Fiber; manganese; carbohydrate|Rinse and soak for even cooking|Naturally gluten-free\nFava bean, dried broad bean|Vicia faba|Mediterranean; West Asian; North African|Earthy, buttery|Creamy when cooked|Ful medames, puree, soups|Protein; fiber; folate; iron|Soak and cook thoroughly; remove skins if desired|Fava allergy; contraindicated for people with G6PD deficiency`,
    'Spices & Herbs': `Za'atar herb (Origanum syriacum)|Origanum syriacum dried leaves|Levantine|Thyme-like, savory, resinous|Dried leaf flakes|Za'atar blend, breads, eggs|Aromatic compounds|Blend with sesame and sumac as desired|Sesame allergen in common blends\nAjí panca, dried pepper|Capsicum baccatum var. pendulum|Peruvian; Andean|Mild heat, fruity, smoky|Leathery dried chile|Adobo, anticuchos, sauces|Carotenoids; vitamin C in fresh form|Soak and blend; heat is usually mild|Vegan`,
    'Protein & Dairy': `Pork blood curd|Cooked coagulated pork blood|Southeast Asian; East Asian|Iron-rich, mild, savory|Soft, custard-like cubes|Soups, stir-fries, stews|Protein; heme iron|Buy from inspected source; cook thoroughly|Pork; blood product\nGhee, cultured-butter style|Clarified butterfat|South Asian; Indian|Nutty, toasted, buttery|Clear fat, solid when cool|Tadka, sauté, baking|Predominantly fat; fat-soluble vitamins|Water and milk solids removed to varying degrees|Dairy; trace milk proteins may remain`,
    'Pantry & Condiments': `Achuete annatto oil|Bixa orellana seed-infused oil|Filipino|Earthy, peppery, lightly floral|Red-orange infused oil|Pancit palabok, rice, marinades|Fat; carotenoid pigments|Strain seeds; do not scorch|Vegan\nKashmiri ver, spice cake|Kashmiri chili and spice paste/cake|Kashmiri; Himalayan|Warm, aromatic, mild heat|Dry cake dissolves to paste|Rogan josh, nadru, meat stews|Plant compounds; salt varies|Toast or dissolve in cooking fat|May contain asafoetida or shared-equipment allergens`,
  },
];
// Further ten-round pass: focus on established regional foods and named cultivars.
refreshRounds.push(
  {
    'Vegetables': `African baobab leaves, dried|Adansonia digitata leaves|Sahelian; West African|Mild, tart, leafy|Dried flakes rehydrate into a soft thickener|Soups, sauces, grain dishes|Fiber; minerals; composition varies by drying|Cook thoroughly; use food-grade dried leaves|Vegan\nMoringa pods, immature|Moringa oleifera pods|South Asian; East African|Green, mild, slightly sweet|Firm ridged pods soften with simmering|Curries, stews, sambar|Fiber; vitamin C; folate|Trim ends; cook until tender|Vegan`,
    'Fruit & Citrus': `White sapote|Casimiroa edulis|Mexican; Central American|Sweet, custardy, pear-like|Soft creamy flesh|Fresh, chilled dessert, puree|Carbohydrate; fiber; vitamin C|Remove seeds; eat ripe flesh only|Vegan\nBlack sapote|Diospyros nigra|Mexican; Central American|Mild sweet, chocolate-pudding-like when ripe|Soft dark pulp|Desserts, smoothies, baking|Fiber; carbohydrate|Eat only ripe soft pulp; unripe fruit is astringent|Vegan`,
    'Grains & Pulses': `Tepary bean, brown|Phaseolus acutifolius|Sonoran Desert; Southwestern United States; Northern Mexico|Nutty, earthy|Firm skin, creamy center|Beans, soups, stews, flour|Protein; fiber; iron|Soak optional; simmer until tender|Legume sensitivity possible\nCowpea, yardlong bean dried seed|Vigna unguiculata subsp. sesquipedalis seed|South and Southeast Asian|Earthy, mild|Creamy cooked pulse|Dal, stew, fritters|Protein; fiber; folate|Cook dried seed thoroughly|Legume sensitivity possible`,
    'Spices & Herbs': `Black cardamom, Nepal type|Amomum subulatum|Nepalese; Himalayan; North Indian|Smoky, resinous, camphoraceous|Large smoky pods|Curries, rice, braises|Aromatic compounds|Crush pod; remove before serving if desired|Vegan\nGrains of paradise, whole seed|Aframomum melegueta|West African; Ghanaian|Peppery, citrus, ginger-like|Small aromatic seeds|Seasoning, soups, fish|Aromatic compounds|Use culinary quantities; grind fresh|Vegan`,
    'Protein & Dairy': `Buffalo milk, fresh|Water buffalo milk|South Asian; Italian|Rich, mildly sweet|Creamier than many cow milks|Mozzarella, yogurt, drinks|Protein; calcium; fat varies|Pasteurization and composition vary|Dairy\nQuail eggs|Coturnix japonica eggs|East Asian; European|Rich, mild|Small firm white and yolk|Pickle, garnish, steam, fry|Protein; choline; micronutrients|Cook thoroughly; refrigerate|Egg allergen`,
    'Pantry & Condiments': `Verjuice (verjus)|Unripe grape juice|French; Persian; Mediterranean|Gently sour, fruity|Pourable juice|Dressings, sauces, deglazing|Organic acids; small carbohydrate amount|Use as mild acid; products vary in sweetness|Vegan\nPreserved lemon, Moroccan-style|Salt-cured Citrus limon|Moroccan; North African|Salty, intensely citrus, fermented|Soft peel; tender pulp|Tagines, salads, dressings|Sodium; citrus oils|Rinse peel if needed; recipe salt may need adjustment|Vegan`,
  },
  {
    'Vegetables': `Winged bean tuber|Psophocarpus tetragonolobus tuber|Papua New Guinea; Southeast Asian|Nutty, starchy|Dense tuber|Boil, roast, curry|Starch; protein; fiber|Cook until tender; cultivar and maturity affect texture|Vegan\nElephant foot yam|Amorphophallus paeoniifolius corm|South Asian; Southeast Asian|Earthy, starchy|Dense, firm; softens with cooking|Curry, roast, chips|Carbohydrate; fiber; potassium|Peel with gloves if sensitive; cook thoroughly to reduce irritation|Vegan`,
    'Fruit & Citrus': `Lucuma, Peruvian|Pouteria lucuma|Peruvian; Andean|Maple-like, sweet, pumpkin-like|Dry, dense yellow flesh|Ice cream, custard, powder|Carbohydrate; fiber; carotenoids|Usually eaten ripe or as dried powder|Vegan\nCamu camu|Myrciaria dubia|Amazonian; Peruvian; Brazilian|Very tart, berry-like|Small juicy fruit|Powder, juice, preserves|Vitamin C; polyphenols|Often blended or sweetened due to acidity|Vegan`,
    'Grains & Pulses': `Grass pea, white seed|Lathyrus sativus|South Asian; Mediterranean; East African|Earthy, nutty|Firm pulse, flour-producing|Flatbread, porridge, dal|Protein; fiber|Use as part of varied diet; prolonged heavy reliance on inadequately processed grass pea is unsafe|Legume; preparation and diet diversity matter\nLentil, black beluga|Lens culinaris cultivar|Canadian Prairies; South Asian|Earthy, peppery|Small glossy seeds hold shape|Salad, pilaf, soups|Protein; fiber; folate; iron|Rinse; simmer until tender|Legume sensitivity possible`,
    'Spices & Herbs': `Mastic, Chios tears|Pistacia lentiscus var. chia resin|Chios, Greece|Piney, resinous, subtly sweet|Hard resin tears|Breads, sweets, beverages|Aromatic resin; used in small amounts|Grind chilled resin with sugar|Pistachio-family botanical; consider allergy history\nAleppo pepper (pul biber)|Capsicum annuum dried flake|Syrian; Turkish|Fruity, mild heat, slightly salty|Oily coarse flakes|Eggs, kebab, dips, vegetables|Carotenoids; sodium varies by product|Check salt and oil in packaged versions|Vegan`,
    'Protein & Dairy': `Goat whey ricotta|Whey cheese|Italian; Mediterranean|Mild, milky, slightly tangy|Soft, moist curds|Pasta, pastry, spreads|Protein; calcium; fat varies|Use chilled; moisture varies|Dairy\nVenison, red deer|Cervus elaphus meat|European; New Zealand|Lean, mineral, gamey|Fine-grained, low-fat meat|Roast, stew, mince|Protein; iron; zinc|Cook with care to avoid drying; follow local safe-handling guidance|Mammalian meat`,
    'Pantry & Condiments': `Miso, hatcho style|Long-aged soybean miso|Aichi, Japan|Deep umami, earthy, salty|Dense dark paste|Soup, glaze, braise|Protein; sodium|Use modestly; salt differs by brand|Soy; may contain barley\nTamarind seed powder|Tamarindus indica seed kernel powder|South Asian; Thai|Mild, nutty, tart|Fine thickening powder|Soups, sauces, bakery|Fiber; protein; composition varies|Use food-grade processed powder|Legume/seed sensitivity possible`,
  },
  {
    'Vegetables': `Amaranth leaves, red cultivar|Amaranthus tricolor cultivar|South Asian; East African|Green, mildly earthy|Tender leaves; stems crisp|Saag, stir-fry, soup|Folate; vitamins A and C; iron|Cook leaves; nutrient content varies|Vegan\nPurslane, golden type|Portulaca oleracea cultivar|Mediterranean; West Asian|Lemony, lightly salty|Succulent crisp leaves and stems|Salad, stew, sauté|Omega-3 ALA; vitamin C; minerals|Wash carefully; use food-grade cultivated source|Vegan`,
    'Fruit & Citrus': `African monkey orange (wild medlar)|Strychnos cocculoides|Southern African|Sweet-tart, aromatic|Dryish edible pulp around seeds|Fresh pulp, juice, preserves|Vitamin C; fiber; composition varies|Use correctly identified ripe fruit; remove seeds|Vegan\nMarang|Artocarpus odoratissimus|Borneo; Malaysian; Bruneian|Sweet, creamy, aromatic|Soft white segments|Fresh fruit, dessert|Carbohydrate; fiber|Eat ripe segments; discard seeds unless cooked for food use|Vegan`,
    'Grains & Pulses': `Rice, Kalijira aromatic type|Oryza sativa aromatic landrace|Bangladesh; Bengal|Fragrant, sweet, nutty|Tiny slender grains|Pilaf, pulao, desserts|Carbohydrate; small amounts of B vitamins|Rinse gently; short grain cooks quickly|Naturally gluten-free\nMung bean, black gram whole urad|Vigna mungo|South Asian|Earthy, robust|Small dark seed; creamy cooked interior|Dal makhani, sprouts, batter|Protein; fiber; folate|Soak and cook fully|Legume sensitivity possible`,
    'Spices & Herbs': `Dill seed, Indian sowa|Anethum graveolens seed|South Asian; Bengali|Warm, caraway-like, citrusy|Small aromatic seeds|Pickles, lentils, breads|Aromatic compounds|Toast lightly; use whole or ground|Vegan\nSichuan pepper, green husk|Zanthoxylum species dried pericarp|Chinese; Sichuan|Citrusy, tingling, floral|Husks with numbing sensation|Mapo tofu, oil, seasoning|Aromatic compounds|Remove seeds; toast husks gently|Vegan`,
    'Protein & Dairy': `Rabbit, domestic meat|Oryctolagus cuniculus|European; global husbandry|Mild, lean, poultry-like|Fine-grained, low-fat|Stew, roast, braise|Protein; iron; B vitamins|Cook thoroughly; low fat can dry out|Mammalian meat\nSheep yogurt, strained (kishk base)|Cultured sheep milk|Levantine; Eastern Mediterranean|Tangy, rich|Thick strained dairy|Dips, sauces, dried kishk|Protein; calcium; fat varies|Keep refrigerated; fermentation varies|Dairy`,
    'Pantry & Condiments': `Saba banana vinegar|Fermented banana vinegar|Philippines|Fruity, gently sharp|Pourable vinegar|Dipping sauces, adobo, marinades|Acetic acid; negligible macros per use|Acidity and sweetness differ by producer|Vegan\nKecap manis, Indonesian sweet soy sauce|Soy sauce sweetened with palm sugar|Indonesia|Sweet, salty, umami|Thick syrupy sauce|Nasi goreng, satay, marinades|Carbohydrate; sodium|Check soy and wheat ingredients|Soy; often wheat`,
  },
  {
    'Vegetables': `Cassava leaves (pondu/saka-saka)|Manihot esculenta leaves|Central African; Congolese; West African|Green, earthy|Leaves break down into thick stew|Saka-saka, pondu, stews|Protein; fiber; micronutrients vary|Must be pounded and thoroughly cooked to reduce cyanogenic compounds; never eat raw|Vegan when properly prepared\nPumpkin leaves (ugu)|Telfairia occidentalis leaves|Nigerian; West African|Green, mild, slightly sweet|Tender leaves, firm stems|Soups, sauté, stew|Folate; iron; vitamins A and C|Wash; cook stems longer than leaves|Vegan`,
    'Fruit & Citrus': `African pear (bush butter fruit)|Dacryodes edulis|Central African; West African|Buttery, mild, resinous|Firm raw; soft and oily after warming|Boiled/roasted snack, side dish|Fat; fiber; composition varies|Usually heated/softened before eating|Vegan\nMonkey kola (Cola lepidota)|Cola lepidota|West and Central African|Sweet-tart, juicy|Fibrous juicy pulp|Fresh fruit, local drinks|Vitamin C; fiber|Eat ripe pulp; availability and names vary regionally|Vegan`,
    'Grains & Pulses': `Fonio, black fonio|Digitaria iburua|West African; Guinea; Nigeria|Nutty, earthy|Tiny grains cook light and fluffy|Couscous-style side, porridge|Carbohydrate; protein; minerals|Rinse carefully; cooks quickly|Naturally gluten-free\nBambara groundnut, cream seed|Vigna subterranea landrace|West African; Sahelian|Nutty, mildly sweet|Dense, chestnut-like|Boil, roast, flour, stew|Protein; carbohydrate; fiber|Cook until fully tender|Legume sensitivity possible`,
    'Spices & Herbs': `African basil (scent leaf)|Ocimum gratissimum|West African; Nigerian|Clove-like, herbal, peppery|Aromatic leaves|Pepper soup, stews, tea|Aromatic compounds|Add late for fresher aroma|Vegan\nUziza leaf|Piper guineense leaf|Nigerian; West African|Peppery, warm, resinous|Aromatic leaves|Soups, stews, pepper soup|Aromatic compounds|Use modestly; fresh and dried profiles differ|Vegan`,
    'Protein & Dairy': `Snail, African giant land snail|Achatina/Archachatina species|West African; Central African|Mild, earthy, savory|Chewy, firm|Pepper soup, stew, grill|Protein; iron; minerals|Use legally sourced food-grade snails; clean and cook thoroughly|Mollusk allergen\nGoat milk, fermented yogurt|Cultured goat milk|Middle Eastern; South Asian|Tangy, goaty, sharp|Pourable to thick|Drink, sauces, marinades|Protein; calcium; fat varies|Culture and fermentation vary|Dairy`,
    'Pantry & Condiments': `Ogbono seed powder|Irvingia gabonensis kernel|West African; Nigerian|Nutty, mild|Strong mucilaginous thickener|Ogbono soup, stews|Fat; protein; fiber|Use food-grade kernels; cook into soup|Tree seed sensitivity possible\nEru (okok) dried leaves|Gnetum africanum leaves|Central African; Cameroonian|Green, earthy|Fibrous leaves soften with long cooking|Eru soup, vegetables|Fiber; protein; micronutrients vary|Use food-grade leaves; cook thoroughly|Vegan`,
  },
  {
    'Vegetables': `Chicory, witloof endive|Cichorium intybus var. foliosum|Belgian; French|Bitter, crisp, mildly sweet when cooked|Tight pale leaves|Braise, gratin, salad|Fiber; folate; vitamin K|Trim base; cover-grown leaves are less bitter|Vegan\nCeleriac root|Apium graveolens var. rapaceum|Northern European; French|Celery-like, nutty|Dense, creamy when cooked|Remoulade, mash, soup, roast|Fiber; potassium; vitamin K|Peel thick skin; acidulate cut surface|Vegan`,
    'Fruit & Citrus': `Calamondin (calamansi), variegated cultivar|Citrus × microcarpa cultivar|Philippine; Southeast Asian|Tart, floral, aromatic|Small juicy citrus with edible peel|Juice, dipping sauce, marmalade|Vitamin C; fiber|Wash peel if using whole fruit|Vegan\nKumquat, Meiwa|Citrus japonica cultivar|Chinese; Japanese|Sweet peel, tart pulp|Small oval fruit, edible peel|Fresh, candied, marmalade|Vitamin C; fiber|Eat whole after washing; remove seeds if desired|Vegan`,
    'Grains & Pulses': `Rice, fragrant basmati Dehraduni type|Oryza sativa basmati landrace|Uttarakhand, India; Himalayan foothills|Floral, nutty|Long slender grains stay separate|Biryani, pilaf|Carbohydrate; small amounts of B vitamins|Rinse and soak; regional GI claims require substantiation|Naturally gluten-free\nChickpea, desi brown|Cicer arietinum desi type|South Asian; Ethiopian|Earthy, nutty|Smaller, firmer seed|Chana masala, roasted snack, flour|Protein; fiber; folate|Soak dried seed; cook fully|Legume sensitivity possible`,
    'Spices & Herbs': `Kaffir lime leaf (makrut lime leaf)|Citrus hystrix leaf|Thai; Indonesian; Malaysian|Floral, lime, resinous|Double-lobed aromatic leaf|Curry, soup, rice infusion|Aromatic compounds|Bruise or tear; remove tough leaf before serving|Vegan\nFresh turmeric leaf|Curcuma longa leaf|South Asian; Indonesian|Green, ginger-like, earthy|Broad flexible leaf|Wraps, steamed rice, desserts|Aromatic compounds|Use culinary leaf, not substitute for rhizome|Vegan`,
    'Protein & Dairy': `Camel meat, young camel|Camelus dromedarius|Arabian Peninsula; North/East African|Lean, beef-like, mildly sweet|Lean muscle|Stew, grilled cuts, rice dishes|Protein; iron; zinc|Cook thoroughly; cut and age affect tenderness|Mammalian meat\nYak meat|Bos grunniens|Tibetan Plateau; Himalayan|Lean, mineral, beef-like|Dense lean muscle|Stew, dried meat, grilled dishes|Protein; iron; zinc|Cook thoroughly; avoid overcooking lean cuts|Mammalian meat`,
    'Pantry & Condiments': `Aonori flakes|Ulva prolifera and related edible green algae|Japanese; East Asian|Briny, green, umami|Fine dried flakes|Okonomiyaki, noodles, rice|Iodine; minerals; composition varies|Use food-grade seaweed; iodine content varies|Seaweed; check iodine-sensitive diets\nKatsuobushi (bonito flakes)|Dried, smoked, fermented skipjack tuna|Japanese|Smoky, savory, umami|Thin dry flakes|Dashi, topping, sauces|Protein; sodium varies|Use food-grade product; contains fish|Fish allergen`,
  },
  {
    'Vegetables': `Sea beet leaves|Beta vulgaris subsp. maritima|Mediterranean; Atlantic coasts|Mineral, salty, spinach-like|Tender leaves, crisp stems|Sauté, soup, pie filling|Fiber; folate; vitamins A and C|Wash thoroughly; cultivated food-grade source|Vegan\nNew Zealand yam (oca)|Oxalis tuberosa|Andean; New Zealand cultivation|Tangy, mildly sweet|Small waxy tubers; sweeten after curing|Roast, boil, stew|Carbohydrate; vitamin C|Cook; raw tubers are strongly tart and oxalate-rich|Vegan`,
    'Fruit & Citrus': `Cupuacu|Theobroma grandiflorum|Amazonian; Brazilian|Sour-sweet, chocolate-pineapple aroma|Creamy aromatic pulp|Juice, confectionery, ice cream|Fiber; vitamin C; fats in seeds|Use ripe pulp; seeds are processed separately|Vegan\nPequi|Caryocar brasiliense|Brazilian Cerrado|Strongly aromatic, buttery, savory|Soft yellow pulp with spiny inner endocarp|Rice, stew, liqueur|Fat; carotenoids|Never bite into the hard spiny pit; scrape pulp carefully|Vegan`,
    'Grains & Pulses': `Andean lupin (tarwi), debittered|Lupinus mutabilis|Andean; Peruvian; Bolivian|Nutty, mildly bitter|Firm, creamy seed|Salad, stew, flour, snack|Protein; fiber; unsaturated fat|Traditional debittering/leaching is essential; use processed food-grade product|Lupin/legume allergen\nCanary seed, food-grade dehulled|Phalaris canariensis|Canadian; South American cultivation|Mild, nutty|Small pale seeds|Porridge, flour, beverage|Protein; unsaturated fat; fiber|Use food-grade dehulled variety only|Check gluten cross-contact`,
    'Spices & Herbs': `Huacatay (black mint)|Tagetes minuta|Peruvian; Andean|Minty, anise, citrus, basil|Aromatic leaves|Ocopa sauce, potatoes, marinades|Aromatic compounds|Use culinary herb in modest amounts|Vegan\nAchiote (annatto), whole seed|Bixa orellana seed|Mesoamerican; Caribbean; Filipino|Earthy, peppery, mildly floral|Hard red seeds|Oil infusions, marinades, rice|Carotenoid pigment|Infuse in oil; strain seeds|Vegan`,
    'Protein & Dairy': `Alpaca meat|Vicugna pacos|Andean; Peruvian; Bolivian|Lean, mild, slightly sweet|Fine lean muscle|Grill, stew, roast|Protein; iron|Cook thoroughly; lean cuts benefit from brief cooking|Mammalian meat\nCuy (guinea pig), food-grade|Cavia porcellus|Andean; Peruvian; Ecuadorian|Savory, mild|Small carcass; crisp skin when roasted|Roast, stew|Protein; iron|Use inspected food-grade source; cook thoroughly|Mammalian meat`,
    'Pantry & Condiments': `Ají amarillo paste|Capsicum baccatum var. pendulum paste|Peruvian|Fruity, bright, medium heat|Smooth thick paste|Ceviche, sauces, stews|Carotenoids; sodium varies|Packaged heat and salt vary|Vegan\nChicha de jora, cooking beverage|Fermented maize beverage|Peruvian; Andean|Sour, malty, yeasty|Cloudy liquid|Stews, marinades, beverage|Carbohydrate; fermentation-dependent alcohol|Alcohol content and fermentation vary; use food-grade product|May contain alcohol`,
  },
  {
    'Vegetables': `Taro corm, Bun-long cultivar|Colocasia esculenta cultivar|Chinese; Southeast Asian cultivation|Mild, nutty, starchy|Firm, becomes creamy|Stew, chips, poi-style mash|Carbohydrate; fiber; potassium|Cook thoroughly to reduce calcium oxalate irritation|Vegan\nLotus stem, young runner|Nelumbo nucifera young rhizome|Chinese; Japanese|Mild, sweet, vegetal|Crisp, tender|Stir-fry, pickle, soup|Fiber; vitamin C|Wash cavities and cook as desired|Vegan`,
    'Fruit & Citrus': `Kiwano (horned melon)|Cucumis metuliferus|Southern African; cultivated globally|Tart, cucumber-banana, refreshing|Jelly-like seed pulp|Fresh, drinks, salads|Vitamin C; water; fiber|Cut ripe fruit; pulp and seeds are edible|Vegan\nBael, Bengal quince|Aegle marmelos|Indian; South Asian|Aromatic, sweet, resinous|Fibrous pulp in hard shell|Sherbet, preserve, chutney|Fiber; carbohydrates|Crack ripe fruit; strain pulp|Vegan`,
    'Grains & Pulses': `Rice, sticky black (Laotian cultivar)|Oryza sativa glutinous pigmented landrace|Laos; Northern Thailand|Nutty, sweet|Very sticky, chewy|Khao niew dam, desserts|Anthocyanins; fiber; carbohydrate|Soak and steam for traditional texture|Naturally gluten-free\nPea, pigeon pea, toor landrace|Cajanus cajan cultivar|India; Caribbean; East Africa|Nutty, earthy|Creamy when split and cooked|Dal, rice dishes, soups|Protein; fiber; folate|Rinse and cook thoroughly|Legume sensitivity possible`,
    'Spices & Herbs': `Grains of paradise, fermented seed|Aframomum melegueta processed seed|West African|Peppery, citrusy, warm|Aromatic small seeds|Spice blends, stews|Aromatic compounds|Fermentation is a local processing variation; source labeled food-grade|Vegan\nJamaican allspice leaf|Pimenta dioica leaf|Caribbean; Central American|Clove, cinnamon, pepper|Aromatic leaf|Jerk seasoning, soups, infusions|Aromatic compounds|Use fresh or dried leaf; remove after infusion|Vegan`,
    'Protein & Dairy': `Ostrich meat|Struthio camelus|Southern African; global farming|Mild, beef-like, very lean|Fine lean muscle|Steak, mince, kebab|Protein; iron|Cook carefully; lean cuts dry quickly|Bird meat\nEmu meat|Dromaius novaehollandiae|Australian native; farmed|Mild, lean, beef-like|Fine lean muscle|Steak, sausage, stew|Protein; iron|Use inspected food-grade meat; cook thoroughly|Bird meat`,
    'Pantry & Condiments': `Kombu, Rausu type|Saccharina japonica regional product|Hokkaido, Japan|Deep umami, mineral|Thick dried kelp|Dashi, simmered dishes|Iodine; glutamates|Iodine varies; use culinary amounts|Seaweed; iodine-sensitive diets\nKombu, Ma-kombu type|Saccharina japonica regional product|Hokkaido, Japan|Sweet umami, mineral|Broad dried frond|Dashi, tsukudani|Iodine; glutamates|Do not treat as a fixed iodine source; content varies by species and batch|Seaweed`,
  },
  {
    'Vegetables': `Gundruk greens, fermented mustard leaf|Brassica juncea fermented leaves|Nepali; Himalayan|Sour, earthy, savory|Soft fermented leaves|Soup, pickle, side dish|Fiber; sodium varies|Fermentation and salt vary; cook if desired|Vegan\nAchu (cocoyam) leaf|Xanthosoma sagittifolium leaf|Cameroonian; Central African|Green, mild|Leaves soften into thick sauce|Achu soup, stews|Fiber; micronutrients vary|Cook thoroughly; never eat raw|Vegan`,
    'Fruit & Citrus': `Rambai|Baccaurea motleyana|Borneo; Malaysian; Indonesian|Sweet-sour, aromatic|Segmented juicy flesh|Fresh, preserves, curries|Vitamin C; fiber|Peel; remove seeds|Vegan\nSalak, pondoh type|Salacca zalacca cultivar|Yogyakarta, Indonesia|Sweet, tangy, aromatic|Crisp, dry, segmented|Fresh, sweets, pickles|Fiber; carbohydrate|Peel scaly skin; seed not eaten|Vegan`,
    'Grains & Pulses': `Rice, Kalanamak|Oryza sativa aromatic landrace|Uttar Pradesh, India|Floral, pandan-like|Short-medium grain; cohesive|Rice meals, kheer|Carbohydrate; aroma compounds|Rinse gently; source region/variety claims need traceability|Naturally gluten-free\nMoth bean|Vigna aconitifolia|Rajasthan; South Asian|Earthy, nutty|Small pulse; sprouts readily|Matki usal, sprouts, dal|Protein; fiber; iron|Soak or sprout; cook thoroughly|Legume sensitivity possible`,
    'Spices & Herbs': `Dried lime leaf, makrut|Citrus hystrix leaf|Thai; Southeast Asian|Floral, citrus, resinous|Dried aromatic leaf|Curry paste, soup, infusions|Aromatic compounds|Rehydrate or infuse; discard tough leaf|Vegan`,
    'Protein & Dairy': `Buffalo mozzarella di bufala|Water buffalo milk pasta filata cheese|Campania, Italy|Milky, tangy, rich|Soft, moist, elastic|Caprese, pizza, pasta|Protein; calcium; fat|Keep chilled; moisture differs by aging|Dairy\nKashkaval, Bulgarian kashkaval|Sheep/cow milk aged cheese|Bulgaria; Balkans|Salty, buttery, tangy|Firm, sliceable, melts|Grilled, baked, sandwiches|Protein; calcium; sodium|Milk species and aging vary|Dairy`,
    'Pantry & Condiments': `Douchi, Yangjiang fermented black beans|Fermented soybean|Guangdong, China|Salty, pungent, umami|Small dark beans|Steamed fish, stir-fry, sauces|Protein; sodium|Rinse or use sparingly; salt varies|Soy\nPixian doubanjiang|Fermented broad bean and chili paste|Sichuan, China|Hot, salty, fermented umami|Coarse red-brown paste|Mapo tofu, braises, sauces|Protein; sodium|Salt and wheat ingredients vary by brand|Legume; may contain wheat`,
  },
  {
    'Vegetables': `Radicchio di Chioggia|Cichorium intybus cultivar|Veneto, Italy|Bitter, crisp, mildly sweet when cooked|Compact red-purple head|Salad, grill, risotto|Fiber; vitamin K; folate|Grill or roast to mellow bitterness|Vegan\nCima di rapa (turnip greens, Italian type)|Brassica rapa subsp. sylvestris|Southern Italian; Apulian|Bitter, mustardy, pungent|Leafy with tender flowering shoots|Orecchiette, sauté, soup|Vitamins A, C, K; folate|Blanch if desired to reduce bitterness|Vegan`,
    'Fruit & Citrus': `Finger lime, red champagne cultivar|Citrus australasica cultivar|Australian native; Queensland|Tart, floral|Pink vesicles, crisp pop|Seafood, desserts, drinks|Vitamin C; fiber|Use food-grade fruit; color varies by cultivar|Vegan\nMango, Alphonso cultivar|Mangifera indica cultivar|Ratnagiri and Devgad, India|Rich sweet, floral, resinous|Smooth fiber-light flesh|Fresh, aamras, desserts|Vitamin C; carotenoids; fiber|Ripeness and origin determine aroma|Vegan`,
    'Grains & Pulses': `Rice, Sona Masuri|Oryza sativa medium-grain cultivar|Andhra Pradesh; Telangana, India|Mild, lightly aromatic|Medium slender grains, soft|Everyday rice, idli, dosa|Carbohydrate; small amounts of B vitamins|Rinse and cook with appropriate water ratio|Naturally gluten-free\nRice, Gobindobhog|Oryza sativa aromatic landrace|West Bengal, India|Fragrant, buttery, sweet|Short-grain, soft and cohesive|Bhog, pulao, kheer|Carbohydrate; aroma compounds|Gentle rinse preserves aroma|Naturally gluten-free`,
    'Spices & Herbs': `Green cardamom, Alleppey type|Elettaria cardamomum cultivar/grade|Kerala, India|Floral, sweet, eucalyptus-like|Green aromatic pods|Chai, biryani, desserts|Aromatic compounds|Quality depends on harvest and curing|Vegan\nBlack pepper, Tellicherry garbled extra bold|Piper nigrum grade|Kerala, India|Hot, woody, citrusy|Large whole peppercorns|Grind for finishing, marinades|Piperine and aromatic compounds|Trade grade is size/quality descriptor, not separate species|Vegan`,
    'Protein & Dairy': `Paneer, buffalo milk|Acid-set buffalo milk cheese|South Asian; Indian|Mild, milky|Firm, holds shape in cooking|Palak paneer, tikka, curry|Protein; calcium; fat varies|Refrigerate; fresh cheese is perishable|Dairy\nChicken, country hen (desi type)|Gallus gallus domesticus|South Asian; Indian|Savory, more pronounced poultry flavor|Firm, leaner texture often|Curry, roast, stew|Protein; B vitamins; iron|Cook to safe internal temperature; breed/age affect texture|Poultry`,
    'Pantry & Condiments': `Kashmiri chili paste|Capsicum annuum paste|Kashmir, India|Mild to moderate heat, fruity|Smooth red paste|Rogan josh, marinades, sauces|Carotenoids; sodium varies|Commercial and homemade salt/oil differ|Vegan\nKokum agal|Garcinia indica concentrated extract|Konkan, India|Sour, fruity|Dark liquid concentrate|Sol kadhi, curries, drinks|Organic acids|Dilute to taste; product concentration varies|Vegan`,
  },
);
for (const [roundIndex, additions] of refreshRounds.entries()) {
  for (const [sheetName, block] of Object.entries(additions)) {
    const existing = new Set(data[sheetName].map(row => row[0].trim().toLocaleLowerCase()));
    const fresh = parse(block).filter(row => !existing.has(row[0].trim().toLocaleLowerCase()));
    data[sheetName].push(...fresh);
  }
  console.log(`Completed new expansion round ${roundIndex + 1} of ${refreshRounds.length}`);
}

const specialtyRows = parse(`
Maitake mushroom|Edible fungus|Japan; East Asia|Earthy, peppery, rich umami|Frilly clusters, crisp edges|Roast, saute, soup, grain dishes|B vitamins; fiber; minerals|Tear into clusters and sear|Vegan; mushroom sensitivity possible
Enoki mushroom|Edible fungus|East Asia|Mild, delicate|Thin crisp-tender strands|Hot pot, soup, noodle dish|B vitamins; fiber|Trim root base and cook thoroughly|Vegan; mushroom sensitivity possible
Shimeji mushroom|Edible fungus|Japan; East Asia|Nutty, mildly bitter raw, savory cooked|Firm small caps|Stir-fry, soup, rice, pasta|B vitamins; fiber|Cook before eating; raw bitterness softens|Vegan; mushroom sensitivity possible
Beech mushroom / buna-shimeji|Edible fungus|Japan; East Asia|Nutty, savory|Small clustered caps|Soup, saute, noodle dishes|B vitamins; fiber|Cook thoroughly|Vegan; mushroom sensitivity possible
Nameko mushroom|Edible fungus|Japan|Mild, earthy|Small caps with natural gelatinous coating|Miso soup, soba, rice|B vitamins; fiber|Rinse lightly if desired; cook|Vegan; mushroom sensitivity possible
Wood ear mushroom|Edible fungus|China; East Asia|Mild, absorbs seasoning|Crunchy, gelatinous|Stir-fry, soup, salad, dumplings|Fiber; iron|Rehydrate dried product fully and cook|Vegan; mushroom sensitivity possible
Cloud ear fungus|Edible fungus|East Asia|Mild, earthy|Thin, crisp gelatinous texture|Soup, stir-fry, salad|Fiber; minerals|Often sold dried; rehydrate and cook|Vegan; mushroom sensitivity possible
Tremella / snow fungus|Edible fungus|China; East Asia|Mild, delicate|Soft, gelatinous|Sweet soup, dessert, stew|Fiber; polysaccharides|Soak dried fungus and remove hard base|Vegan; mushroom sensitivity possible
Lion's mane mushroom|Edible fungus|East Asia; North American cultivation|Mild, seafood-like|Shreddable, tender|Saute, roast, soup, plant-based filling|Fiber; protein; micronutrients vary|Cook thoroughly; food use differs from supplements|Vegan; mushroom sensitivity possible
Black trumpet mushroom|Edible fungus|Europe; North America|Smoky, earthy, rich|Thin, delicate funnel|Pasta, soup, sauce, dried seasoning|Fiber; minerals|Wild identification should be by an expert|Vegan; mushroom sensitivity possible
Hedgehog mushroom|Edible fungus|Europe; North America|Nutty, peppery|Firm, tooth-like underside|Saute, roast, soup|Fiber; minerals|Wild foraging requires expert identification|Vegan; mushroom sensitivity possible
Chicken of the woods|Edible fungus|Europe; North America|Savory, chicken-like|Layered, tender when young|Saute, fry, stew, sandwich|Protein; fiber; micronutrients vary|Only eat positively identified, young, properly cooked specimens|Vegan; mushroom sensitivity possible
Hen of the woods|Edible fungus|Japan; North America|Earthy, nutty|Layered, firm fronds|Roast, saute, soup, rice|Fiber; B vitamins|Clean crevices and cook thoroughly|Vegan; mushroom sensitivity possible
Truffle, black|Edible fungus|France; Italy; Spain|Earthy, cocoa, forest aroma|Firm aromatic fungus|Pasta, egg, butter, sauce|Small culinary portions|Species and aroma vary; use sparingly|Vegan; mushroom sensitivity possible
Truffle, white|Edible fungus|Piedmont, Italy; Balkans|Garlic, cheese, honey, intense aroma|Firm aromatic fungus|Egg, pasta, risotto, potato|Small culinary portions|Usually shaved raw over warm food|Vegan; mushroom sensitivity possible
Matsutake mushroom|Edible fungus|Japan; Korea; China; Himalayas|Piney, spicy, earthy|Firm, aromatic|Rice, soup, grill|Fiber; minerals|Wild species trade and identification vary|Vegan; mushroom sensitivity possible
Sea lettuce / Ulva|Edible seaweed|Coastal regions worldwide|Briny, green|Thin, tender sheet|Soup, salad, seasoning|Iodine; minerals|Species, salinity and iodine vary; use food-grade product|Vegan; iodine varies
Dulse|Edible red seaweed|North Atlantic; Pacific|Salty, smoky, mineral|Soft dried fronds|Snack, soup, seasoning, bread|Iodine; iron; potassium|Dried and flaked forms differ|Vegan; iodine varies
Nori, aonori flakes|Edible seaweed|Japan|Roasted, briny, delicate umami|Fine flakes|Okonomiyaki, noodles, rice|Iodine; minerals|Use as finishing seasoning|Vegan; iodine varies
Hijiki|Edible seaweed|Japan; Korea; China|Mineral, briny|Thin wiry strands|Simmered side dish, salad|Fiber; minerals|Inorganic arsenic levels vary; follow local food authority advice and product instructions|Vegan
Arame|Edible seaweed|Japan|Mild, sweet, mineral|Thin wiry strands|Salad, vegetable side, rice|Fiber; iodine|Soak and cook per package|Vegan; iodine varies
Irish moss / carrageen moss|Edible red seaweed|Ireland; Caribbean|Mild, marine|Gel-forming seaweed|Pudding, drink, thickener|Fiber; minerals|Food-grade processing and species matter|Vegan
Sea grapes / umibudo|Edible seaweed|Okinawa; Southeast Asia|Briny, fresh|Small bursting beads|Salad, rice, seafood|Minerals; iodine|Keep chilled; brief rinse immediately before serving|Vegan; iodine varies
Hijiki, young shoots|Edible seaweed form|Japan; Korea|Mineral, mildly sweet|Tender strands|Simmered side dish|Fiber; minerals|Same arsenic consideration as dried hijiki products|Vegan
Samphire / sea asparagus|Halophyte vegetable|Europe; Mediterranean; coastal regions|Salty, green, crisp|Juicy succulent stems|Seafood, salad, pickle, saute|Fiber; minerals|Rinse well; salinity varies|Vegan
Sea purslane|Halophyte vegetable|Mediterranean; coastal Europe|Salty, tangy, succulent|Crisp leaves|Salad, garnish, pickle|Fiber; minerals|Use food-grade cultivated/identified plants|Vegan
Elderflower|Edible flower|Europe; West Asia|Floral, honeyed, muscat-like|Delicate flower clusters|Syrup, fritter, drink, dessert|Negligible nutrition per serving|Use correctly identified flowers; avoid stems/leaves|Vegan
Hibiscus calyx / roselle|Edible calyx|West Africa; Caribbean; South Asia|Tart, cranberry-like|Fleshy red calyx|Drink, jam, sauce, tea|Vitamin C; anthocyanins|Dried calyx is concentrated and tart|Vegan
Squash blossom|Edible flower|Mesoamerica; Mediterranean|Mild, sweet, vegetal|Delicate petals|Stuff, fry, quesadilla, pasta|Small amounts of vitamin C|Use fresh, remove inner stamens if desired|Vegan
Nasturtium flower|Edible flower|Andean heritage; global gardens|Peppery, watercress-like|Tender petals and leaves|Salad, garnish, compound butter|Vitamin C; carotenoids|Use only pesticide-free food-grade flowers|Vegan
Borage flower|Edible flower|Mediterranean; European|Cool cucumber-like|Soft blue star-shaped petals|Drink, salad, garnish|Small culinary portions|Use only food-grade flowers; avoid excess consumption|Vegan
Daylily buds|Edible flower bud|East Asian|Mild, sweet, vegetal|Tender, crisp-tender|Soup, stir-fry, dried in hot pot|Fiber; carotenoids|Use food-grade Hemerocallis; never substitute toxic lilies|Vegan
Banana flower|Edible flower bud|South/Southeast Asia|Mild, artichoke-like|Layered and tender|Salad, curry, fritter|Fiber; potassium|Prepare promptly; soak cut pieces to limit browning|Vegan
Purslane seed|Edible seed|Mediterranean; West Asian|Mild, nutty|Tiny crisp seed|Flatbread, porridge, garnish|Fiber; minerals|Use food-grade seed; not common in retail|Vegan
Ant larvae / escamoles|Edible insect|Central Mexico|Nutty, buttery|Soft, granular|Taco, omelet, saute|Protein; fat; minerals|Seasonal specialty; source from regulated suppliers|Insect allergy and possible cross-reactivity
Weaver ant eggs|Edible insect|Southeast Asia|Sour, tangy, citrus-like|Soft eggs and larvae|Salad, soup, omelet|Protein; fat; micronutrients vary|Traditional food; use food-grade supply and cook appropriately|Insect allergy and possible cross-reactivity
Silkworm pupae|Edible insect|Korea; China; Thailand|Nutty, earthy|Soft, chewy|Snack, stew, garnish|Protein; fat; minerals|Cooked or canned products are common; follow label|Insect allergy and possible cross-reactivity
Grasshopper / chapulines|Edible insect|Mexico; Central America|Toasted, nutty, savory|Crisp dried insect|Snack, taco, salsa, garnish|Protein; iron; fiber/chitin|Use only food-grade products; allergens possible|Insect allergy and possible crustacean cross-reactivity
Black soldier fly larvae ingredient|Edible insect ingredient|Global regulated food production|Nutty, mild|Whole dried or milled|Animal feed and some regulated human foods|Protein; fat; minerals vary|Human-food approval varies by jurisdiction; verify legal status|Insect allergy; possible shellfish cross-reactivity
Bee pollen|Bee-collected food|Global|Floral, sweet, waxy|Granules|Smoothie, yogurt, garnish|Protein; micronutrients vary|Allergy risk can be serious; use food-grade product|Potential bee/pollen allergy
Royal jelly|Bee-derived food|Global|Tart, pungent, creamy|Viscous secretion|Supplement/food ingredient|Composition varies|Not a conventional cooking ingredient; avoid unsupported health claims|Potential bee allergy
Edible lichen / reindeer moss|Lichen|Arctic and sub-Arctic traditional foods|Earthy, mineral|Dried, rehydrated|Traditional dishes and specialty garnish|Carbohydrate; minerals vary|Only specially prepared food-grade species; many lichens are unsafe raw|Use expert identification and proper processing
`);
specialtyRows.push(...parse(`
Geoduck clam|Mollusk seafood|Pacific Northwest, North America|Sweet, clean, oceanic|Crisp siphon; tender body|Crudo, sashimi, chowder, stir-fry|Protein; minerals|Keep cold; separate parts have different textures|Mollusk allergen
Jellyfish, edible prepared|Marine animal|East and Southeast Asia|Mild, saline|Crunchy, crisp|Salad, cold appetizer, pickle|Protein; sodium varies|Use food-grade processed product; rinse per label|Seafood allergy possible
Whelk|Marine gastropod|Europe; East Asia|Briny, mildly sweet|Firm, chewy|Salad, stew, braise, skewer|Protein; iron; B12|Cook to tender; species and prep vary|Mollusk allergen
Conch|Marine gastropod|Caribbean; Atlantic tropics|Mildly sweet, oceanic|Firm, needs tenderizing|Fritter, salad, stew, chowder|Protein; minerals|Respect local harvest rules; tenderize and cook|Mollusk allergen
Smelt, freshwater|Small fish|Northern Europe; Great Lakes; East Asia|Mild, delicate, slightly sweet|Small, tender bones|Fry, grill, pickle, smoke|Protein; omega-3; calcium if eaten whole|Cook promptly; bones may be eaten when small and tender|Fish allergen
Sturgeon roe / caviar|Fish roe|Caspian region; farmed globally|Salty, buttery, marine|Delicate popping eggs|Garnish, blini, egg, pasta|Protein; omega-3; sodium|Keep very cold; verify species and responsible source|Fish allergen
Emu meat|Game bird meat|Australia; farmed elsewhere|Lean, beef-like, mild game note|Fine-grained, low fat|Steak, mince, sausage|Protein; iron; B vitamins|Cook lean cuts carefully to avoid drying|Animal product
Ostrich meat|Game bird meat|Southern Africa; farmed globally|Lean, beef-like, slightly sweet|Fine-grained, low fat|Steak, mince, kebab|Protein; iron; B vitamins|Cook safely; avoid overcooking lean cuts|Animal product
`));

specialtyRows.push(...parse(`
Sago grubs|Edible insect larva|Papua New Guinea; Indonesia|Rich, nutty, buttery|Soft, creamy|Roast, grill, sago palm dishes|Protein; fat; minerals|Traditional food; use regulated food-grade supply and cook fully|Insect allergy; possible cross-reactivity
Giant water bug / maeng da|Edible insect|Thailand; Southeast Asia|Pungent, floral, savory|Firm, aromatic|Nam prik, seasoning paste, snack|Protein; minerals|Use food-grade product; species and processing matter|Insect allergy; possible cross-reactivity
Palm weevil larvae|Edible insect larva|West Africa; Latin America; Southeast Asia|Nutty, rich, buttery|Soft, fatty|Roast, grill, stew|Protein; fat; minerals|Only regulated food-grade supply; cook thoroughly|Insect allergy; possible cross-reactivity
Termite alates|Edible insect|Africa; South Asia; Latin America|Nutty, earthy|Crisp roasted insect|Snack, flour, relish|Protein; fat; minerals|Seasonal traditional food; identify and source safely|Insect allergy; possible cross-reactivity
Alligator meat|Reptile meat|Southern United States|Mild, chicken-like, lightly gamey|Lean, firm|Fry, grill, stew, sausage|Protein; low fat; minerals|Use inspected farmed supply; cook thoroughly|Animal product
Sea squirt / tunicate|Marine invertebrate|Korea; Japan; Chile|Briny, iodine-rich, pungent|Soft, gelatinous|Raw seafood, rice, stew|Protein; minerals|Use fresh, regulated product; flavor is strong|Seafood allergy possible
Sea snail, turban shell|Marine gastropod|Japan; Korea; Pacific|Sweet, briny, mildly bitter|Firm, chewy|Sashimi, braise, stew|Protein; minerals|Remove hard parts as instructed; cook safely|Mollusk allergen
Freshwater mussel|Bivalve seafood|Global freshwater regions|Earthy, briny, mineral|Tender to chewy|Stew, soup, smoked foods|Protein; iron; B12|Use approved waters and inspected sources; cook fully|Mollusk allergen
Cockle|Bivalve seafood|North Atlantic; East Asia|Sweet, briny|Small, tender, chewy|Pasta, soup, rice, grill|Protein; iron; B12|Purge grit and cook until shells open|Mollusk allergen
Periwinkle / winkles|Marine gastropod|North Atlantic; West Africa|Salty, sea-sweet|Small, firm|Snack, stew, seafood plate|Protein; iron|Cook fully; species and collection rules vary|Mollusk allergen
Sea beet / wild coastal beet greens|Coastal wild green|Europe; Mediterranean|Mineral, salty, spinach-like|Tender leaf, firm stem|Saute, soup, side dish|Fiber; minerals|Use correctly identified food-grade plants; rinse sand and salt|Vegan
Ice plant / crystalline herb|Succulent vegetable|South Africa; Mediterranean cultivation|Salty, crisp, mildly tart|Watery, crunchy leaves|Salad, garnish, pickle|Water; minerals|Use cultivated food-grade plants; rinse lightly|Vegan
Monk's beard / agretti|Coastal vegetable|Italy; Mediterranean|Salty, grassy, gently sour|Succulent thin stems|Pasta, saute, salad, seafood|Fiber; minerals|Trim roots and blanch briefly|Vegan
`));

specialtyRows.push(...parse(`
Crayfish / crawfish|Freshwater crustacean|Southern United States; Europe; East Asia|Sweet, briny, lightly earthy|Tender tail meat|Boil, stew, bisque, rice|Protein; B12; minerals|Use regulated harvest/farmed supply; cook thoroughly|Crustacean shellfish allergen
Freshwater prawn, giant river type|Freshwater crustacean|South and Southeast Asia|Sweet, delicate, oceanic|Firm, succulent flesh|Curry, grill, soup, rice|Protein; selenium; B12|Source legally and cook thoroughly|Crustacean shellfish allergen
Mud crab|Crustacean seafood|South and Southeast Asia; Pacific|Rich, sweet, briny|Firm claw and body meat|Curry, chili crab, steamed dishes|Protein; zinc; B12|Keep chilled and cook fully|Crustacean shellfish allergen
Snow crab|Crustacean seafood|North Pacific; North Atlantic|Sweet, clean, briny|Flaky leg meat|Steamed, soup, salad|Protein; selenium; B12|Thaw and heat safely if pre-cooked|Crustacean shellfish allergen
Langoustine / Norway lobster|Crustacean seafood|North Atlantic; Mediterranean|Sweet, delicate, lobster-like|Tender, fine-grained|Grill, pasta, bisque|Protein; selenium; B12|Highly perishable; cook briefly|Crustacean shellfish allergen
Sea pineapple / Korean sea squirt|Marine invertebrate|Korea; Japan|Briny, pungent, sweet finish|Soft, crisp outer tunic|Rice, seafood, raw preparations|Protein; minerals|Use very fresh regulated product and keep cold|Seafood allergy possible
Sea fig / anemone, prepared|Marine invertebrate|Spain; Mediterranean|Iodine-rich, briny, marine|Soft, gelatinous|Fried fritter, seafood dishes|Protein; minerals|Regional specialty; use only legally marketed prepared product|Seafood allergy possible
Purple laver / gim|Edible seaweed|Korea|Roasted, savory, briny|Thin crisp sheet|Rice, soup, snack, garnish|Iodine; minerals|Seasoned and unseasoned forms differ in salt and oil|Vegan; iodine varies
Mozuku seaweed|Edible seaweed|Okinawa, Japan|Mild, tart, marine|Fine slippery strands|Vinegar side dish, soup|Fiber; iodine|Salted and fresh products require different preparation|Vegan; iodine varies
Wakame stem / mekabu|Edible seaweed part|Japan; Korea|Briny, mineral, mild umami|Slick, crunchy fronds|Soup, vinegar salad, rice|Fiber; iodine|Rinse salted product; iodine content varies|Vegan; iodine varies
`));

specialtyRows.push(...parse(`
Abalone, canned dried style|Mollusk seafood product|East Asia; Pacific|Briny, concentrated, sweet umami|Dense, chewy|Soup, braise, banquet dishes|Protein; sodium varies|Soak or rinse according to product instructions|Mollusk allergen
Sea squirt, salted fermented|Marine invertebrate product|Korea; Japan|Salty, pungent, marine umami|Soft, intensely flavored|Rice, side dish, stew|Protein; sodium|Use commercially prepared product and refrigerate|Seafood allergy possible
Sea urchin gonad, frozen|Echinoderm seafood product|Japan; Chile; Mediterranean|Briny, sweet, creamy|Soft, delicate|Pasta, rice, sauce|Protein; omega-3; sodium|Thaw under refrigeration; use promptly|Seafood allergy possible
Razor clam, dried|Bivalve seafood product|China; East Asia|Concentrated sweet, briny, umami|Chewy after rehydration|Soup, congee, stir-fry|Protein; sodium|Rehydrate and cook fully|Mollusk allergen
Whitebait|Small fish|New Zealand; Mediterranean; East Asia|Mild, oceanic|Tiny, tender fish|Fritter, omelet, pasta, salad|Protein; calcium; omega-3|Use legal sustainable catch and cook thoroughly|Fish allergen
Eel, freshwater|Freshwater fish|Japan; China; Europe|Rich, sweet, savory|Silky, fatty flesh|Grill, rice, stew|Protein; omega-3; B12|Buy inspected, responsibly sourced product; cook fully|Fish allergen
Smoked eel|Smoked freshwater fish|Northern Europe; Japan|Smoky, rich, sweet-savory|Tender, oily|Sandwich, salad, rice, appetizer|Protein; omega-3; sodium|Ready-to-eat handling depends on package; refrigerate|Fish allergen
Sea palm kelp|Edible brown seaweed|Pacific coast, North America|Mineral, salty, umami|Crunchy stalk|Pickle, salad, garnish|Iodine; minerals|Use food-grade product; iodine varies|Vegan
Gracilaria seaweed|Edible red seaweed|Southeast Asia; Caribbean|Mild, marine|Firm, crisp strands|Salad, soup, agar source|Fiber; minerals|Species and processing vary; use food-grade product|Vegan
Caulerpa seaweed|Edible green seaweed|Southeast Asia; Pacific|Salty, fresh, oceanic|Tiny bead-like branches|Salad, rice, seafood|Minerals; iodine varies|Use only approved edible species and food-grade supply|Vegan
Wild garlic mustard greens|Edible leafy plant|Europe; North America|Garlicky, peppery, bitter|Tender leaf; fibrous older stems|Pesto, greens, soup|Vitamin C; fiber|Correct identification and food-grade foraging are essential|Vegan
Nettles, cooked leaves|Leafy green|Europe; West Asia; North America|Green, spinach-like, mineral|Tender after cooking|Soup, puree, pasta, filling|Iron; calcium; vitamin C|Handle with gloves; cook thoroughly to neutralize stinging hairs|Vegan
`));

specialtyRows.push(...parse(`
Fish milt / shirako|Fish organ ingredient|Japan; Northern Europe|Creamy, briny, mild umami|Soft, custard-like|Hot pot, tempura, sushi-style dishes|Protein; omega-3; minerals|Use very fresh regulated fish and cook appropriately|Fish allergen
Cod liver|Fish organ ingredient|Northern Europe; North Atlantic|Rich, marine, savory|Soft, oily|Smoked spread, preserve, pate|Vitamin A; vitamin D; omega-3|Portion carefully due to high vitamin A; product salt varies|Fish allergen
Chicken feet|Poultry ingredient|East Asia; Caribbean; West Africa|Mild, savory|Gelatin-rich, collagenous|Soup, braise, stock|Protein; collagen; minerals|Clean carefully and simmer until tender|Animal product
Beef tendon|Meat connective tissue|East Asia; Southeast Asia|Mild, absorbs broth|Gelatinous, chewy|Soup, braise, noodle dish|Protein; collagen|Long simmer until tender; source inspected product|Animal product
Pork blood curd|Blood ingredient|China; Vietnam; Philippines; Europe|Iron-rich, savory, mild|Soft, custard-like|Soup, stew, sausage|Protein; iron|Use inspected source and cook thoroughly|Animal product
Duck gizzard|Poultry organ meat|East Asia; Europe; West Africa|Deep savory, mineral|Firm, chewy|Braise, grill, stew, confit|Protein; iron; B12|Clean and slow-cook or slice thinly|Animal product
Beef tripe, honeycomb|Ruminant stomach|West Africa; East Asia; Mediterranean|Mild, absorbs sauce|Springy, porous|Soup, stew, braise|Protein; B12; minerals|Clean well and simmer until tender|Animal product
Fermented soybean curd, red preserved|Fermented soybean product|China; East Asia|Salty, creamy, pungent|Soft cubes in brine|Rice, greens, marinade, sauce|Protein; sodium|Use small amounts; fermentation and salt differ|Soy allergen
Sago flour|Palm starch|Papua New Guinea; Indonesia; Pacific|Neutral, mildly sweet|Fine starch; elastic when cooked|Porridge, pancake, dumpling, pudding|Mostly carbohydrate|Hydrate and cook thoroughly|Vegan; naturally gluten-free
`));

specialtyRows.push(...parse(`
Termite mushrooms / Termitomyces|Edible fungus|Africa; South Asia; Southeast Asia|Savory, earthy, rich umami|Firm cap and stem|Stew, grill, rice, soup|Protein; fiber; minerals|Wild identification by expert only; use food-grade supply|Vegan; mushroom sensitivity possible
Bamboo fungus / veiled lady mushroom|Edible fungus|China; Southeast Asia|Mild, absorbs broth|Spongy, delicate|Soup, banquet dishes, braise|Fiber; minerals|Use cultivated food-grade fungus; rehydrate and cook|Vegan; mushroom sensitivity possible
Termite queen, food-grade|Edible insect|Central Africa; Southeast Asia|Rich, fatty, nutty|Soft-bodied|Traditional seasonal food|Protein; fat; minerals|Highly specialized food; source regulated products and cook appropriately|Insect allergy; possible cross-reactivity
Crab roe|Crustacean ingredient|East Asia; Europe; North America|Briny, rich, sweet umami|Soft, granular|Sauce, soup, rice, spread|Protein; cholesterol; minerals|Use cooked, regulated seafood product|Crustacean shellfish allergen
Lobster tomalley|Crustacean organ ingredient|North Atlantic|Rich, marine, creamy|Soft paste|Sauce, bisque, stuffing|Protein; fat; minerals|Food safety advisories and harvest conditions vary; use regulated products|Crustacean shellfish allergen
Fish maw / swim bladder|Fish collagen ingredient|China; Southeast Asia|Mild, absorbs broth|Gelatinous, spongy|Soup, braise, banquet dishes|Protein; collagen|Dried forms require soaking and long cooking|Fish allergen
Pandanus fruit keys|Fruit/seed segment|Pacific Islands; Southeast Asia|Nutty, sweet, mildly resinous|Fibrous segments|Traditional chew, food coloring, regional dishes|Carbohydrate; fiber|Use known food-grade species and local preparation|Vegan
Nipa palm fruit|Palm fruit|Southeast Asia|Mild, sweet, jelly-like|Translucent soft flesh|Dessert, drink, preserve|Carbohydrate; water|Use ripe edible fruit from food-grade source|Vegan
Palmyra sprout / ice apple|Palm seed product|South Asia; Southeast Asia|Mild, cooling, subtly sweet|Tender translucent jelly|Fresh snack, dessert, drink|Water; carbohydrate|Highly perishable; use fresh food-grade product|Vegan
Sago grub palm pith|Palm starch ingredient|Pacific Islands; Indonesia|Neutral, lightly sweet|Fine starch or moist paste|Porridge, pancake, pudding|Mostly carbohydrate|Processing and species vary; use food-grade sago|Vegan; naturally gluten-free
`));

specialtyRows.push(...parse(`
Edible acorn flour|Wild seed flour|Europe; Mediterranean; Indigenous North America|Nutty, tannic, earthy|Fine flour or meal|Bread, porridge, thickener|Carbohydrate; fiber; minerals|Acorns require leaching to remove tannins; use properly processed flour|Vegan
Pine pollen|Tree pollen|East Asia; North America|Sweet, resinous, floral|Fine yellow powder|Baking, porridge, drink|Protein; micronutrients vary|Use food-grade product; pollen allergy may occur|Potential pollen allergen
Bee brood, food-grade|Insect brood|Southeast Asia; Africa|Nutty, creamy|Soft larvae and pupae|Roast, soup, omelet|Protein; fat; minerals|Use inspected food-grade source and cook thoroughly|Insect allergy; possible cross-reactivity
Bamboo rat meat, farmed|Rodent meat|Southeast Asia|Mild, pork-like, savory|Lean, firm meat|Stew, braise, roast|Protein; iron|Species laws and zoonotic controls vary; use only regulated farmed products|Animal product
Crocodile meat, farmed|Reptile meat|Southeast Asia; Southern Africa|Mild, chicken-like, lightly fishy|Lean, firm|Grill, curry, fry|Protein; low fat|Use inspected farmed supply and cook thoroughly|Animal product
Fermented palm sap / toddy|Fermented plant beverage|South Asia; Africa; Pacific|Sweet-sour, yeasty, lightly alcoholic|Thin cloudy beverage|Drink, cooking, fermentation base|Sugars; alcohol varies|Fermentation begins quickly; use regulated food-grade product|May contain alcohol
Raw cacao pulp|Cacao fruit pulp|West Africa; Latin America|Sweet-sour, floral, tropical|Juicy white pulp|Drink, dessert, fresh fruit|Carbohydrate; fiber|Fresh pulp is perishable; source food-grade cacao pods|Vegan
Carob pod pulp|Tree pod|Mediterranean; West Asia|Sweet, roasted, cocoa-like|Dry pod or syrup|Powder, syrup, baking|Fiber; sugars|Seeds are very hard; use processed food-grade products|Vegan
Kelp flour|Seaweed ingredient|Japan; Korea; North Atlantic|Mineral, briny, umami|Fine green-brown powder|Seasoning, noodles, broth|Iodine; minerals|Iodine levels can be high and vary widely; use modestly|Vegan
Edible bamboo fungus eggs / immature sporocarp|Edible fungus|East Asia|Mild, earthy|Soft, gelatinous|Soup, braise|Fiber; minerals|Only cultivated, positively identified edible species|Vegan; mushroom sensitivity possible
`));

const tasteRows = [
['Flavor lever','What it contributes','Common ingredients','Balance / pairing cue'],
['Salty','Enhances flavor; moderates bitterness','Salt, soy sauce, miso, cheese, olives, capers','Add in stages; salty ingredients also add umami'],
['Sour / acidic','Brightness; cuts richness','Lemon, lime, vinegar, tamarind, yogurt, tomato','Add to taste near the end; balance with fat or sweetness'],
['Sweet','Rounds heat and acidity; supports browning','Onion, carrot, fruit, honey, sugar, coconut milk','Use enough to balance; avoid masking savory depth'],
['Bitter','Complexity and contrast','Kale, radicchio, coffee, cocoa, fenugreek, citrus pith','Pair with fat, salt, sweet or acid'],
['Umami / savory','Brothy depth and persistence','Mushrooms, tomato, aged cheese, soy, fish sauce, meat','Layer small amounts; reduce sodium elsewhere'],
['Heat / pungency','Warmth or sharp aromatic lift','Chili, pepper, ginger, mustard, raw garlic','Control by dose, seed/pith removal and cooking time'],
['Fat / richness','Body, aroma carrying, mouthfeel','Oil, butter, cream, nuts, avocado, coconut','Brighten rich dishes with acid, herbs or crisp texture'],
['Herbal / floral','Freshness, perfume and lift','Basil, cilantro, mint, parsley, cardamom, saffron','Fresh herbs often go late; dried herbs can cook longer'],
['Toast / smoke','Roasted depth','Toasted spices, smoked paprika, grill, browned butter','Use a small amount; pair with acid and fresh elements'],
['Texture contrast','Crunch, creaminess, chew or crispness','Seeds, nuts, croutons, legumes, grains, pickles','Build at least two textures in composed plates'],
['Aromatic base','Foundation for layered savory dishes','Onion, garlic, ginger, celery, carrot, leek','Sweat or brown appropriately before liquid ingredients']
];
tasteRows.push(...parse(`
Citrus acidity|Bright, fragrant sourness|Lemon, lime, yuzu, calamansi, bergamot|Zest adds aroma; juice adds acidity and moisture
Vinegar acidity|Sharp, clean sourness|Rice, sherry, wine, malt, balsamic vinegar|Different vinegars vary in sweetness, aroma and intensity
Fermented sourness|Tangy, savory depth|Yogurt, kimchi, sauerkraut, sourdough, fermented pickles|Fermentation adds more than acidity; salt and texture also change
Berry tartness|Fruity, bright acidity|Cranberry, currant, raspberry, sour cherry|Pair with sugar or fat when acidity is strong
Green bitterness|Fresh, peppery, leafy bitter|Arugula, radicchio, mustard greens, watercress|Fat, salt, sweet fruit or acid can soften perceived bitterness
Root earthiness|Soil-like, sweet, mineral|Beet, celeriac, parsnip, mushroom, lentils|Roasting adds sweetness; fresh herbs and acid lift the finish
Smoky flavor|Wood smoke, char or cured depth|Chipotle, smoked paprika, smoked fish, grilled eggplant|Use a small amount; balance with citrus or fresh herbs
Floral aroma|Perfumed, delicate, sweet|Saffron, cardamom, rose, orange blossom, lychee|Heat can flatten delicate floral notes; add late where suitable
Anise / licorice|Sweet herbal spice|Fennel, anise, star anise, tarragon, Thai basil|Works with citrus, tomato, pork, fish and stone fruit
Peppery warmth|Dry pungency and gentle heat|Black pepper, mustard, ginger, radish, horseradish|Freshly grated or ground versions taste sharper
Numbing tingle|Citrusy mouthfeel change|Sichuan pepper, sansho|Pair with chili, garlic, fermented bean or citrus flavors
Chili heat, grassy|Fresh vegetal heat|Jalapeño, serrano, green chili, shishito|Heat varies by pepper and plant; add gradually
Chili heat, fruity|Sweet-fruit aroma with heat|Habanero, Scotch bonnet, ají amarillo|Small amounts contribute strong aroma and heat
Chili heat, smoky|Dried, toasted, smoke-backed heat|Chipotle, smoked chili, ancho, Urfa biber|Rehydrate dried chilies or bloom powder gently in oil
Chili heat, clean|Direct, sharp capsaicin heat|Cayenne, chile de árbol, bird's eye chili|Use carefully; remove seeds only if desired, since heat varies by variety
Mild sweet spice|Warm, round baking aroma|Cinnamon, nutmeg, allspice, mace|Pairs with fruit, grains, dairy and chocolate
Savory toasted notes|Roasted, nutty, browned aromas|Toasted sesame, cumin, browned butter, roasted nuts|Toast briefly; burnt spices turn bitter
Mineral / briny taste|Sea-like salinity and mineral depth|Oysters, seaweed, anchovy, capers, olives|Acid and fresh herbs help keep briny foods lively
Creamy richness|Soft, coating mouthfeel|Cream, coconut milk, avocado, tahini, cashew|Acid, chili and crisp textures offset richness
Buttery richness|Round, dairy or nut-like fat notes|Butter, ghee, brioche, avocado, some nuts|Browned forms gain toasted, caramel notes
Caramel sweetness|Cooked sugar and browned-fruit notes|Date, molasses, roasted onion, maple, browned butter|Can pair with bitter, smoky, sour or salty notes
Fresh green aroma|Cut grass, leaves, stems|Cilantro, parsley, basil, scallion, pea shoots|Add fresh herbs late and keep them dry before chopping
Savory brothiness|Long, lingering savory taste|Dashi, stock, mushrooms, aged cheese, miso|Layer umami ingredients but monitor sodium
Texture: crisp|Clean snap and audible crunch|Jicama, apple, cucumber, raw pepper, pickle|Keep cold and dry; dressing softens crisp items
Texture: creamy|Smooth, soft, cohesive|Pureed beans, yogurt, ricotta, potato, squash|Add crisp or acidic accents for contrast
Texture: chewy|Resilient bite|Farro, wheat berries, dried fruit, seitan, dried mushrooms|Useful contrast with soft vegetables or sauces
Texture: gelatinous|Silky, sticky or thickening quality|Okra, chia, tapioca, gelatin, some seaweeds|Acid or crisp garnishes can offset a thick mouthfeel
Texture: flaky|Layered, separable bite|White fish, puff pastry, phyllo, laminated dough|Protect from excess moisture until service
Texture: juicy|Moist release when bitten|Tomato, citrus, melon, stone fruit|Salt can draw water from produce; dress close to service
`));
tasteRows.push(...parse(`
Astringent / tannic|Drying, puckering sensation|Strong tea, unripe persimmon, walnut skin, red wine|Fat, sweetness and cooking can soften astringency
Sulfurous allium|Savory, sharp, sulfur-rich aroma|Onion, garlic, leek, chive, scallion|Cooking sweetens and softens; raw forms are sharper
Cooling menthol|Cool, minty aromatic sensation|Peppermint, spearmint, menthol herbs|Use with fruit, chocolate, lamb, peas or yogurt
Resinous / piney|Evergreen, balsamic aroma|Juniper, rosemary, spruce tips, pine nuts|Use sparingly; pair with game, citrus or berries
Funky fermented|Strong, aged, fermented aroma|Blue cheese, natto, fish sauce, shrimp paste, ripe durian|Small amounts can deepen dishes; allow for strong aroma
Mushroom umami|Earthy, brothy savory depth|Shiitake, porcini, dried mushroom, yeast extract|Dried forms concentrate flavor; rehydrate and use liquor
Fruity ester aroma|Juicy, floral fruit-like fragrance|Ripe banana, pineapple, fermented fruit, some cheeses|Ripeness and fermentation stage change aroma
Starchy sweetness|Mild sweetness with powdery or creamy starch|Potato, rice, corn, cassava, plantain|Roasting and frying create browned notes; balance with acid
Oily / fatty mouthfeel|Coating, rich, lingering texture|Sesame oil, olive oil, nuts, fatty fish, cream|Acid, bitter greens and crisp texture refresh the palate
Pungent mustard heat|Nasal, sharp, fast fading heat|Mustard, wasabi, horseradish, radish|Freshly grated forms lose pungency over time
Saline taste|Direct saltiness|Salt, brined cheese, olives, cured foods|Account for salt already present in condiments and stocks
Sweet-sour balance|Fruit-like sweet and tart together|Tamarind, sour plum, pomegranate molasses, some pickles|Can balance fatty, smoky or chili-hot foods
Charred bitterness|Roast, char, smoke with bitter edge|Grilled eggplant, charred onion, coffee, toasted spices|Avoid excessive charring; fresh acid brightens the dish
Creamy tang|Dairy-like richness with acidity|Sour cream, cultured cream, yogurt, creme fraiche|Works as a cooling counterpoint to spice and smoke
Oceanic umami|Briny, mineral, sea-derived savoriness|Seaweed, shellfish, dried fish, oyster sauce|Concentrated products can be salty; dose gradually
Alkaline / mineral|Soft, mineral, slightly soapy if excessive|Lye noodles, alkaline water, some mineral salts|Use food-grade ingredients and measured quantities
Metallic|Coin-like, mineral aftertaste|Some greens, iron-rich foods, certain water|Acid and aromatics may balance; avoid assuming it signals a specific nutrient
Medicinal / camphoraceous|Cool, balsamic, sharp herbal note|Mastic, cubeb, some herbs, clove|Use a light hand; too much can dominate
Nutty / roasted seed|Toasted, warm, fatty aroma|Sesame, peanut, hazelnut, roasted chickpea|Toast gently and protect from scorching
Green pepper / capsicum aroma|Fresh, vegetal, sometimes grassy|Green peppercorn, bell pepper, fresh chili|Pair with citrus, onion, herbs or char
Stone-fruit kernel aroma|Almond, cherry, marzipan-like|Mahleb, apricot kernel aroma, some stone fruits|Kernel-like notes may be allergen-sensitive; use food-safe products
Salty-fermented funk|Salt, aged protein, strong umami|Kashk, shrimp paste, fish sauce, fermented tofu|Use sparingly and account for sodium
Oceanic sweetness|Sweet brine and clean sea aroma|Scallop, crab, lobster, geoduck|Avoid overcooking; acid and herbs brighten
Animalic / gamey|Earthy, mineral, wild-meat aroma|Venison, squab, goat, some aged cheeses|Use herbs, alliums, smoke or fruit to complement
Cooling herbal bitterness|Cool leafy aroma with bitter edge|Dandelion greens, chicory, mint, some herbs|Fat, salt and gentle cooking soften the edge
`));

tasteRows.push(...parse(`
Tropical resin|Sweet, floral, resin-like fruit aroma|Salak, marang, langsat, some tropical fruits|Pair with citrus, coconut, spice or cultured dairy
Wet-earth petrichor|Fresh soil, rain, mineral aroma|Beet, some mushrooms, root vegetables|Roasting, acid and aromatic herbs can shift the earthy impression
Pickled sour|Sharp acid with salt and fermentation notes|Suan cai, takuan, curtido, sour bamboo shoot|Use as a condiment and account for sodium
Tea-like bitterness|Dry, leafy, tannic bitterness|Tea leaf, chicory, some greens|Fat, dairy, sweetness and shorter steeping can soften it
Fermented leaf aroma|Leafy, sour, savory fermentation|Laphet, fermented mustard greens, kimchi|Small portions add complexity and salt
Freshwater mineral|Clean, earthy, sweet-water savoriness|Carp, catfish, freshwater mussels|Use aromatics, acid and clean cooking liquids to complement
Bone-in savory|Deep, gelatin-rich, roasted-meat savoriness|Oxtail, poultry backs, fish frames|Long cooking extracts gelatin; skim and balance with acidity
Green seed / pod flavor|Fresh, grassy, legume-like|Cowpea pods, drumstick pods, winged bean|Quick cooking preserves green notes and crispness
Sour-fruit perfume|Tartness with fragrant fruit aroma|Phalsa, tamarillo, bael, fingered citron|Use peel or juice in measured amounts; bitterness differs by fruit
`));

tasteRows.push(...parse(`
Seaweed iodine / mineral|Marine, mineral, sometimes metallic|Wakame, kombu, nori, mozuku|Iodine varies widely; use food products in reasonable portions
Smoked seafood|Smoke layered over brine and umami|Smoked eel, mussels, fish, roe|Pair with acid, fresh herbs or crisp vegetables
Lactic tang|Cultured dairy sourness, creamy finish|Kefir, skyr, filmjölk, cultured cheese|Fat and fermentation time change perceived sharpness
Fermented seed savor|Nutty, pungent, savory fermentation|Iru, ogiri, fermented locust bean|Use a little at a time; account for salt and aroma
Botanical lemon aroma|Citrus fragrance from leaves or peel|Lemon myrtle, makrut lime leaf, fingered citron|Add late or infuse gently to protect volatile aroma
Sweet grain aroma|Warm cereal, malt, toasted rice|Kañiwa, fonio, Job's tears, roasted grain|Toasting deepens flavor; avoid scorching
Oceanic creaminess|Rich, soft, sweet marine flavor|Crab, langoustine, freshwater prawn|Cook gently and serve with mild acid
Piquant fermented greens|Sharp mustard, sourness, salt|Nukazuke, suan cai, kiam chye|Rinse or temper with rice and fat when salt is strong
`));

tasteRows.push(...parse(`
Berry-scented sweetness|Floral, fruity sweetness with gentle acidity|Cloudberry, salmonberry, jujube, yangmei|Pair with dairy, grain, citrus or mild herbs
Resinous tropical aroma|Sweet fruit with balsamic, pine-like notes|Marula, imbe, salak, some wild berries|Use ripe fruit and balance with citrus or spice
Fermented-fish pungency|Salty marine funk with savory depth|Iru, fermented mackerel, fish paste, laphet accompaniments|Use small amounts; cook or dilute as appropriate
Bran and husk bitterness|Earthy, dry, whole-grain finish|Red rice, black rice, buckwheat, whole pulses|Rinsing, soaking and cooking soften harshness
Sweet cultured dairy|Creamy tang with mild sweetness|Skyr, filmjölk, kefir, buffalo curd|Fruit, herbs and warm spices can complement the acidity
Fresh ocean salinity|Clean brine and sea aroma|Kelp, cockles, whitebait, sea lettuce|Use fresh products and keep cold; iodine and salt vary
Smoky vegetal depth|Charred or smoked green flavor|Smoked pepper, greens, grilled vegetables|Acid and herbs refresh the finish
Floral-citrus peel|Bright aromatic citrus without strong juice acidity|Fingered citron, sudachi, kabosu, makrut leaf|Zest or infuse gently; too much pith can be bitter
`));

tasteRows.push(...parse(`
Sour cereal fermentation|Tangy, cereal-forward, lightly yeasty|Ogi, kenkey dough, injera starter|Fermentation time and grain affect sourness
Mineral organ richness|Deep, iron-like, savory|Liver, cod liver, blood curd, gizzard|Use small portions and balance with acid or herbs
Gelatinous collagen texture|Silky, sticky, broth-rich mouthfeel|Tendon, chicken feet, tripe, long-simmered stock|Long cooking extracts gelatin and softens connective tissue
Sour green fruit|Tart, grassy, sometimes tannic|Unripe mango, kei apple, sour berries|Salt, sugar or fat can soften sharp acidity
Tropical butteriness|Rich, creamy, mild fruit fat|Safou, avocado, some palm fruits|Acid and fresh herbs can balance richness
Smoky pod spice|Dry smoke, fruit and chili warmth|Pasilla Oaxaca, smoked paprika, dried chilies|Toast or soak gently; charred spices can become bitter
Nutty seed oil|Toasted, earthy, oily aroma|Baobab seed oil, sesame, groundnut oils|Storage and freshness affect aroma; add delicate oils late
Sour dairy culture|Fresh lactic tang and creamy aroma|Skyr, kefir, fermented milk, labneh|Salt and fat change perceived acidity
`));

tasteRows.push(...parse(`
Wild mushroom savor|Earthy, mineral, rich umami|Termitomyces, bamboo fungus, wild cultivated mushrooms|Use identified species; dried forms concentrate flavor
Palm-fruit creaminess|Mild, buttery, sweet tropical flesh|Safou, nipa palm fruit, ice apple|Serve fresh or gently cook; heat alters delicate texture
Fermented fish intensity|Pungent marine salt and deep umami|Ngari, hentak, kapi, fermented fish pastes|Use sparingly and reduce added salt
Bitter green floral|Leafy bitterness with floral herbal aroma|Chrysanthemum greens, sesbania flowers, papalo|Fat, acid and short cooking can soften bitterness
Smoked sourness|Fermented acidity layered with smoke|Sukuti, bamboo shoot pickle, smoked fish|Balance with fresh herbs, grain or mild dairy
Root starch earthiness|Dense, earthy, gently sweet starch|Suran, taro, yam, sweet potato|Cook thoroughly; acid, spice and fat can balance
Cooling palm jelly|Watery, delicate sweetness with soft gel texture|Ice apple, nipa fruit, young palm products|Use chilled and fresh; texture is easily damaged by heat
Savory legume ferment|Nutty, salty, fermented umami|Iru, ogiri, fermented soybean products|Small amounts add depth; account for sodium and legume allergens
`));

tasteRows.push(...parse(`
Tannic nut sweetness|Nutty, dry, subtly sweet, mildly bitter|Acorn flour, chestnut, cacao, tea|Leaching and roasting change tannin and sweetness
Fermented fruit acidity|Sweet-sour, yeasty, fruity|Date vinegar, palm toddy, fermented fruit|Fermentation stage changes sweetness, acid and alcohol
Green tropical sourness|Citrus-like tartness with grassy aroma|Naranjilla, calamondin, roselle leaf, unripe fruit|Use as an acid element and balance with fat or sweet notes
Sour dairy richness|Lactic tang wrapped in butterfat|Yak yogurt, buffalo buttermilk, goat cheese|Salt, fat level and culture change perceived sharpness
Pepper-leaf aroma|Citrus, anise, pepper, herbal pungency|Hoja santa, Piper sarmentosum, culantro|Use fresh leaves sparingly; prolonged heat fades aroma
Palm starch neutrality|Soft, neutral, subtly sweet starch|Sago, palm flour, root starch|Pair with aromatic sauces, fruit or fermented condiments
Insect nutty richness|Roasted, earthy, fatty umami|Bee brood, grubs, termites|Use food-grade product and account for allergen cross-reactivity
Oceanic plant salinity|Fresh brine, mineral, seaweed umami|Kelp stems, sea grapes, ogonori|Iodine and salinity vary; avoid over-seasoning
Smoky dairy tang|Cultured sourness with wood smoke|Smoked goat cheese, smoked sheep cheese|Fruit, honey and mild grains soften intensity
Sweet aromatic fruit flesh|Floral tropical sweetness with soft texture|Pawpaw, gac, kepel, jujube|Acid, spice and chilled dairy can add contrast
`));

tasteRows.push(...parse(`
Deep mineral green|Green, earthy, iron-like finish|Chenopodium greens, chaya, cassava leaves|Cook thoroughly and balance with citrus or aromatics
Sour fermented grain|Tangy, yeasty, cereal aroma|Enset kocho, injera, ogi, kenkey|Fermentation duration and grain blend change acidity
Tropical resin-fruit|Fragrant, sweet, balsamic fruit notes|Kaffir plum, marula, kepel, gac|Use ripe fruit; acid and spice can lift the profile
Peppery root spice|Warm, sharp, earthy pungency|Black cumin, chitrak, ginger, long pepper|Toast gently; overuse can dominate the dish
Fruity souring agent|Bright, sharp, fruit-derived acidity|Batuan, tamarind, kokum, sour plum|Use gradually; acidity varies by harvest and preparation
Smoky fermented savor|Aged, smoky, salty umami|Ngari, smoked cheese, fermented black beans|Use modest amounts and adjust salt elsewhere
Fatty organ richness|Soft, mineral, creamy meat flavor|Sweetbreads, liver, kidney, heart|Acid, herbs and crisp texture offer contrast
Clean cultured tang|Mild lactic sourness with creamy finish|Kumis curd, laban, yogurt, kefir|Culture, temperature and milk source alter tang
Popping saline|Small bursts of briny liquid and marine aroma|Sea grapes, roe, some fresh seaweeds|Add at service; long soaking softens texture
Sour berry astringency|Tart, drying, tannic fruit finish|Bilberry, chokecherry, rose hip, sour plums|Sweetness and cooking mellow astringency
`));

const originsRows = [
['Region / cuisine pattern','Common pantry & flavor cues','Common building blocks / dishes','Notes'],
['South Asian','Cumin, coriander, turmeric, chili, ginger, garlic, mustard seed, curry leaves','Rice, lentils, flatbreads, vegetables, yogurt, chutneys','Huge regional diversity; ingredient combinations vary by household and region'],
['North Indian / Mughlai','Garam masala, cardamom, saffron, ginger, garlic, yogurt, nuts','Biryani, korma-style sauces, kebabs, dal, breads','Richness and spice levels vary widely'],
['South Indian / Sri Lankan','Curry leaves, mustard seed, tamarind, coconut, chilies, black pepper','Rice, sambar, coconut curries, sambols, fermented batters','Coconut, sourness and tempering styles differ across regions'],
['Persian / Iranian','Saffron, dried lime, mint, parsley, pomegranate, walnuts','Rice, herbs, stews, grilled meats, yogurt','Herbs and sour-fruit notes are frequent, with regional variation'],
['Levantine','Lemon, sumac, tahini, parsley, mint, chickpeas, olive oil','Mezze, flatbreads, grilled foods, legumes, salads','Shared across multiple national food traditions with distinct local forms'],
['Maghrebi / North African','Cumin, coriander, paprika, preserved lemon, harissa, mint','Couscous, tagines, chickpeas, grains, roasted vegetables','Harissa and spice mixes vary by country and maker'],
['Italian','Olive oil, tomato, garlic, basil, oregano, rosemary, cheese','Pasta, risotto, beans, polenta, antipasti','Regional ingredients and techniques are highly diverse'],
['French','Butter, cream, shallot, thyme, bay, mustard, wine vinegar','Stocks, braises, sauces, gratins, pastries','Classic techniques are a reference; everyday regional cooking is broader'],
['Greek / Aegean','Olive oil, lemon, oregano, dill, yogurt, feta','Grilled foods, beans, salads, pies, seafood','Local produce and island/coastal traditions vary'],
['Spanish','Olive oil, garlic, paprika, saffron, sherry vinegar','Rice dishes, stews, cured foods, seafood, legumes','Regional cuisines differ substantially across Spain'],
['Mexican','Dried chilies, corn, beans, tomato, lime, cilantro, epazote','Tortillas, moles, salsas, stews, beans','Chili species and sauces carry distinct regional identities'],
['Central American','Corn, beans, plantain, chilies, herbs, citrus','Masa-based foods, rice and beans, stews, pickles','Broad regional label; country-specific dishes should be named where possible'],
['Caribbean','Allspice, thyme, Scotch bonnet, citrus, coconut, scallion','Rice and peas, stews, grilled foods, plantain, seafood','Distinct island and diaspora traditions; blends vary'],
['Southern US / Cajun-Creole','Celery, onion, pepper, paprika, thyme, bay, cayenne','Gumbo, jambalaya, beans, cornbread','Cajun and Creole traditions are related but not interchangeable'],
['Chinese','Ginger, garlic, scallion, soy, rice vinegar, sesame, Sichuan pepper','Rice, noodles, dumplings, stir-fries, braises','Chinese regional cuisines differ greatly in seasoning and technique'],
['Japanese','Soy, miso, rice vinegar, kombu, bonito, sesame, mirin','Rice, noodles, soups, pickles, grilled foods','Dashi sources vary; vegetarian alternatives use kombu and mushrooms'],
['Korean','Gochujang, doenjang, soy, sesame, garlic, scallion, chili flakes','Rice, soups, fermented vegetables, noodles, grilled dishes','Fermented pastes and kimchi vary by region and producer'],
['Thai','Lemongrass, galangal, lime, fish sauce, chilies, basil, coconut','Rice, noodle dishes, salads, curries, soups','Balance often combines sour, salty, sweet, hot and aromatic elements'],
['Vietnamese','Fish sauce, lime, herbs, chili, rice noodles, aromatics','Noodle soups, rolls, rice plates, fresh herb salads','Fresh herb sets and regional styles differ'],
['Indonesian','Kecap manis, sambal, coconut, turmeric, galangal, lemongrass','Rice, satay, rendang-style braises, sambals','Many island cuisines; broad label should not erase local distinctions'],
['Filipino','Vinegar, soy, garlic, black pepper, coconut, calamansi','Adobo-style braises, rice, noodles, stews','Sweet-sour-savory balances vary by region and household'],
['Ethiopian / Eritrean','Berbere, fenugreek, ginger, garlic, clarified spiced butter','Injera, lentils, vegetable stews, slow-cooked meats','Berbere blends and fasting menus vary; teff is a key grain association'],
['West African','Peanuts, chilies, ginger, tomato, leafy greens, palm oil','Peanut stews, grains, yams, plantain, rice dishes','Many distinct national and regional traditions; avoid treating as one cuisine'],
['Nordic','Dill, mustard, rye, berries, potatoes, fish, cultured dairy','Open sandwiches, pickles, soups, baked goods, seafood','Preservation, seasonality and local ingredients are common themes'],
['British / Irish','Potatoes, oats, mustard, herbs, dairy, apples','Pies, roasts, soups, puddings, breads','Regional and contemporary traditions are broad'],
['Andean','Potatoes, corn, quinoa, ají chilies, herbs','Potato dishes, stews, grain bowls, roasted foods','Crop diversity and local ingredients vary by altitude and country'],
['Balkan / Eastern European','Paprika, dill, cabbage, beet, sour cream, caraway','Soups, dumplings, pickles, stews, breads','Overlapping influences with substantial country-level variation'],
['North American contemporary pantry','Chilies, maple, corn, beans, barbecue spices, dairy','Grill, salads, baked dishes, smoked foods, fusion menus','A broad modern pantry category rather than a single traditional cuisine']
];
originsRows.push(...[
['Bengali','Mustard oil, nigella, panch phoron, green chili, fish, rice','Rice, fish curry, lentils, sweets, vegetable sides','West Bengal and Bangladesh have shared and distinct regional traditions'],
['Gujarati','Cumin, mustard, sesame, jaggery, lentils, peanuts','Thali, dal, breads, farsan, sweet-sour vegetables','Sweet-sour balance varies by district and household'],
['Tamil','Tamarind, curry leaves, mustard, pepper, coconut, rice','Rice, sambar, rasam, dosa, vegetable curries','Tamil cuisine spans Tamil Nadu, Sri Lanka and diaspora forms'],
['Kashmiri','Dried chilies, fennel, ginger, saffron, yogurt','Rice, slow braises, greens, breads','Kashmiri Pandit and Muslim food traditions differ'],
['Goan','Coconut, kokum, vinegar, chilies, seafood','Fish curry, rice, breads, pork preparations','Coastal ingredients and Portuguese-era influences vary by community'],
['Nepali / Himalayan','Millet, buckwheat, lentils, chili, fermented greens','Dal-bhat, dumplings, pickles, soups','Mountain ecology and ethnic traditions create great diversity'],
['Tibetan','Barley, yak dairy/meat, chili, tea, dried foods','Tsampa, momos, soups, butter tea','High-altitude foodways; regional and diaspora practices vary'],
['Georgian / Caucasus','Walnut, tkemali, herbs, coriander, fenugreek, wine','Khachapuri, dumplings, stews, sauces','Caucasus cuisines are diverse; Georgian examples are not region-wide rules'],
['Turkish / Anatolian','Yogurt, pepper paste, sumac, eggplant, herbs','Meze, pilaf, kebab, breads, beans','Anatolian regional cuisines vary across climate and history'],
['Armenian','Pomegranate, apricot, herbs, lamb, wheat, yogurt','Dolma, breads, stews, grilled foods','Shared borderland influences with distinctive Armenian food traditions'],
['Azerbaijani','Saffron, herbs, pomegranate, lamb, rice','Plov, dolma, soups, kebabs','Regional variation spans Caspian coast and inland areas'],
['Uzbek / Central Asian','Cumin, coriander, lamb, carrots, rice, flatbread','Plov, noodles, dumplings, kebabs','Silk Road influences; local preparations differ'],
['Kazakh / steppe','Meat, dairy, fermented mare milk, wheat','Noodles, dumplings, preserved meats, dairy','Traditional diet reflects pastoral ecology and contemporary change'],
['Mongolian','Meat, dairy, flour, fermented milk','Dumplings, noodles, dairy foods, soups','Ingredients reflect pastoral traditions; city diets are broader'],
['North African Amazigh','Couscous, olive oil, herbs, legumes, preserved foods','Tagine, couscous, flatbreads, seasonal vegetables','Amazigh traditions are diverse across Morocco, Algeria, Tunisia and beyond'],
['Egyptian','Fava beans, cumin, garlic, coriander, rice, molokhia','Ful, koshari, stews, breads, pickles','Local dishes and ingredients vary across Nile, delta and coastal regions'],
['Sudanese','Peanuts, sesame, okra, sorghum, hibiscus, spices','Stews, porridge, flatbread, grilled foods','Regional and ethnic traditions vary across a large country'],
['Ghanaian','Cassava, yam, plantain, palm oil, chili, fermented fish','Soups, fufu, rice dishes, stews','Coastal, forest and northern cuisines differ'],
['Senegalese','Fish, rice, tamarind, bissap, peanuts, herbs','Thieboudienne-style rice, yassa, stews','Senegalese traditions include many local ingredients and techniques'],
['Nigerian','Yam, cassava, rice, melon seed, palm oil, chilies','Soups, swallows, rice dishes, bean cakes','Hundreds of ethnic groups; broad national label spans many cuisines'],
['Congo Basin','Cassava leaves, manioc, fish, palm oil, peanuts','Stews, fermented cassava, grilled fish','Central African food traditions vary by river, forest and country'],
['Madagascar','Rice, greens, zebu, ginger, vanilla, coconut','Rice plates, stews, broths, seafood','Island regions and highlands have varied ingredients and practices'],
['Polynesian','Breadfruit, taro, coconut, fish, banana','Earth-oven foods, poi-style preparations, seafood','Island-specific food traditions and crops differ'],
['Māori / Aotearoa New Zealand','Kūmara, seafood, native greens, wild foods','Hāngī, seafood, root vegetables','Traditional foods are locally specific; use culturally accurate names'],
['Inuit / Arctic','Fish, marine mammals, caribou, berries, wild plants','Dried, frozen, fermented and fresh foods','Food systems are locally specific and shaped by climate and sovereignty'],
['Amazonian Indigenous foodways','Cassava, fish, palm fruits, cacao, chili, wild greens','Farinha, fish stews, roasted foods, beverages','Hundreds of communities; avoid treating as one cuisine'],
['Brazilian regional','Cassava, beans, rice, tropical fruit, fish, herbs','Feijoada, moqueca, farofa, regional stews','Brazilian food varies across Amazon, northeast, south and urban regions'],
['Peruvian coastal','Ají peppers, citrus, seafood, corn, potatoes','Ceviche, rice dishes, stews, seafood','Coastal, Andean and Amazonian cuisines differ'],
['Chilean','Seafood, potatoes, corn, beans, herbs, berries','Stews, empanadas, seafood, breads','Long geography creates strongly regional food patterns'],
['Argentine / Pampas','Beef, wheat, yerba mate, herbs, grilled foods','Asado, empanadas, pasta, mate','Indigenous, immigrant and regional traditions intersect'],
['New Orleans Creole','Seafood, celery, onion, peppers, tomato, herbs','Gumbo, rice dishes, sauces','Distinct from Cajun traditions, with many family variations'],
['Hawaiian','Taro, fish, seaweed, coconut, breadfruit','Poi, poke, laulau, seafood','Native Hawaiian traditions and contemporary local food are distinct but connected'],
['Jewish Ashkenazi','Rye, onion, dill, cabbage, schmaltz, pickles','Breads, soups, dumplings, braises','Religious practice and diaspora shaped varied regional dishes'],
['Jewish Sephardi / Mizrahi','Rice, legumes, herbs, citrus, eggplant, spices','Rice dishes, fish, stews, pastries','Many distinct communities across Mediterranean and Middle East'],
['Edo / Japanese regional','Miso, soy, kombu, rice, seafood','Regional soups, noodles, pickles, grilled foods','Japan has many regional food cultures; this label is not a single flavor profile'],
['Sichuan','Chili, Sichuan pepper, fermented bean paste, garlic','Hot pot, mapo tofu, noodles, braises','Sichuan flavor includes many styles beyond numbing heat'],
['Cantonese','Seafood, ginger, scallion, soy, rice wine','Dim sum, roast meats, steamed dishes, soups','Cantonese cuisine is broad and differs across Guangdong and diaspora'],
['Hunan','Chili, smoked foods, garlic, fermented ingredients','Braised dishes, stir-fries, pickles','Often hot and aromatic, with regional variation'],
['Okinawan','Bitter melon, pork, tofu, sweet potato, seaweed','Stir-fries, soups, tofu, rice dishes','Distinct Ryukyuan food traditions and island ingredients'],
['Korean Jeolla','Fermented pastes, seafood, chili, greens','Banchan, rice, stews, fermented vegetables','Regional reputation and practices vary within the province'],
['Thai Isan','Sticky rice, herbs, chili, lime, fermented fish','Som tam, grilled meats, larb-style salads','Distinct northeast Thai food culture with Lao connections'],
['Vietnamese Mekong Delta','Rice, fish, coconut, herbs, tropical fruit','Freshwater fish dishes, noodles, soups','Southern regional ingredients differ from northern and central cuisines']
]);
originsRows.push(...[
['Odia / Odisha','Rice, mustard, panch phutana, greens, seafood','Pakhala, dalma, fish, pitha','Coastal and inland food practices differ across Odisha'],
['Assamese','Rice, mustard, bamboo shoot, herbs, freshwater fish','Rice meals, sour fish curries, greens, fermented foods','Brahmaputra valley and hill communities have distinct traditions'],
['Manipuri','Rice, fermented fish, herbs, bamboo shoot, chilies','Eromba, fish, seasonal greens, rice','Ingredient use differs by community and locality'],
['Ladakhi','Barley, apricot, yak dairy, turnip, dried foods','Thukpa, skyu, butter tea, breads','High-altitude seasonal practices vary by village and household'],
['Sindhi','Rice, wheat, tamarind, chickpea flour, greens','Sai bhaji, kadhi, fish, breads','Regional and diaspora practices vary'],
['Balochi','Lamb, wheat, dates, dairy, regional herbs','Sajji, flatbreads, meat stews','Pastoral and coastal foodways differ across Balochistan'],
['Omani / Gulf coast','Dates, lime, fish, rice, spices','Shuwa, rice dishes, seafood, halwa','Coastal and inland traditions use different ingredients'],
['Yemeni','Fenugreek, hilbeh, hawaij, bread, meat','Saltah, stews, breads, rice','Regional practices differ across highlands and coast'],
['Kurdish','Wheat, rice, herbs, yogurt, lamb, fruit','Dolma, kebab, soups, breads','Food traditions span several countries and vary by region'],
['Ainu / Hokkaido','Salmon, deer, mountain vegetables, kelp','Grilled fish, stews, preserved foods','Indigenous foodways are specific and should be represented respectfully'],
['Ryukyuan / Yaeyama','Seafood, pork, sugarcane, tropical greens','Noodles, tofu, seaweed, island vegetables','Island-specific practices vary within Okinawa prefecture'],
['Lao','Sticky rice, herbs, padaek, chilies, river fish','Larb, jeow, soups, grilled foods','Strong regional and ethnic diversity across Laos'],
['Cambodian / Khmer','Prahok, fish, rice, lemongrass, palm sugar','Amok, soups, fermented fish dishes','Coastal, Tonle Sap and inland ingredients differ'],
['Burmese / Myanmar','Fermented tea leaf, fish paste, chickpea flour, oil','Rice, salads, noodles, curries','Many ethnic food traditions; national label spans substantial diversity'],
['Malay coastal','Coconut, turmeric, galangal, pandan, seafood','Rice, curries, sambals, grilled fish','Malay foodways vary across Malaysia, Indonesia, Brunei and diaspora'],
['Sámi / Nordic Arctic','Reindeer, fish, berries, wild herbs','Dried, smoked, preserved and fresh foods','Traditional food systems are locally specific and living'],
['Basque','Cod, peppers, beans, cider, seafood','Pintxos, stews, grilled fish, sauces','Spanish and French Basque regions share links and retain local variation'],
['Sicilian','Citrus, eggplant, capers, tuna, wheat, almonds','Pasta, seafood, sweets, vegetable dishes','Island regions show Arab, Greek and broader Mediterranean influences'],
['Sardinian','Sheep dairy, durum wheat, herbs, seafood','Pasta, breads, pecorino, roasted foods','Pastoral, coastal and inland cooking differ'],
['Oaxacan','Corn, beans, chilies, cacao, herbs, squash seed','Mole, tlayudas, tamales, market foods','Many Indigenous nations and local traditions; not one uniform cuisine']
]);

originsRows.push(...[
['Bengali Muslim coastal','Mustard, rice, fish, lime, aromatics','Fish curries, rice, festive dishes','Regional practices overlap with and differ from nearby Bengali traditions'],
['Tripuri / Northeast India','Rice, fermented fish, herbs, bamboo shoot, pork','Stews, rice, fermented ingredients','Local foodways vary among Indigenous communities'],
['Konkani coastal','Coconut, kokum, rice, fish, spices','Fish curries, rice, vegetable dishes','Konkan food traditions vary from Goa through Maharashtra and Karnataka'],
['Chettinad','Pepper, fennel, kalpasi, chilies, aromatics','Spiced meat, rice, vegetable curries','A Tamil regional tradition with household variation'],
['Bihari','Mustard oil, sattu, greens, rice, lentils','Litti, chokha, fish, seasonal vegetables','District and community traditions vary widely'],
['Burmese Shan','Rice noodles, tomato, herbs, fermented soybean','Noodles, rice dishes, pickles','Shan cuisine is one of many distinct Myanmar food traditions'],
['Hmong','Herbs, chilies, greens, pork, rice','Herb-rich soups, grilled foods, rice','Ingredients and practices vary across Hmong communities and diaspora'],
['Taiwanese Hakka','Preserved mustard greens, dried seafood, pork, rice','Braised dishes, rice, noodles','Hakka food traditions are diverse across Taiwan and diaspora'],
['Japanese Tohoku','Rice, miso, mountain vegetables, seafood','Rice meals, pickles, soups, preserved foods','Climate and coastal/inland geography shape local ingredients'],
['Japanese Okinawa islands','Pork, tofu, bitter melon, seaweed, tropical greens','Stir-fries, soups, noodles, rice dishes','Island-specific cuisines vary beyond mainland Japanese patterns'],
['Southern Italian Calabria','Chili, citrus, eggplant, pork, seafood','Pasta, preserves, grilled foods','Calabrian pantry varies between inland and coastal communities'],
['Greek Pontic','Wheat, dairy, greens, fish, spices','Pasta, breads, dumplings, dairy dishes','Pontic heritage includes regional and diaspora variations'],
['Haitian','Scotch bonnet, thyme, citrus, rice, beans, root crops','Rice, stews, fritters, griot-style dishes','Local ingredients and diaspora adaptations differ'],
['Garifuna','Cassava, coconut, fish, plantain, herbs','Hudut, cassava bread, seafood','Distinct Afro-Indigenous traditions across Central American Caribbean coasts'],
['Cajun Acadiana','Rice, seafood, pork, celery, peppers, herbs','Gumbo, jambalaya, rice dishes','Distinct regional tradition; recipes vary by family and parish']
]);

originsRows.push(...[
['Mizo','Rice, fermented soybean, bamboo shoot, greens, pork','Stews, rice, fermented foods','Mizo food traditions vary by locality and community'],
['Naga','Smoked meat, fermented bamboo shoot, chilies, herbs','Rice, smoked foods, chutneys, stews','Many Naga communities have distinct food practices and ingredients'],
['Garhwali / Uttarakhand hills','Millets, pulses, greens, local herbs, dairy','Phaanu, chainsoo, breads, seasonal vegetables','Mountain ecology and district traditions shape the pantry'],
['Kumaoni','Mandua, bhatt, seasonal greens, local spices','Pulses, breads, chutneys, rice','Regional and household variations are substantial'],
['Malvani coastal','Coconut, kokum, chilies, fish, rice','Fish curry, sol kadhi, seafood, breads','Coastal Maharashtra and Goa-related traditions vary locally'],
['Bhojpuri','Wheat, rice, sattu, mustard, seasonal vegetables','Litti, chokha, dal, pickles','Shared across eastern Uttar Pradesh, Bihar and diaspora communities'],
['Kashmiri Wazwan','Fennel, dry ginger, meat, yogurt, spices','Multi-course ceremonial meat preparations','Wazwan is a specialist culinary tradition with regional and household variation'],
['Bengali Sylheti','Citrus, fish, rice, herbs, mustard','Rice, fish dishes, sour preparations','Sylheti traditions span Bangladesh, India and diaspora communities'],
['Korean Jeju','Seafood, buckwheat, citrus, pork, greens','Noodles, soups, seafood, local citrus','Island ecology shapes distinct Jeju ingredients and dishes'],
['Chinese Fujian / Hokkien','Seafood, rice wine, mushrooms, red fermented rice','Soups, noodles, seafood, braises','Coastal and inland Fujian cooking differs'],
['Chinese Yunnan','Mushrooms, herbs, chilies, rice, preserved foods','Rice noodles, mushroom dishes, pickles','High biodiversity and ethnic diversity create substantial variety'],
['Chinese Xinjiang','Lamb, wheat, cumin, peppers, fruit','Noodles, breads, pilaf, grilled foods','Regional cuisines reflect local and Central Asian influences'],
['Taiwanese Indigenous cuisines','Millet, taro, wild greens, fish, game','Steamed grains, preserved foods, grilled dishes','Distinct Indigenous communities and food sovereignties; avoid homogenizing'],
['Andean Quechua highlands','Potato, quinoa, maize, chuño, herbs','Soups, stews, roasted foods, grain dishes','Local crop varieties and preservation practices vary by altitude'],
['Amazonian riverine','Freshwater fish, cassava, palm fruits, chili','Fish stews, farinha, grilled foods','Foodways are specific to river basin and community']
]);

originsRows.push(...[
['Garo hills','Rice, millet, pork, fermented bamboo shoot, greens','Rice meals, smoked foods, seasonal stews','Local food traditions vary across Garo communities'],
['Khasi hills','Rice, pork, fermented soy, bamboo shoot, greens','Jadoh, stews, chutneys','Khasi household and district food practices vary'],
['Sikkimese Bhutia','Buckwheat, barley, yak dairy, fermented greens','Momos, thukpa, breads, preserved foods','Mountain and community traditions differ'],
['Tamang','Millet, maize, buckwheat, greens, fermented foods','Dhindo, soups, dumplings, pickles','Foodways vary by locality and season'],
['Banjara / Lambadi','Millet, pulses, chilies, dairy, seasonal greens','Flatbreads, stews, preserved foods','Community practices are diverse and change with region'],
['Tulu coastal','Rice, coconut, fish, tamarind, chilies','Neer dosa, fish curry, coastal vegetables','Coastal Karnataka communities have varied traditions'],
['Kumaoni Kumaon hills','Bhatt, mandua, local greens, ghee, spices','Pulse dishes, breads, chutneys','Distinct from neighboring Garhwali traditions though ingredients overlap'],
['Japanese Kansai','Dashi, kombu, soy, wheat, seasonal vegetables','Udon, tofu, simmered dishes, pickles','Kansai includes multiple regional food cultures'],
['Japanese Kyushu','Barley miso, sweet soy, citrus, pork, fish','Noodles, hot pots, grilled dishes','Island and prefectural foodways vary substantially'],
['Okinawa Miyako islands','Seafood, sugarcane, pork, tropical greens','Island vegetables, noodle soups, preserved foods','Local island ingredients and names differ within the Ryukyus'],
['Chinese Guangxi','Rice noodles, herbs, pickled bamboo, river fish','Rice noodle soups, sour dishes, stews','Regional ethnic diversity shapes local ingredients'],
['Chinese Hubei','Freshwater fish, lotus, rice, sesame','Fish dishes, lotus-root soup, noodles','Lake and river ingredients are prominent but vary locally'],
['Taiwanese Hakka mountain','Tea, preserved mustard greens, pork, rice','Braised dishes, tea snacks, rice meals','Mountain communities and diaspora adaptations vary'],
['Sichuan Zigong','Chili, Sichuan pepper, ginger, preserved ingredients','Salt-baked and braised dishes, hot-pot styles','A city-level Sichuan pattern; household recipes differ'],
['Basque coastal','Cod, anchovy, peppers, seafood, cider','Fish stews, grilled seafood, sauces','Coastal products and inland pastoral foodways contrast']
]);

originsRows.push(...[
['Akan / Asante','Plantain, cassava, yam, groundnut, palm oil, fermented fish','Fufu, stews, rice dishes, snacks','Akan food practices vary by region and household'],
['Ewe','Maize, cassava, fish, chilies, herbs','Akple, seafood, soups, fermented staples','Communities span Ghana, Togo and Benin with local variations'],
['Ga coastal','Fish, maize, pepper, palm oil, coconut','Kenkey, seafood, pepper sauces','Coastal Accra traditions vary by family and neighborhood'],
['Mossi / Burkina Faso','Sorghum, millet, shea, leafy greens, legumes','Tô, sauces, porridges','Foodways vary by season and local crop systems'],
['Tuareg / Saharan','Millet, dates, goat milk, meat, tea','Grains, stews, breads, tea service','Nomadic and settled communities have distinct practices'],
['Eritrean highlands','Teff, berbere, legumes, clarified butter','Injera, tsebhi, lentils','Food traditions overlap with but differ from Ethiopian regions'],
['Oromo','Maize, barley, coffee, dairy, greens','Breads, porridge, meat, coffee ceremonies','Large community with diverse regional cuisines'],
['Somali coastal','Rice, fish, coconut, cardamom, spices','Rice dishes, seafood, breads, tea','Coastal and inland Somali foodways differ'],
['Zanzibar Swahili','Coconut, clove, fish, rice, tamarind','Pilau, seafood curries, breads','Island and coastal traditions mix local and Indian Ocean influences'],
['Comorian','Coconut, cassava, rice, fish, vanilla, cloves','Rice, curries, seafood, root crops','Island-specific food practices vary across the archipelago'],
['Samoan','Taro, breadfruit, coconut, fish, banana','Earth oven foods, coconut dishes, seafood','Traditional and contemporary foodways are both diverse'],
['Tongan','Root crops, coconut, seafood, pork','Umukai, lu pulu, root vegetable dishes','Island-specific practices and ingredients vary'],
['Fijian','Taro, cassava, coconut, seafood, leafy greens','Lovo, kokoda, root crops','Indigenous Fijian and Indo-Fijian traditions are distinct and interconnected'],
['Chamorro / Mariana Islands','Rice, coconut, seafood, breadfruit, local herbs','Kelaguen, red rice, root crops','Chamorro food traditions are locally specific and living'],
['Papuan highlands','Sweet potato, greens, pork, pandanus, forest foods','Earth-oven foods, soups, tubers','Many distinct communities; labels should remain locally specific'],
['Māori kūmara-centered','Kūmara, seafood, kūtai, native herbs','Hāngī, preserved foods, seafood','Local iwi traditions and food names vary; use culturally precise sources'],
['Kalaallit / Greenlandic','Fish, marine mammals, berries, seaweed, game','Dried, fermented, fresh and frozen foods','Food systems are community-specific and shaped by season and climate'],
['Tibetan Amdo','Barley, yak dairy, lamb, noodles, fermented foods','Tsampa, thukpa, dairy dishes','Amdo foodways vary within the larger Tibetan region'],
['Uyghur oasis cuisine','Wheat, lamb, cumin, peppers, grapes, apricots','Noodles, breads, pilaf, grilled foods','Oasis, city and household traditions differ'],
['Kazakh southern regional','Rice, wheat, fruit, lamb, dairy, herbs','Pilaf, noodles, breads, fermented dairy','Southern and steppe foodways have distinct crop availability'],
]);

originsRows.push(...[
['Afar / Danakil','Sorghum, millet, camel milk, dates, spices','Flatbreads, stews, dairy, tea','Pastoral and settled food systems vary by locality'],
['Tigrayan','Teff, barley, legumes, berbere, dairy','Injera, stews, breads, lentils','Regional traditions overlap with neighboring highland cuisines and remain distinct'],
['Bemba / Zambia','Maize, cassava, groundnuts, leafy vegetables, fish','Nshima, relish, stews, dried foods','Foodways vary by season and region'],
['Shona / Zimbabwe','Maize, sorghum, groundnuts, greens, beans','Sadza, relishes, stews','Local crops and household preferences vary'],
['Chewa / Malawi','Maize, fish, groundnuts, leafy greens','Nsima, ndiwo, dried fish dishes','Lake and inland food patterns differ'],
['Zulu / South Africa','Maize, sorghum, beans, greens, cultured dairy','Porridge, stews, sour milk foods','Traditional and contemporary foodways are varied'],
['Oromo highland coffee region','Coffee, barley, milk, grains, greens','Coffee ceremony, porridges, breads','Local dishes and coffee traditions vary by community'],
['Sundanese / West Java','Rice, vegetables, herbs, fermented soybean, chili','Fresh vegetable meals, sambal, rice','Sundanese cuisine is distinct within Indonesia and varies by locality'],
['Minangkabau','Coconut, chili, turmeric, beef, cassava leaves','Rendang, rice, curry, leafy dishes','Household and ceremonial recipes vary'],
['Bugis / South Sulawesi','Rice, fish, coconut, spices, seafood','Rice dishes, fish, soups, grilled foods','Coastal and inland traditions differ'],
['Bicolano / Philippines','Coconut milk, chili, fish, taro leaves','Laing, rice, seafood, stews','Regional dishes and heat levels vary by province'],
['Ilocano / Philippines','Bitter greens, fish, vinegar, legumes, pork','Pinakbet, grilled foods, soups','Local vegetables and seasoning practices vary'],
['Maranao / Mindanao','Rice, spices, coconut, beef, herbs','Beef stews, rice, layered ceremonial dishes','Distinct Muslim Filipino tradition; local terms matter'],
['Mapuche / Wallmapu','Maize, potatoes, beans, native fruits, herbs','Grain dishes, stews, breads, seasonal foods','Indigenous food traditions are locally specific and living'],
['Aymara Altiplano','Quinoa, potato, chuño, barley, herbs','Soups, stews, dried potato foods','High-altitude crops and preservation practices vary by community']
]);

originsRows.push(...[
['Sotho / Basotho','Sorghum, maize, beans, greens, sour milk','Pap, stews, breads, cultured dairy','Foodways vary across Lesotho and neighboring communities'],
['Tswana','Sorghum, maize, beans, melons, meat','Bogobe, stews, dried foods','Regional crop and household variation is substantial'],
['Venda','Maize, groundnuts, leafy greens, fermented foods','Porridges, relishes, stews','Local ingredients and seasonal foodways vary'],
['Tsonga / Shangaan','Cassava, maize, groundnuts, fish, greens','Porridges, relishes, stews','Community traditions span multiple countries and differ locally'],
['Makonde','Cassava, maize, coconut, seafood, greens','Porridges, cassava dishes, fish','Coastal and inland foodways differ between Tanzania and Mozambique'],
['Makua','Cassava, maize, coconut, beans, fish','Stews, porridge, seafood','Large community with local food traditions across Mozambique and Tanzania'],
['Bemba, northern Zambia','Maize, cassava, groundnuts, mushrooms, fish','Nshima, relish, dried foods','Seasonality and local ingredients affect meals'],
['Luba / Kasai','Cassava, maize, palm oil, fish, leafy greens','Fufu, stews, fermented cassava','Food systems vary across the Democratic Republic of the Congo'],
['Kongo / Bakongo','Cassava, palm oil, peanuts, fish, greens','Cassava staples, stews, fermented foods','Practices span Angola, Republic of Congo and DRC with local diversity'],
['Bamileke','Maize, cocoyam, beans, vegetables, spices','Fufu, stews, market foods','West Cameroon communities have distinct household traditions'],
['Bassa / Cameroon coastal','Cassava, plantain, fish, palm oil, spices','Cassava dishes, stews, seafood','Coastal and forest ingredients vary by locality'],
['Akan, Ashanti regional','Plantain, yam, cassava, groundnut, palm oil','Fufu, soups, stews, snacks','Regional practices differ within the broader Akan food tradition'],
['Tiv / Benue','Yam, cassava, sesame, leafy greens, fish','Pounded yam, stews, sesame foods','Local crop systems and household cooking vary'],
['Nupe / Niger River','Rice, fish, millet, sorghum, herbs','Rice dishes, fish, porridges','River and inland communities have varied foodways'],
['Dinka / Nilotic','Sorghum, milk, fish, meat, wild greens','Porridge, dairy, dried foods','Food traditions reflect pastoral and riverine ecologies'],
['Hadza / Tanzania traditional foraging','Wild tubers, berries, honey, game','Seasonal hunted and gathered foods','Foodways are community-specific; do not generalize or romanticize'],
['Ainu Sakhalin heritage','Salmon, kelp, berries, wild greens','Preserved fish, soups, seasonal foods','Ainu foodways vary by community and historical context'],
['Ainu Hokkaido coastal','Salmon, cod, seaweed, mountain plants','Fish dishes, soups, preserved foods','Indigenous food culture is living and locally specific'],
['Kanak / New Caledonia','Yam, taro, coconut, seafood, greens','Bougna, root crops, seafood','Customary ingredients and practices vary among Kanak communities'],
['Aboriginal Australian desert foodways','Native seeds, tubers, fruits, game, herbs','Seasonal foods and ground seed preparations','Hundreds of language groups and food traditions; use specific local names']
]);

originsRows.push(...[
['Hadiya / southern Ethiopia','Enset, teff, legumes, greens, spices','Fermented enset, flatbread, stews','Local food practices vary across Hadiya communities'],
['Sidama','Enset, coffee, maize, legumes, greens','Kocho, coffee, stews, fermented staples','Foodways are regionally distinct and evolving'],
['Konso','Sorghum, maize, enset, pulses, vegetables','Porridge, fermented foods, stews','Agricultural and food traditions are locally specific'],
['Kamba / Kenya','Maize, millet, beans, greens, livestock foods','Ugali, stews, roasted foods','Seasonal and household variation is substantial'],
['Kikuyu / Central Kenya','Maize, beans, potatoes, greens, milk','Githeri, mukimo, porridge','Local practices vary by community and season'],
['Luo / Lake Victoria','Fish, maize, millet, greens, groundnuts','Fish stews, ugali, vegetables','Lake, inland and diaspora patterns differ'],
['Taita / Kenya coast hinterland','Maize, beans, greens, roots, fruit','Porridge, stews, seasonal dishes','Local crop ecology shapes regional dishes'],
['Kongo / Angola north','Cassava, palm oil, fish, beans, greens','Funge, stews, fermented cassava','Food traditions differ across language groups and regions'],
['Ovimbundu','Maize, millet, beans, cassava, greens','Funje, stews, dried foods','Regional and household practices vary in central Angola'],
['Herero','Dairy, meat, maize, wild fruits','Porridge, meat and milk dishes','Foodways reflect pastoral history and contemporary change'],
['San / Kalahari regional foods','Wild melons, tubers, berries, game, seeds','Seasonal gathered and hunted ingredients','Many distinct San communities; use precise local sources and names'],
['Rapa Nui / Easter Island','Sweet potato, taro, seafood, banana, sugarcane','Earth-oven dishes, root crops, seafood','Island foodways are specific and living'],
['Tokelauan','Pandanus, coconut, taro, fish','Root crops, seafood, coconut preparations','Local island practices and names vary'],
['Marshallese','Breadfruit, pandanus, coconut, fish','Fermented breadfruit, seafood, root crops','Atoll foodways are locally specific and shaped by ecology'],
['Palauan','Taro, seafood, coconut, tropical greens','Soups, seafood, root crops','Palauan ingredients and customary practices are distinct'],
['Chamorro, Guam regional','Rice, coconut, seafood, breadfruit, lime','Kelaguen, red rice, stews','Guam and Northern Mariana food practices overlap and differ'],
['Karelian','Rye, fish, berries, dairy, mushrooms','Pies, porridges, fish dishes','Karelian food traditions span Russia and Finland with variation'],
['Mari / Volga region','Rye, barley, potatoes, mushrooms, berries','Breads, porridges, soups','Local and seasonal ingredients distinguish household foodways'],
['Bashkir','Honey, dairy, horse meat, grains, berries','Breads, fermented dairy, meat dishes','Pastoral and settled cuisines vary by region'],
['Kalmyk','Dairy, lamb, tea, grains','Dairy foods, soups, meat dishes','Food practices reflect regional ecology and cultural history']
]);

const originVariantRows = [
['Ingredient / variant','Origin or regional association','Taste / sensory variation','Common use','Scope note'],
['Alphonso mango','Maharashtra, India','Perfumed, rich, very sweet','Purees, desserts, drinks','Seasonal cultivar; growing area and ripeness affect flavor'],
['Kesar mango','Gujarat, India','Floral, sweet, saffron-like aroma','Purees, sweets, drinks','Regional cultivar association'],
['Ataulfo mango','Mexico','Buttery, sweet, low-fiber flesh','Fresh eating, salsa, dessert','Names and cultivars vary in markets'],
['Green mango','South Asia and Southeast Asia','Tart, crisp, starchy to juicy by maturity','Pickle, chutney, salad, curry','Used before ripening; local cultivars differ'],
['Bomba rice','Valencia, Spain','Mild, absorbs stock while staying distinct','Paella and rice dishes','Other Spanish rice cultivars are also used'],
['Arborio rice','Northern Italy','Starchy, creamy, resilient center','Risotto, rice pudding','Carnaroli and Vialone Nano are other regional choices'],
['Carnaroli rice','Northern Italy','Creamy, resilient, lightly nutty','Risotto','Cooking behavior varies by milling and age'],
['Basmati rice','Indian subcontinent, especially Himalayan foothills','Long-grain, fragrant, nutty','Biryani, pilaf, steamed rice','Many cultivars and aging grades exist'],
['Jasmine rice','Thailand and Southeast Asia','Floral, soft, lightly sticky','Rice plates, curries','Fragrance and texture vary by cultivar'],
['Glutinous rice','East and Southeast Asia','Sticky, mild, slightly sweet','Mochi, dumplings, sticky rice dishes','Glutinous describes texture; rice is naturally gluten-free'],
['Purple yam / ube','Philippines','Sweet, earthy, vanilla-like','Desserts, breads, drinks','Not interchangeable with purple sweet potato'],
['Taro corm','Pacific Islands, East Asia, South Asia','Nutty, earthy, starchy','Stews, poi-style preparations, desserts','Varieties differ; cook fully before eating'],
['Cassava / yuca','South America, West Africa, Caribbean','Mild, starchy, often floury','Boiled roots, fufu, gari, tapioca','Processing and preparation traditions vary'],
['Fonio','West Africa, especially Sahel','Fine, delicate, nutty','Porridge, pilaf, couscous-style dishes','Cultivars and processing differ by community'],
['Teff','Ethiopia and Eritrea','Earthy, cocoa-like, mildly sweet','Injera, porridge','Flour color and fermentation style vary'],
['Sorghum grain','Africa, South Asia, North America','Mild to nutty, chewy','Porridge, flatbread, flour, popped grain','Many grain and syrup cultivars exist'],
['Amaranth grain','Mesoamerica; cultivated worldwide','Earthy, nutty','Porridge, popped grain, flour','Several species and seed colors are used'],
['Chayote','Mexico and Central America','Mild, crisp to tender','Stew, salad, saute','Local names and shapes vary widely'],
['Napa cabbage','China, Korea, Japan','Sweet, juicy, mild mustard note','Kimchi, dumplings, soup','Korean kimchi varieties are diverse'],
['Bok choy','China; East Asian diaspora cuisines','Juicy stems, mild green leaves','Stir-fry, soup, steamed dishes','Baby and large cultivars differ in texture'],
['Gai lan / Chinese broccoli','Southern China; Cantonese cuisine','Bitter, green, slightly sweet','Blanch, stir-fry, oyster-style sauce','Related Asian brassicas have distinct names and shapes'],
['Amaranth greens','West Africa, East Africa, South Asia, Caribbean','Earthy, mildly mineral','Saute, stew, greens dishes','Multiple edible Amaranthus species and local varieties'],
['Moringa leaves','South Asia and East Africa','Peppery, green','Soup, dal, saute','Fresh and dried leaves differ in intensity'],
['Purslane','Mediterranean, West Asian, Latin American','Lemony, succulent, lightly salty','Salads, stews, pickles','Wild and cultivated forms are eaten'],
['Bitter melon','South Asia, China, Okinawa, Caribbean','Strong bitter, grassy','Stir-fry, curry, stuffed dishes','Several cultivars vary in bitterness and shape'],
['Scotch bonnet chili','Caribbean','Fruity, floral, very hot','Jerk marinades, pepper sauces','Distinct from similarly hot habanero cultivars'],
['Ají amarillo','Peru and Andean region','Fruity, tropical, medium heat','Sauces, stews, marinades','Fresh, frozen and paste forms differ'],
['Poblano / ancho','Mexico','Mild earthy fresh pepper; dried ancho is sweet-fruity','Relleno, mole-style sauces','Ancho is the dried ripe poblano form'],
['Chipotle','Mexico','Smoky, earthy, moderate heat','Adobo, salsa, beans','Typically smoked ripe jalapeño; processing differs'],
['Aleppo pepper','Syria and southern Turkey','Fruity, tangy, mild-medium heat','Finishing spice, kebab, dip','Commercial products vary in salt and oil'],
['Urfa biber','Southeastern Turkey','Raisin-like, smoky, gentle heat','Finishing spice, yogurt, grilled food','Also called isot pepper'],
['Kampot pepper','Kampot, Cambodia','Floral, fruity, bright pepper warmth','Seafood, fruit, finishing','Geographic indication with several color styles'],
['Tellicherry pepper','Malabar Coast, India','Large, mature, citrus-woody peppercorn','Grinder, rubs, sauces','Trade grade usually refers to size/maturity'],
['Ceylon cinnamon','Sri Lanka','Delicate, sweet, floral','Baking, drinks, savory dishes','Distinct from stronger cassia types'],
['Cassia cinnamon','China, Indonesia, Vietnam','Strong, hot, woody','Baking, chai, braise','Several cassia species enter trade'],
['Mexican oregano','Mexico and Central America','Citrusy, floral, peppery','Beans, chili, salsa','Botanically distinct from Mediterranean oregano'],
['Thai basil','Thailand and Southeast Asia','Anise, licorice, peppery','Curry, stir-fry, noodles','Several basil types are used across Southeast Asia'],
['Shiso / perilla','Japan and Korea','Mint, basil, anise; red forms can be cinnamon-like','Wrap, pickle, garnish','Green and red cultivars and preparations differ'],
['Pandan','Southeast Asia','Vanilla-like, grassy, toasted-rice aroma','Rice, custard, cake','Fresh leaf, extract and paste are different forms'],
['Berbere','Ethiopia and Eritrea','Warm, aromatic, chili-hot','Stews, lentils, roast meats','No single fixed recipe; household blends vary'],
['Ras el hanout','Morocco and North Africa','Warm, floral, complex','Tagine, couscous, roast','Blend may contain many spices and differs by maker'],
['Za’atar','Levant','Herbal, tart, sesame, savory','Flatbread, dip, vegetables','Herb and sumac ratios vary regionally'],
['Gochujang','Korea','Sweet-hot, fermented, savory','Stew, glaze, marinade','Fermentation time, grain and heat level vary'],
['Doenjang','Korea','Earthy, salty, fermented umami','Soup, stew, dip','Regional and household recipes differ'],
['Doubanjiang','Sichuan, China','Salty, chili-hot, fermented bean umami','Mapo tofu-style dishes, braises','Bean type, aging and heat vary by producer'],
['Miso','Japan','Sweet-mild white to salty-deep red','Soup, glaze, dressing','Rice, barley and soybean styles differ'],
['Fish sauce','Thailand, Vietnam, Philippines, other coastal regions','Salty, briny, deeply savory','Dipping sauce, curry, dressing','Fish species, fermentation and salt vary'],
['Soy sauce','China, Japan, Korea, Southeast Asia','Salty umami to sweet-dark depending on style','Seasoning, braise, dipping','Light/dark/shoyu/tamari/kecap are not identical products'],
['Kimchi','Korea','Sour, spicy, savory, fermented','Side dish, stew, pancake','Ingredients and fermentation stages vary by region'],
['Mole sauces','Mexico','Complex chili, toasted, bitter-sweet, savory','Sauce for meats, vegetables, enchiladas','Many distinct regional and family recipes'],
['Harissa','Tunisia and North Africa','Chili, garlic, spice, smoky or bright','Paste for stew, marinade, dip','Fresh/dried chili and spice blends vary'],
['Ajvar','Balkans','Roasted pepper, sweet, smoky','Spread, relish, side','Pepper and eggplant proportions vary'],
['Tkemali','Georgia, Caucasus','Tart plum, herbal, fruity','Sauce for grilled foods and potatoes','Green and ripe plum versions differ'],
['Palm oil, red','West and Central Africa','Earthy, rich, carotenoid aroma','Stews, soups, rice dishes','Traditional varieties and processing differ; source sustainably'],
['Ghee','South Asia, Middle East','Nutty, buttery; cultured styles add tang','Tempering, cooking, finishing','Clarification and butter source affect aroma'],
['Yogurt','South Asia, Balkans, Middle East, global','Mild tang to sharp sour; thick or pourable','Marinade, sauce, side, drinks','Milk species, cultures, draining and aging change taste'],
['Feta-style brined cheese','Greece and Eastern Mediterranean','Salty, tangy, briny','Salad, pastry, dip','Milk species and protected designations matter'],
['Halloumi','Cyprus','Salty, tangy, milky; squeaky when grilled','Grill, fry, salad','Milk blend and aging vary'],
['Manchego','La Mancha, Spain','Sheep-milk, nutty, savory','Cheese board, tapas','Aging class changes texture and intensity'],
['Parmigiano Reggiano','Emilia-Romagna, Italy','Savory, nutty, crystalline umami','Grating, finishing, broth rind','Protected designation has defined production area and rules'],
['Gorgonzola','Lombardy/Piedmont, Italy','Blue mold, creamy, sharp to sweet','Sauce, salad, cheese board','Dolce and piccante versions differ'],
['Blue Stilton','England','Salty, tangy, creamy, blue-mold aroma','Cheese board, sauce, soup','Protected regional style; age affects flavor'],
['Roquefort','Aveyron, France','Sharp, salty, sheep-milk blue','Dressing, sauce, cheese board','Protected designation and specific production methods'],
['Sake kasu','Japan','Fermented rice, sweet-sour, savory','Marinade, soup, baking','By-product of sake pressing; aroma varies'],
['Preserved lemon','Morocco and North Africa','Salty, sour, fragrant citrus rind','Tagine, salad, sauce','Salted whole fruit and paste are different forms'],
['Dried shrimp','East and Southeast Asia','Concentrated briny umami','Soup, sambal, stir-fry','Species and drying method affect strength'],
['Dried anchovy','Korea, Japan, Southeast Asia','Salty, oceanic umami','Broth, stir-fry, condiment','Size and species vary; may be head-on or cleaned'],
['Kombu','Japan; East Asian coastal traditions','Mineral, glutamate-rich umami','Dashi, beans, broth','Kelp species and iodine content vary'],
['Nori','Japan, Korea','Roasted, briny, delicate umami','Sushi, rice seasoning, garnish','Sheet grade, roast and seasoning differ'],
['Matsutake mushroom','Japan, Korea, China, Himalayan regions','Spicy pine, earthy, aromatic','Rice, soup, grill','Wild species and local naming differ; availability seasonal'],
['Porcini / king bolete','Europe, Asia, North America','Nutty, earthy, rich umami','Risotto, soup, sauce, dried seasoning','Fresh and dried forms have different intensity'],
['Morel mushroom','Europe, North America, Asia','Earthy, nutty, deep aroma','Cream sauce, saute, stuffing','Must be cooked; species and season matter'],
['Chanterelle mushroom','Europe, North America, Asia','Apricot-like, peppery, earthy','Saute, pasta, sauce','Wild species identification requires expertise'],
['Oyster mushroom','East Asian cultivation; global','Mild, savory; crisp edges when seared','Stir-fry, roast, soup','Several species and cluster forms are sold'],
['Sea grapes / green caviar','Okinawa, Southeast Asia, Pacific','Briny, crisp pop','Salad, sashimi accompaniment','Fresh seaweed handling and salinity vary'],
['Cacao origin: Ghana','West Africa','Cocoa-forward, nutty, fruity notes by lot','Chocolate, baking, drinks','Flavor depends on cultivar, fermentation and roasting'],
['Cacao origin: Ecuador Arriba/Nacional','Ecuador','Floral, fruity, complex cocoa','Chocolate, desserts','Named types and sourcing claims should be verified'],
['Vanilla, Tahitian type','French Polynesia; Pacific cultivation','Floral, fruity, cherry-like aroma','Custards, cream, drinks','Cultivar and curing affect fragrance'],
['Vanilla, Bourbon type','Madagascar and Indian Ocean islands','Creamy, sweet, classic vanilla','Baking, dessert, extract','Regional trade term; origin and processing vary'],
['Honey, manuka style','New Zealand; Australia','Herbal, earthy, strong aroma','Glaze, tea, dressing','Floral source, grading system and processing vary'],
['Honey, acacia style','Europe; Central Asia','Mild, floral, slow to crystallize','Sweetener, cheese pairing','Botanical source and local standards vary'],
['Honey, buckwheat style','North America; Europe','Dark, malty, robust','Baking, glaze, cheese pairing','Color and intensity depend on floral source'],
['Maple syrup, grade/color styles','Northeastern North America','Light delicate to dark robust caramel','Pancake, glaze, baking','Seasonal run and grade affect aroma'],
['Olive oil, Koroneiki cultivar','Greece','Green, peppery, artichoke notes','Finishing, dressing, cooking','Harvest timing changes bitterness and pungency'],
['Olive oil, Arbequina cultivar','Spain; Mediterranean cultivation','Mild, fruity, almond-like','Dressing, finishing, baking','Cultivar and harvest maturity matter'],
['Olive oil, Picual cultivar','Spain','Bitter, peppery, green tomato notes','Cooking, dressing, finishing','High stability and flavor depend on harvest and storage'],
['Tea, matcha','Japan','Vegetal, sweet, umami, mild bitterness','Drink, dessert, noodles, seasoning','Grade, cultivar and preparation affect taste'],
['Coffee, Ethiopian origin styles','Ethiopia','Floral, citrus, berry notes common in some lots','Brew, dessert, rub','Processing and growing area vary greatly'],
['Coffee, Indonesian origin styles','Indonesia','Earthy, spice, cocoa notes common in some lots','Brew, dessert, rub','Processing, island, cultivar and roast change profile'],
];
originVariantRows.push(
['Granny Smith apple','Australia; cultivated worldwide','Tart, crisp, green aroma','Pie, salad, fresh eating','Ripeness changes acidity and firmness'],
['Fuji apple','Japan; cultivated worldwide','Very sweet, crisp, juicy','Fresh eating, salad, baking','Sweetness and storage texture vary by lot'],
['Cox’s Orange Pippin apple','United Kingdom','Aromatic, balanced sweet-tart, complex','Fresh eating, cider, baking','Seasonal heritage cultivar'],
['Bramley apple','United Kingdom','Sharp, high-acid, breaks down when cooked','Pie, sauce, crumble','Cooking apple with distinctive texture'],
['Concord grape','United States, especially New England','Musky, sweet-tart, aromatic','Juice, jelly, baking','Slip-skin cultivar; fresh table use differs'],
['Muscat grape','Mediterranean, West Asia, global vineyards','Floral, grapey, sweet','Fresh eating, raisins, wine','Large family of cultivars and colors'],
['Thompson Seedless grape','Central Asia origin; widely cultivated','Mild sweet, crisp to chewy when dried','Table grape, raisin','Also known as Sultana in some markets'],
['Hass avocado','Mexico/California origin; global cultivation','Rich, buttery, nutty','Guacamole, salad, toast','Oil content changes with maturity and season'],
['Fuerte avocado','Mexico and Central America','Creamy, mild, slightly nutty','Fresh eating, salad','Thin skin and seasonal availability'],
['San Marzano tomato','Campania, Italy','Sweet, low-acid, dense flesh','Sauce, canning','Protected regional product has specific designation rules'],
['Roma tomato','United States/Italy-associated market type','Meaty, mild, lower moisture','Sauce, roasting, canning','Many plum tomato cultivars are sold as Roma type'],
['Cherokee Purple tomato','United States, Tennessee heritage type','Sweet, smoky, rich, soft flesh','Slicing, salad, sauce','Heirloom variety; growing conditions matter'],
['Brandywine tomato','United States heritage cultivar','Rich, sweet-tart, juicy','Slicing, salad','Multiple named strains and colors exist'],
['Yukon Gold potato','Canada','Buttery, creamy, medium starch','Mash, roast, soup','Storage and growing conditions alter texture'],
['Jersey Royal potato','Jersey, Channel Islands','Nutty, waxy, delicate skin','Boil, steam, salad','Protected geographical indication and short season'],
['Purple Majesty potato','United Kingdom breeding; grown globally','Earthy, slightly nutty, purple flesh','Roast, mash, salad','Pigment retention depends on cooking method'],
['Satsuma mandarin','Japan; cultivated in East Asia and worldwide','Sweet, low-acid, tender segments','Fresh eating, juice, dessert','Early-season citrus; seedlessness is common but variable'],
['Yuzu citrus','Japan, Korea, China','Highly aromatic, tart, floral peel','Ponzu, seasoning, dessert','Used mostly for juice and zest, not eaten like an orange'],
['Finger lime','Australia','Citrus pearls, tart, lime aroma','Seafood, drinks, dessert','Many color and flavor cultivars'],
['Meyer lemon','China origin; cultivated globally','Floral, sweet-tart, less sharp than common lemon','Dessert, preserve, dressing','Hybrid citrus; peel is aromatic'],
['Black garlic','Korea, Japan, global production','Sweet, balsamic, mellow umami','Sauce, dressing, spread','Aged under controlled heat and humidity; not a distinct cultivar'],
['Gochugaru pepper flakes','Korea','Fruity, smoky, mild to hot','Kimchi, stew, seasoning','Coarse/fine grind and heat grades differ'],
['Kashmiri chili','Kashmir, India; similar trade types elsewhere','Mild heat, deep red color, fruity','Curries, tandoori marinades','Market substitutions and color strength vary'],
['Byadgi chili','Karnataka, India','Deep red, wrinkled, aromatic, moderate heat','Masala, oil, chutney','Often valued for color; cultivar and drying matter'],
['Dalle khursani chili','Sikkim and Himalayan region','Fruity, very hot, aromatic','Pickle, sauce, condiment','Local cultivars and heat vary widely'],
['Sichuan pepper, green','Sichuan and southwest China','Citrus, floral, numbing tingle','Hot pot, fish, oil','Distinct from red Sichuan pepper in aroma and intensity'],
['Sumac, Syrian style','Levant','Tart, fruity, red berry-like','Salad, grilled meat, dip','Species and grinding affect color and sourness'],
['Maldon sea salt','Essex, England','Clean saline, delicate flakes','Finishing salt','Texture and crystal size change perceived saltiness'],
['Korean pear','Korea, East Asia','Crisp, very juicy, mild sweet','Fresh eating, marinade, salad','Also called Asian pear; multiple cultivars'],
['Nashi pear','Japan, East Asia','Crisp, watery, floral sweet','Fresh eating, salad, dessert','Cultivar shape and sweetness vary'],
['Rainier cherry','Washington State, United States','Very sweet, floral, delicate','Fresh eating, dessert','Short season; bruises easily'],
['Morello sour cherry','Central and Eastern Europe','Tart, deep fruit aroma','Preserve, pie, sauce','Sour cherry varieties differ in acidity'],
['Rambutan, yellow type','Southeast Asia','Sweet, floral, sometimes more acidic','Fresh eating, fruit salad','Color and flavor vary by species and cultivar'],
['Coconut, King coconut','Sri Lanka','Sweet, aromatic, refreshing water','Drink, dessert','Tender fruit and mature coconut have very different uses'],
['Coconut, macapuno','Philippines','Soft, jelly-like, sweet meat','Dessert, preserves','Naturally abnormal endosperm development cultivar'],
['Saffron, Kashmir','Kashmir, India','Floral, honey-like, earthy','Rice, sweets, broth','High-value spice; authenticity and storage matter'],
['Saffron, La Mancha','Castile-La Mancha, Spain','Hay, honey, floral warmth','Rice, soup, dessert','Protected origin; thread quality and storage matter'],
['Turmeric, Lakadong','Meghalaya, India','Earthy, warm, high color intensity','Curry, pickle, drinks','Local cultivar and curcumin claims vary by source'],
['Cardamom, green Malabar','Kerala, India','Sweet, eucalyptus, citrus, floral','Tea, rice, sweets, curry','Pod size, curing, age and origin affect flavor'],
['Vanilla, Mexican planifolia','Mexico','Spicy, woody, creamy floral','Custard, chocolate, sauce','Curing and source authenticity change profile'],
['Paprika, La Vera smoked','Extremadura, Spain','Sweet to hot, deeply smoked','Stew, sausage, roast vegetables','Protected origin with distinct smoke-drying tradition'],
['Coffee, Blue Mountain','Jamaica','Mild, balanced, low harshness','Brew, dessert','Origin label is controlled; grade and roast matter'],
['Coffee, Kona','Hawaii, United States','Mild, nutty, bright, smooth','Brew, dessert','Origin and blend labeling rules vary by market'],
['Tea, Darjeeling first flush','West Bengal, India','Light, floral, muscatel, brisk','Brew, dessert pairing','Season, estate, cultivar and processing vary'],
['Tea, Assam second flush','Assam, India','Malty, rich, brisk','Brew, milk tea','Flush and orthodox/CTC processing create distinct styles'],
['Tea, Uji matcha','Kyoto, Japan','Sweet umami, grassy, lingering','Tea, wagashi, dessert','Shade, milling, grade and preparation affect taste']
);
originVariantRows.push(
['Koshihikari rice','Japan; cultivated globally','Sweet, glossy, tender, lightly sticky','Sushi, rice bowls, steamed rice','Cultivation area and polishing affect texture'],
['Sasanishiki rice','Miyagi, Japan','Clean, delicate, less sticky than some table rices','Sushi, steamed rice','Regional cultivar with a distinct cooking texture'],
['Indica red rice','South Asia; Southeast Asia','Nutty, earthy, bran-rich','Rice plates, porridge, salad','Many landraces and processing forms; cook time varies'],
['Black glutinous rice','Thailand; Southeast Asia','Sweet, nutty, floral','Dessert, porridge, rice blend','Anthocyanin-rich bran; soak for more even cooking'],
['Japonica barley / mochimugi','Japan; East Asia','Nutty, chewy','Rice blend, porridge, soup','Pearled and whole forms cook differently'],
['Tepary bean, white','Sonoran desert','Mild, earthy, creamy','Stew, salad, beans and grains','Landraces differ in color and cooking time'],
['Borlotti bean','Italy; cultivated worldwide','Nutty, creamy, lightly sweet','Pasta e fagioli, salad, stew','Fresh pod and dried bean are different market forms'],
['Flageolet bean','France','Delicate, mild, buttery','Lamb side, salad, stew','Harvest maturity changes color and texture'],
['Gigante bean','Greece; Mediterranean','Mild, creamy, large bean flavor','Baked beans, stew, salad','Several large white bean types are sold under similar names'],
['Cranberry bean','Americas; Italian market forms','Nutty, creamy, rich','Soup, braise, salad','Speckled color fades during cooking'],
['Piquillo pepper','Navarra, Spain','Sweet, roasted, gently tangy','Stuffed pepper, tapas, sauce','Often sold roasted and peeled in jars'],
['Padron pepper','Galicia, Spain','Mild, green, occasionally hot','Blistered tapas, side dish','Heat varies; small peppers are typically mild'],
['Shishito pepper','Japan','Sweet, grassy, occasionally hot','Blistered, tempura, garnish','One in a batch may be hotter; heat is variable'],
['Aji charapita','Peruvian Amazon','Fruity, citrusy, very hot','Salsa, hot sauce, seasoning','Small wild-type chili; use sparingly'],
['Piri piri chili','Southern Africa; Portuguese-speaking regions','Bright, sharp, hot','Sauce, marinade, grill','Name refers to different Capsicum types and sauces'],
['Kampot white pepper','Kampot, Cambodia','Floral, clean, warm pepper','Fish, sauces, finishing','Ripe berries are soaked and hulled; origin claims vary by grade'],
['Timut pepper','Nepal','Grapefruit, citrus, numbing','Fish, chocolate, fruit, spice blends','Often grouped with related Zanthoxylum peppers'],
['Aleppo olive cultivar','Syria; Eastern Mediterranean','Fruity, buttery, gently bitter','Table olive, oil, preserve','Curing method and harvest stage change taste'],
['Nocellara del Belice olive','Sicily, Italy','Buttery, mild, green-fruited','Table olive, oil','Protected origin product and cultivar use vary by processing'],
['Comté, young','Jura, France','Milky, nutty, fruity','Table cheese, sandwich, gratin','Aging develops stronger savory and crystalline notes'],
['Pecorino Sardo','Sardinia, Italy','Sheep-milk, salty, grassy, sharp with age','Grating, cheese board, pasta','Aging class and local production affect intensity'],
['Ossau-Iraty','Basque Country and Béarn, France','Sheep-milk, nutty, gentle fruit','Cheese board, cooking','Protected regional cheese; aging varies'],
['Manuka honey, East Cape style','New Zealand','Herbal, earthy, sometimes medicinal','Glaze, drink, cheese pairing','Floral source and grading require verified labeling'],
['Darjeeling second flush, muscatel lot','West Bengal, India','Muscat grape, floral, brisk','Tea, dessert pairing','Season and estate shape the profile; lot variation is large'],
['Gyokuro tea','Japan','Sweet umami, marine, shaded green aroma','Brew, rice, dessert pairing','Shade duration and brewing temperature matter'],
['Oolong, Dong Ding','Taiwan','Floral, toasted, creamy','Tea service, dessert pairing','Roast level and harvest season vary by maker'],
['Cacao, Madagascar Sambirano','Madagascar','Red fruit, citrus, lively acidity','Chocolate, pastry, drink','Fermentation and roast strongly shape flavor'],
['Cacao, São Tomé','Gulf of Guinea islands','Cocoa, red fruit, earthy notes','Chocolate, baking','Farm, fermentation and drying distinguish lots'],
['Black lime, Omani dried lime','Oman; Gulf region','Tart, fermented, smoky citrus','Rice, broth, stew','Whole or ground forms have different intensity']
);

originVariantRows.push(
['Sona masuri rice','Andhra Pradesh and Telangana, India','Light, mild, medium-grain, less aromatic','Everyday rice, idli batter, pongal','Milling and age alter softness'],
['Gobindobhog rice','West Bengal, India','Short-grain, fragrant, buttery aroma','Khichuri, pulao, sweets','Aromatic landrace; crop region affects quality'],
['Kalanamak rice','Uttar Pradesh, India','Fragrant, nutty, soft','Steamed rice, kheer','Traditional aromatic variety; milling changes aroma'],
['Joha rice','Assam, India','Aromatic, delicate, floral','Steamed rice, dessert','Multiple local cultivars use this regional name'],
['Matta rice / Kerala red rice','Kerala, India','Nutty, earthy, robust','Rice meals, puttu, porridge','Parboiling and bran retention affect texture'],
['Kavuni rice','Tamil Nadu, India','Nutty, earthy, dark bran aroma','Sweet pudding, rice dishes','Traditional black rice type; soak and cook thoroughly'],
['Kalijira rice','Bangladesh','Tiny-grain, aromatic, delicate','Pulao, steamed rice','Small grain size gives fast cooking'],
['Katarni rice','Bihar, India','Fragrant, soft, mildly sweet','Rice dishes, festive meals','Named regional variety; availability is seasonal'],
['Konde rice, red landrace','Goa; Konkan, India','Earthy, nutty, firm','Rice plates, porridge','Local landraces and milling vary'],
['Black turtle bean','Central America; Latin America','Earthy, dense, slightly sweet','Rice and beans, soup, refried beans','Skin color and cooking time vary by lot'],
['Mayocoba bean / peruano','Mexico','Buttery, mild, creamy','Refried beans, soup, side dish','Yellow bean type; dried age affects cook time'],
['Anasazi bean','Southwestern United States; Indigenous cultivation','Sweet, earthy, quick-cooking','Stew, chili, bean salad','Named market type; seed coat pattern varies'],
['Jacob’s cattle bean','United States heritage type','Earthy, hearty, creamy','Baked beans, soup, stew','Heirloom seed strain and region vary'],
['Christmas lima bean','Peru; cultivated in the Americas','Chestnut-like, buttery','Stew, braise, salad','Large flat bean; cook thoroughly'],
['Scarlet runner bean','Central America; European gardens','Rich, earthy, slightly sweet','Stew, soup, fresh pod dishes','Cook dried seeds thoroughly; raw beans are unsafe'],
['Cavolo nero / Tuscan kale','Tuscany, Italy','Deep green, mildly bitter, sweetens with cooking','Soup, pasta, braise','Leaf age changes tenderness'],
['Romanesco cauliflower','Italy; cultivated worldwide','Nutty, mild, slightly sweet','Roast, steam, pasta, crudite','Broccoli-cauliflower relative; florets cook evenly when separated'],
['Rapa da foglia / cime di rapa','Southern Italy','Bitter, peppery, green','Pasta, saute, sausage pairing','Young leaves and flower buds have different bitterness'],
['Naga king chili','Nagaland, India; Northeast South Asia','Fruity, floral, extremely hot','Pickle, chutney, sauce','Use tiny quantities; heat varies and can be intense'],
['Biquinho pepper','Brazil','Sweet, fruity, very mild','Pickle, garnish, sauce','Often sold pickled; fresh varieties vary'],
['Rocoto pepper','Andes, especially Peru and Bolivia','Fruity, crisp, hot; black seeds','Rocoto relleno, salsa, sauce','Thick flesh; hotter than many common peppers'],
['Aji limo','Peru','Citrus, floral, hot','Ceviche, salsa, marinade','Several colors and local types exist'],
['Tosa red ginger','Kochi, Japan','Fresh, floral, mildly spicy','Pickles, garnish, sushi','Young stems and rhizome are used differently'],
['Wasabi, Shizuoka type','Shizuoka, Japan','Fresh green heat, sweet finish','Grated condiment, soba, seafood','Authentic rhizome differs from many horseradish blends'],
['Cacao, Chuao origin','Aragua, Venezuela','Nutty, red fruit, balanced cocoa','Chocolate, couverture, pastry','Origin name and lot quality require traceability'],
['Coffee, Yirgacheffe washed','Gedeo, Ethiopia','Floral, citrus, tea-like','Pour-over, espresso blend','Variety, washing and roast affect cup profile'],
['Coffee, Mandheling style','Sumatra, Indonesia','Earthy, herbal, cocoa, full-bodied','Brew, espresso, dessert','Trade style; processing varies by producer'],
['Tea, Lapsang Souchong','Wuyi, China','Smoky, resinous, sweet black tea','Brew, cooking, marinade','Smoke intensity differs by producer'],
['Tea, Ceylon Uva seasonal','Uva, Sri Lanka','Brisk, floral, sometimes minty','Brew, iced tea','Season, elevation and processing shape the cup'],
['Olive, Kalamata type','Peloponnese, Greece','Fruity, wine-like, mildly bitter','Table olive, salad, tapenade','Protected name applies to defined origin/product specifications'],
['Olive, Taggiasca cultivar','Liguria, Italy','Sweet, delicate, almond-like','Table olive, oil, sauce','Small fruit; brining and oil styles differ'],
['Cheese, Brin d’Amour','Corsica, France','Sheep-milk, herbal, floral rind','Cheese board','Rind herbs and aging vary by maker'],
['Cheese, Queso de Cabrales','Asturias, Spain','Intense blue, salty, sharp','Cheese board, sauce','Protected origin; cave aging shapes flavor'],
['Cheese, Taleggio','Lombardy, Italy','Fruity, tangy, buttery, washed-rind aroma','Polenta, sandwich, risotto','Ripeness changes softness and pungency'],
['Honey, sidr floral source','Yemen; Arabian Peninsula','Floral, fruity, deep sweetness','Sweetener, drink, cheese pairing','Botanical identity and geographic claims need verification'],
['Chili, Kashmiri Degi type','North India; South Asia trade','Mild heat, bright red, earthy','Tandoori marinade, curry, oil','Commercial blends may mix cultivars; check product claims']
);

originVariantRows.push(
['Koshikari black rice','Japan','Sweet, earthy, dark bran, mildly sticky','Rice bowls, porridge, dessert','Regional cultivar names and products vary'],
['Swarna rice','Eastern India; Bangladesh','Mild, soft, everyday rice aroma','Steamed rice, khichuri, daily meals','Cultivation and parboiling affect texture'],
['Gobindobhog, aromatic parboiled type','West Bengal, India','Fragrant, nutty, richer cooked aroma','Festive rice, khichuri, sweets','Processing and cultivar identity vary by producer'],
['Ponni rice','Tamil Nadu, India','Soft, mild, lightly fragrant','Everyday rice, idli, dosa batter','Raw and parboiled forms cook differently'],
['Jeera samba rice','Tamil Nadu, India','Aromatic, small-grain, savory','Biryani, pulao','Regional name refers to fragrant small-grain rice types'],
['Kavuni, red-black landrace','Tamil Nadu, India','Nutty, earthy, mildly sweet','Pudding, rice meals','Color and cooking behavior vary by local seed stock'],
['Korean black rice, Heukmi','Korea','Nutty, sweet, earthy','Rice blend, porridge','Often mixed with white rice; cultivar and polish matter'],
['Red cargo rice, Thai type','Thailand','Nutty, earthy, bran-rich','Rice plates, salad, porridge','Soaking and parboiling alter cook time'],
['Sticky rice, purple Lao type','Laos; Northeast Thailand','Sweet, floral, earthy','Steamed sticky rice, dessert','Pigmented glutinous rice; soak before steaming'],
['Pigeon pea, Congo brown type','Caribbean; East Africa','Earthy, mild, creamy','Rice and peas, stew, dhal','Fresh and dried forms have different cooking times'],
['Split pea, yellow type','South Asia; Europe','Mild, sweet, earthy','Soup, dal, puree','Split skinless peas cook faster than whole peas'],
['Black soybean, Korean seoritae','Korea','Sweet, nutty, rich','Rice mix, braise, snack','Green interior when mature; cook thoroughly'],
['Anasazi bean, white-spotted type','Southwestern United States','Mild, sweet, earthy','Stew, chili, bean salad','Seed pattern and flavor vary by seed line'],
['Peanut, Valencia type','New Mexico, United States','Sweet, rich, nutty','Roast, boiled peanut, sauce','Usually multiple kernels per shell; allergen'],
['Peanut, Spanish type','United States; Spain trade type','Sweet, rich, red skin, high oil','Confectionery, roast, oil','Market type; not necessarily Spanish-grown'],
['Tomato, Costoluto Fiorentino','Tuscany, Italy','Sweet, tangy, rich aroma','Sauce, salad, canning','Ribbed heritage cultivar; growing conditions matter'],
['Tomato, Green Zebra','United States breeding; grown globally','Tart, bright, green-fruited','Salad, salsa, roast','Green when ripe; color is not a ripeness indicator alone'],
['Tomato, Oxheart type','Italy; cultivated worldwide','Sweet, mild, low seed gel','Slicing, sauce','Multiple oxheart cultivars exist'],
['Eggplant, Rosa Bianca','Sicily, Italy; global gardens','Mild, creamy, gently sweet','Roast, grill, braise','Cultivar color and bitterness depend on maturity'],
['Eggplant, Thai green','Thailand; Southeast Asia','Bitter, vegetal, firm','Curry, stir-fry, pickle','Small round fruits remain firm in sauces'],
['Carrot, Chantenay','France; cultivated worldwide','Sweet, earthy, aromatic','Roast, braise, soup','Short, broad-root cultivar; soil affects sweetness'],
['Carrot, Nantes type','France; cultivated worldwide','Sweet, crisp, mild','Fresh eating, juice, roast','Market type includes several cultivars'],
['Melon, Galia','Israel breeding; Mediterranean cultivation','Floral, sweet, tropical','Fresh eating, dessert, salad','Netting and aroma help indicate maturity'],
['Melon, Hami','Xinjiang, China','Sweet, crisp, floral','Fresh eating, fruit plate','Many regional Hami cultivars and harvest times'],
['Melon, Charentais','Charente, France; grown globally','Very aromatic, floral, sweet','Fresh eating, dessert','Ripeness and harvest timing strongly affect aroma'],
['Citrus, sudachi','Tokushima, Japan','Sharp, green, aromatic','Squeeze over fish, noodles, sauce','Used mostly as juice and zest'],
['Citrus, kabosu','Oita, Japan','Mild tart, floral, fragrant','Seafood, ponzu, hot pot','Harvest maturity changes color and acidity'],
['Chili, Aleppo Halaby type','Aleppo region; Levant','Fruity, raisin-like, gentle heat','Finishing spice, grilled foods','Trade names can include blended or substitute peppers'],
['Chili, Tien Tsin','China; East Asian market type','Clean, sharp, medium-hot','Stir-fry, chili oil, dry-fried dishes','Heat varies with age and growing conditions'],
['Sichuan pepper, red Da Hong Pao type','Sichuan, China','Citrus, floral, numbing, warm','Hot pot, chili oil, spice blends','Cultivar and harvest maturity affect aroma'],
['Cumin, Gujarat type','Gujarat, India','Warm, earthy, citrus edge','Tempering, spice blends, breads','Seed size and drying alter aroma'],
['Coriander seed, Moroccan type','Morocco','Citrusy, floral, warm','Ras el hanout, stew, bread','Whole and ground forms differ in freshness'],
['Mint, spearmint Nana','Morocco; North Africa','Sweet, cool, mild menthol','Tea, salad, lamb, drink','Cultivar and drying change intensity'],
['Thyme, Za’atar wild oregano type','Levant','Herbal, resinous, savory','Za’atar blend, bread, grilled foods','Common names cover multiple species; verify botanical identity'],
['Coffee, Huehuetenango washed','Guatemala','Citrus, floral, crisp acidity','Brew, espresso, dessert','Farm elevation, cultivar, processing and roast vary'],
['Coffee, Sidamo natural','Ethiopia','Berry, floral, sweet fruit','Brew, dessert','Natural processing gives fruit notes that vary by lot'],
['Tea, Assam orthodox orthodox leaf','Assam, India','Malty, brisk, sometimes honeyed','Brew, milk tea','Orthodox leaf differs from CTC; grade and estate vary'],
['Tea, Nilgiri frost tea','Tamil Nadu, India','Fragrant, bright, brisk','Brew, iced tea','Season and elevation affect aroma'],
['Tea, Puer sheng raw aged cake','Yunnan, China','Bitter-sweet, floral to earthy with age','Tea service, pairing','Storage and age produce large sensory changes'],
['Cheese, Livarot','Normandy, France','Savory, pungent, creamy','Cheese board, cooking','Washed-rind aroma strengthens with age'],
['Cheese, Reblochon','Haute-Savoie, France','Creamy, nutty, gently pungent','Tartiflette, cheese board','Protected regional style; aging affects texture'],
['Cheese, Idiazabal smoked','Basque Country and Navarre, Spain','Sheep-milk, smoky, nutty','Cheese board, grated dishes','Smoke level and aging vary within designation rules'],
['Cheese, Cotija añejo','Mexico','Salty, dry, sharp, crumbly','Beans, elote, soups','Age and moisture determine grating texture'],
['Olive, Arbequina table green','Catalonia, Spain','Mild, fruity, buttery','Table olive, snack, oil','Brining styles create additional flavor differences'],
['Honey, chestnut floral source','Mediterranean; Europe','Dark, bitter-sweet, woody','Cheese, baking, glaze','Floral source and geography affect intensity'],
['Honey, orange blossom floral source','Mediterranean; California; global','Citrus, floral, light sweetness','Dessert, tea, dressing','Botanical source and season influence aroma']
);

originVariantRows.push(
['Tetsukabuto squash','Japan; Brazil cultivation','Sweet, dense, nutty','Roast, soup, curry, puree','Hybrid winter squash; storage changes sweetness'],
['Red kuri squash','Japan; East Asian cultivation','Chestnut-like, sweet, earthy','Roast, soup, puree','Thin edible skin when cooked; cultivar size varies'],
['Jarrahdale pumpkin','Australia','Mild, sweet, dense orange flesh','Roast, soup, pie','Blue-grey skin; maturity affects flavor'],
['Delicata squash','North America','Sweet, corn-like, mild nutty notes','Roast, stuff, saute','Thin skin is edible when cooked'],
['Tromboncino squash','Italy','Mild, sweet, firm','Saute, roast, pasta, pickle','Young fruit is used as summer squash; mature fruit stores'],
['Futsu black squash','Japan','Nutty, sweet, chestnut-like','Roast, soup, puree','Rind darkens as fruit matures'],
['Pattypan squash, yellow type','Americas; global cultivation','Mild, buttery, slightly sweet','Grill, stuff, roast, saute','Young fruit is most tender'],
['Culantro, long-leaf strain','Caribbean; Southeast Asia','Strong cilantro, pungent, citrusy','Soup, sofrito, seasoning paste','Leaf size and aroma vary by cultivar'],
['Pak choi, Shanghai green type','China; East Asia','Mild, sweet, juicy','Stir-fry, soup, steam','Green-stemmed form differs from white-stem types'],
['Cabbage, pointed hispi type','Northern Europe','Sweet, mild, delicate mustard','Roast, braise, slaw','Conical head; cooks faster than dense storage cabbage'],
['Radicchio, Chioggia','Veneto, Italy','Bitter, peppery, lightly sweet','Salad, grill, risotto','Round red head; bitterness varies with season'],
['Endive, Belgian witloof','Belgium; France','Bitter, crisp, slightly sweet when cooked','Salad, braise, gratin','Grown blanched; light exposure increases bitterness'],
['Corn, Oaxacan green','Oaxaca, Mexico','Earthy, sweet, mineral','Tortilla, atole, tamale','Local landrace; color and texture vary'],
['Corn, Hopi blue flour type','Southwestern United States','Sweet, earthy, robust corn flavor','Blue cornmeal, bread, porridge','Indigenous seed varieties; avoid treating as one uniform cultivar'],
['Corn, glass gem type','United States breeding','Mild, sweet, decorative multicolor kernels','Meal, popcorn, ornamental food use','Kernel quality varies; not every ear is ideal for fresh eating'],
['Bean, ayocote negro','Mexico','Earthy, rich, creamy','Bean stew, salad, soup','Large runner bean type; cooking time varies'],
['Bean, ayocote morado','Mexico','Nutty, earthy, mildly sweet','Soup, stew, bean salad','Color fades during cooking'],
['Bean, tepary, brown landrace','Sonoran desert','Nutty, earthy, dense','Stew, beans and grains','Landraces differ in drought adaptation and cook time'],
['Pea, marrowfat type','United Kingdom; Europe','Sweet, starchy, creamy','Mushy peas, soup, snack','Maturity and drying influence texture'],
['Pea, pigeon pea, Congo type','Caribbean; Central Africa','Earthy, nutty, lightly sweet','Rice, stew, dhal','Fresh green and dried pulses differ greatly'],
['Potato, La Ratte fingerling','France','Nutty, buttery, waxy','Roast, salad, saute','Small elongated cultivar; harvest age changes waxiness'],
['Potato, Désirée','Netherlands; Europe','Earthy, creamy, red-skinned','Mash, roast, bake','Starch level shifts with growing and storage'],
['Potato, Andean papas nativas, purple type','Peru; Bolivia','Earthy, nutty, sometimes floral','Boil, roast, mash, stew','A broad group of landraces rather than one cultivar'],
['Sweet potato, Okinawan purple type','Okinawa; Japan','Sweet, dry, earthy, floral','Roast, dessert, noodles','Distinct from ube; flesh color and moisture vary'],
['Sweet potato, Beauregard','United States; global cultivation','Sweet, moist, orange-fleshed','Bake, mash, fry, pie','Modern cultivar; storage increases sweetness'],
['Yam, water yam white type','West Africa; Caribbean','Mild, starchy, slightly sweet','Boil, pound, stew','Species and varieties differ in texture'],
['Yam, yellow guinea type','West Africa','Earthy, sweet, dense','Boil, roast, pounded yam','Cook fully; cultivar and maturity affect texture'],
['Ginger, Cochin type','Kerala, India','Citrus, warm, sweet-spicy','Pickle, tea, curry, preserve','Fresh ginger differs from dried spice; age changes fiber'],
['Garlic, Rocambole type','Europe; North America','Rich, complex, pungent','Roast, sauce, pickle','Hardneck cultivar group; storage and harvest affect pungency'],
['Onion, Tropea red type','Calabria, Italy','Sweet, mild, juicy','Fresh salad, pickle, roast','Geographic name and cultivar use vary by product'],
['Chili, Aji dulce','Caribbean','Fruity, aromatic, nearly no heat','Sofrito, stew, sauce','Heat varies by strain; distinct from hot ají types'],
['Chili, Guajillo','Mexico','Mild heat, berry, tea, tannic notes','Mole, adobo, sauce','Dried ripe Mirasol pepper; rehydrate before blending'],
['Chili, Pasilla Oaxaca','Oaxaca, Mexico','Smoky, raisin-like, moderate heat','Mole, salsa, bean dishes','Smoke-dried regional chili; not the same as standard pasilla'],
['Coffee, Tarrazú washed','Costa Rica','Bright, citrus, chocolate, balanced','Brew, espresso','Microregion and lot influence cup profile'],
['Coffee, Antigua Guatemala washed','Guatemala','Cocoa, spice, balanced acidity','Brew, espresso','Volcanic soils and processing vary across estates'],
['Tea, Longjing / Dragon Well, West Lake','Zhejiang, China','Chestnut, sweet, green, delicate','Brew, tea cuisine','Origin designation and harvest grade matter'],
['Tea, Tieguanyin, Anxi traditional roast','Fujian, China','Orchid, toasted, creamy','Brew, dessert pairing','Roast and oxidation vary by maker'],
['Tea, Ceylon Dimbula high-grown','Sri Lanka','Bright, brisk, citrus, floral','Brew, iced tea','Season and elevation create batch variation'],
['Cacao, Piura white criollo type','Piura, Peru','Floral, red fruit, delicate cocoa','Chocolate, dessert','Genetic and origin claims require lot traceability'],
['Honey, heather floral source','Scotland; Northern Europe','Herbal, smoky, thick, bittersweet','Cheese, glaze, baking','Floral source, geography and crystallization affect texture'],
['Cheese, Garrotxa','Catalonia, Spain','Goat-milk, nutty, herbaceous','Cheese board, salads','Aging and rind treatment vary by producer'],
['Cheese, Mahón-Menorca semi-cured','Balearic Islands, Spain','Salty, buttery, tangy','Table cheese, sandwich, gratin','Protected regional cheese; age affects sharpness']
);

originVariantRows.push(
['Fonio, Guinea landrace','Guinea; West Africa','Delicate, nutty, lightly floral','Porridge, couscous, pilaf','Local seed and threshing methods vary'],
['Sorghum, red kafir type','Southern Africa','Earthy, mild, slightly sweet','Porridge, meal, malt','Many regional sorghum landraces share trade names'],
['Millet, kodo brown type','South Asia','Nutty, earthy, robust','Porridge, pilaf, fermented foods','Seed coat and processing affect cooking'],
['Millet, proso yellow type','Eurasia; North America','Mild, sweet, slightly grassy','Porridge, pilaf, flour','Cultivar and dehulling change texture'],
['Barley, hulless Tibetan type','Tibet; Himalayan regions','Nutty, earthy','Tsampa, porridge, flour','Landraces vary; barley contains gluten'],
['Chickpea, desi red type','South Asia; Ethiopia','Earthy, robust, thick skin','Chana, flour, stew','Desi and kabuli chickpeas differ in size and texture'],
['Cowpea, black-eyed cream type','West Africa; United States South','Mild, creamy, earthy','Rice, stew, fritter','Seed eye color and cooking time vary by cultivar'],
['Cowpea, purple hull type','Southern United States','Sweet, fresh, bean-like','Fresh shell peas, stew, side','Harvested fresh; not the same as dried black-eyed peas'],
['Yam, water yam purple type','West Africa; Caribbean','Mild, starchy, lightly sweet','Boil, pound, stew','Cook fully; species and cultivar vary'],
['Taro, Bun-long type','Hawaii; Pacific cultivation','Nutty, mildly sweet','Poi, stew, chips','Corm and leaf preparation differ; cook thoroughly'],
['Taro, dasheen type','Caribbean','Earthy, starchy, mildly sweet','Callaloo, soup, boil','Names vary regionally and may refer to different cultivars'],
['Cassava, bitter cultivar','South America; Africa','Neutral, starchy','Flour, gari, fufu, tapioca','Requires traditional detoxification and processing; never eat raw'],
['Cassava, sweet cultivar','South America; Caribbean','Mild, sweet, starchy','Boil, roast, fry','Still requires correct cooking and preparation'],
['Plantain, French horn type','West Africa; Caribbean','Starchy green; sweet when fully ripe','Boil, roast, fry, fufu','Ripeness changes use and sweetness'],
['Plantain, Nendran type','Kerala, India','Starchy, sweet when ripe','Chips, steamed, dessert','Regional cultivar; cooking behavior varies by maturity'],
['Banana, Pisang Awak','Southeast Asia','Sweet, aromatic, slightly tangy','Fresh, cooked dessert, fritter','Market names and cultivars vary regionally'],
['Mango, Kachamitha','Bangladesh; eastern India','Very sour green, fibrous','Pickle, chutney, curry','Used unripe; cultivar names vary by locality'],
['Mango, Langra','North India; Bangladesh','Aromatic, tangy, low fiber','Fresh eating, puree','Seasonal cultivar; ripeness changes acidity'],
['Mango, Sindhri','Sindh, Pakistan','Very sweet, floral, juicy','Fresh eating, drinks','Seasonal cultivar; growing conditions affect aroma'],
['Mango, Nam Dok Mai','Thailand','Floral, sweet, smooth flesh','Fresh eating, sticky rice dessert','Several cultivars use similar names'],
['Citrus, calamondin','Philippines; Southeast Asia','Sharp sour, aromatic peel','Condiment, drink, marinade','Fruit color and acidity vary with maturity'],
['Citrus, bergamot, Reggio type','Calabria, Italy','Floral, bitter, citrus perfume','Tea, confection, zest, marmalade','Oil aroma is intense; mostly used as zest or flavoring'],
['Citrus, fingered lime, green pearls','Australia','Citrus burst, tart, fresh','Seafood, drink, garnish','Pearl color and acidity vary by cultivar'],
['Pepper, Sarawak white','Sarawak, Malaysia','Clean, earthy, warm pepper','Soup, fish, sauce','Processing and aging change aroma'],
['Pepper, Penja white','Cameroon','Earthy, woody, aromatic heat','Fish, meat, sauce','Regional product; origin labels should be verified'],
['Clove, Zanzibar type','Zanzibar; Tanzania','Sweet, warm, pungent','Rice, spice blends, pickles','Bud maturity and drying alter oil content'],
['Cardamom, black Nepal type','Nepal; Himalayan regions','Smoky, resinous, camphoraceous','Stew, rice, spice mix','Large black cardamoms are distinct from green types'],
['Cacao, Beniano wild type','Bolivia, Amazon basin','Fruit-forward, earthy, complex cocoa','Chocolate, drink, dessert','Wild-origin and cultivar claims need traceability'],
['Coffee, Bourbon Pointu','Réunion Island','Light, floral, gentle acidity','Filter brew, dessert','Rare cultivar; availability and roast profile vary'],
['Tea, Kenyan purple leaf','Kenya','Fruity, brisk, floral, mild astringency','Brew, iced tea','Anthocyanin-bearing cultivar; processing changes flavor'],
['Tea, Nepal Ilam orthodox','Ilam, Nepal','Floral, muscatel, bright','Brew, tea pairing','Estate, season and oxidation strongly affect profile'],
['Cheese, Tête de Moine','Bernese Jura, Switzerland','Nutty, floral, savory','Rosettes, cheese board','Aging changes aroma and texture'],
['Cheese, Fynbo','Denmark','Mild, buttery, nutty','Table cheese, sandwich','Aging and producer styles vary'],
['Cheese, Vacherin Mont d’Or','Jura, France and Switzerland','Creamy, woodsy, buttery','Baked cheese, spooned service','Seasonal product; protected designations have specific rules'],
['Honey, tupelo floral source','Southeastern United States','Light, buttery, floral','Tea, dressing, cheese pairing','Floral claims and sourcing should be verified'],
['Honey, manuka floral source, multifloral','New Zealand; Australia','Earthy, herbal, robust','Glaze, drink, cheese pairing','Grading and floral source are distinct product claims']
);

originVariantRows.push(
['Teff, Magna type','Ethiopia','Nutty, mild, earthy','Injera, porridge, flour','Teff seed color and local landrace influence flavor'],
['Sorghum, durra type','Northeast Africa; South Asia','Mild, earthy, slightly sweet','Flatbread, porridge, malt','Durra is a broad crop race with many local varieties'],
['Sorghum, bicolor kafir type','Southern Africa','Mild, cereal-like, earthy','Porridge, flour, brewing','Cultivar and malting determine taste and texture'],
['Pearl millet, Nigerian gero type','West Africa','Nutty, earthy, robust','Tuwo, porridge, fermented foods','Local name covers multiple landraces'],
['Pearl millet, bajra green type','Rajasthan, India','Earthy, sweet, grassy','Flatbread, porridge, fermented batter','Fresh flour has stronger aroma than aged flour'],
['Yam, Dioscorea alata, white type','West Africa; Caribbean; Pacific','Mild, starchy, sweet','Boil, pound, roast','White-flesh varieties differ in moisture and starch'],
['Yam, Dioscorea alata, purple type','Pacific Islands; Caribbean','Earthy, mildly sweet, purple flesh','Boil, mash, dessert','Distinct from purple sweet potato and ube types'],
['Cassava, gari yellow type','West Africa','Sour, nutty, toasted','Eba, porridge, snack','Fermentation, roasting and palm oil change color and flavor'],
['Cassava, farinha d’água','Amazon Basin, Brazil','Tangy, toasted, crisp','Farofa, fish, rice dishes','Traditional fermentation and drying vary by region'],
['Plantain, French horn, green stage','West Africa; Caribbean','Starchy, mild, savory','Boil, roast, fufu','Green stage is less sweet than ripe fruit'],
['Plantain, French horn, ripe stage','West Africa; Caribbean','Sweet, caramelized, fruity','Fry, roast, dessert','Ripeness transforms starch to sugars'],
['Banana, Gros Michel','Caribbean heritage; global historical trade','Sweet, creamy, classic banana aroma','Fresh eating, dessert','Current availability and disease pressure vary by region'],
['Banana, Blue Java','Southeast Asia; cultivated in tropics','Sweet, creamy, vanilla-like aroma','Fresh eating, dessert','Common name covers a cultivar group; flavor varies by maturity'],
['Mango, Keitt','Florida, United States; global cultivation','Sweet, mild, low fiber','Fresh eating, pickle, puree','Late-season cultivar; green skin may remain when ripe'],
['Mango, Tommy Atkins','Florida, United States; global trade','Mild sweet, firm, lightly floral','Fresh eating, salsa, export fruit','Trade cultivar; fiber and aroma depend on maturity'],
['Pineapple, Queen type','South Africa; Australia; Indian Ocean','Sweet, aromatic, firm, tangy','Fresh eating, preserve, grill','Small fruit, concentrated aroma; cultivar varies'],
['Pineapple, MD-2 type','Hawaii breeding; global cultivation','Sweet, low acidity, uniform gold flesh','Fresh fruit, juice, export','Commercial cultivar; harvest maturity changes sweetness'],
['Papaya, Red Lady type','Taiwan breeding; global cultivation','Sweet, musky, orange-red flesh','Fresh eating, drink, salad','Cultivar is widely grown; maturity affects aroma'],
['Guava, Allahabad Safeda','Uttar Pradesh, India','Sweet, mild, aromatic, white flesh','Fresh eating, juice, dessert','Seed count and firmness vary by maturity'],
['Guava, Thai pink type','Thailand; Southeast Asia','Crisp, mild, sweet, pink flesh','Fresh snack, salad, juice','Multiple cultivars sold under market names'],
['Citrus, yuzu, Hana Yuzu type','Japan','Floral, intense citrus peel, tart juice','Ponzu, zest, drinks','Cultivar selection and harvest time change aroma'],
['Citrus, kabosu, green harvest','Oita, Japan','Sharp, herbal citrus, high acidity','Seafood, ponzu, noodle dishes','Green and yellow ripe fruit differ in aroma'],
['Citrus, sudachi, Tokushima type','Tokushima, Japan','Bright, sharp, floral citrus','Fish, noodles, sauce','Regional cultivar and season affect acidity'],
['Chili, Piment d’Espelette','Basque Country, France','Mild, fruity, gently smoky','Egg, fish, sauce, finishing','Protected regional product; heat is moderate and variable'],
['Chili, Aji Panca','Peru','Smoky, berry-like, mild heat','Adobo, stew, marinade','Fresh and dried forms differ in pungency'],
['Chili, Aji Mirasol','Peru','Fruity, bright, medium heat','Sauce, stew, rehydrated chili paste','Dried form of aji amarillo types; names vary'],
['Pepper, Kampot red','Kampot, Cambodia','Fruity, sweet, warm pungency','Fruit, seafood, finishing','Fully ripe dried pepper; handling affects aroma'],
['Pepper, Lampong black','Sumatra, Indonesia','Sharp, earthy, robust heat','Grinder, rub, pickle','Trade origin style; crop year affects pungency'],
['Nutmeg, Banda Islands type','Maluku, Indonesia','Sweet, woody, warm, resinous','Baking, stew, beverage','Mace and nutmeg are separate seed coverings'],
['Vanilla, Papua New Guinea type','Papua New Guinea','Woody, cocoa-like, fruity','Dessert, beverage, baking','Cultivar and curing distinguish lots'],
['Vanilla, Uganda type','Uganda','Chocolate, woody, rich aroma','Baking, chocolate, extract','Curing and origin are variable; traceability matters'],
['Coffee, Harrar natural','Ethiopia','Blueberry, wine-like, spice','Brew, espresso, dessert','Natural processing can produce strong fruit notes, varying by lot'],
['Coffee, Nyeri washed SL28/SL34','Kenya','Blackcurrant, citrus, bright acidity','Filter brew, espresso','Variety, elevation, process and roast shape flavor'],
['Tea, Assam orthodox second flush, Mangalam style','Assam, India','Malty, honeyed, bold','Black tea, milk tea','Estate examples are illustrative; batch characteristics vary'],
['Tea, Yunnan Dianhong gold tips','Yunnan, China','Sweet, cocoa, honey, gentle tannin','Brew, dessert pairing','Bud ratio and grade vary among producers'],
['Tea, Hojicha, Kyoto roasted stem style','Kyoto, Japan','Toasted, nutty, low grassy aroma','Brew, dessert, cooking','Roast intensity and leaf/stem blend differ'],
['Honey, eucalyptus floral source','Australia; Mediterranean cultivation','Herbal, menthol, dark sweetness','Glaze, tea, cheese','Floral source and geography change intensity'],
['Honey, acacia floral source, Robinia','Europe; Central Asia','Light, floral, mild sweetness','Tea, cheese, baking','Common name and botanical source should be checked on label'],
['Cheese, Oscypek','Polish Carpathians','Smoky, salty, sheep-milk tang','Grill, snack, cheese board','Protected regional style; seasonal production varies'],
['Cheese, Sirene brined','Balkans','Salty, tangy, crumbly','Salad, pastry, pastry filling','Milk blend and brine strength vary by country'],
['Cheese, Halloumi matured','Cyprus','Salty, tangy, firmer, savory','Grill, fry, salad','Aging makes flavor stronger than fresh halloumi'],
['Olive, Castelvetrano Nocellara green','Sicily, Italy','Buttery, mild, green-fruited','Table olive, antipasto','Curing process affects salinity and crunch'],
['Olive, Amfissa black','Central Greece','Fruity, soft, wine-like','Table olive, salad, tapenade','Ripeness and curing produce different styles'],
['Sea salt, Guérande grey sel','Brittany, France','Mineral, moist saline','Cooking, finishing','Moisture and grain size change salinity perception'],
['Sea salt, Maldon smoked flakes','Essex, England','Clean salinity with smoke','Finishing, grilled foods','Smoke intensity varies by producer']
);

originVariantRows.push(
['Wheat, Senatore Cappelli durum','Sardinia; Southern Italy','Nutty, grassy, robust wheat aroma','Pasta, bread, couscous','Heritage durum cultivar; gluten-containing'],
['Wheat, Khorasan Kamut brand grain','North America; marketed from Khorasan type','Buttery, sweet, nutty','Pasta, bread, pilaf','Brand name applies to specific certified products'],
['Rye, Pumpernickel meal','Germany; Central Europe','Deep, malty, sour-sweet','Dense bread, porridge','Coarse rye meal and long baking produce darker flavor'],
['Rye, Finnish ruis type','Finland','Sour, earthy, robust','Rye bread, porridge','Sourdough culture and milling change acidity'],
['Buckwheat, Tartary type','Himalayas; Central Asia; East Asia','Bitter, earthy, strong rutin notes','Noodle, tea, porridge','Stronger bitterness than common buckwheat'],
['Rice, Koshihikari Niigata','Niigata, Japan','Sweet, glossy, soft, aromatic','Steamed rice, sushi','Regional growing conditions influence taste'],
['Rice, Hitomebore Miyagi','Miyagi, Japan','Balanced sweetness, soft but distinct grains','Steamed rice, sushi','Cultivar and polishing affect texture'],
['Rice, Calrose medium-grain','California, United States','Mild, soft, slightly sticky','Rice bowls, sushi-style dishes','Market cultivar; cook-water ratio varies by brand'],
['Rice, Wehani red aromatic','California, United States','Nutty, earthy, popcorn-like aroma','Pilaf, salad, side dish','Aromatic red rice developed from Indian basmati lineage'],
['Rice, Carolina Gold','South Carolina, United States','Sweet, buttery, delicate','Pilaf, rice dishes, grits','Heritage variety; milling style changes cook time'],
['Rice, Carolina Red','South Carolina, United States','Earthy, nutty, bran-rich','Rice bowls, pilaf, porridge','Regional red rice variety; wholegrain texture'],
['Maize, Cuzco giant white','Cusco, Peru','Sweet, mild, nutty','Choclo, soup, toasted kernels','Large kernels; fresh and dried forms differ'],
['Maize, Morado purple','Peru','Earthy, mildly sweet, anthocyanin-rich','Chicha morada, dessert, flour','Varieties vary in color intensity and acidity'],
['Maize, Criollo de Oaxaca yellow','Oaxaca, Mexico','Sweet corn, rich masa aroma','Tortilla, tamale, atole','Local landrace; grain and nixtamal process affect flavor'],
['Bean, Mayflower shelling type','North America; Caribbean heritage','Nutty, mild, creamy','Baked beans, soup, stew','Heirloom type; seed color and regional seed lines vary'],
['Bean, Vermont cranberry heirloom','New England, United States','Sweet, earthy, creamy','Soup, baked beans, salad','Heritage regional seed type; cranberry pattern varies'],
['Bean, flageolet verte type','France','Delicate, buttery, mild green-bean note','Lamb side, braise, salad','Harvest stage determines green color and texture'],
['Pea, marrowfat green type','United Kingdom; Northern Europe','Sweet, starchy, creamy','Mushy peas, soup, snack','Maturity and drying create floury texture'],
['Mango, Himsagar','West Bengal, India','Fragrant, very sweet, low fiber','Fresh eating, dessert, puree','Seasonal cultivar with regional naming'],
['Mango, Dasheri','North India','Floral, sweet, aromatic','Fresh eating, drink, dessert','Cultivar season and ripeness change aroma'],
['Mango, Banganapalli / Safeda','Andhra Pradesh, India','Sweet, mild, smooth flesh','Fresh eating, slice, juice','Regional cultivar; market names can vary'],
['Mango, Neelam','South India','Sweet, aromatic, tangy finish','Fresh eating, dessert','Late-season cultivar; flesh fiber varies'],
['Guava, Lucknow-49','North India','Sweet, aromatic, creamy white flesh','Fresh eating, juice, preserve','Cultivar and seed development depend on maturity'],
['Guava, Thai Maroon','Thailand','Sweet, crisp, pink flesh','Fresh snack, salad, juice','Named market type; color intensity varies'],
['Papaya, Kapoho Solo','Hawaii, United States','Sweet, mild, orange-yellow flesh','Fresh eating, salad, dessert','Solo-type cultivar; growing conditions affect flavor'],
['Papaya, Formosa type','Brazil; Latin America','Sweet, mild, large fruit','Fresh eating, drink, puree','Market type includes multiple cultivars'],
['Pineapple, Sugarloaf','Benin; West Africa; Caribbean cultivation','Very sweet, low acid, white flesh','Fresh eating, juice','Regional cultivars and trade names vary'],
['Passion fruit, Purple Panama type','Australia; Pacific cultivation','Intensely aromatic, tart-sweet','Drink, dessert, sauce','Purple skin cultivar type; acidity varies'],
['Cacao, Nacional Arriba, Ecuador','Ecuador','Floral, jasmine, fruit, delicate cocoa','Chocolate, dessert','Genetic identity and origin require traceability'],
['Cacao, Trinitario, Trinidad heritage','Trinidad and Tobago; global cultivation','Balanced cocoa, fruit, spice','Chocolate, couverture','Hybrid group with broad flavor variation'],
['Coffee, Bourbon Sidra','Ecuador; Colombia cultivation','Floral, sweet, tropical fruit','Filter brew, espresso','Cultivar identity and processing should be verified by lot'],
['Coffee, SL28 variety','Kenya; East Africa','Blackcurrant, citrus, bright acidity','Filter brew, espresso','Variety is one factor; elevation and process also matter'],
['Tea, Assamica Yabukita Japanese black tea','Japan','Sweet, vegetal, light malt','Black tea, beverage pairing','Cultivar typically used for green tea; oxidized styles vary'],
['Tea, Taiwan Ruby 18 / Hong Yu','Taiwan','Mint, cinnamon, ripe fruit','Black tea, dessert pairing','Cultivar and oxidation create distinctive aroma'],
['Olive, Picual Jaén early harvest','Jaén, Spain','Green tomato, bitter, peppery','Finishing, dressing, cooking','Early harvest generally increases green pungency'],
['Olive, Koroneiki early harvest','Crete, Greece','Green herb, peppery, artichoke','Dressing, finishing','Harvest maturity changes bitterness and fruitiness'],
['Cheese, Queso de Flor de Guía','Gran Canaria, Spain','Tangy, herbal, sheep-milk, complex','Cheese board, sauce','Protected regional cheese; plant rennet tradition varies'],
['Cheese, São Jorge','Azores, Portugal','Sharp, salty, spicy, firm','Cheese board, grating','Protected origin and aging class affect intensity'],
['Cheese, Tulum','Turkey','Salty, tangy, sheep/goat milk','Breakfast, pastry, meze','Curing vessel and milk blend vary regionally'],
['Honey, leatherwood floral source','Tasmania, Australia','Floral, spicy, woody, intense','Cheese, glaze, baking','Floral source and season affect aroma'],
['Honey, acacia Robinia floral source, Hungary','Hungary','Very light, delicate, floral','Tea, desserts, cheese','Botanical source and labeling standards vary'],
['Chili, Guntur Sannam','Andhra Pradesh, India','Bright red, medium-hot, pungent','Chili powder, curry, oil','Cultivar and drying affect pungency and color'],
['Chili, Ramnad Mundu','Tamil Nadu, India','Round-fruited, fruity, medium-hot','Chutney, curry, chili oil','Regional chili with distinctive round fruit shape'],
['Chili, Bhoot Jolokia red ripe','Assam; Northeast India','Fruity, smoky, extremely hot','Pickle, sauce, seasoning','Heat can be very high and varies by pepper; handle with care']
);

originVariantRows.push(
['Rye, Danko type','Central Europe','Robust, earthy, mildly sweet','Bread, porridge, flour','Cultivar and milling determine flavor'],
['Rye, Dankowskie Złote','Poland','Sweet, mild, nutty rye aroma','Bread, flour, grain dishes','Regional grain variety; grain storage affects freshness'],
['Oat, Avena nuda naked type','China; East Asia','Mild, nutty, creamy','Porridge, flour, oat rice','Hull-less type; cooking time varies'],
['Oat, Scottish pinhead','Scotland','Nutty, hearty, slightly sweet','Porridge, broth, baking','Steel-cut grain retains texture'],
['Barley, Bere type','Scotland; Orkney','Earthy, malty, chewy','Brose, broth, bread','Ancient landrace with seasonal variation'],
['Barley, Black barley Qingke','Tibet; Himalayan regions','Nutty, toasted, earthy','Tsampa, noodles, porridge','Color and processing vary; barley contains gluten'],
['Millet, kodo, varagu landrace','South India','Earthy, nutty, firm','Porridge, pilaf, fermented batter','Local names and varieties differ'],
['Millet, teff Magna variety','Ethiopia','Mild, nutty, smooth','Injera, porridge','Seed color and milling change aroma'],
['Lentil, Eston green cultivar','Canada','Mild, earthy, peppery','Soup, salad, stew','Small seed holds shape when cooked'],
['Lentil, Beluga black type','Canada; North America','Earthy, rich, slightly mineral','Salad, pilaf, caviar-style garnish','Small black lentil; cooking time varies by age'],
['Lentil, Castelluccio di Norcia','Umbria, Italy','Delicate, earthy, peppery','Soup, stew, side dish','Protected geographical indication; harvest varies'],
['Bean, Tongues of Fire borlotti type','Italy; Americas','Nutty, creamy, sweet','Soup, braise, salad','Speckled pods and seeds vary by strain'],
['Bean, Tarbais','Southwest France','Mild, buttery, creamy','Cassoulet, stew','Protected regional bean; harvest and drying affect texture'],
['Bean, Coco de Paimpol','Brittany, France','Mild, delicate, creamy','Stew, salad, braise','Protected regional product; fresh and dried forms differ'],
['Bean, Scarlet Emperor runner','United Kingdom; Central America heritage','Earthy, rich, slightly sweet','Stew, soup, fresh pods','Cook mature seeds thoroughly'],
['Pea, Marrowfat processed type','United Kingdom','Sweet, floury, creamy','Mushy peas, snack','Processing and drying affect texture'],
['Peanut, Kintoki type','Japan; South America ancestry','Sweet, nutty, high oil','Roasted snack, confectionery','Cultivar and roasting strongly change flavor'],
['Pine nut, Korean pine seed','Korea; Northeast Asia','Sweet, buttery, resinous','Porridge, garnish, sauce','Species differs from Mediterranean pine nuts'],
['Pumpkin seed, Styrian oilseed type','Styria, Austria','Nutty, green, rich','Roast, oil, garnish','Hull-less seed type; roasting level changes flavor'],
['Sesame, Benne seed','West Africa; Southern United States','Nutty, sweet, aromatic','Baking, sauce, oil','Benne and sesame names reflect diverse seed types'],
['Sesame, black Japanese type','Japan; East Asia','Toasty, slightly bitter, intense nut aroma','Dessert, rice, sauce','Roasting deepens aroma; seed color alone does not define cultivar'],
['Chili, Peruvian Ají Mochero','Northern Peru','Citrus, tropical fruit, medium heat','Salsa, seafood, stew','Regional landrace; fresh and dried forms differ'],
['Chili, Aji Limo amarillo','Peru','Floral, citrus, hot','Ceviche, sauce, marinade','Color and heat vary among local types'],
['Chili, Siling labuyo','Philippines','Sharp, clean, high heat','Vinegar condiment, stew, sauce','Many local small-fruited cultivars use this name'],
['Chili, Wiri Wiri','Guyana; Caribbean','Fruity, bright, very hot','Pepper sauce, stew','Small round fruit; heat varies widely'],
['Chili, Datil','St. Augustine, Florida, United States','Sweet-fruity, intense heat','Hot sauce, relish, marinade','Regional cultivar; fruitiness distinguishes it from generic habanero'],
['Chili, Fatalii','Central Africa; global cultivation','Citrus, fruity, very hot','Sauce, seasoning','Yellow and red forms differ slightly; use sparingly'],
['Chili, Aji Pineapple','South America; global gardens','Pineapple-like fruit, medium heat','Salsa, sauce, fresh condiment','Cultivar identity and heat vary by seed source'],
['Cucumber, Armenian yard-long type','West Asia; Mediterranean','Mild, sweet, crisp','Salad, pickle, fresh eating','Botanically a muskmelon relative; harvest young for tenderness'],
['Cucumber, Lemon type','United States heritage cultivar','Mild, sweet, low bitterness','Fresh eating, pickle','Round yellow fruit; maturity can increase seediness'],
['Eggplant, Listada de Gandia','Spain','Mild, creamy, lightly sweet','Roast, grill, braise','Striped heritage cultivar; bitterness varies by maturity'],
['Eggplant, Thai pea type','Thailand; Southeast Asia','Bitter, green, vegetal','Curry, pickle, relish','Small fruits are often cooked whole'],
['Okra, Clemson Spineless','United States; global cultivation','Green, mild, lightly sweet','Stew, fry, grill','Pod tenderness depends on harvest size'],
['Okra, Burgundy red','United States; West Africa market types','Mild, grassy, subtle sweetness','Roast, saute, stew','Red pigment may fade during cooking'],
['Tomato, San Marzano dell’Agro Sarnese-Nocerino','Campania, Italy','Sweet, low acidity, dense flesh','Sauce, canning','Protected designation; similar market types may not share the designation'],
['Tomato, Datterino','Italy','Sweet, concentrated, low acid','Sauce, roast, salad','Small plum type; sweetness varies by growing conditions'],
['Tomato, Moneymaker','United Kingdom heritage type','Balanced sweet-tart, juicy','Slicing, cooking','Cultivar and harvest conditions affect flavor'],
['Tomato, Black Krim','Crimea; cultivated worldwide','Smoky, rich, sweet-tart','Slicing, salad, sauce','Heirloom type; color and saltiness vary by growing conditions'],
['Apple, Bramley Seedling','England','Very tart, aromatic, breaks down when cooked','Pie, sauce, baking','Cooking apple; acidity softens with sugar and heat'],
['Apple, Pink Lady / Cripps Pink','Australia; global cultivation','Sweet-tart, crisp, aromatic','Fresh eating, salad, baking','Trade names and licensed cultivars differ by market'],
['Apple, Ashmead’s Kernel','England','Sharp, nutty, aromatic','Fresh eating, cider, baking','Heritage cultivar; storage changes sweetness'],
['Pear, Conference','United Kingdom; Europe','Sweet, mild, juicy','Fresh eating, poach, bake','Ripeness strongly affects texture'],
['Pear, Comice','France; global cultivation','Very sweet, floral, buttery','Fresh eating, dessert, cheese pairing','Delicate flesh bruises easily'],
['Pear, Forelle','Germany; South Africa cultivation','Sweet, crisp, lightly spicy','Fresh eating, salad','Small speckled cultivar; season varies by hemisphere'],
['Plum, Damson','United Kingdom; Europe','Sharp, tannic, aromatic','Jam, gin, sauce','Usually cooked; ripe fruit varies in acidity'],
['Plum, Greengage','France; Europe','Very sweet, honeyed, floral','Fresh eating, preserve, dessert','Green skin does not indicate unripe fruit'],
['Apricot, Moorpark','England; Europe','Rich, aromatic, sweet-tart','Fresh eating, preserve, baking','Ripeness and drying change flavor concentration'],
['Peach, Saturn / donut type','China; global cultivation','Sweet, floral, low acid','Fresh eating, dessert','Flat fruit type; skin and flesh color vary'],
['Grape, Muscat of Alexandria','Mediterranean; West Asia','Floral, muscat, sweet','Table grape, raisin, wine','Dual-purpose cultivar; flavor varies by ripeness'],
['Grape, Kyoho','Japan; East Asia','Sweet, aromatic, large berry','Fresh eating, juice, dessert','Skin often separated; seeded and seedless forms exist'],
['Grape, Shine Muscat','Japan; East Asia cultivation','Sweet, floral, crisp, muscat aroma','Fresh eating, dessert','Cultivar and production standards vary by country'],
['Citrus, Dekopon / Shiranui','Japan','Very sweet, low acidity, fragrant','Fresh eating, dessert','Name and trademark usage vary by region'],
['Citrus, Kishu mandarin','Japan; East Asia','Sweet, delicate, seedless, small','Fresh snack, lunchbox fruit','Small fruit; peel and sweetness vary by harvest'],
['Citrus, Bergamot Femminello','Calabria, Italy','Floral, bitter, aromatic peel','Tea, zest, marmalade','Oil and fruit aroma vary with harvest and processing'],
['Citrus, Finger lime Alstonville','Australia','Citrus pearls, lime, tart','Seafood, dessert, drink','Pearl color and juice acidity vary by clone'],
['Cacao, Chuncho','Cusco, Peru','Floral, nutty, mild fruit, cocoa','Chocolate, beverage','Native cacao group; fermentation and traceability matter'],
['Cacao, Nacional Arriba, Los Ríos','Ecuador','Floral, fruit, delicate cocoa','Chocolate, couverture','Origin and genetic labels should be traceable'],
['Coffee, Geisha / Gesha, Panama type','Panama; Ethiopia ancestry','Jasmine, bergamot, tropical fruit','Filter brew, specialty espresso','Cultivar, altitude, processing and roast drive large variation'],
['Coffee, Pacamara','El Salvador; Central America','Large bean, citrus, chocolate, floral','Filter brew, espresso','Hybrid cultivar; processing alters flavor'],
['Tea, Wuyi Rou Gui oolong','Fujian, China','Cinnamon, mineral, roasted fruit','Brew, tea pairing','Roast, oxidation and origin determine profile'],
['Tea, Phoenix Dancong Mi Lan Xiang','Guangdong, China','Honey orchid, fruit, floral','Brew, dessert pairing','Cultivar and processing vary by producer'],
['Tea, Japanese Kabusecha','Japan','Sweet, grassy, moderate umami','Brew, rice pairing','Partial shading distinguishes it from sencha and gyokuro'],
['Cheese, Asiago Pressato','Veneto, Italy','Mild, milky, sweet','Table cheese, sandwich','Younger style; aging is shorter than d’Allevo'],
['Cheese, Asiago d’Allevo','Veneto, Italy','Nutty, savory, sharper with age','Grating, cheese board','Aged style; texture and flavor change over time'],
['Cheese, Queso de Valdeón','Castile and León, Spain','Blue, creamy, sharp, salty','Sauce, cheese board','Protected regional cheese; aging affects pungency'],
['Cheese, Torta del Casar','Extremadura, Spain','Sheep-milk, buttery, tangy','Spoonable cheese, bread','Protected regional cheese; seasonal texture varies'],
['Olive, Leccino','Tuscany, Italy','Mild, fruity, almond-like','Oil, table olive','Harvest maturity and curing affect profile'],
['Olive, Frantoio','Tuscany, Italy','Herbal, peppery, artichoke','Extra virgin olive oil','Cultivar and early harvest intensify pungency'],
['Olive, Hojiblanca','Andalusia, Spain','Herbal, almond, gentle bitterness','Oil, table olive','Dual-purpose cultivar; curing changes taste'],
['Honey, leatherwood, Tasmania floral source','Tasmania, Australia','Floral, spicy, woody, intense','Tea, glaze, cheese','Season and region affect aroma'],
['Honey, manuka, Leptospermum scoparium floral source','New Zealand; Australia','Earthy, herbal, resinous','Glaze, tea, cheese','Floral source and grading claims require verified labeling'],
['Honey, sourwood floral source','Appalachian region, United States','Light, floral, buttery','Tea, dessert, cheese','Harvest location and bloom conditions alter flavor'],
['Honey, acacia Robinia floral source, Romania','Romania','Light, floral, mild','Tea, baking, cheese','Botanical source and geographical claims vary by producer']
);

const nutritionRows = [
['Food group / examples','Typical nutrient pattern','Menu-planning use','Interpretation note'],
['Non-starchy vegetables: leafy greens, brassicas, peppers','Often fiber, folate, vitamin C, vitamin K and carotenoids; energy density usually low','Add color, crunch, freshness and volume across meals','Nutrients vary by variety, edible portion, storage and cooking'],
['Starchy vegetables: potato, sweet potato, corn, squash','More carbohydrate; potassium and fiber; orange flesh often has beta-carotene','Use as a substantial side, base or thickener','Preparation and added fats change total energy substantially'],
['Fruit and citrus','Carbohydrate plus fiber; vitamin C in citrus and many berries; carotenoids in some fruits','Fresh, acidic, sweet or aromatic contrast; dessert and sauce roles','Juice has less intact fiber than whole fruit; dried fruit is concentrated'],
['Whole grains: oats, barley, brown rice, millet','Carbohydrate, fiber, B vitamins and minerals; oats/barley contain beta-glucan','Build meal bowls, sides, porridge and texture','Refining removes bran/germ; fortification differs by market'],
['Pulses: lentils, chickpeas, beans, peas','Plant protein, fiber, folate, iron, potassium and magnesium','Affordable protein, hearty texture and menu variety','Canned forms can be high in sodium; rinse if desired'],
['Nuts and seeds','Unsaturated fats, protein, vitamin E or minerals depending on type','Crunch, richness, sauces and garnish','Energy dense; allergen status matters'],
['Oils and fats','Mostly fat; olive and some seed oils are rich in unsaturated fats','Carry aroma, cook, dress and finish dishes','A small volume contributes substantial energy; oils differ in flavor and heat tolerance'],
['Fish, especially oily fish','Protein; oily fish can provide omega-3 fats and vitamin D','Main protein, umami and coastal menu options','Species, feed, cut and preparation alter values'],
['Poultry and lean meat','Protein; B vitamins; iron and zinc vary by meat type','Main dish protein and savory depth','Cut, skin and cooking method change fat and energy'],
['Eggs and dairy','Protein; eggs provide choline; dairy commonly provides calcium and B12','Binding, emulsifying, enriching, baking and main proteins','Fat, sodium and added sugar vary strongly by product'],
['Fermented foods: yogurt, miso, tempeh, kimchi','Patterns vary; can contribute protein, cultures and sodium','Tang, umami, preservation and contrast','Live cultures are product- and process-specific; heating may reduce them'],
['Herbs and spices','Used in small quantities; contribute aroma and some micronutrients/plant compounds','High flavor impact with little added energy','Do not treat culinary amounts as a therapeutic dose'],
['Condiments and sauces','Often contribute sodium, sugar, acid or umami','Fast flavor building and menu identity','Check labels; serving sizes can be small but sodium/sugar concentrated'],
['Nutrition data sources','USDA FoodData Central and FAO/INFOODS composition tables','Use for quantitative nutrient calculations or ingredient comparisons','Select a matching raw/cooked/processed record and edible portion; record database entry and date']
];

for (const [sheetName, rows] of Object.entries(data)) {
  const malformed = rows.findIndex(row => row.length !== headers.length);
  if (malformed >= 0) throw new Error(`${sheetName} row ${malformed + 1} has ${rows[malformed].length} fields; expected ${headers.length}`);
}
const badOriginVariant = originVariantRows.findIndex(row => row.length !== 5);
if (badOriginVariant >= 0) throw new Error(`Origin Variants row ${badOriginVariant + 1} has ${originVariantRows[badOriginVariant].length} fields; expected 5`);
const badTaste = tasteRows.findIndex(row => row.length !== 4);
if (badTaste >= 0) throw new Error(`Taste & Pairing row ${badTaste + 1} has ${tasteRows[badTaste].length} fields; expected 4`);
const badSpecialty = specialtyRows.findIndex(row => row.length !== headers.length);
if (badSpecialty >= 0) throw new Error(`Specialty Foods row ${badSpecialty + 1} has ${specialtyRows[badSpecialty].length} fields; expected ${headers.length}`);
const ingredientKey = name => name.normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const seenIngredientKeys = new Set();
const removedIngredientDuplicates = [];
for (const [sheetName, rows] of [...Object.entries(data), ['Specialty Foods', specialtyRows]]) {
  const retained = rows.filter(row => {
    const key = ingredientKey(row[0]);
    if (!key || !seenIngredientKeys.has(key)) { seenIngredientKeys.add(key); return true; }
    removedIngredientDuplicates.push(`${row[0]} (${sheetName})`);
    return false;
  });
  if (sheetName === 'Specialty Foods') specialtyRows.splice(0, specialtyRows.length, ...retained);
  else data[sheetName] = retained;
}
console.log(`Removed ${removedIngredientDuplicates.length} duplicate ingredient names${removedIngredientDuplicates.length ? `: ${removedIngredientDuplicates.join('; ')}` : ''}`);
const ingredientNames = new Map();
for (const [sheetName, rows] of [...Object.entries(data), ['Specialty Foods', specialtyRows]]) {
  for (const row of rows) {
    const key = ingredientKey(row[0]);
    if (ingredientNames.has(key)) throw new Error(`Duplicate ingredient name: ${row[0]} (${ingredientNames.get(key)} and ${sheetName})`);
    ingredientNames.set(key, sheetName);
  }
}

const wb = Workbook.create();
const overview = wb.worksheets.add('Start Here');
const categories = Object.entries(data);
for (const [name] of categories) wb.worksheets.add(name);
wb.worksheets.add('Taste & Pairing');
wb.worksheets.add('Cuisine Origins');
wb.worksheets.add('Origin Variants');
wb.worksheets.add('Specialty Foods');
wb.worksheets.add('Nutrition Guide');
wb.worksheets.add('Sources & Notes');

overview.getRange('A1:F1').merge();
overview.getRange('A1').values = [['INGREDIENT LIBRARY | DISH & MENU PLANNING']];
overview.getRange('A2:F2').merge();
overview.getRange('A2').values = [['A broad, filterable pantry reference spanning produce, grains, pulses, seasonings, proteins and cooking essentials.']];
overview.getRange('A4:C4').values = [['Browse by sheet','Approx. entries','Use it for']];
const summary = categories.map(([name, rows]) => [name, rows.length, `Find ${name.toLowerCase()} by flavor, cuisine association, role or nutrition cue`]);
summary.push(['Taste & Pairing', tasteRows.length - 1, 'Build flavor balance and ingredient combinations']);
summary.push(['Cuisine Origins', originsRows.length - 1, 'Explore pantry patterns with regional diversity notes']);
summary.push(['Origin Variants', originVariantRows.length - 1, 'Compare named regional ingredients, cultivars, and taste differences']);
summary.push(['Specialty Foods', specialtyRows.length, 'Browse fungi, sea vegetables, edible flowers, insects, and other specialties']);
summary.push(['Nutrition Guide', nutritionRows.length - 1, 'Plan varied plates and find accurate nutrition sources']);
overview.getRange(`A5:C${4 + summary.length}`).values = summary;
overview.getRange('A20:F20').merge();
overview.getRange('A20').values = [['How to use this library']];
overview.getRange('A21:F24').merge(true);
overview.getRange('A21:A24').values = [
  ['1. Filter category sheets by type, cuisine association, taste, use or dietary note.'],
  ['2. Use Taste & Pairing to balance salty, sour, sweet, bitter, umami, heat, fat, aroma and texture.'],
  ['3. Cuisine labels are illustrative associations, not exclusive origins or rules; many ingredients travel across regions.'],
  ['4. Nutrition highlights are qualitative cues. For exact nutrient values, match the food form and edible portion in USDA FoodData Central or FAO/INFOODS.']
];
overview.getRange('A26:F26').merge();
overview.getRange('A26').values = [['* Cuisine associations indicate common culinary use or historical association; they are not claims of exclusive origin.']];
overview.getRange('A27:F27').merge();
overview.getRange('A27').values = [['† Nutrition highlights are qualitative and not medical advice. Values change with cultivar, brand, edible portion, storage, processing and cooking.']];

for (const [name, rows] of categories) {
  const sh = wb.worksheets.getItem(name);
  sh.getRange('A1:I1').values = [headers];
  sh.getRange(`A2:I${rows.length + 1}`).values = rows;
  sh.tables.add(`A1:I${rows.length + 1}`, true, `${name.replace(/[^A-Za-z]/g,'')}Table`);
  sh.freezePanes.freezeRows(1);
  sh.getRange(`A1:I${rows.length + 1}`).format.wrapText = true;
  sh.getRange(`A2:I${rows.length + 1}`).format.verticalAlignment = 'top';
  sh.getRange('A:A').format.columnWidth = 24;
  sh.getRange('B:B').format.columnWidth = 20;
  sh.getRange('C:C').format.columnWidth = 33;
  sh.getRange('D:D').format.columnWidth = 27;
  sh.getRange('E:E').format.columnWidth = 25;
  sh.getRange('F:F').format.columnWidth = 34;
  sh.getRange('G:G').format.columnWidth = 34;
  sh.getRange('H:H').format.columnWidth = 34;
  sh.getRange('I:I').format.columnWidth = 27;
  sh.getRange(`A2:I${rows.length + 1}`).format.rowHeight = 56;
}

const specialty = wb.worksheets.getItem('Specialty Foods');
specialty.getRange('A1:I1').values = [headers];
specialty.getRange(`A2:I${specialtyRows.length + 1}`).values = specialtyRows;
specialty.tables.add(`A1:I${specialtyRows.length + 1}`, true, 'SpecialtyFoodsTable');
specialty.freezePanes.freezeRows(1);
specialty.getRange(`A1:I${specialtyRows.length + 1}`).format.wrapText = true;
specialty.getRange(`A2:I${specialtyRows.length + 1}`).format.verticalAlignment = 'top';
specialty.getRange('A:A').format.columnWidth = 24;
specialty.getRange('B:B').format.columnWidth = 20;
specialty.getRange('C:C').format.columnWidth = 33;
specialty.getRange('D:D').format.columnWidth = 27;
specialty.getRange('E:E').format.columnWidth = 25;
specialty.getRange('F:F').format.columnWidth = 34;
specialty.getRange('G:G').format.columnWidth = 34;
specialty.getRange('H:H').format.columnWidth = 34;
specialty.getRange('I:I').format.columnWidth = 27;
specialty.getRange(`A2:I${specialtyRows.length + 1}`).format.rowHeight = 56;

const taste = wb.worksheets.getItem('Taste & Pairing');
taste.getRange(`A1:D${tasteRows.length}`).values = tasteRows;
taste.tables.add(`A1:D${tasteRows.length}`, true, 'TastePairingTable');
taste.freezePanes.freezeRows(1);
taste.getRange(`A1:D${tasteRows.length}`).format.wrapText = true;
taste.getRange(`A:A`).format.columnWidth = 24;
taste.getRange(`B:B`).format.columnWidth = 42;
taste.getRange(`C:C`).format.columnWidth = 46;
taste.getRange(`D:D`).format.columnWidth = 48;
taste.getRange(`A2:D${tasteRows.length}`).format.rowHeight = 50;

const origins = wb.worksheets.getItem('Cuisine Origins');
origins.getRange(`A1:D${originsRows.length}`).values = originsRows;
origins.tables.add(`A1:D${originsRows.length}`, true, 'CuisineOriginsTable');
origins.freezePanes.freezeRows(1);
origins.getRange(`A1:D${originsRows.length}`).format.wrapText = true;
origins.getRange('A:A').format.columnWidth = 30;
origins.getRange('B:B').format.columnWidth = 48;
origins.getRange('C:C').format.columnWidth = 46;
origins.getRange('D:D').format.columnWidth = 54;
origins.getRange(`A2:D${originsRows.length}`).format.rowHeight = 56;

const originVariants = wb.worksheets.getItem('Origin Variants');
originVariants.getRange(`A1:E${originVariantRows.length}`).values = originVariantRows;
originVariants.tables.add(`A1:E${originVariantRows.length}`, true, 'OriginVariantsTable');
originVariants.freezePanes.freezeRows(1);
originVariants.getRange(`A1:E${originVariantRows.length}`).format.wrapText = true;
originVariants.getRange('A:A').format.columnWidth = 34;
originVariants.getRange('B:B').format.columnWidth = 34;
originVariants.getRange('C:C').format.columnWidth = 42;
originVariants.getRange('D:D').format.columnWidth = 38;
originVariants.getRange('E:E').format.columnWidth = 48;
originVariants.getRange(`A2:E${originVariantRows.length}`).format.rowHeight = 52;

const nutrition = wb.worksheets.getItem('Nutrition Guide');
nutrition.getRange(`A1:D${nutritionRows.length}`).values = nutritionRows;
nutrition.tables.add(`A1:D${nutritionRows.length}`, true, 'NutritionGuideTable');
nutrition.freezePanes.freezeRows(1);
nutrition.getRange(`A1:D${nutritionRows.length}`).format.wrapText = true;
nutrition.getRange('A:A').format.columnWidth = 33;
nutrition.getRange('B:B').format.columnWidth = 56;
nutrition.getRange('C:C').format.columnWidth = 44;
nutrition.getRange('D:D').format.columnWidth = 58;
nutrition.getRange(`A2:D${nutritionRows.length}`).format.rowHeight = 56;

const sources = wb.worksheets.getItem('Sources & Notes');
sources.getRange('A1:C9').values = [
  ['Topic','Reference','Use / limitation'],
  ['Food nutrient composition','USDA FoodData Central — https://fdc.nal.usda.gov/','Search for matching food form and preparation; database includes multiple data types and records.'],
  ['Food composition systems','FAO/INFOODS — https://www.fao.org/food-composition','International network and food composition resources; values depend on country tables and methods.'],
  ['Per-100 g values / portions','USDA FoodData Central Foundation Foods documentation — https://fdc.nal.usda.gov/Foundation_Foods_Documentation/','Foundation documentation explains data and portion conversion; match edible portion and food state.'],
  ['Global food vocabulary','FoodOn — https://foodon.org/','A global food ontology helps classify foods and processes; this workbook is a practical culinary subset, not a complete ontology export.'],
  ['Culinary origins and use','Ingredient-level associations are broad culinary context, not exclusive or definitive historical claims.','Regional names, cultivars, preparation and taste vary; verify specialist sources for historical research.'],
  ['Nutrition caution','Qualitative highlights are for menu planning only.','For labels, clinical diets or claims, calculate from verified product/recipe data and a qualified professional as appropriate.'],
  ['Coverage note','Food diversity includes species, varieties within species, and locally used foods.','No finite menu-planning workbook can list every local cultivar, wild food, edible organism or commercial product on Earth.'],
  ['Neglected and underutilized crops','FAO & FARA, Compendium of Forgotten Foods in Africa — https://openknowledge.fao.org/3/cc5044en/cc5044en.pdf','Used as a discovery reference for regional crop diversity; does not establish that every listed food is available or safe in every location.']
];
sources.tables.add('A1:C9', true, 'SourcesNotesTable');
sources.freezePanes.freezeRows(1);
sources.getRange('A1:C9').format.wrapText = true;
sources.getRange('A:A').format.columnWidth = 28;
sources.getRange('B:B').format.columnWidth = 72;
sources.getRange('C:C').format.columnWidth = 72;
sources.getRange('A2:C9').format.rowHeight = 48;

// Cohesive workbook styling.
for (let i = 0; i < 13; i++) {
  const sh = wb.worksheets.getItemAt(i);
  sh.showGridLines = false;
  sh.getUsedRange().format.font = { name: 'Aptos', size: 10, color: '#243244' };
  const hdr = sh.getRange(sh.name === 'Start Here' ? 'A4:C4' : (sh.name === 'Sources & Notes' ? 'A1:C1' : (sh.name === 'Origin Variants' ? 'A1:E1' : (sh.name === 'Taste & Pairing' || sh.name === 'Cuisine Origins' || sh.name === 'Nutrition Guide' ? 'A1:D1' : 'A1:I1'))));
  hdr.format = { fill: '#155E63', font: { name: 'Aptos Display', size: 10, bold: true, color: '#FFFFFF' }, wrapText: true, verticalAlignment: 'center' };
  hdr.format.rowHeight = 34;
}
overview.getRange('A1:F1').format = { fill: '#123B43', font: { name: 'Aptos Display', size: 18, bold: true, color: '#FFFFFF' }, verticalAlignment: 'center' };
overview.getRange('A1:F1').format.rowHeight = 42;
overview.getRange('A2:F2').format = { fill: '#E8F2F0', font: { name: 'Aptos', size: 11, color: '#35545A' }, wrapText: true, verticalAlignment: 'center' };
overview.getRange('A2:F2').format.rowHeight = 38;
overview.getRange('A:A').format.columnWidth = 30;
overview.getRange('B:B').format.columnWidth = 18;
overview.getRange('C:C').format.columnWidth = 74;
overview.getRange('A20:F20').format = { fill: '#155E63', font: { name: 'Aptos Display', size: 11, bold: true, color: '#FFFFFF' } };
overview.getRange('A21:F24').format = { wrapText: true, verticalAlignment: 'center' };
overview.getRange('A21:F24').format.rowHeight = 30;
overview.getRange('A26:F27').format = { fill: '#F2F5F4', font: { name: 'Aptos', size: 9, color: '#50636A' }, wrapText: true, verticalAlignment: 'center' };
overview.getRange('A26:F27').format.rowHeight = 30;
overview.freezePanes.freezeRows(4);

wb.recalculate();
const summaryCheck = await wb.inspect({ kind: 'sheet', include: 'id,name', maxChars: 2000 });
console.log(summaryCheck.ndjson);
const tableCheck = await wb.inspect({ kind: 'table', sheetId: 'Vegetables', range: 'A1:I6', maxChars: 1800, tableMaxRows: 6, tableMaxCols: 9 });
console.log(tableCheck.ndjson);
const preview = await wb.render({ sheetName: 'Start Here', range: 'A1:F27', scale: 1, format: 'png' });
await fs.writeFile(`${outDir}/preview.png`, new Uint8Array(await preview.arrayBuffer()));
const visualChecks = [
  ['Vegetables', 'A1:I8'], ['Fruit & Citrus', 'A1:I8'], ['Grains & Pulses', 'A1:I8'],
  ['Spices & Herbs', 'A1:I8'], ['Protein & Dairy', 'A1:I8'], ['Pantry & Condiments', 'A1:I8'],
  ['Specialty Foods', 'A1:I8'], ['Taste & Pairing', 'A1:D135'], ['Cuisine Origins', 'A1:D8'],
  ['Nutrition Guide', 'A1:D15'], ['Sources & Notes', 'A1:C9'], ['Origin Variants', 'A1:E10']
];
for (const [sheetName, range] of visualChecks) {
  const img = await wb.render({ sheetName, range, scale: 1, format: 'png' });
  await fs.writeFile(`${outDir}/preview-${sheetName.replace(/[^A-Za-z]+/g, '-')}.png`, new Uint8Array(await img.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outDir}/Ingredients_Menu_Planning.xlsx`);
console.log(`Saved ${outDir}/Ingredients_Menu_Planning.xlsx`);
