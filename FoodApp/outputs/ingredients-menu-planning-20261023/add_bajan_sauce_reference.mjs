import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const path='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const sh=wb.worksheets.getItem('Caribbean Foods & Sources');
const rows=sh.getUsedRange().values;
const url='https://shopbarbados.co/product-category/lifestyle/pepper-sauces/';
if(!rows.some(r=>r[1]===url)){
 const t=sh.tables.items[0]; if(!t)throw new Error('Source table missing');
 t.rows.add(null,[['Shop Barbados — Pepper Sauces',url,'Regional Bajan pepper-sauce descriptions; examples identify vinegar, Scotch bonnet peppers and spice blends.']]);
 const n=sh.getUsedRange().values.length; sh.getRange(`A${n}:C${n}`).format.wrapText=true;sh.getRange(`A${n}:C${n}`).format.verticalAlignment='top';sh.getRange(`A${n}:C${n}`).format.rowHeight=48;
}
wb.recalculate();const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);console.log('Added Bajan pepper-sauce reference.');
