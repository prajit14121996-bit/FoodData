import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const path='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const p=wb.worksheets.getItem('Packaged Products');
p.getRange('A174:I177').format.wrapText=true;
p.getRange('A174:I177').format.verticalAlignment='top';
p.getRange('A174:I177').format.rowHeight=82;
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);
console.log('Formatted the four appended product rows for wrapped, readable text.');
