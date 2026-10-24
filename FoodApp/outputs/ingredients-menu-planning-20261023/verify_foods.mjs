import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
import fs from 'node:fs/promises';
const p='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(p));
const c=await wb.inspect({kind:'table',sheetId:'Caucasus Foods',range:'A1:G8',tableMaxRows:8,tableMaxCols:7,maxChars:10000});
console.log(c.ndjson);
const im=await wb.render({sheetName:'Caucasus Foods',range:'A1:G8',scale:1,format:'png'}); await fs.writeFile('/private/tmp/foods_batch2.png',new Uint8Array(await im.arrayBuffer()));
const prod=await wb.render({sheetName:'Packaged Products',range:'A178:I181',scale:1,format:'png'}); await fs.writeFile('/private/tmp/products_caucasus.png',new Uint8Array(await prod.arrayBuffer()));
