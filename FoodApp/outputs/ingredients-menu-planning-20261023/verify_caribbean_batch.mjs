import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const p='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(p));
for(const [name,range] of [['Caribbean Foods','A1:G16'],['Caribbean Foods & Sources','A1:C8'],['Packaged Products','A182:I185']]){const r=await wb.inspect({kind:'table',sheetId:name,range,include:'values',tableMaxRows:20,tableMaxCols:12,maxChars:18000});await fs.writeFile(`/private/tmp/${name.replaceAll(' ','_').replaceAll('&','and')}.ndjson`,r.ndjson);}
for(const [sheetName,range,file] of [['Caribbean Foods','A1:G16','/private/tmp/caribbean_foods.png'],['Packaged Products','A181:I185','/private/tmp/caribbean_products.png']]){const b=await wb.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(file,new Uint8Array(await b.arrayBuffer()));}
