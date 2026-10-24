import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';import fs from 'node:fs/promises';
const p='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(p));
for(const [sheetName,range] of [['Saudi Date Cultivars','A1:G9'],['Saudi Date Cultivars & Sources','A1:C8'],['Packaged Products','A187:I190']]){
 const result=await wb.inspect({kind:'table',sheetId:sheetName,range,tableMaxRows:12,tableMaxCols:9,maxChars:12000});console.log(result.ndjson);
 const blob=await wb.render({sheetName,range,scale:1,format:'png'});await fs.writeFile(`/private/tmp/${sheetName.replaceAll(' ','_').replaceAll('&','and')}.png`,new Uint8Array(await blob.arrayBuffer()));
}
