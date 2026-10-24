import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const path='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
wb.worksheets.getItem('Saudi Date Cultivars').getRange('A7').values=[['Sagai (Qassim-associated date cultivar)']];
wb.recalculate();const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);console.log('Corrected a repeated alias in the Sagai cultivar label.');
