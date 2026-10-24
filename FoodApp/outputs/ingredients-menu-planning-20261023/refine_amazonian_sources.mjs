import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const path='/Users/prajitkumarlaha/Desktop/DB_owners/FoodApp/outputs/ingredients-menu-planning-20261023/Chips_Condiments_Sauces_Bottled_Products.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const sh=wb.worksheets.getItem('Amazonian Brazil Foods & Sources');
sh.getRange('A10:C10').values=[['Embrapa — Pará cupuaçu and regional cuisine','https://www.embrapa.br/en/embrapa-no-cirio/cupuacu','Pará cupuaçu aroma and regional food forms; locally used named clones.']];
sh.getRange('A13:C13').values=[['Embrapa — Tucupi safety protocol','https://www.embrapa.br/pt/agencia-de-noticias-embrapa/busca-de-noticias/-/noticia/25540013/pesquisa-estabelece-protocolo-de-seguranca-para-a-fabricacao-de-tucupi','Embrapa research on safe cassava-liquid processing and cyanogenic compound control for tucupi.']];
wb.recalculate();const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);console.log('Refined two Embrapa source references.');
