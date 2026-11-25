import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const input=await FileBlob.load('大学俄语1_逐课生词分类与变位.xlsx');
const wb=await SpreadsheetFile.importXlsx(input);
const wb2=await SpreadsheetFile.importXlsx(input);
console.log((await wb.inspect({kind:'sheet',include:'id,name',maxChars:2000})).ndjson);
console.log(wb.help('worksheet.delete', {include:'index,examples,notes',maxChars:2000}).ndjson);
try { wb.worksheets.getItem('名词').getRange('A501:B501').copyFrom(wb2.worksheets.getItem('名词').getRange('A5:B5'),'all'); console.log('COPY_OK'); } catch(e) { console.log('COPY_ERR',String(e)); }
const out=await SpreadsheetFile.exportXlsx(wb);
await out.save('tmp/artifact/probe.xlsx');
