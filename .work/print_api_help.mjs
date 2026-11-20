import {Workbook} from '@oai/artifact-tool';
const wb=Workbook.create();
console.log(wb.help('worksheet.pageSetup',{include:'index,examples,notes'}).ndjson);
