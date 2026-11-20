import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const root=process.cwd();
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(`${root}/大学俄语1_逐课生词分类与变位.xlsx`));
const output=await SpreadsheetFile.exportXlsx(wb);
await output.save(`${root}/.work/print_base.xlsx`);
console.log('print_base.xlsx');
