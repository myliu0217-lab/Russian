import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const root=process.cwd();
const labels=JSON.parse(await fs.readFile(`${root}/.work/conjugation_labels.json`,'utf8'));
const input=await FileBlob.load(`${root}/大学俄语1_逐课生词分类与变位_修订版.xlsx`);
const wb=await SpreadsheetFile.importXlsx(input);
const sh=wb.worksheets.getItem('动词');
sh.getRange('Y3').values=[['变位法']];
sh.getRange('Y4:Y195').values=labels.map(x=>[x]);
sh.getRange('Y3').format={fill:'#2875A4',font:{name:'Microsoft YaHei',size:10,bold:true,color:'#FFFFFF'},rowHeight:22,verticalAlignment:'center'};
sh.getRange('Y4:Y195').format={fill:'#FFFFFF',font:{name:'Microsoft YaHei',size:10,color:'#172B4D'},rowHeight:20,verticalAlignment:'center'};
sh.getRange('Y3:Y195').format.borders={preset:'all',style:'thin',color:'#BFDCEB'};
sh.getRange('Y:Y').format.columnWidth=18;
const output=await SpreadsheetFile.exportXlsx(wb);
await output.save(`${root}/.work/conjugation_base.xlsx`);
console.log('wrote conjugation_base.xlsx');
