import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const root=process.cwd();
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(`${root}/大学俄语1_逐课生词分类与变位.xlsx`));
const verb=wb.worksheets.getItem('动词');
let changed=0;
for(let r=4;r<=195;r++){
  const first=verb.getRange(`A${r}`).values[0][0];
  if(typeof first==='string'&&first.startsWith('Урок'))continue;
  const cell=verb.getRange(`F${r}`);
  const current=cell.values[0][0];
  const next=current==='第一变位法'?'一':current==='第二变位法'?'二':null;
  cell.values=[[next]];
  changed++;
}
const nouns=wb.worksheets.getItem('名词变格');
let unbold=0;
for(let r=4;r<=92;r++){
  const first=nouns.getRange(`A${r}`).values[0][0];
  if(typeof first==='string'&&first.startsWith('Урок'))continue;
  nouns.getRange(`A${r}:O${r}`).format.font={name:'Microsoft YaHei',size:10,color:'#172B4D',bold:false};
  unbold++;
}
const output=await SpreadsheetFile.exportXlsx(wb);
await output.save(`${root}/.work/base_after_cleanup.xlsx`);
console.log({changed,unbold});
