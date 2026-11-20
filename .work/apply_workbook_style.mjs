import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const root=process.cwd();
const source=`${root}/大学俄语1_逐课生词分类与变位.xlsx`;
const output=`${root}/.work/styled_base.xlsx`;
const data=JSON.parse(await fs.readFile(`${root}/.work/noun_rows.json`,'utf8'));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const specs=[
 ['名词',501,2],['名词变格',data.rows.length,15],['动词',195,24],
 ['副词',107,2],['形容词',106,2],['代词',58,2],['前置词',22,2],
 ['数词',43,2],['连接词',16,2],['语气词',14,2],['感叹词',18,2],
];
const col=(n)=>{let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s};
for(const [name,count,width] of specs){
 const sh=wb.worksheets.getItem(name), last=col(width);
 const note=sh.getRange('A2').values[0][0];
 if(typeof note==='string'&&note.includes('俄语词条去重音符号'))
   sh.getRange('A2').values=[[note.replace('俄语词条去重音符号','重音元音以红字标出')]];
 if(name==='名词变格'){
   for(let r=4;r<=45;r++){
     const v=sh.getRange(`A${r}`).values[0][0];
     if(typeof v==='string'&&v.startsWith('Урок'))sh.getRange(`A${r}:O${r}`).unmerge();
   }
   sh.getRange(`A4:O${count}`).values=data.rows.slice(3);
   sh.getRange(`A4:O${count}`).format.font={name:'Microsoft YaHei',size:10,color:'#172B4D'};
   sh.getRange(`A4:O${count}`).format.fill='#FFFFFF';
   for(let i=3;i<data.rows.length;i++){
     if(!String(data.rows[i][0]??'').startsWith('Урок'))continue;
     const r=i+1;
     sh.getRange(`A${r}:O${r}`).merge();
     sh.getRange(`A${r}:O${r}`).format={fill:'#DCEBF5',font:{name:'Microsoft YaHei',size:10,bold:true,color:'#123F66'},rowHeight:20,verticalAlignment:'center'};
   }
 }
 sh.showGridLines=false;
 sh.getRange(`A4:${last}${count}`).format.rowHeight=20;
 sh.getRange(`A4:${last}${count}`).format.verticalAlignment='center';
 sh.getRange(`A3:${last}${count}`).format.borders={preset:'all',style:'thin',color:'#BFDCEB'};
 sh.getRange(`A1:${last}1`).format.rowHeight=30;
 sh.getRange(`A2:${last}2`).format.rowHeight=22;
 sh.getRange(`A3:${last}3`).format.rowHeight=22;
 console.log(name,count);
}
const note=wb.worksheets.add('说明');
note.showGridLines=false;
note.getRange('A1:B1').merge();
note.getRange('A1').values=[['编制说明与重音来源']];
note.getRange('A2:B5').values=[
 ['词汇范围','仅取教材逐课“生词 / Новые слова”栏；未使用书末生词表。'],
 ['重音标记','俄语词条及表内词形的重音元音以红字和重音符标出；ё 仅以红字标出。'],
 ['重音数据','eugenysilin/russian-ukrainian-stress-data（基于 Wiktionary），CC BY-SA 4.0'],
 ['项目地址','https://github.com/eugenysilin/russian-ukrainian-stress-data'],
];
note.getRange('A1:B1').format={fill:'#123F66',font:{name:'Microsoft YaHei',size:14,bold:true,color:'#FFFFFF'},rowHeight:30};
note.getRange('A2:B5').format={font:{name:'Microsoft YaHei',size:10,color:'#172B4D'},rowHeight:20,verticalAlignment:'center'};
note.getRange('A2:B5').format.borders={preset:'all',style:'thin',color:'#BFDCEB'};
note.getRange('A:A').format.columnWidth=23;
note.getRange('B:B').format.columnWidth=92;
const x=await SpreadsheetFile.exportXlsx(wb);
await x.save(output);
console.log(output);
