import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const input='大学俄语1_逐课生词分类与变位.xlsx';
const blob=await FileBlob.load(input);
const original=await SpreadsheetFile.importXlsx(blob);
const wb=await SpreadsheetFile.importXlsx(blob);
const src=original.worksheets.getItem('名词变格');
const dst=wb.worksheets.getItem('名词变格');
const rows=src.getUsedRange().values;
const normalize=s=>String(s??'').toLowerCase().normalize('NFD').replaceAll('\u0301','').normalize('NFC').replaceAll('ё','е').replace(/[^а-я]/g,'');
const remove=new Set(['имя','время']);
const additions=[
  [5,['о́вощи','蔬菜','教材给复数；单数原型 о́вощ（阳）；复数间接格重音后移','о́вощ','о́воща','о́вощу','о́вощ','о́вощем','о́воще','о́вощи','овоще́й','овоща́м','о́вощи','овоща́ми','овоща́х']],
  [7,['ту́фли','布鞋','教材给复数；单数原型 ту́фля（阴）；复数二格 ту́фель','ту́фля','ту́фли','ту́фле','ту́флю','ту́флей','ту́фле','ту́фли','ту́фель','ту́флям','ту́фли','ту́флями','ту́флях']],
  [7,['стихи́','诗','教材给复数；单数原型 стих（阳）；复数一格重音后移','стих','стиха́','стиху́','стих','стихо́м','стихе́','стихи́','стихо́в','стиха́м','стихи́','стиха́ми','стиха́х']],
  [10,['роди́тели','父母；双亲','教材给复数；单数原型 роди́тель（阳）','роди́тель','роди́теля','роди́телю','роди́теля','роди́телем','роди́теле','роди́тели','роди́телей','роди́телям','роди́телей','роди́телями','роди́телях']],
  [12,['фру́кты','水果','教材给复数；单数原型 фрукт（阳）','фрукт','фру́кта','фру́кту','фрукт','фру́ктом','фру́кте','фру́кты','фру́ктов','фру́ктам','фру́кты','фру́ктами','фру́ктах']]
];
const groups=new Map();
const headings=new Map();
let lesson=null,template=null,removed=0;
for(let i=3;i<rows.length;i++){
  const word=String(rows[i][0]??'');
  const heading=word.match(/Урок\s+(\d+)/i);
  if(heading){lesson=Number(heading[1]);headings.set(lesson,i);if(!groups.has(lesson))groups.set(lesson,[]);continue;}
  if(!lesson||!word)continue;
  if(template===null)template=i;
  if(remove.has(normalize(word))){removed++;continue;}
  groups.get(lesson).push({source:i,values:rows[i]});
}
if(removed!==2)throw Error(`Expected to remove 2 rows, found ${removed}`);
for(const [l,values] of additions){
  if(values.length!==15)throw Error(`Wrong form count: ${values[0]}`);
  const group=groups.get(l);
  if(!group)throw Error(`Lesson ${l} absent`);
  if(group.some(e=>normalize(e.values[0])===normalize(values[0])))throw Error(`Already present: ${values[0]}`);
  group.push({source:null,values});
}
for(let i=0;i<rows.length;i++){
  const first=String(rows[i][0]??'');
  if(i===0||i===1||/Урок\s+\d+/i.test(first))dst.getRangeByIndexes(i,0,1,15).unmerge();
}
dst.getUsedRange().clear({applyTo:'all'});
for(let i=0;i<3;i++)dst.getRangeByIndexes(i,0,1,15).copyFrom(src.getRangeByIndexes(i,0,1,15),'all');
dst.getRangeByIndexes(1,0,1,15).merge();
let out=3;
for(const [l,entries] of [...groups.entries()].sort((a,b)=>a[0]-b[0])){
  if(!entries.length)continue;
  const h=dst.getRangeByIndexes(out,0,1,15);
  h.copyFrom(src.getRangeByIndexes(headings.get(l),0,1,15),'all');
  h.merge();h.format.rowHeight=25;out++;
  for(const e of entries){
    const dest=dst.getRangeByIndexes(out,0,1,15);
    if(e.source!==null){dest.copyFrom(src.getRangeByIndexes(e.source,0,1,15),'all');dest.format.rowHeight=22;}
    else {dest.copyFrom(src.getRangeByIndexes(template,0,1,15),'all');dest.values=[e.values];dest.getCell(0,2).format.wrapText=true;dest.format.rowHeight=56;}
    out++;
  }
}
dst.getRange('A2').values=[['依据 2020 年版逐课生词栏；列出特殊变化、教材以复数列词及仅复数名词。-мя 中性名词的规则变化见《俄语1review》。格序为一至六格。']];
await wb.recalculate();
const output=await SpreadsheetFile.exportXlsx(wb);
await output.save('tmp/artifact/plural_nouns_revised.xlsx');
console.log(JSON.stringify({removed,added:additions.length,finalRows:out}));
