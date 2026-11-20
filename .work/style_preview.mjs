import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {Workbook} from '@oai/artifact-tool';
const wb=Workbook.create();
const sh=wb.worksheets.add('名词样式预览');
sh.showGridLines=false;
const rows=[
 ['名词｜按课次整理',null],
 ['仅显示前两课的部分生词，用于确认行高和分割线。',null],
 ['俄语生词','中文意思'],
 ['Урок 1｜第 1 课',null],
 ['папа','爸爸'],['мама','妈妈'],['томат','番茄'],['поэт','诗人'],['ум','智慧；智力'],['Эмма','爱玛'],
 ['Урок 2｜第 2 课',null],['кот','猫'],['окно','窗户'],['утка','鸭子'],['касса','收款台'],
];
sh.getRange('A1:B15').values=rows;
sh.getRange('A1:B15').format.font={name:'Microsoft YaHei',size:10,color:'#172B4D'};
sh.getRange('A1:B15').format.verticalAlignment='center';
sh.getRange('A1:B15').format.rowHeight=20;
sh.getRange('A1:B1').merge();
sh.getRange('A1:B1').format={fill:'#123F66',font:{name:'Microsoft YaHei',size:15,bold:true,color:'#FFFFFF'},rowHeight:30,verticalAlignment:'center'};
sh.getRange('A2:B2').merge();
sh.getRange('A2:B2').format={fill:'#F4F8FC',font:{name:'Microsoft YaHei',size:10,color:'#425466'},rowHeight:22,verticalAlignment:'center'};
sh.getRange('A3:B3').format={fill:'#2875A4',font:{name:'Microsoft YaHei',size:10,bold:true,color:'#FFFFFF'},rowHeight:22,verticalAlignment:'center'};
for(const r of [4,11]){
  sh.getRange(`A${r}:B${r}`).merge();
  sh.getRange(`A${r}:B${r}`).format={fill:'#DCEBF5',font:{name:'Microsoft YaHei',size:10,bold:true,color:'#123F66'},rowHeight:20,verticalAlignment:'center'};
}
sh.getRange('A3:B15').format.borders={preset:'all',style:'thin',color:'#BFDCEB'};
sh.getRange('A:A').format.columnWidth=34;
sh.getRange('B:B').format.columnWidth=35;
const img=await wb.render({sheetName:'名词样式预览',range:'A1:B15',scale:2,format:'png'});
await fs.writeFile(fileURLToPath(new URL('style_preview.png',import.meta.url)),new Uint8Array(await img.arrayBuffer()));
console.log('rendered');
