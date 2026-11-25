import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const sourcePath='大学俄语1_逐课生词分类与变位.xlsx';
const blob=await FileBlob.load(sourcePath);
const sourceBook=await SpreadsheetFile.importXlsx(blob);
const book=await SpreadsheetFile.importXlsx(blob);
const src=sourceBook.worksheets.getItem('动词');
const dst=book.worksheets.getItem('动词');
const rows=src.getUsedRange().values;
const norm=s=>String(s??'').toLowerCase().normalize('NFD').replaceAll('\u0301','').normalize('NFC').replaceAll('ё','е').replace(/[^а-я]/g,'');

const place='в/на+⑥（地点）';
const motion='в/на+④（去向）；из/с+②（来源）';
const patterns={
  стоять:place,
  жить:'в/на+⑥（居住地）；у+②（住在某人处）',
  работать:'кем+⑤（职业）；в/на+⑥（地点）',
  отдыхать:'无固定接格；可接 в/на+⑥（地点）',
  обедать:'无固定接格；可单独使用',
  идти:motion,
  ехать:motion,
  разговаривать:'с кем+⑤（交谈对象）；о ком/чём+⑥（话题）',
  мечтать:'о ком/чём+⑥；+不定式',
  отвечать:'кому+③（回答对象）；на что+④（问题）',
  заниматься:'кем/чем+⑤（从事、学习）',
  вставать:'无固定接格；可接 с+②（从某处起来）',
  завтракать:'无固定接格；可单独使用',
  быть:'кем/чем+⑤（身份）；в/на+⑥（地点）',
  учиться:'чему+③（学习内容）；в+⑥（学校）',
  собираться:'в/на+④（去向）；+不定式（打算）',
  гулять:'по+③（沿途）；в/на+⑥（地点）',
  фотографироваться:'с кем+⑤（合影对象）；в/на+⑥（地点）',
  лежать:place,
  родиться:'в/на+⑥（出生地、时间）',
  сидеть:place,
  думать:'о ком/чём+⑥；+不定式（打算）',
  мочь:'+不定式（能够做）',
  смочь:'+不定式（能够做）',
  стать:'кем/чем+⑤（成为）；+不定式（开始）',
  пойти:motion,
  ходить:motion,
  казаться:'кому+③（在谁看来）；каким+⑤（显得）',
  остаться:'в/на+⑥（留下的地点）；с кем+⑤',
  оставаться:'в/на+⑥（留下的地点）；с кем+⑤',
  болеть:'чем+⑤（患某病）；у кого+②（某人身体疼）',
  познакомиться:'с кем+⑤（认识谁）',
  знакомиться:'с кем+⑤（认识谁）',
  находиться:place,
  висеть:place,
  дозвониться:'до кого+②（打通电话给谁）',
  нравиться:'кому+③（谁喜欢）',
  понравиться:'кому+③（谁喜欢）',
  приехать:motion,
  приезжать:motion,
  поехать:motion,
  прийти:motion,
  приходить:motion,
  исполниться:'кому+③（年龄）；сколько лет',
  исполняться:'кому+③（年龄）；сколько лет',
  отправляться:'из/с+②（出发地）；в/на+④（去向）',
  отправиться:'из/с+②（出发地）；в/на+④（去向）',
  поступить:'в+④（入学）；с кем/чем+⑤（对待）',
  поступать:'в+④（入学）；с кем/чем+⑤（对待）',
  верить:'кому+③（相信谁）；в кого/что+④（信仰）',
  привыкнуть:'к кому/чему+③；+不定式',
  привыкать:'к кому/чему+③；+不定式',
  случиться:'с кем/чем+⑤（发生在谁身上）',
  случаться:'с кем/чем+⑤（发生在谁身上）',
  становиться:'кем/чем+⑤（成为）',
  советовать:'кому+③（给谁建议）；что+④／+不定式',
  посоветовать:'кому+③（给谁建议）；что+④／+不定式',
  ездить:motion,
  заболеть:'чем+⑤（患某病）',
  заболевать:'чем+⑤（患某病）',
  наступить:'无固定接格（时间到来）',
  наступать:'无固定接格（时间到来）',
  умываться:'无固定接格；可接 чем+⑤（用什么洗）',
  умыться:'无固定接格；可接 чем+⑤（用什么洗）',
  лететь:motion,
  опаздывать:'в/на+④（迟到的地点或活动）',
  опоздать:'в/на+④（迟到的地点或活动）',
  бегать:motion,
  бежать:motion,
  произойти:'от+②（原因）；с кем/чем+⑤（发生对象）',
  происходить:'от+②（原因）；с кем/чем+⑤（发生对象）',
  летать:motion,
  ждать:'кого/что+④；кого/чего+②（不定对象）',
  вернуться:'из/с+②（来源）；в/на+④（返回处）',
  помочь:'кому+③（帮助谁）；чем+⑤（帮助方式）',
  помогать:'кому+③（帮助谁）；чем+⑤（帮助方式）',
  интересоваться:'кем/чем+⑤（对什么感兴趣）',
  бояться:'кого/чего+②；за кого/что+④；+不定式',
  дуть:'на кого/что+④（朝……吹）；из+②（风从何处吹来）',
  мешать:'кому/чему+③（妨碍谁/什么）；+不定式',
  помешать:'кому/чему+③（妨碍谁/什么）；+不定式',
  говорить:'о ком/чём+⑥（谈论）；с кем+⑤（交谈）；кому+③ что+④（告诉）',
  рассказывать:'кому+③ что+④（讲述）；о ком/чём+⑥（话题）',
  преподавать:'кому+③（学生）；что+④（课程）'
};
const correctedTypes={
  вернуться:'不及物', помочь:'不及物', помогать:'不及物', интересоваться:'不及物', бояться:'不及物',
  дуть:'不及物（风吹义）', мешать:'不及物（妨碍义）', помешать:'不及物（妨碍义）',
  говорить:'兼及物/不及物', советовать:'兼及物/不及物', посоветовать:'兼及物/不及物',
  ждать:'及物（④/②）'
};
const targeted=new Set();
for(let i=0;i<rows.length;i++){
  const first=String(rows[i][0]??'');
  if(i<2 || /Урок\s+\d+/i.test(first)) dst.getRangeByIndexes(i,0,1,25).unmerge();
}
dst.getUsedRange().clear({applyTo:'all'});
for(let i=0;i<rows.length;i++){
  const first=String(rows[i][0]??'');
  const heading=i===1 || /Урок\s+\d+/i.test(first);
  if(heading){
    dst.getRangeByIndexes(i,0,1,25).copyFrom(src.getRangeByIndexes(i,0,1,25),'all');
    dst.getRangeByIndexes(i,0,1,26).merge();
  }else{
    dst.getRangeByIndexes(i,0,1,4).copyFrom(src.getRangeByIndexes(i,0,1,4),'all');
    dst.getRangeByIndexes(i,5,1,21).copyFrom(src.getRangeByIndexes(i,4,1,21),'all');
    dst.getRangeByIndexes(i,4,1,1).copyFrom(src.getRangeByIndexes(i,3,1,1),'all');
    if(i===2){
      dst.getRangeByIndexes(i,4,1,1).values=[['接格关系（常见搭配）']];
    }else if(first){
      const key=norm(first);
      const pattern=patterns[key];
      dst.getRangeByIndexes(i,4,1,1).values=[[null]];
      if(pattern){dst.getRangeByIndexes(i,4,1,1).values=[[pattern]];targeted.add(key);}
      const type=correctedTypes[key];
      if(type)dst.getRangeByIndexes(i,3,1,1).values=[[type]];
      if(pattern && pattern.length>38)dst.getRangeByIndexes(i,0,1,26).format.rowHeight=35;
      else dst.getRangeByIndexes(i,0,1,26).format.rowHeight=23;
    }
  }
  if(heading || i===0){
    const height=src.getRangeByIndexes(i,0,1,25).format.rowHeight;
    if(height)dst.getRangeByIndexes(i,0,1,26).format.rowHeight=height;
  }
}
dst.getRange('A2').values=[['接格关系列按 2020 年版所列词义给出常见搭配：②/③/④/⑤/⑥为格；“无固定接格”表示动词可单独使用，地点/去向等是常见可选成分。']];
dst.getRange('E3').format.rowHeight=30;
dst.getRange('E:E').format.columnWidth=47;
dst.getRangeByIndexes(2,4,rows.length-2,1).format.wrapText=true;
for(let j=4;j<25;j++){
  const width=src.getRangeByIndexes(2,j,1,1).format.columnWidth;
  if(width)dst.getRangeByIndexes(2,j+1,1,1).format.columnWidth=width;
}
const missing=Object.keys(patterns).filter(k=>!targeted.has(k));
if(missing.length)throw Error('Pattern words absent from sheet: '+missing.join(','));
const intransitiveMissing=[];
for(let i=3;i<rows.length;i++){
  if(rows[i][3]==='不及物' && rows[i][0] && !patterns[norm(rows[i][0])])intransitiveMissing.push(rows[i][0]);
}
if(intransitiveMissing.length)throw Error('Intransitive verbs without patterns: '+intransitiveMissing.join(','));
await book.recalculate();
const output=await SpreadsheetFile.exportXlsx(book);
await output.save('tmp/artifact/verb_government_revised.xlsx');
console.log(JSON.stringify({rows:rows.length,patternWords:targeted.size,intransitiveMissing:intransitiveMissing.length}));
