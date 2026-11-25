import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const input='大学俄语1_逐课生词分类与变位.xlsx';
const edits=JSON.parse(await fs.readFile('tmp/artifact/edits.json','utf8'));
const blob=await FileBlob.load(input);
const original=await SpreadsheetFile.importXlsx(blob);
const wb=await SpreadsheetFile.importXlsx(blob);
const normalize=s=>String(s??'').toLowerCase().normalize('NFD').replaceAll('\u0301','').normalize('NFC').replaceAll('ё','е').replace(/[^а-я]/g,'');
const remove=new Set(edits.remove.map(([l,w])=>`${l}|${normalize(w)}`));
const moves=new Map(edits.move.map(([l,w,to])=>[`${l}|${normalize(w)}`,to]));
const sheets=['名词','名词变格','动词','副词','形容词','代词','前置词','数词','连接词','语气词','感叹词'];
const verbData={
  лежать:['未','不及物','现在','二',['лежу́','лежи́шь','лежи́т','лежи́м','лежи́те','лежа́т'],['лежа́л','лежа́ла','лежа́ло','лежа́ли'],['лежи́','лежи́те']],
  сидеть:['未','不及物','现在','二',['сижу́','сиди́шь','сиди́т','сиди́м','сиди́те','сидя́т'],['сиде́л','сиде́ла','сиде́ло','сиде́ли'],['сиди́','сиди́те']],
  зарабатывать:['未','及物','现在','一',['зараба́тываю','зараба́тываешь','зараба́тывает','зараба́тываем','зараба́тываете','зараба́тывают'],['зараба́тывал','зараба́тывала','зараба́тывало','зараба́тывали'],['зараба́тывай','зараба́тывайте']],
  навещать:['未','及物','现在','一',['навеща́ю','навеща́ешь','навеща́ет','навеща́ем','навеща́ете','навеща́ют'],['навеща́л','навеща́ла','навеща́ло','навеща́ли'],['навеща́й','навеща́йте']],
  строить:['未','及物','现在','二',['стро́ю','стро́ишь','стро́ит','стро́им','стро́ите','стро́ят'],['стро́ил','стро́ила','стро́ило','стро́или'],['стро́й','стро́йте']],
  построить:['完','及物','简单将来','二',['постро́ю','постро́ишь','постро́ит','постро́им','постро́ите','постро́ят'],['постро́ил','постро́ила','постро́ило','постро́или'],['постро́й','постро́йте']],
  сдавать:['未','及物','现在','一',['сдаю́','сдаёшь','сдаёт','сдаём','сдаёте','сдаю́т'],['сдава́л','сдава́ла','сдава́ло','сдава́ли'],['сдава́й','сдава́йте']],
  сдать:['完','及物','简单将来',null,['сда́м','сда́шь','сда́ст','сдади́м','сдади́те','сдаду́т'],['сдал','сдала́','сда́ло','сда́ли'],['сдай','сда́йте']],
  принести:['完','及物','简单将来','一',['принесу́','принесёшь','принесёт','принесём','принесёте','принесу́т'],['принёс','принесла́','принесло́','принесли́'],['принеси́','принеси́те']],
  приносить:['未','及物','现在','二',['приношу́','прино́сишь','прино́сит','прино́сим','прино́сите','прино́сят'],['приноси́л','приноси́ла','приноси́ло','приноси́ли'],['приноси́','приноси́те']],
  верить:['未','不及物','现在','二',['ве́рю','ве́ришь','ве́рит','ве́рим','ве́рите','ве́рят'],['ве́рил','ве́рила','ве́рило','ве́рили'],['ве́рь','ве́рьте']],
  умываться:['未','不及物','现在','一',['умыва́юсь','умыва́ешься','умыва́ется','умыва́емся','умыва́етесь','умыва́ются'],['умыва́лся','умыва́лась','умыва́лось','умыва́лись'],['умыва́йся','умыва́йтесь']],
  умыться:['完','不及物','简单将来','一',['умо́юсь','умо́ешься','умо́ется','умо́емся','умо́етесь','умо́ются'],['умы́лся','умы́лась','умы́лось','умы́лись'],['умо́йся','умо́йтесь']],
  выбирать:['未','及物','现在','一',['выбира́ю','выбира́ешь','выбира́ет','выбира́ем','выбира́ете','выбира́ют'],['выбира́л','выбира́ла','выбира́ло','выбира́ли'],['выбира́й','выбира́йте']],
  выбрать:['完','及物','简单将来','一',['вы́беру','вы́берешь','вы́берет','вы́берем','вы́берете','вы́берут'],['вы́брал','вы́брала','вы́брало','вы́брали'],['вы́бери','вы́берите']],
  искать:['未','及物','现在','一',['ищу́','и́щешь','и́щет','и́щем','и́щете','и́щут'],['иска́л','иска́ла','иска́ло','иска́ли'],['ищи́','ищи́те']],
  стать:['完','不及物','简单将来','一',['ста́ну','ста́нешь','ста́нет','ста́нем','ста́нете','ста́нут'],['стал','ста́ла','ста́ло','ста́ли'],['стань','ста́ньте']],
  пойти:['完','不及物','简单将来','一',['пойду́','пойдёшь','пойдёт','пойдём','пойдёте','пойду́т'],['пошёл','пошла́','пошло́','пошли́'],['пойди́','пойди́те']],
  ходить:['未','不及物','现在','二',['хожу́','хо́дишь','хо́дит','хо́дим','хо́дите','хо́дят'],['ходи́л','ходи́ла','ходи́ло','ходи́ли'],['ходи́','ходи́те']],
  учить:['未','及物','现在','二',['учу́','у́чишь','у́чит','у́чим','у́чите','у́чат'],['учи́л','учи́ла','учи́ло','учи́ли'],['учи́','учи́те']],
  выучить:['完','及物','简单将来','二',['вы́учу','вы́учишь','вы́учит','вы́учим','вы́учите','вы́учат'],['вы́учил','вы́учила','вы́учило','вы́учили'],['вы́учи','вы́учите']]
};
const nounSpecial=[
  [8,['календа́рь','[阳]日历','复数一格重音移动：календа́рь → календари́','календа́рь','календаря́','календарю́','календа́рь','календарём','календаре́','календари́','календаре́й','календаря́м','календари́','календаря́ми','календаря́х']],
  [10,['дворе́ц','宫殿','单数间接格词干 е 脱落：дворе́ц → дворца́','дворе́ц','дворца́','дворцу́','дворе́ц','дворцо́м','дворце́','дворцы́','дворцо́в','дворца́м','дворцы́','дворца́ми','дворца́х']],
  [16,['ребёнок','孩子','复数异词根：ребёнок → де́ти','ребёнок','ребёнка','ребёнку','ребёнка','ребёнком','ребёнке','де́ти','дете́й','де́тям','дете́й','детьми́','де́тях']]
];
const correctedWord=new Map([
  ['полка́','по́лка'],
  ['Са́нкт-Петербу́рг','Санкт-Петербу́рг'],
  ['по́-кита́йски','по-кита́йски'],
  ['по́-ру́сски','по-ру́сски'],
  ['по́-англи́йски','по-англи́йски'],
  ['по́-моему́','по-мо́ему']
]);
const correctedNounDeclension=new Map([
  ['полка',['по́лка','书架','复数二格插入 о：по́лка → по́лок','по́лка','по́лки','по́лке','по́лку','по́лкой','по́лке','по́лки','по́лок','по́лкам','по́лки','по́лками','по́лках']]
]);
const metrics=[];
for (const name of sheets) {
  const src=original.worksheets.getItem(name);
  const dst=wb.worksheets.getItem(name);
  const rows=src.getUsedRange().values;
  const cols=rows[2].length;
  const groups=new Map();
  const headings=new Map();
  let lesson=null;
  let template=null;
  let moved=0,deleted=0,added=0;
  for(let i=3;i<rows.length;i++) {
    const word=String(rows[i][0]??'');
    const match=word.match(/Урок\s+(\d+)/i);
    if(match) { lesson=Number(match[1]); headings.set(lesson,i); if(!groups.has(lesson)) groups.set(lesson,[]); continue; }
    if(!lesson || !word) continue;
    if(template===null) template=i;
    const key=`${lesson}|${normalize(word)}`;
    if(remove.has(key)){deleted++;continue;}
    const target=moves.get(key)??lesson;
    if(target!==lesson)moved++;
    if(!groups.has(target))groups.set(target,[]);
    groups.get(target).push({source:i,values:rows[i]});
  }
  if(name!=='名词变格') for(const [l,category,word,meaning] of edits.add) {
    if(category!==name)continue;
    if(!groups.has(l))groups.set(l,[]);
    if(groups.get(l).some(e=>normalize(e.values[0])===normalize(word)))continue;
    const values=Array(cols).fill(null); values[0]=word; values[1]=meaning;
    if(name==='动词'){
      const data=verbData[normalize(word)];
      if(!data)throw Error(`No forms for ${word}`);
      values[2]=data[0];values[3]=data[1];values[4]=data[2];values[5]=data[3];
      for(let k=0;k<6;k++) values[6+k]=data[4][k];
      if(data[0]==='未'){
        const aux=['бу́ду','бу́дешь','бу́дет','бу́дем','бу́дете','бу́дут'];
        for(let k=0;k<6;k++)values[12+k]=`${aux[k]} ${word}`;
      }
      for(let k=0;k<4;k++)values[18+k]=data[5][k];
      for(let k=0;k<2;k++)values[22+k]=data[6][k];
    }
    groups.get(l).push({source:null,values}); added++;
  }
  if(name==='名词变格') for(const [l,values] of nounSpecial){
    if(!groups.has(l))groups.set(l,[]);
    if(groups.get(l).some(e=>normalize(e.values[0])===normalize(values[0])))continue;
    groups.get(l).push({source:null,values});added++;
  }
  for(let i=0;i<rows.length;i++){
    const a=String(rows[i][0]??'');
    if(i===0 || /Урок\s+\d+/i.test(a)) dst.getRangeByIndexes(i,0,1,cols).unmerge();
  }
  dst.getUsedRange().clear({applyTo:'all'});
  for(let i=0;i<3;i++) dst.getRangeByIndexes(i,0,1,cols).copyFrom(src.getRangeByIndexes(i,0,1,cols),'all');
  let out=3;
  const allLessons=[...groups.keys()].sort((a,b)=>a-b);
  for(const l of allLessons){
    const entries=groups.get(l);
    if(!entries?.length)continue;
    const headingSource=headings.get(l)??headings.values().next().value;
    const h=dst.getRangeByIndexes(out,0,1,cols);
    h.copyFrom(src.getRangeByIndexes(headingSource,0,1,cols),'all');
    h.getCell(0,0).values=[[`Урок ${l}｜第 ${l} 课`]];
    h.merge();
    h.format.rowHeight=25;
    out++;
    for(const e of entries){
      const dest=dst.getRangeByIndexes(out,0,1,cols);
      if(e.source!==null){dest.copyFrom(src.getRangeByIndexes(e.source,0,1,cols),'all');}
      else {dest.copyFrom(src.getRangeByIndexes(template,0,1,cols),'all');dest.values=[e.values];}
      const fix=correctedWord.get(String(e.values[0]??''));
      if(fix)dest.getCell(0,0).values=[[fix]];
      if(name==='名词变格'){
        const declensionFix=correctedNounDeclension.get(normalize(e.values[0]));
        if(declensionFix)dest.values=[declensionFix];
      }
      dest.format.rowHeight=22;
      out++;
    }
  }
  if(name==='名词变格')dst.getRange('A2').values=[['依据 2020 年版逐课生词栏；列出特殊变化及仅有复数的名词。格序为一至六格。']];
  else if(name==='动词')dst.getRange('A2').values=[['依据 2020 年版逐课“生词”栏；完成体人称形式是简单将来，未完成体另列复合将来。']];
  else dst.getRange('A2').values=[['来源：2020 年版《大学俄语》学生用书 1 各课“生词 / Новые слова”栏。多音节词用字母正上方的重音符标注；单音节词不标。']];
  metrics.push({sheet:name,original:rows.length,final:out,moved,deleted,added});
}
const notes=wb.worksheets.getItem('说明');
notes.getRange('B2').values=[['2020 年 2 月第 1 版《大学俄语》（东方新版）学生用书 1，各课“生词 / Новые слова”栏；未使用书末词汇表。']];
notes.getRange('A6').values=[['教材版本']];
notes.getRange('B6').values=[['2020 年 2 月第 1 版（史铁强、张金兰）']];
await wb.recalculate();
const out=await SpreadsheetFile.exportXlsx(wb);
await out.save('tmp/artifact/revised.xlsx');
console.log(JSON.stringify(metrics));
