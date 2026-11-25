import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('tmp/artifact/verb_government_revised.xlsx'));
for(const [range,name,scale] of [['A1:Z10','header',0.9],['A194:F208','late',1.2]]){
  const preview=await wb.render({sheetName:'动词',range,scale,format:'png'});
  await fs.writeFile(`tmp/artifact/verb_government_${name}.png`,new Uint8Array(await preview.arrayBuffer()));
}
