import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('tmp/artifact/plural_nouns_revised.xlsx'));
const preview=await wb.render({sheetName:'名词变格',range:'A64:O69',scale:1.2,format:'png'});
await fs.writeFile('tmp/artifact/plural_nouns_preview.png',new Uint8Array(await preview.arrayBuffer()));
