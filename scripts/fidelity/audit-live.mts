import {readFileSync,writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {renderDocumentPdf} from '../../src/render.js';
import {closeBrowser} from '../../src/browser.js';
const fixtures=process.env.FIDELITY_FIXTURE_DIR,out=process.env.FIDELITY_OUTPUT_DIR;
if(!fixtures||!out)throw new Error('Set FIDELITY_FIXTURE_DIR and FIDELITY_OUTPUT_DIR');
try {for(const id of process.argv.slice(2)){
 if(!/^[a-f0-9-]{36}$/.test(id))throw new Error('Invalid fixture UUID');
 const f=JSON.parse(readFileSync(join(fixtures,id+'.json'),'utf8'));
 const result=await renderDocumentPdf({documentId:id,workspaceId:f.workspaceId});
 writeFileSync(join(out,id+'-live.pdf'),result.buffer);
 console.log(JSON.stringify({id,bytes:result.buffer.length,fonts:result.fonts}));
}}finally{await closeBrowser()}
