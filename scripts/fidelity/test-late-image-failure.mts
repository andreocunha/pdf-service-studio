import puppeteer from 'puppeteer-core';
import assert from 'node:assert/strict';
const chromium=await import('@sparticuz/chromium');
const executablePath=process.env.CHROMIUM_EXECUTABLE_PATH||(process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':await chromium.default.executablePath());
const url=process.argv[2];if(!url)throw new Error('Pass an export fixture URL');
const b=await puppeteer.launch({executablePath,headless:true,args:['--no-sandbox']});
try {const p=await b.newPage();await p.goto(url);await p.waitForFunction(()=>window.__PDF_READY,{timeout:60000});
 const error=await p.evaluate(async()=>{const img=new Image();img.src='data:image/png;base64,broken';document.querySelector('.pdf-render-host').appendChild(img);try{await window.__PDF_SETTLE();return ''}catch(e){return e.message}});
 assert.match(error,/imagem/);console.log('PASS: final settle rejects a broken image introduced after READY');
}finally{await b.close()}
