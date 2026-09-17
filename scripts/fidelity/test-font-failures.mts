import assert from 'node:assert/strict';
import puppeteer from 'puppeteer-core';
const chromium=await import('@sparticuz/chromium');
const executablePath=process.env.CHROMIUM_EXECUTABLE_PATH||(process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':await chromium.default.executablePath());
const url=process.argv[2];if(!url)throw new Error('Pass the export fixture URL');
const browser=await puppeteer.launch({executablePath,headless:true,args:['--no-sandbox']});
const resources=process.argv.slice(3).length?process.argv.slice(3):['/api/fonts','/font-metrics/manifest.json','/fonts/'];
try {for(const resource of resources) {
 const context=await browser.createBrowserContext();const page=await context.newPage();await page.setRequestInterception(true);
 page.on('request',r=>r.url().includes(resource)?r.respond({status:503,body:'test unavailable'}):r.continue());
 const start=Date.now();await page.goto(url);
 await page.waitForFunction(()=>window.__PDF_ERROR||window.__PDF_READY,{timeout:90000});
 const state=await page.evaluate(()=>({error:window.__PDF_ERROR,ready:!!window.__PDF_READY}));
 assert.ok(state.error,resource+' must report failure');
 const expected=resource==='/api/fonts'?'catálogo':resource==='/font-metrics/manifest.json'?'preparação das páginas':resource==='/fonts/'?'fonte':resource==='missing-template'?'bloco de design':'imagem';
 assert.ok(state.error.toLowerCase().includes(expected),'must fail for the injected fault, not an unrelated asset: '+state.error);assert.equal(state.ready,false,'must never authorize PDF capture with missing document resources');
 console.log(JSON.stringify({resource,state,ms:Date.now()-start}));await context.close();
}}finally{await browser.close()}
