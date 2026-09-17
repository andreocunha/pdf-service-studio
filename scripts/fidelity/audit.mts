import puppeteer from 'puppeteer-core';
import {preparePdfMasks} from '../../src/pdf-masks.js';
import {writeFileSync,mkdirSync} from 'node:fs';
const out=process.env.FIDELITY_OUTPUT_DIR; if(!out)throw new Error('Set FIDELITY_OUTPUT_DIR');mkdirSync(out,{recursive:true});
const baseUrl=process.env.FIDELITY_BASE_URL||'http://localhost:3002';
const id=process.argv[2]; if(!/^[a-f0-9-]{36}$/.test(id||''))throw new Error('Pass a fixture UUID');const modes=process.argv.slice(3).length?process.argv.slice(3):['editor','export'];
const chromium=await import('@sparticuz/chromium');
const executablePath=process.env.CHROMIUM_EXECUTABLE_PATH||(process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':await chromium.default.executablePath());
const browser=await puppeteer.launch({executablePath,headless:true,args:['--no-sandbox'],protocolTimeout:300000});
try {for(const mode of modes){
 const context=await browser.createBrowserContext();const page=await context.newPage();await page.setViewport({width:1200,height:1000,deviceScaleFactor:1});await page.emulateMediaType('screen'); await page.evaluateOnNewDocument('window.__name = function(fn) { return fn; }');
 const errors=[];page.on('requestfailed',r=>errors.push((r.failure()?.errorText||'request failed')+' '+r.url().split('?')[0]));page.on('pageerror',e=>errors.push(String(e)));page.on('response',r=>{if(r.status()>=400)errors.push(r.status()+' '+r.url().split('?')[0])});
 await page.evaluateOnNewDocument((delayMs)=>{const original=window.fetch;window.fetch=async(...args)=>{if(String(args[0]).includes('/api/fonts'))await new Promise(r=>setTimeout(r,delayMs));return original(...args)}},Number(process.env.FIDELITY_FONT_DELAY_MS??1500));
 console.log(JSON.stringify({id,mode,event:'start'}));const start=Date.now();
 await page.goto(`${baseUrl}/render-pdf/validation/${id}?mode=${mode}`,{waitUntil:'domcontentloaded',timeout:180000});
 await page.waitForFunction(mode==='editor'?()=>window.__FIDELITY_READY||window.__FIDELITY_ERROR:()=>window.__PDF_READY||window.__PDF_ERROR,{timeout:240000,polling:200});
 const status=await page.evaluate(()=>({error:window.__FIDELITY_ERROR||window.__PDF_ERROR,pending:!!document.querySelector('[data-font-layout-pending]'),fonts:document.fonts.status}));
 console.log(JSON.stringify({id,mode,event:'readiness',status,errors}));
 if(status.error||status.pending||errors.length)throw new Error('Incomplete fixture: '+JSON.stringify({status,errors}));
 await page.evaluate(async()=>{await document.fonts.ready;document.querySelectorAll('[data-page-index]').forEach(p=>p.style.contentVisibility='visible');await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));});
 const snapshot=await page.evaluate(()=>{
  const pages=Array.from(document.querySelectorAll('[data-page-index]'));
  return {dpr:devicePixelRatio,pages:pages.map(p=>{const pr=p.getBoundingClientRect(),scale=pr.width/parseFloat(getComputedStyle(p).width);const rect=r=>({x:(r.x-pr.x)/scale,y:(r.y-pr.y)/scale,width:r.width/scale,height:r.height/scale});
   return {width:pr.width/scale,height:pr.height/scale,images:Array.from(p.querySelectorAll('img')).filter(e=>e.getBoundingClientRect().width&&e.getBoundingClientRect().height).map(e=>({src:e.currentSrc||e.src,complete:e.complete,naturalWidth:e.naturalWidth,rect:rect(e.getBoundingClientRect())})),navigation:Array.from(p.querySelectorAll('[data-target-block-id]')).filter(e=>e.getBoundingClientRect().width&&e.getBoundingClientRect().height).map(e=>({target:e.getAttribute('data-target-block-id'),text:e.textContent,rect:rect(e.getBoundingClientRect())})),blocks:Array.from(p.querySelectorAll('[data-block-id]')).map(b=>{
    const nodes=[];const walker=document.createTreeWalker(b,NodeFilter.SHOW_TEXT);let n;
    while(n=walker.nextNode()){
     if(!n.textContent?.trim())continue;const e=n.parentElement;if(!e||e.closest('button,[data-editor-only],[aria-hidden="true"],svg,script,style'))continue;
     const range=document.createRange();range.selectNodeContents(n);const rs=Array.from(range.getClientRects());if(!rs.length||rs.every(r=>!r.height||!r.width))continue;
     const s=getComputedStyle(e);nodes.push({text:n.textContent,font:s.fontFamily,size:s.fontSize,weight:s.fontWeight,lineHeight:s.lineHeight,transform:s.textTransform,rects:rs.map(rect)});
    }
    return {id:b.getAttribute('data-block-id'),type:b.getAttribute('data-block-type'),rect:rect(b.getBoundingClientRect()),areas:Array.from(b.querySelectorAll('[data-editable]')).map(e=>({key:e.getAttribute('data-editable'),text:e.textContent,rect:rect(e.getBoundingClientRect()),lineHeight:getComputedStyle(e).lineHeight,size:getComputedStyle(e).fontSize,minHeight:getComputedStyle(e).minHeight})),nodes};
   })};
  })};
 });
 if(snapshot.pages.some(p=>p.images.some(i=>!i.complete||!i.naturalWidth)))throw new Error('An image is missing from the rendered document');
 writeFileSync(`${out}/${id}-${mode}.json`,JSON.stringify({snapshot,status,errors},null,2));
 console.log(JSON.stringify({id,mode,event:'snapshot',ms:Date.now()-start,pages:snapshot.pages.length,blocks:snapshot.pages.flatMap(p=>p.blocks).length,nodes:snapshot.pages.flatMap(p=>p.blocks.flatMap(b=>b.nodes)).length,status,errors}));
 if(mode==='export'){
  const meta=await page.evaluate(()=>window.__PDF_META);
  await page.setViewport({width:meta.pageWidthPx,height:meta.pageHeightPx,deviceScaleFactor:1});await page.evaluate(()=>window.__PDF_SETTLE?.());
  const opts={width:`${meta.pageWidthPx}px`,height:`${meta.pageHeightPx}px`,printBackground:true,tagged:true,margin:{top:0,right:0,bottom:0,left:0}};
  await page.pdf({...opts,path:`${out}/${id}-before.pdf`});
  const masks=await preparePdfMasks(page);
  await page.pdf({...opts,path:`${out}/${id}-after.pdf`});
  console.log(JSON.stringify({id,mode,event:'pdf',masks,ms:Date.now()-start}));
 }
 if(process.env.FIDELITY_SCREENSHOTS==='1'){const pages=await page.$$('[data-page-index]');for(let i=0;i<pages.length;i++)await pages[i].screenshot({path:`${out}/${id}-${mode}-${i+1}.png`});}
 await context.close();
}}finally{await browser.close()}
