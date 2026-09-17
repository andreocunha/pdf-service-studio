import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { execFileSync } from 'node:child_process';
import puppeteer from 'puppeteer-core';
import chromium from '@sparticuz/chromium';
import { preparePdfMasks } from '../src/pdf-masks.js';

const executablePath = process.env.CHROMIUM_EXECUTABLE_PATH
  || (process.platform === 'darwin' ? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' : await chromium.executablePath());
const bundled = !process.env.CHROMIUM_EXECUTABLE_PATH && process.platform !== 'darwin';
const browser = await puppeteer.launch({
  executablePath,
  headless: bundled ? 'shell' : true,
  args: bundled
    ? puppeteer.defaultArgs({ args: chromium.args, headless: 'shell' }).filter(arg => arg !== '--font-render-hinting=none')
    : ['--no-sandbox'],
});
const out = mkdtempSync(join(tmpdir(), 'lex-pdf-mask-test-'));
try {
  const page = await browser.newPage();
  await page.setViewport({ width: 794, height: 1123 });
  await page.setContent(`<style>
    body{margin:30px;background:#eee;font:18px sans-serif}
    .stamp{box-sizing:border-box;width:343.5px;height:301.25px;padding:8px;background:#3886b6;color:white;
      mask:radial-gradient(circle 8px at 12px 12px,#0000 97%,#000) -12px -12px/24px 24px round,linear-gradient(#000 0 0) content-box !important;}
    .small{width:160.75px;height:30.5px;padding:3px;background:#ffe11e;color:black;
      mask:radial-gradient(circle 3px at 4.5px 4.5px,#0000 97%,#000) -4.5px -4.5px/9px 9px round,linear-gradient(#000 0 0) content-box !important;}
    .soft{mask:linear-gradient(transparent,black);width:100px;height:40px;background:green}
  </style>
  <div id="card" class="stamp"><h1>ENTRADA GRATUITA</h1><a href="https://example.com/">Texto selecionável</a></div>
  <div id="label" class="stamp small">03 · Inscrições</div>
  <div id="soft" class="soft">Soft fade</div>
  <div id="clipped" class="stamp small" style="clip-path:inset(2px)">Existing clip</div>
  <div id="hidden" class="stamp" style="display:none">Hidden</div>`);
  const before = await page.screenshot();
  const layoutBefore = await page.$eval('#card', e => JSON.stringify(e.getBoundingClientRect()));
  await page.pdf({ path: join(out, 'before.pdf'), width: '794px', height: '1123px', printBackground: true });
  assert.equal(await preparePdfMasks(page), 2, 'card and small stamp convert; soft/hidden/already-clipped masks stay intact');
  assert.equal(await preparePdfMasks(page), 0, 'preparation is idempotent');
  assert.equal(await page.$eval('#card', e => JSON.stringify(e.getBoundingClientRect())), layoutBefore, 'layout is unchanged');
  assert.equal(await page.$eval('#card', e => getComputedStyle(e).maskImage), 'none');
  assert.equal(await page.$eval('#card a', e => e.getAttribute('href')), 'https://example.com/');
  assert.equal(await page.$eval('#card a', e => e.textContent), 'Texto selecionável');
  const after = await page.screenshot();
  const error = await page.evaluate(async (sources) => {
    const samples: Uint8ClampedArray[] = [];
    for (const src of sources) {
      const image = new Image(); image.src = src; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
      const ctx = canvas.getContext('2d')!; ctx.drawImage(image, 0, 0);
      samples.push(ctx.getImageData(0, 0, canvas.width, canvas.height).data);
    }
    let error = 0;
    for (let i = 0; i < samples[0].length; i++) error += Math.abs(samples[0][i] - samples[1][i]);
    return error / samples[0].length;
  }, [before, after].map(b => `data:image/png;base64,${Buffer.from(b).toString('base64')}`));
  assert.ok(error < 0.5, `screen appearance changed: mean channel error ${error}`);
  await page.pdf({ path: join(out, 'after.pdf'), width: '794px', height: '1123px', printBackground: true });

  if (process.platform === 'darwin') {
    const renderer = join(out, 'render-quartz');
    execFileSync('swiftc', ['-module-cache-path', join(out, 'swift-cache'), 'scripts/render-quartz.swift', '-o', renderer]);
    execFileSync(renderer, [join(out, 'after.pdf'), join(out, 'quartz'), '1']);
    const source = `data:image/png;base64,${readFileSync(join(out, 'quartz-1.png')).toString('base64')}`;
    const samples = await page.evaluate(async (src) => {
      const image = new Image(); image.src = src; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
      const ctx = canvas.getContext('2d')!; ctx.drawImage(image, 0, 0);
      // Transparent notches at each edge plus solid interior. The original
      // Quartz rendering fails the right/bottom and sometimes interior probes.
      return [[54,31],[31,54],[372,54],[54,330],[330,290]].map(([x,y]) =>
        Array.from(ctx.getImageData(x, y, 1, 1).data).slice(0,3));
    }, source);
    for (const [index, rgb] of samples.entries()) {
      const expected = index < 4 ? [238,238,238] : [56,134,182];
      assert.ok(rgb.every((c,i) => Math.abs(c - expected[i]) < 10), `Quartz probe ${index}: ${rgb}, expected ${expected}`);
    }
    const textBefore = execFileSync('pdftotext', [join(out, 'before.pdf'), '-'], { encoding: 'utf8' });
    const textAfter = execFileSync('pdftotext', [join(out, 'after.pdf'), '-'], { encoding: 'utf8' });
    assert.equal(textAfter, textBefore, 'PDF text remains selectable and unchanged');
  }
  console.log(`PDF masks passed: layout, pixels (${error.toFixed(3)}), text, links, small labels, soft-mask exclusion, idempotence${process.platform === 'darwin' ? ', Quartz edges' : ''}. Artifacts: ${out}`);
} finally {
  await browser.close();
}
