/**
 * npm run test:docx-fonts -- <pdf> <docx-do-ilovepdf> [saida.docx]
 * (WORD_FONTS_URL=https://app.lexstudio.ai/word-fonts/ pra usar as de produção)
 *
 * Roda só a etapa de fontes sobre um par PDF + docx já convertido (sem gastar
 * iLovePDF) e confere o resultado: toda face usada no PDF tem que aparecer no
 * docx e estar embutida quando o manifest tem o arquivo.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { strict as assert } from 'node:assert';

import { strFromU8, unzipSync } from 'fflate';

import { restoreDocxFonts } from '../src/docx-fonts.js';

const [pdfPath, docxPath, outPath = docxPath.replace(/\.docx$/, '-fontes.docx')] = process.argv.slice(2);
if (!pdfPath || !docxPath) {
  console.error('uso: npm run test:docx-fonts -- <pdf> <docx> [saida.docx]');
  process.exit(1);
}

const t0 = Date.now();
// Fontes do Word servidas pelo app (lex-studio-v2 rodando local, por padrão).
const fontsUrl = process.env.WORD_FONTS_URL ?? 'http://localhost:3000/word-fonts/';
const { docx, report } = await restoreDocxFonts(readFileSync(docxPath), readFileSync(pdfPath), fontsUrl);
writeFileSync(outPath, docx);
console.log(JSON.stringify({ ...report, elapsedMs: Date.now() - t0, out: outPath }, null, 1));

const files = unzipSync(new Uint8Array(docx));
const fontTable = strFromU8(files['word/fontTable.xml']);
const rels = strFromU8(files['word/_rels/fontTable.xml.rels'] ?? new Uint8Array());
for (const [path] of Object.entries(files).filter(([p]) => p.startsWith('word/fonts/'))) {
  assert(rels.includes(path.replace('word/', '')), `${path} sem relationship`);
}
const embeds = [...fontTable.matchAll(/r:id="([^"]+)"/g)].map((m) => m[1]);
for (const id of embeds) assert(rels.includes(`Id="${id}"`), `fontTable aponta pra ${id} inexistente`);
// Ordem do schema — o Word recusa o arquivo com filho fora de lugar.
for (const part of Object.keys(files).filter((p) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(p))) {
  const xml = strFromU8(files[part]);
  for (const [pPr] of xml.matchAll(/<w:pPr>[\s\S]*?<\/w:pPr>/g)) {
    const top = pPr.replace(/<w:rPr\s*\/>|<w:rPr\b[\s\S]*?<\/w:rPr>/g, '<RPR/>');
    assert((top.match(/<RPR\/>/g) ?? []).length <= 1, `${part}: pPr com dois rPr`);
    const spacing = top.indexOf('<w:spacing');
    assert(spacing < 0 || top.indexOf('<RPR/>') < 0 || spacing < top.indexOf('<RPR/>'), `${part}: spacing depois do rPr`);
  }
  for (const [rPr] of xml.matchAll(/<w:rPr>[\s\S]*?<\/w:rPr>/g)) {
    assert(!/<w:spacing[^>]*w:line=/.test(rPr), `${part}: w:line dentro de rPr`);
    const fonts = rPr.indexOf('<w:rFonts');
    const size = rPr.indexOf('<w:sz ');
    assert(fonts < 0 || size < 0 || fonts < size, `${part}: rFonts depois de sz`);
  }
}
assert(report.embedded.length > 0, 'nenhuma fonte embutida');
assert(report.matched / Math.max(1, report.runs) > 0.9, `só ${report.matched}/${report.runs} runs achados no PDF`);
console.log('ok');
