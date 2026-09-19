import assert from 'node:assert/strict';
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFString } from 'pdf-lib';
import { normalizePdfLinks } from '../src/pdf-links.js';

const name = PDFName.of;
const pdf = await PDFDocument.create();
const first = pdf.addPage([600, 800]);
const last = pdf.addPage([600, 800]);
first.drawText('Menu');
last.drawText('Destination');
const c = pdf.context;
const target = c.obj([last.ref, 'XYZ', 6, 350.88, 0]);
const samePage = c.obj([first.ref, 'XYZ', 6, 123.45, null]);
pdf.catalog.set(name('Dests'), c.obj({ legacy: target, wrapped: { D: samePage }, cycle: PDFName.of('cycle') }));
const leaf = c.obj({ Names: [PDFString.of('unicode-ç'), { D: target }, PDFHexString.fromText('hex'), target] });
const tree = c.obj({ Kids: [c.register(leaf)] });
pdf.catalog.set(name('Names'), c.obj({ Dests: c.register(tree) }));

const link = (extra: Record<string, unknown>) => c.obj({ Type: 'Annot', Subtype: 'Link', Rect: [10, 20, 200, 45], Border: [0, 0, 0], ...extra });
const annotations = [
  link({ Dest: name('legacy') }),
  link({ Dest: name('wrapped') }),
  link({ Dest: PDFString.of('unicode-ç') }),
  link({ A: { S: 'GoTo', D: PDFHexString.fromText('hex') } }),
  link({ Dest: target }),
  link({ A: { S: 'GoTo', D: target } }),
  link({ A: { S: 'URI', URI: PDFString.of('https://example.com/#menu') } }),
  link({ A: { S: 'URI', URI: PDFString.of('https://wa.me/5511999999999') } }),
  link({ A: { S: 'GoToR', F: PDFString.of('other.pdf'), D: name('legacy') } }),
  link({ Dest: name('missing') }),
  link({ Dest: name('cycle') }),
  link({ Dest: c.obj([c.register(c.obj({})), 'Fit']) }),
  link({ A: { S: 'GoTo', D: [last.ref, 'XYZ', 6, 350.88, 2] } }),
  link({ A: { S: 'GoTo', D: [last.ref, 'XYZ', 6, null, 0] } }),
];
first.node.set(name('Annots'), c.obj(annotations.map(a => c.register(a))));
const input = Buffer.from(await pdf.save());
const result = await normalizePdfLinks(input);
assert.equal(result.converted, 6);
assert.equal(result.unresolved, 3);
const output = await PDFDocument.load(result.buffer);
assert.equal(output.getPageCount(), 2);
const actual = output.getPage(0).node.lookup(name('Annots'), PDFArray);
for (let i = 0; i < 6; i++) {
  const annotation = actual.lookup(i, PDFDict);
  assert.equal(annotation.has(name('Dest')), false);
  const action = annotation.lookup(name('A'), PDFDict);
  assert.equal(action.get(name('S'))?.toString(), '/GoTo');
  const expected = i === 1 ? c.obj([first.ref, 'XYZ', null, 123.45, null]) : c.obj([last.ref, 'XYZ', null, 350.88, null]);
  assert.equal(action.lookup(name('D'), PDFArray).toString(), expected.toString());
  assert.equal(annotation.lookup(name('Rect'), PDFArray).toString(), '[ 10 20 200 45 ]');
}
assert.equal(output.catalog.lookup(name('Dests'), PDFDict).lookup(name('legacy'), PDFArray).toString(), target.toString(), 'shared named destinations stay untouched');
for (let i = 6; i < annotations.length; i++) {
  assert.equal(actual.lookup(i, PDFDict).toString(), annotations[i].toString(), `unrelated/unresolved link ${i} preserved`);
}
const again = await normalizePdfLinks(result.buffer);
assert.equal(again.converted, 0);
assert.strictEqual(again.buffer, result.buffer, 'no rewriting when already normalized');
const externalOnly = await PDFDocument.create();
externalOnly.addPage().node.set(name('Annots'), externalOnly.context.obj([{
  Type: 'Annot', Subtype: 'Link', Rect: [0, 0, 100, 20], A: { S: 'URI', URI: PDFString.of('https://maps.google.com/') },
}]));
const externalBytes = Buffer.from(await externalOnly.save());
assert.strictEqual((await normalizePdfLinks(externalBytes)).buffer, externalBytes);

// Previously generated FitH links must retain their exact section coordinates.
const fitPdf = await PDFDocument.create();
const fitPage = fitPdf.addPage([600, 800]);
fitPage.node.set(name('Annots'), fitPdf.context.obj([750, 350, null].map(top => ({
  Type: 'Annot', Subtype: 'Link', Rect: [0, 0, 100, 20],
  A: { S: 'GoTo', D: [fitPage.ref, 'FitH', top] },
}))));
const repaired = await normalizePdfLinks(Buffer.from(await fitPdf.save()));
assert.equal(repaired.converted, 2);
const repairedPdf = await PDFDocument.load(repaired.buffer);
const repairedLinks = repairedPdf.getPage(0).node.lookup(name('Annots'), PDFArray);
for (const [index, top] of [750, 350].entries()) {
  assert.equal(repairedLinks.lookup(index, PDFDict).lookup(name('A'), PDFDict)
    .lookup(name('D'), PDFArray).toString(),
  fitPdf.context.obj([fitPage.ref, 'XYZ', null, top, null]).toString());
}
assert.equal(repairedLinks.lookup(2, PDFDict).lookup(name('A'), PDFDict)
  .lookup(name('D'), PDFArray).toString(), fitPdf.context.obj([fitPage.ref, 'FitH', null]).toString());
assert.strictEqual((await normalizePdfLinks(repaired.buffer)).buffer, repaired.buffer);
console.log('PASS: legacy/name-tree links, XYZ destinations, same-page coordinates, FitH conversion, explicit zooms, external actions, missing/cyclic targets and idempotence. Reader viewport alignment requires device testing.');
