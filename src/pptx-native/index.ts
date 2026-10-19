/**
 * PowerPoint nativo: os slides montados a partir do layout real do Studio
 * (a mesma página que vira PDF), em vez de reconstruídos a partir do PDF
 * como faz o iLovePDF. Leitura do layout, imagens e fontes são as mesmas do
 * Word nativo (src/office); só a montagem muda (build.ts).
 */
import { strToU8, zipSync, type Zippable } from 'fflate';

import { makeAssets, SCALE } from '../office/assets.js';
import { ttfToEot } from '../office/eot.js';
import { embeddableFamilies, fontSlot, loadFontResolver, loadFontFile, type Face } from '../office/fonts.js';
import { readLayout } from '../office/layout.js';
import { pageBackgrounds } from '../office/pages.js';
import { withRenderPage } from '../render.js';
import { buildDeck, emu, esc, type Deck } from './build.js';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.presentationml';

const rels = (items: { id: string; type: string; target: string; external?: boolean }[]) =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  items.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('') +
  '</Relationships>';

/** Tema do Office com a fonte principal do documento (texto novo já sai nela). */
const theme = (font: string) => {
  const f = esc(font);
  const fontScheme = (tag: string) => `<a:${tag}><a:latin typeface="${f}"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`;
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const line = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${fill}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  return (
    `${XML}<a:theme xmlns:a="${NS_A}" name="Lex"><a:themeElements>` +
    '<a:clrScheme name="Lex"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
    `<a:fontScheme name="Lex">${fontScheme('majorFont')}${fontScheme('minorFont')}</a:fontScheme>` +
    `<a:fmtScheme name="Lex"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst><a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>` +
    '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
  );
};

const EMPTY_TREE =
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>';

const master = () =>
  `${XML}<p:sldMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${EMPTY_TREE}</p:cSld>` +
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
  '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="3200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>' +
  '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle>' +
  '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>';

const layoutXml = () =>
  `${XML}<p:sldLayout xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" preserve="1" type="blank"><p:cSld name="Em branco">${EMPTY_TREE}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

/** Fontes usadas embutidas (EOT), pra abrir certo em quem não tem a fonte instalada. */
const embedFonts = async (files: Zippable, used: Set<Face>, baseUrl: string) => {
  const families = embeddableFamilies(used);
  const fontRels: { id: string; target: string }[] = [];
  const entries: string[] = [];
  let n = 0;
  for (const [family, faces] of families) {
    const slots: string[] = [];
    for (const face of faces) {
      try {
        const ttf = await loadFontFile(new URL(face.file!, baseUrl).href);
        n++;
        files[`ppt/fonts/font${n}.fntdata`] = new Uint8Array(ttfToEot(ttf));
        const id = `rIdFont${n}`;
        fontRels.push({ id, target: `fonts/font${n}.fntdata` });
        slots.push(`<p:${fontSlot(face)} r:id="${id}"/>`);
      } catch {
        // Sem o arquivo: a fonte fica só pelo nome (abre certo onde estiver instalada).
      }
    }
    if (slots.length) entries.push(`<p:embeddedFont><p:font typeface="${esc(family)}" pitchFamily="2" charset="0"/>${slots.join('')}</p:embeddedFont>`);
  }
  return { fontRels, list: entries.length ? `<p:embeddedFontLst>${entries.join('')}</p:embeddedFontLst>` : '' };
};

const pack = async (deck: Deck, title: string, fontsUrl: string): Promise<Buffer> => {
  const files: Zippable = {};
  const put = (path: string, xml: string) => (files[path] = strToU8(xml));
  const fonts = await embedFonts(files, deck.used, fontsUrl);
  const n = deck.slides.length;

  put(
    '[Content_Types].xml',
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/>' +
      '<Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="fntdata" ContentType="application/x-fontdata"/>' +
      `<Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentation.main+xml"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT}.slideMaster+xml"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${CT}.slideLayout+xml"/>` +
      deck.slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${CT}.slide+xml"/>`).join('') +
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      `<Override PartName="/ppt/presProps.xml" ContentType="${CT}.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${CT}.viewProps+xml"/>` +
      `<Override PartName="/ppt/tableStyles.xml" ContentType="${CT}.tableStyles+xml"/>` +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',
  );
  put(
    '_rels/.rels',
    rels([
      { id: 'rId1', type: `${REL}/officeDocument`, target: 'ppt/presentation.xml' },
      { id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' },
      { id: 'rId3', type: `${REL}/extended-properties`, target: 'docProps/app.xml' },
    ]),
  );
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  put(
    'docProps/core.xml',
    `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${esc(title)}</dc:title><dc:creator>Lex Studio</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  put(
    'docProps/app.xml',
    `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Lex Studio</Application><Slides>${n}</Slides></Properties>`,
  );

  const slideRels = deck.slides.map((_, i) => ({ id: `rId${i + 2}`, type: `${REL}/slide`, target: `slides/slide${i + 1}.xml` }));
  put(
    'ppt/_rels/presentation.xml.rels',
    rels([
      { id: 'rId1', type: `${REL}/slideMaster`, target: 'slideMasters/slideMaster1.xml' },
      ...slideRels,
      { id: `rId${n + 2}`, type: `${REL}/theme`, target: 'theme/theme1.xml' },
      { id: `rId${n + 3}`, type: `${REL}/presProps`, target: 'presProps.xml' },
      { id: `rId${n + 4}`, type: `${REL}/viewProps`, target: 'viewProps.xml' },
      { id: `rId${n + 5}`, type: `${REL}/tableStyles`, target: 'tableStyles.xml' },
      ...fonts.fontRels.map((f) => ({ id: f.id, type: `${REL}/font`, target: f.target })),
    ]),
  );
  // Tamanho do slide = tamanho da página do Studio (o PowerPoint aceita de 1 a 56 polegadas).
  const clamp = (v: number) => Math.min(51206400, Math.max(914400, emu(v))); // renderNativePptx já recusa fora disso
  put(
    'ppt/presentation.xml',
    `${XML}<p:presentation xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" saveSubsetFonts="1"${fonts.list ? ' embedTrueTypeFonts="1"' : ''}${deck.firstSlideNum !== 1 ? ` firstSlideNum="${deck.firstSlideNum}"` : ''}>` +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      `<p:sldIdLst>${slideRels.map((r, i) => `<p:sldId id="${256 + i}" r:id="${r.id}"/>`).join('')}</p:sldIdLst>` +
      `<p:sldSz cx="${clamp(deck.size.w)}" cy="${clamp(deck.size.h)}"/><p:notesSz cx="6858000" cy="9144000"/>${fonts.list}` +
      '<p:defaultTextStyle><a:lvl1pPr><a:defRPr lang="pt-BR"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>',
  );
  put('ppt/presProps.xml', `${XML}<p:presentationPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"/>`);
  put(
    'ppt/viewProps.xml',
    `${XML}<p:viewPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`,
  );
  put('ppt/tableStyles.xml', `${XML}<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
  put('ppt/theme/theme1.xml', theme(deck.mainFace?.family ?? 'Calibri'));
  put('ppt/slideMasters/slideMaster1.xml', master());
  put(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    rels([
      { id: 'rId1', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' },
      { id: 'rId2', type: `${REL}/theme`, target: '../theme/theme1.xml' },
    ]),
  );
  put('ppt/slideLayouts/slideLayout1.xml', layoutXml());
  put('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([{ id: 'rId1', type: `${REL}/slideMaster`, target: '../slideMasters/slideMaster1.xml' }]));

  deck.slides.forEach((slide, i) => {
    put(`ppt/slides/slide${i + 1}.xml`, slide.xml);
    put(
      `ppt/slides/_rels/slide${i + 1}.xml.rels`,
      rels([
        { id: 'rId1', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' },
        ...slide.rels.map((r) => ({ id: r.id, type: `${REL}/${r.type}`, target: r.target, external: r.type === 'hyperlink' })),
      ]),
    );
  });
  for (const m of deck.media) files[`ppt/media/${m.name}`] = [new Uint8Array(m.data), { level: 0 }];

  return Buffer.from(zipSync(files, { level: 6 }));
};

export const renderNativePptx = async (args: { documentId: string; workspaceId: string; appUrl: string; signal?: AbortSignal }): Promise<{ buffer: Buffer; title: string }> => {
  const fonts = await loadFontResolver(args.appUrl);
  return withRenderPage(args, async ({ page, meta }) => {
    // O PowerPoint só aceita slide de 1 a 56 polegadas: página contínua
    // muito longa sairia cortada — falha aqui e o /pptx cai no iLovePDF.
    for (const v of [meta.pageWidthPx, meta.pageHeightPx]) {
      if (emu(v) < 914400 || emu(v) > 51206400) throw new Error(`página de ${Math.round(v)}px não cabe num slide`);
    }
    const layout = await readLayout(page, meta);
    // Nitidez das capturas (ícones, gradientes): só a densidade muda, o layout não.
    await page.setViewport({ width: meta.pageWidthPx, height: meta.pageHeightPx, deviceScaleFactor: SCALE });
    const assets = makeAssets(page);
    // Sem capa "em foto": no PowerPoint tudo fica no lugar, editável.
    const background = pageBackgrounds(page, layout, new Set(), assets);
    const deck = await buildDeck({ layout, fonts, assets, background });
    const title = meta.title ?? 'document';
    return { buffer: await pack(deck, title, `${args.appUrl}/word-fonts/`), title };
  });
};

/** Só pra testes: monta o pacote a partir de um deck pronto. */
export const packDeck = pack;
