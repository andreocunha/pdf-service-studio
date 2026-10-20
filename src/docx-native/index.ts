/**
 * Word nativo: o documento montado a partir do layout real do Studio (a mesma
 * página que vira PDF), no padrão da Lex — em vez de reconstruído a partir
 * do PDF como faz o iLovePDF. Ver build.ts.
 */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { Packer } from 'docx';

import { embedFonts } from '../docx-fonts.js';
import { deflateSync } from 'node:zlib';
import { makeAssets, SCALE, type ImageData } from '../office/assets.js';
import { loadFontResolver } from '../office/fonts.js';
import { readLayout, withoutGlass } from '../office/layout.js';
import { coverPagesOf, pageBackgrounds, type PageOffsets } from '../office/pages.js';
import { withRenderPage } from '../render.js';
import { buildDocument, PAGE_FIELD } from './build.js';

/**
 * Troca o marcador do rodapé pelos campos do Word: página atual e total, com
 * a mesma defasagem do Studio ("02/28" na terceira página) e dois dígitos.
 */
const pageFields = (xml: string, offsets: PageOffsets | null): string =>
  xml.replace(
    new RegExp(`<w:r>(<w:rPr>(?:(?!</w:r>).)*?</w:rPr>)?<w:t[^>]*>${PAGE_FIELD}</w:t></w:r>`, 'g'),
    (_, rPr = '') => {
      const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`;
      // "01/12" leva zero à esquerda (como no Studio); "1" e "Página 1 de 12", não.
      const padded = (offsets?.format ?? 'page-total') === 'page-total';
      const pic = padded ? ' \\# "00"' : '';
      const field = (name: string, offset: number) => {
        const inner = offset
          ? [
              run('<w:fldChar w:fldCharType="begin"/>'),
              run('<w:instrText xml:space="preserve"> = </w:instrText>'),
              run('<w:fldChar w:fldCharType="begin"/>'),
              run(`<w:instrText xml:space="preserve"> ${name} </w:instrText>`),
              run('<w:fldChar w:fldCharType="separate"/>'),
              run('<w:t>1</w:t>'),
              run('<w:fldChar w:fldCharType="end"/>'),
              run(`<w:instrText xml:space="preserve"> ${offset > 0 ? '+' : '-'} ${Math.abs(offset)}${pic} </w:instrText>`),
            ]
          : [run('<w:fldChar w:fldCharType="begin"/>'), run(`<w:instrText xml:space="preserve"> ${name}${pic} </w:instrText>`)];
        return [...inner, run('<w:fldChar w:fldCharType="separate"/>'), run(`<w:t>${padded ? '01' : '1'}</w:t>`), run('<w:fldChar w:fldCharType="end"/>')].join('');
      };
      const page = field('PAGE', offsets?.current ?? 0);
      const total = () => field('NUMPAGES', offsets?.total ?? 0);
      if (offsets?.format === 'page') return page;
      if (offsets?.format === 'page-of-total') return run('<w:t xml:space="preserve">Página </w:t>') + page + run('<w:t xml:space="preserve"> de </w:t>') + total();
      return page + run('<w:t>/</w:t>') + total();
    },
  );

/** PNG 8×8 de uma cor (o Word estica até a página). */
const solidPng = (hex: string): ImageData => {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (buf: Buffer) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type), data]);
    const out = Buffer.alloc(body.length + 8);
    out.writeUInt32BE(data.length, 0);
    body.copy(out, 4);
    out.writeUInt32BE(crc(body), body.length + 4);
    return out;
  };
  const size = 8;
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // 8 bits
  ihdr[9] = 2; // RGB
  const rgb = Buffer.from(hex, 'hex');
  const row = Buffer.concat([Buffer.from([0]), ...Array.from({ length: size }, () => rgb)]);
  const raw = Buffer.concat(Array.from({ length: size }, () => row));
  const data = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
  return { data, type: 'png' };
};

export const renderNativeDocx = async (args: { documentId: string; workspaceId: string; appUrl: string; signal?: AbortSignal }): Promise<{ buffer: Buffer; title: string }> => {
  const fonts = await loadFontResolver(args.appUrl);
  return withRenderPage(args, async ({ page, meta }) => {
    const layout = withoutGlass(await readLayout(page, meta));
    // Nitidez das capturas (capa, ícones): só a densidade muda, o layout não.
    await page.setViewport({ width: meta.pageWidthPx, height: meta.pageHeightPx, deviceScaleFactor: SCALE });
    const assets = makeAssets(page, args.signal);
    const covers = coverPagesOf(layout);

    // Página de cor lisa (sem imagem de fundo): a cor vai como imagem no
    // cabeçalho, igual ao fundo em imagem — o Word não tem cor por página, e
    // as caixas brancas em cima dela sumiam na página branca.
    const pageArt = pageBackgrounds(page, layout, covers, assets);
    const background = async (index: number) => {
      const art = await pageArt(index);
      const color = layout.pages[index]?.bgColor;
      return art ?? (color && color !== 'ffffff' ? solidPng(color) : null);
    };

    const { doc, used, pageOffsets } = await buildDocument({ layout, fonts, assets, background, coverPages: covers });
    args.signal?.throwIfAborted();
    const packed = await Packer.toBuffer(doc);
    const files = unzipSync(new Uint8Array(packed)) as Zippable;
    for (const path of Object.keys(files).filter((p) => /^word\/(header|footer)\d+\.xml$/.test(p))) {
      files[path] = strToU8(pageFields(strFromU8(files[path] as Uint8Array), pageOffsets));
    }
    // Padrão da equipe: controle de órfãs/viúvas ligado e nada de "manter
    // com o próximo" (ele jogava cards inteiros pra página seguinte quando o
    // texto crescia).
    files['word/styles.xml'] = strToU8(
      strFromU8(files['word/styles.xml'] as Uint8Array).replace('<w:pPrDefault><w:pPr>', '<w:pPrDefault><w:pPr><w:widowControl/>'),
    );
    // Parágrafo que carrega a quebra de seção: mínimo. Do tamanho de uma
    // linha, numa página cheia ele caía na folha seguinte — e a página nova
    // ficava com o cabeçalho da seção anterior.
    files['word/document.xml'] = strToU8(
      strFromU8(files['word/document.xml'] as Uint8Array).replace(
        /<w:p><w:pPr><w:sectPr/g,
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:sectPr',
      ),
    );
    for (const path of Object.keys(files).filter((p) => /^word\/(document|header\d+|footer\d+)\.xml$/.test(p))) {
      // Ordem de empilhamento pela ordem no documento: o fundo do card de fora
      // vem antes e fica embaixo do de dentro (a lib docx usa a altura da
      // imagem como ordem — a caixa maior cobria a menor).
      let z = 1;
      files[path] = strToU8(strFromU8(files[path] as Uint8Array).replace(/relativeHeight="\d+"/g, () => `relativeHeight="${z++}"`));
    }
    if (!files['word/fontTable.xml']) {
      files['word/fontTable.xml'] = new TextEncoder().encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></w:fonts>',
      );
    }
    await embedFonts(files, used, `${args.appUrl}/word-fonts/`, args.signal);
    args.signal?.throwIfAborted();
    return { buffer: Buffer.from(zipSync(files, { level: 6 })), title: meta.title ?? 'document' };
  });
};
