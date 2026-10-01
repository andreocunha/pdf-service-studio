/**
 * Word nativo: o documento montado a partir do layout real do Studio (a mesma
 * página que vira PDF), no padrão da Lex — em vez de reconstruído a partir
 * do PDF como faz o iLovePDF. Ver build.ts.
 */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { Packer } from 'docx';
import type { Page } from 'puppeteer-core';

import { getBrowser } from '../browser.js';
import { embedFonts } from '../docx-fonts.js';
import { withRenderPage } from '../render.js';
import { buildDocument, PAGE_FIELD, type Assets } from './build.js';
import { loadFontResolver } from './fonts.js';
import type { Box, DocumentLayout } from './layout.js';
import { readFileSync } from 'node:fs';

const EXTRACT = readFileSync(new URL('./extract.js', import.meta.url), 'utf8');

type ImageData = { data: Buffer; type: 'png' | 'jpg' | 'gif' | 'bmp'; svg?: Buffer; aspect?: number };

/** Proporção do desenho de um SVG (viewBox, ou largura/altura). */
const svgAspect = (svg: Buffer): number | undefined => {
  const head = svg.toString('utf8', 0, 4000).match(/<svg[^>]*>/)?.[0] ?? '';
  const vb = head.match(/viewBox=["']\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)/);
  if (vb && +vb[1] && +vb[2]) return +vb[1] / +vb[2];
  const w = parseFloat(head.match(/\swidth=["']([\d.]+)/)?.[1] ?? '');
  const h = parseFloat(head.match(/\sheight=["']([\d.]+)/)?.[1] ?? '');
  return w && h ? w / h : undefined;
};

/** Resolução das imagens rasterizadas (capa, fallback de SVG). */
const SCALE = 2;

const typeOf = (contentType: string, url: string): ImageData['type'] | 'svg' | null => {
  if (/svg/.test(contentType) || /\.svg(\?|$)/i.test(url)) return 'svg';
  if (/png/.test(contentType)) return 'png';
  if (/jpe?g/.test(contentType)) return 'jpg';
  if (/gif/.test(contentType)) return 'gif';
  return null;
};

/** SVG → PNG no próprio Chromium (fundo transparente), pro Word que não lê SVG. */
const svgToPng = async (svg: Buffer, w: number, h: number): Promise<Buffer> => {
  const browser = await getBrowser();
  const page = await browser.newPage();
  try {
    await page.setViewport({ width: Math.ceil(w), height: Math.ceil(h), deviceScaleFactor: SCALE });
    const src = `data:image/svg+xml;base64,${svg.toString('base64')}`;
    await page.setContent(
      `<html><body style="margin:0;background:transparent"><img src="${src}" style="display:block;width:${w}px;height:${h}px"></body></html>`,
      { waitUntil: 'load' },
    );
    return Buffer.from(await page.screenshot({ type: 'png', omitBackground: true, clip: { x: 0, y: 0, width: w, height: h } }));
  } finally {
    await page.close().catch(() => undefined);
  }
};

const makeAssets = (page: Page): Assets & { cache: Map<string, Promise<ImageData | null>> } => {
  const cache = new Map<string, Promise<ImageData | null>>();
  const image = (src: string, w: number, h: number) => {
    const key = `${src}|${Math.round(w)}x${Math.round(h)}`;
    let hit = cache.get(key);
    if (!hit) {
      hit = (async () => {
        try {
          const res = await fetch(src, { signal: AbortSignal.timeout(15_000) });
          if (!res.ok) return null;
          const data = Buffer.from(await res.arrayBuffer());
          const type = typeOf(res.headers.get('content-type') ?? '', src);
          // O navegador desenha o SVG sem distorcer (preserveAspectRatio); o Word estica até a caixa.
          if (type === 'svg') return { svg: data, data: await svgToPng(data, w, h), type: 'png', aspect: svgAspect(data) };
          return type ? { data, type } : null;
        } catch {
          return null;
        }
      })();
      cache.set(key, hit);
    }
    return hit;
  };
  /**
   * Captura só o elemento, com fundo transparente: o resto da página some
   * (o texto por cima da arte de um banner sairia duplicado — ele já vai
   * editável no Word; o fundo da página já está no cabeçalho). `alone`: só a
   * pintura da própria caixa, sem os filhos; `notext`: tudo menos o texto
   * (bloco complexo — o texto vai editável por cima da captura).
   */
  const capture = async (id: number, mode: 'all' | 'alone' | 'notext', clip?: Box): Promise<ImageData | null> => {
    const alone = mode === 'alone';
    const el = await page.$(`[data-dx="${id}"]`);
    if (!el) return null;
    await page.evaluate(`(() => {
      const target = document.querySelector('[data-dx="${id}"]');
      const set = (e, v) => { if (!e.hasAttribute('data-dx-vis')) e.setAttribute('data-dx-vis', e.style.visibility); e.style.visibility = v; };
      for (const e of document.querySelectorAll('body *')) set(e, 'hidden');
      set(target, 'visible');
      for (const e of target.querySelectorAll('*')) set(e, ${alone} && !e.closest('[data-dx-baked]') ? 'hidden' : 'visible');
      // A página do PDF força "html, body { background: white !important }": só inline !important ganha.
      for (const e of [document.documentElement, document.body]) { e.setAttribute('data-dx-bg', e.style.background); e.style.setProperty('background', 'transparent', 'important'); }
      if (${mode === 'notext'}) {
        const style = document.createElement('style');
        style.id = 'dx-notext';
        style.textContent = '[data-dx-text], [data-dx-text] * { color: transparent !important; -webkit-text-fill-color: transparent !important; text-decoration-color: transparent !important; text-shadow: none !important; }';
        document.head.appendChild(style);
      }
    })()`);
    try {
      const data = Buffer.from(
        clip
          ? await page.screenshot({ type: 'png', omitBackground: true, captureBeyondViewport: true, clip: { x: clip.x, y: clip.y, width: clip.w, height: clip.h } })
          : await el.screenshot({ type: 'png', omitBackground: true }),
      );
      return { data, type: 'png' as const };
    } finally {
      await page.evaluate(`(() => {
        for (const e of document.querySelectorAll('[data-dx-vis]')) { e.style.visibility = e.getAttribute('data-dx-vis'); e.removeAttribute('data-dx-vis'); }
        for (const e of [document.documentElement, document.body]) { e.style.removeProperty('background'); e.style.background = e.getAttribute('data-dx-bg') || ''; e.removeAttribute('data-dx-bg'); }
        document.getElementById('dx-notext')?.remove();
      })()`);
    }
  };
  const raster = (id: number, _box: Box) => capture(id, 'all');
  const paint = (id: number) => capture(id, 'alone');
  const backdrop = (id: number, clip: Box) => capture(id, 'notext', clip);
  return { image, raster, paint, backdrop, cache };
};

/**
 * Troca o marcador do rodapé pelos campos do Word: página atual e total, com
 * a mesma defasagem do Studio ("02/28" na terceira página) e dois dígitos.
 */
const pageFields = (xml: string, offsets: { current: number; total: number } | null): string =>
  xml.replace(
    new RegExp(`<w:r>(<w:rPr>(?:(?!</w:r>).)*?</w:rPr>)?<w:t[^>]*>${PAGE_FIELD}</w:t></w:r>`, 'g'),
    (_, rPr = '') => {
      const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`;
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
              run(`<w:instrText xml:space="preserve"> ${offset > 0 ? '+' : '-'} ${Math.abs(offset)} \\# "00" </w:instrText>`),
            ]
          : [run('<w:fldChar w:fldCharType="begin"/>'), run(`<w:instrText xml:space="preserve"> ${name} \\# "00" </w:instrText>`)];
        return [...inner, run('<w:fldChar w:fldCharType="separate"/>'), run('<w:t>01</w:t>'), run('<w:fldChar w:fldCharType="end"/>')].join('');
      };
      return field('PAGE', offsets?.current ?? 0) + run('<w:t>/</w:t>') + field('NUMPAGES', offsets?.total ?? 0);
    },
  );

/**
 * Capa (e contracapa): vai inteira como imagem no cabeçalho, não editável,
 * como no Studio. É capa a página cujo único bloco ocupa a página toda, ou a
 * primeira/última página com fundo próprio (diferente da vizinha) — montada
 * com vários blocos sobre a arte, encostados no pé da página.
 */
const coverPagesOf = (layout: DocumentLayout): Set<number> => {
  const covers = new Set<number>();
  const { pages } = layout;
  for (const p of pages) {
    const blocks = layout.blocks.filter((b) => b.page === p.index && b.tree);
    if (blocks.length === 1 && blocks[0].box.h >= layout.meta.pageHeightPx * 0.9) covers.add(p.index);
  }
  const ownBackground = (i: number, neighbour: number) =>
    pages[i]?.bgImage && pages[neighbour] && pages[i].bgImage !== pages[neighbour].bgImage;
  if (pages.length > 1 && ownBackground(0, 1)) covers.add(0);
  // Contracapa só com pouco texto: a última página do anexo, com fundo
  // próprio mas conteúdo corrido, não é capa (ia inteira pro cabeçalho).
  const chars = (i: number) => {
    let n = 0;
    const walk = (t: import('./layout.js').LayoutNode | null): void => {
      if (!t) return;
      if (t.k === 'text') n += t.paras.reduce((a, p) => a + p.runs.reduce((b, r) => b + (r.t?.length ?? 0), 0), 0);
      else if (t.k === 'box') t.kids.forEach(walk);
    };
    layout.blocks.filter((b) => b.page === i).forEach((b) => walk(b.tree));
    return n;
  };
  const last = pages.length - 1;
  if (pages.length > 2 && ownBackground(last, last - 1) && chars(last) <= 150) covers.add(last);
  return covers;
};

export const renderNativeDocx = async (args: { documentId: string; workspaceId: string; appUrl: string }): Promise<{ buffer: Buffer; title: string }> => {
  const fonts = await loadFontResolver(args.appUrl);
  return withRenderPage(args, async ({ page, meta }) => {
    const layout = (await page.evaluate(EXTRACT)) as DocumentLayout;
    layout.meta = { ...layout.meta, ...meta };
    // Nitidez das capturas (capa, ícones): só a densidade muda, o layout não.
    await page.setViewport({ width: meta.pageWidthPx, height: meta.pageHeightPx, deviceScaleFactor: SCALE });
    const assets = makeAssets(page);
    const covers = coverPagesOf(layout);

    const background = async (index: number): Promise<ImageData | null> => {
      const p = layout.pages[index];
      if (!p) return null;
      if (covers.has(index)) {
        // Capa: a arte da página inteira vira fundo no cabeçalho (travada);
        // os textos saem da captura e vão editáveis no corpo, por cima — os
        // campos "CONTRATANTE: Dados" precisam ser preenchidos.
        await page.evaluate(`(() => {
          const style = document.createElement('style');
          style.id = 'dx-cover';
          style.textContent = '[data-dx-text], [data-dx-text] * { color: transparent !important; -webkit-text-fill-color: transparent !important; text-decoration-color: transparent !important; text-shadow: none !important; }';
          document.head.appendChild(style);
        })()`);
        try {
          const shot = await page.screenshot({
            type: 'png',
            clip: { x: p.box.x, y: p.box.y, width: p.box.w, height: p.box.h },
            captureBeyondViewport: true,
          });
          return { data: Buffer.from(shot), type: 'png' };
        } finally {
          await page.evaluate(`document.getElementById('dx-cover')?.remove()`);
        }
      }
      if (!p.bgImage) return null;
      // Fundo vai como PNG desenhado pelo Chromium: o Word não desenha certo
      // todo SVG (os desenhos do pé do Quadro resumo sumiam) e fundo não se edita.
      const img = await assets.image(p.bgImage, p.box.w, p.box.h);
      return img?.svg ? { data: img.data, type: img.type } : img;
    };

    const { doc, used, pageOffsets } = await buildDocument({ layout, fonts, assets, background, coverPages: covers });
    const packed = await Packer.toBuffer(doc);
    const files = unzipSync(new Uint8Array(packed)) as Zippable;
    for (const path of Object.keys(files).filter((p) => /^word\/footer\d+\.xml$/.test(p))) {
      files[path] = strToU8(pageFields(strFromU8(files[path] as Uint8Array), pageOffsets));
    }
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
    await embedFonts(files, used, `${args.appUrl}/word-fonts/`);
    return { buffer: Buffer.from(zipSync(files, { level: 6 })), title: meta.title ?? 'document' };
  });
};
