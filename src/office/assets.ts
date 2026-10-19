/**
 * Imagens do Word e do PowerPoint nativos: download (um por URL), SVG e WebP
 * redesenhados pelo Chromium, recorte "cover" como no Studio e capturas de
 * pedaços da página de render (pintura que o Office não desenha).
 */
import type { Page } from 'puppeteer-core';

import { getBrowser } from '../browser.js';
import type { Box } from './layout.js';

export type ImageData = { data: Buffer; type: 'png' | 'jpg' | 'gif' | 'bmp'; svg?: Buffer; aspect?: number };

export type Assets = {
  image: (src: string, w: number, h: number) => Promise<ImageData | null>;
  raster: (id: number, box: Box) => Promise<ImageData | null>;
  /** Só a pintura da caixa (fundo, gradiente, borda, sombra, cantos), sem o conteúdo. */
  paint: (id: number) => Promise<ImageData | null>;
  /** Tudo do elemento menos o texto (bloco complexo). */
  backdrop: (id: number, clip: Box) => Promise<ImageData | null>;
  /** Fundo de página raster: recorte "cover" como no Studio. */
  pageImage: (src: string, w: number, h: number) => Promise<ImageData | null>;
};

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
export const SCALE = 2;

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

/** Largura × altura de PNG/JPEG pelo cabeçalho (sem decodificar). */
const rasterSize = (data: Buffer, type: string): { w: number; h: number } | null => {
  if (type === 'png' && data.length > 24) return { w: data.readUInt32BE(16), h: data.readUInt32BE(20) };
  if (type !== 'jpg') return null;
  for (let i = 2; i + 9 < data.length; ) {
    if (data[i] !== 0xff) return null;
    const marker = data[i + 1];
    const len = data.readUInt16BE(i + 2);
    if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
      return { w: data.readUInt16BE(i + 7), h: data.readUInt16BE(i + 5) };
    }
    i += 2 + len;
  }
  return null;
};

/**
 * Imagem raster redesenhada pelo Chromium no tamanho da caixa, a partir dos
 * bytes já baixados (data URL — sem segundo download nem URL no HTML), com o
 * enquadramento do Studio: o Word estica a imagem até a caixa (o editor
 * recorta com cover) e não lê WebP (a importação do Studio converte fotos
 * para WebP — o fundo sumia no .docx). `cover` sai em JPEG (fundo opaco,
 * arquivo menor); `contain` em PNG (preserva transparência).
 */
const redraw = async (data: Buffer, mime: string, w: number, h: number, fit: 'cover' | 'contain'): Promise<ImageData | null> => {
  const browser = await getBrowser();
  const tab = await browser.newPage();
  try {
    await tab.setViewport({ width: Math.ceil(w), height: Math.ceil(h), deviceScaleFactor: SCALE });
    await tab.setContent(
      `<html><body style="margin:0;background:transparent"><img style="display:block;width:${w}px;height:${h}px;object-fit:${fit};object-position:center"></body></html>`,
    );
    const ok = await tab.evaluate(
      `new Promise((done) => { const i = document.querySelector('img'); i.onload = () => done(i.naturalWidth > 0); i.onerror = () => done(false); i.src = ${JSON.stringify(`data:${mime};base64,${data.toString('base64')}`)}; setTimeout(() => done(false), 15000); })`,
    );
    if (!ok) return null;
    const clip = { x: 0, y: 0, width: w, height: h };
    return fit === 'cover'
      ? { data: Buffer.from(await tab.screenshot({ type: 'jpeg', quality: 88, clip })), type: 'jpg' }
      : { data: Buffer.from(await tab.screenshot({ type: 'png', omitBackground: true, clip })), type: 'png' };
  } catch {
    return null;
  } finally {
    await tab.close().catch(() => undefined);
  }
};

export const makeAssets = (page: Page): Assets => {
  const cache = new Map<string, Promise<ImageData | null>>();
  // Um download por URL, mesmo com a imagem em vários tamanhos/usos.
  const downloads = new Map<string, Promise<{ data: Buffer; mime: string } | null>>();
  const download = (src: string) => {
    let hit = downloads.get(src);
    if (!hit) {
      hit = (async () => {
        try {
          const res = await fetch(src, { signal: AbortSignal.timeout(15_000) });
          if (!res.ok) return null;
          return { data: Buffer.from(await res.arrayBuffer()), mime: (res.headers.get('content-type') ?? '').split(';')[0].trim() };
        } catch {
          return null;
        }
      })();
      downloads.set(src, hit);
    }
    return hit;
  };
  const cached = (key: string, make: () => Promise<ImageData | null>) => {
    let hit = cache.get(key);
    if (!hit) {
      hit = make().catch(() => null);
      cache.set(key, hit);
    }
    return hit;
  };
  const image = (src: string, w: number, h: number) =>
    cached(`${src}|${Math.round(w)}x${Math.round(h)}`, async () => {
      const got = await download(src);
      if (!got) return null;
      const type = typeOf(got.mime, src);
      // O navegador desenha o SVG sem distorcer (preserveAspectRatio); o Word estica até a caixa.
      if (type === 'svg') return { svg: got.data, data: await svgToPng(got.data, w, h), type: 'png', aspect: svgAspect(got.data) };
      if (type) return { data: got.data, type };
      // WebP e afins: o Word não lê — vai como PNG desenhado pelo Chromium.
      return /^image\//.test(got.mime) ? redraw(got.data, got.mime, w, h, 'contain') : null;
    });
  /** Fundo de página raster: recorte "cover" como no Studio. PNG/JPEG que já
   *  têm a proporção da página vão como estão (sem reamostrar). */
  const pageImage = (src: string, w: number, h: number) =>
    cached(`page|${src}|${Math.round(w)}x${Math.round(h)}`, async () => {
      const got = await download(src);
      if (!got) return null;
      const type = typeOf(got.mime, src);
      if (type === 'svg') return image(src, w, h);
      // JPEG com EXIF pode estar girado (orientação): só o Chromium desenha certo.
      const rotated = type === 'jpg' && got.data.subarray(0, 65536).includes(Buffer.from('Exif\0\0'));
      if ((type === 'png' || type === 'jpg') && !rotated) {
        const size = rasterSize(got.data, type);
        if (size && Math.abs(size.w / size.h - w / h) <= 0.02 * (w / h)) return { data: got.data, type };
      }
      return /^image\//.test(got.mime) ? redraw(got.data, got.mime, w, h, 'cover') : null;
    });
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
      for (const e of target.querySelectorAll('*')) set(e, ${alone} ? 'hidden' : 'visible');
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
  return { image, raster, paint, backdrop, pageImage };
};
