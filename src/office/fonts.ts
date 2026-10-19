/**
 * Fontes do Office (Word e PowerPoint): as faces geradas pelo app em
 * `<app>/word-fonts/` (manifest + métricas + TTF), e a ponte "fonte do CSS
 * (família + peso + itálico, como o Chromium resolveu no layout) → face do
 * Office" (família legada + negrito/itálico), pelas mesmas fontes que o
 * editor usa: catálogo do app (/api/fonts) + manifest.
 */
import { brotliDecompressSync } from 'node:zlib';

import type { Run } from './layout.js';

export type Face = {
  family: string;
  bold: boolean;
  italic: boolean;
  lineHeight: number;
  file?: string;
  /** Peso e estilo do desenho (a família do Word já não diz: "CAIXA Std Bold" não é negrito). */
  weight?: number;
  style?: 'italic' | 'normal';
  /** Arquivo de public/fonts de onde a face saiu ("caixa-std/CAIXA_Std-Bold.woff2"). */
  source?: string;
  /** Instância de fonte variável: arquivo de origem e FontBBox (milésimos de em). */
  variable?: string;
  bbox?: number[];
};

export type FontData = {
  baseUrl: string;
  faces: Record<string, Face>;
  metrics: Map<Face, Record<string, number>>;
  instances: Map<string, Face[]>;
};

/** O manifest muda a cada deploy do app com fonte nova; relê de tempos em tempos. */
const MANIFEST_TTL_MS = 5 * 60_000;
const FETCH_TIMEOUT_MS = 8_000;

const fetchBytes = async (url: string): Promise<Buffer> => {
  const res = await fetch(url, { signal: AbortSignal.timeout(FETCH_TIMEOUT_MS) });
  if (!res.ok) throw new Error(`word-fonts: ${res.status} em ${url}`);
  return Buffer.from(await res.arrayBuffer());
};

/**
 * Os arquivos são brotli (.br). Se algum CDN no caminho servir com
 * Content-Encoding: br, o fetch já entrega descomprimido — aceita os dois.
 */
const unbrotli = (buf: Buffer): Buffer => {
  try {
    return brotliDecompressSync(buf);
  } catch {
    return buf;
  }
};

let fontData: { data: FontData; at: number } | null = null;
export const loadFonts = async (baseUrl: string): Promise<FontData> => {
  if (fontData && fontData.data.baseUrl === baseUrl && Date.now() - fontData.at < MANIFEST_TTL_MS) return fontData.data;
  try {
    const [manifest, metricsBr] = await Promise.all([
      fetchBytes(new URL('manifest.json', baseUrl).href),
      fetchBytes(new URL('metrics.json.br', baseUrl).href),
    ]);
    const faces = JSON.parse(manifest.toString('utf8')) as Record<string, Face>;
    const raw = JSON.parse(unbrotli(metricsBr).toString('utf8')) as Record<string, Record<string, number>>;
    const metrics = new Map<Face, Record<string, number>>();
    const instances = new Map<string, Face[]>();
    for (const [ps, face] of Object.entries(faces)) {
      if (raw[ps]) metrics.set(face, raw[ps]);
      if (face.variable) instances.set(face.variable, [...(instances.get(face.variable) ?? []), face]);
    }
    fontData = { data: { baseUrl, faces, metrics, instances }, at: Date.now() };
  } catch (err) {
    // App fora do ar por um instante: a última versão lida serve.
    if (fontData?.data.baseUrl !== baseUrl) throw err;
  }
  return fontData!.data;
};

/** Arquivos de fonte têm hash no nome — nunca mudam, podem ficar em memória. */
const fontFiles = new Map<string, Promise<Buffer>>();
/** ~100–300 KB por face: 60 cobrem os documentos do dia (≈15 MB no máximo). */
const MAX_CACHED_FONTS = 60;
export const loadFontFile = (url: string): Promise<Buffer> => {
  let file = fontFiles.get(url);
  if (!file) {
    if (fontFiles.size >= MAX_CACHED_FONTS) fontFiles.clear();
    file = fetchBytes(url).then(unbrotli);
    file.catch(() => fontFiles.delete(url));
    fontFiles.set(url, file);
  }
  return file;
};

/**
 * Espera a promessa, mas desiste no cancelamento (a exportação estourou o
 * prazo). O download segue pro cache compartilhado — outra exportação pode
 * estar esperando o mesmo arquivo.
 */
export const untilAborted = <T>(p: Promise<T>, signal?: AbortSignal): Promise<T> => {
  if (!signal) return p;
  signal.throwIfAborted();
  return new Promise<T>((resolve, reject) => {
    const onAbort = () => reject(signal.reason);
    signal.addEventListener('abort', onAbort, { once: true });
    p.then(resolve, reject).finally(() => signal.removeEventListener('abort', onAbort));
  });
};

type CatalogVariant = { file: string; weight: number; style: string; isVariable?: boolean };
type CatalogFamily = { name: string; folder: string; variants: CatalogVariant[] };

const WEIGHT_BY_NAME: Record<string, number> = {
  thin: 100, hairline: 100, extralight: 200, ultralight: 200, light: 300, book: 350, regular: 400,
  medium: 500, semibold: 600, demibold: 600, bold: 700, extrabold: 800, ultrabold: 800, black: 900, heavy: 900,
};

/** Peso de uma face do manifest (instância variável não guarda o número). */
const faceWeight = (face: Face): number => {
  if (face.weight) return face.weight;
  const last = face.family.replace(/ Italic$/, '').split(' ').at(-1)!.toLowerCase().replace(/[^a-z]/g, '');
  return WEIGHT_BY_NAME[last] ?? (face.bold ? 700 : 400);
};

export type FontResolver = {
  data: FontData;
  resolve: (cssFamily: string, weight: number, italic: boolean) => Face | null;
};

export const loadFontResolver = async (appUrl: string): Promise<FontResolver> => {
  const data = await loadFonts(`${appUrl}/word-fonts/`);
  const res = await fetch(`${appUrl}/api/fonts`, { signal: AbortSignal.timeout(8000) });
  const catalog = ((await res.json()) as { families: CatalogFamily[] }).families;
  const bySource = new Map<string, Face[]>();
  for (const face of Object.values(data.faces)) {
    if (!face.source) continue;
    const list = bySource.get(face.source) ?? [];
    if (!list.includes(face)) list.push(face);
    bySource.set(face.source, list);
  }
  const cache = new Map<string, Face | null>();
  const resolve = (cssFamily: string, weight: number, italic: boolean): Face | null => {
    const key = `${cssFamily}|${weight}|${italic}`;
    if (cache.has(key)) return cache.get(key)!;
    // getCssFontFamily desambigua fontes de sistema com " Custom" (Georgia Custom).
    const name = cssFamily.replace(/ Custom$/, '');
    const family = catalog.find((f) => f.name === name) ?? catalog.find((f) => f.name.replace(/\s/g, '') === name.replace(/\s/g, ''));
    let face: Face | null = null;
    if (family) {
      const style = italic ? 'italic' : 'normal';
      const variants = family.variants.filter((v) => v.style === style);
      const pool = variants.length ? variants : family.variants;
      const variant = pool.reduce((a, b) => (Math.abs(b.weight - weight) < Math.abs(a.weight - weight) ? b : a));
      const faces = (bySource.get(variant.file) ?? []).filter((f) => !f.style || (f.style === 'italic') === (variant.style === 'italic'));
      face = faces.length
        ? faces.reduce((a, b) => (Math.abs(faceWeight(b) - weight) < Math.abs(faceWeight(a) - weight) ? b : a))
        : null;
    }
    cache.set(key, face);
    return face;
  };
  return { data, resolve };
};


/** Lugar da face na família legada (regular, negrito, itálico, negrito itálico) — é assim que Word e PowerPoint embutem. */
export type FontSlot = 'regular' | 'bold' | 'italic' | 'boldItalic';
export const FONT_SLOTS: FontSlot[] = ['regular', 'bold', 'italic', 'boldItalic'];
export const fontSlot = (f: Face): FontSlot => (f.bold && f.italic ? 'boldItalic' : f.bold ? 'bold' : f.italic ? 'italic' : 'regular');

/** Faces usadas que têm arquivo, por família, uma por lugar, em ordem de lugar. */
export const embeddableFamilies = (used: Iterable<Face>): Map<string, Face[]> => {
  const byFamily = new Map<string, Face[]>();
  for (const face of used) {
    if (!face.file) continue;
    const list = byFamily.get(face.family) ?? [];
    if (!list.some((f) => fontSlot(f) === fontSlot(face))) list.push(face);
    byFamily.set(face.family, list);
  }
  for (const list of byFamily.values()) list.sort((a, b) => FONT_SLOTS.indexOf(fontSlot(a)) - FONT_SLOTS.indexOf(fontSlot(b)));
  return byFamily;
};

/**
 * Espaçamento entre letras (px por caractere) que faz o trecho ocupar no
 * Office a mesma largura do Studio (o Chromium arredonda o avanço de cada
 * letra; o Office não). Inclui o letter-spacing do CSS. `drawnSize`: o
 * tamanho que o Office desenha (o Word só aceita meio ponto). Sem métrica:
 * só o letter-spacing do CSS.
 */
export const fitLetterSpacing = (r: Run, face: Face | null, fonts: FontData, drawnSize: number, justified = false): number | undefined => {
  const css = r.letterSpacing || undefined;
  if (!face || !r.t) return css;
  const metric = fonts.metrics.get(face);
  if (!metric) return css;
  const text = r.caps ? r.t.toUpperCase() : r.t;
  let studio = 0;
  let drawn = 0;
  let count = 0;
  for (const ch of text) {
    const adv = metric[ch.codePointAt(0)!];
    if (adv === undefined) return css;
    studio += adv * r.size;
    drawn += adv * drawnSize;
    count++;
  }
  if (!count || !text.trim()) return undefined;
  // Largura de referência: a medida no Studio (pega arredondamento e kerning
  // do Chromium), somando o espaço onde a linha quebrou — ele não entra na
  // medida. Justificado com quebra estica os espaços: aí vale a largura
  // natural + letter-spacing.
  const space = (metric[32] ?? 0.25) * r.size;
  const lines = r.lines ?? 1;
  const target =
    r.width && !(justified && lines > 1) ? r.width + (lines - 1) * space : studio + (r.letterSpacing ?? 0) * count;
  const perChar = (target - drawn) / count;
  // Diferença grande demais não é arredondamento — é outra coisa.
  if (Math.abs(perChar) > r.size * 0.25) return css;
  return perChar;
};
