/**
 * Fonte do CSS (família + peso + itálico, como o Chromium resolveu no layout)
 * → face do Word (família legada + negrito/itálico), pelas mesmas fontes que o
 * editor usa: catálogo do app (/api/fonts) + manifest das fontes do Word.
 */
import { loadFonts, type Face, type FontData } from '../docx-fonts.js';

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
