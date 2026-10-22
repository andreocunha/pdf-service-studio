/**
 * Layout do documento lido na página de render (extract.js) — o mesmo layout
 * que vira PDF. É a base do Word e do PowerPoint nativos.
 */
import { readFileSync } from 'node:fs';

import type { Page } from 'puppeteer-core';

import type { RenderMeta } from '../render.js';

/** Formato do número do rodapé do Studio (`data-document-footer-format`). */
export type FooterNumberFormat = 'page-total' | 'page' | 'page-of-total';

/** Cor com transparência (alfa de 0 a 1). */
export type Tint = { color: string; alpha: number };
export type Border = {
  w: number;
  /** Cor já misturada com o que está atrás (o Word não tem transparência). */
  color: string;
  /** A cor transparente de verdade (PowerPoint). */
  raw?: Tint;
  style?: 'dashed' | 'dotted';
};
/** box-shadow por fora. */
export type Shadow = { x: number; y: number; blur: number; spread: number; color: string; alpha: number };
export type Box = { x: number; y: number; w: number; h: number };
export type Run = {
  t?: string;
  br?: boolean;
  font: string;
  weight: number;
  italic: boolean;
  size: number;
  color: string;
  /** Transparência do texto (PowerPoint). */
  alpha?: number;
  caps: boolean;
  lower: boolean;
  letterSpacing: number;
  underline: boolean;
  strike: boolean;
  highlight: string | null;
  verticalAlign: 'super' | 'sub' | null;
  link: string | null;
  /** Largura real do trecho no Studio (px), somando as linhas. */
  width?: number;
  /** Em quantas linhas o trecho ficou no Studio. */
  lines?: number;
  /** Espaços em sequência viram um só (white-space normal). */
  collapse?: boolean;
  /** Onde começa cada linha nova no Studio (índice no texto). */
  breaks?: number[];
  /** text-shadow (px, já na escala). */
  shadow?: { x: number; y: number; blur: number; color: string; alpha: number };
};
export type Para = {
  align: 'left' | 'center' | 'right' | 'both';
  lineHeight: number | null;
  fontSize: number;
  runs: Run[];
  list?: { kind: string; level: number };
  /** Item de lista: onde o texto começa (px, página) e o vão do marcador antes dele. */
  x?: number;
  indent?: number;
  /** Meio da primeira e da última linha (px, página). */
  first?: number;
  last?: number;
  /** Altura da área do texto na primeira linha (ascendente + descendente, px). */
  area?: number;
  /** Marcador do item de lista (::marker). */
  marker?: { color: string; alpha?: number; size?: number };
};
type Deco = {
  fill?: string;
  /** Preenchimento transparente de verdade (PowerPoint); `fill` é a mistura. */
  fillRaw?: Tint;
  shadow?: Shadow[];
  /** Existe só pelo `fillRaw` (vidro sobre imagem): o Word não vê. */
  glass?: boolean;
  bgImage?: string;
  border?: Partial<Record<'top' | 'right' | 'bottom' | 'left', Border>>;
  pad?: [number, number, number, number];
  /** Maior raio de canto (px). */
  radius?: number;
  /** Cantos com raios diferentes. */
  uneven?: boolean;
  /** Pintura que o Word não tem (gradiente, sombra, imagem): vai capturada, atrás do texto. */
  paint?: boolean;
  abs?: boolean;
};
export type LayoutNode =
  | ({
      k: 'text';
      id: number;
      box: Box;
      paras: Para[];
      /** Onde as linhas terminam de verdade (px, página). */
      inkRight?: number | null;
      /** Texto centralizado na vertical dentro da própria caixa. */
      middle?: boolean;
      autonumber: string | null;
      autonumberStyle?: (Omit<Run, 't' | 'br' | 'link'> & { lineHeight: number | null }) | null;
      /** Estilo do número no Studio (`a`, `A`, `i`, `I`, `1`); null = caminho 1.2.3. */
      autonumberFormat?: string | null;
      /** O número como aparece ("c", "IV"). */
      autonumberText?: string | null;
      /** Onde o número aparece (px, página). */
      autonumberBox?: Box | null;
      /** Vão entre o número e o texto que vem depois dele na mesma linha (px). */
      autonumberGap?: number | null;
      /** Fundo do próprio número (selo), quando opaco. */
      autonumberFill?: { color: string; box: Box } | null;
    } & Deco)
  | { k: 'img'; id: number; box: Box; src: string; abs?: boolean; radius?: number; clip?: Box }
  | { k: 'raster'; id: number; box: Box; abs?: boolean; clip?: Box }
  | ({ k: 'shape'; id: number; box: Box } & Deco)
  | ({
      k: 'box';
      id: number;
      box: Box;
      layout: { display: string; dir: string; alignItems: string; justify: string; textAlign: string };
      kids: LayoutNode[];
    } & Deco);
export type LayoutBlock = { blockId: string; type: string; page: number; box: Box; tree: LayoutNode | null };
export type DocumentLayout = {
  meta: RenderMeta & { paddingTopPx?: number };
  pages: {
    index: number;
    box: Box;
    bgImage: string | null;
    bgColor: string | null;
    footer: {
      text: Extract<LayoutNode, { k: 'text' }> | null;
      number: { text: string; format?: FooterNumberFormat; box: Box; style: Omit<Run, 't' | 'br' | 'link'>; align: string } | null;
    } | null;
    nav: { tree: LayoutNode; sig: string } | null;
  }[];
  blocks: LayoutBlock[];
};

/**
 * Texto de uma linha só no Studio: pelas linhas medidas de cada trecho (a
 * caixa pode ser mais alta que o texto — rótulo centralizado numa linha de
 * 41px); sem medida, pela altura.
 */
export const isOneLine = (n: Extract<LayoutNode, { k: 'text' }>): boolean => {
  if (n.paras.length !== 1) return false;
  if (n.paras[0].runs.some((r) => r.br)) return false;
  const runs = n.paras[0].runs.filter((r) => r.t);
  if (runs.length && runs.every((r) => r.lines !== undefined)) return runs.every((r) => (r.lines ?? 1) <= 1);
  return n.box.h < (n.paras[0].lineHeight ?? n.paras[0].fontSize * 1.2) * 1.5;
};

const EXTRACT = readFileSync(new URL('./extract.js', import.meta.url), 'utf8');

/** Lê o layout da página de render já pronta (depois do __PDF_READY). */
export const readLayout = async (page: Page, meta: RenderMeta): Promise<DocumentLayout> => {
  const layout = (await page.evaluate(EXTRACT)) as DocumentLayout;
  layout.meta = { ...layout.meta, ...meta };
  return layout;
};

/**
 * Tira as caixas que só existem pela transparência (`glass`): o Word não tem
 * preenchimento transparente e montaria a tabela diferente por causa delas.
 */
export const withoutGlass = (layout: DocumentLayout): DocumentLayout => {
  const strip = (n: LayoutNode): LayoutNode | null => {
    if (n.k === 'shape') return n.glass ? null : n;
    if (n.k !== 'box') return n;
    const kids = n.kids.map(strip).filter((k): k is LayoutNode => k !== null);
    if (n.glass && kids.length === 1 && kids[0].k === 'text') return { ...kids[0], abs: n.abs };
    // Ficou vazia: a regra do extrator pra caixa sem filhos (forma, se pinta algo; senão some).
    if (!kids.length && n.kids.length) {
      if (!n.fill && !n.border && !n.bgImage) return null;
      const { kids: _kids, layout: _layout, ...deco } = n;
      return { ...deco, k: 'shape' };
    }
    return { ...n, kids };
  };
  return {
    ...layout,
    pages: layout.pages.map((p) => (p.nav ? { ...p, nav: { ...p.nav, tree: strip(p.nav.tree) ?? p.nav.tree } } : p)),
    blocks: layout.blocks.map((b) => (b.tree ? { ...b, tree: strip(b.tree) } : b)),
  };
};
