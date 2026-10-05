/**
 * Leitura do layout do documento na página de render — o mesmo layout que
 * vira PDF — pra montar o Word nativo.
 */
import { readFileSync } from 'node:fs';

import { withRenderPage, type RenderMeta } from '../render.js';

export type Border = { w: number; color: string };
export type Box = { x: number; y: number; w: number; h: number };
export type Run = {
  t?: string;
  br?: boolean;
  font: string;
  weight: number;
  italic: boolean;
  size: number;
  color: string;
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
};
type Deco = {
  fill?: string;
  bgImage?: string;
  border?: Partial<Record<'top' | 'right' | 'bottom' | 'left', Border>>;
  pad?: [number, number, number, number];
  /** Maior raio de canto (px). */
  radius?: number;
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
    } & Deco)
  | { k: 'img'; id: number; box: Box; src: string; abs?: boolean }
  | { k: 'raster'; id: number; box: Box; abs?: boolean }
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
      number: { text: string; box: Box; style: Omit<Run, 't' | 'br' | 'link'>; align: string } | null;
    } | null;
    nav: { tree: LayoutNode; sig: string } | null;
  }[];
  blocks: LayoutBlock[];
};

const EXTRACT = readFileSync(new URL('./extract.js', import.meta.url), 'utf8');

export const readLayout = (args: { documentId: string; workspaceId: string }) =>
  withRenderPage(args, async ({ page }) => {
    const layout = (await page.evaluate(EXTRACT)) as DocumentLayout;
    return { layout, page };
  });
