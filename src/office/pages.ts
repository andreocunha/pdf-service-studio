/**
 * O que vale por página, igual no Word e no PowerPoint: quais são capa, a
 * imagem de fundo de cada uma e a numeração do rodapé.
 */
import type { Page } from 'puppeteer-core';

import type { Assets, ImageData } from './assets.js';
import type { DocumentLayout, FooterNumberFormat, LayoutNode } from './layout.js';

/**
 * Capa (e contracapa): vai inteira como imagem no cabeçalho, não editável,
 * como no Studio. É capa a página cujo único bloco ocupa a página toda, ou a
 * primeira/última página com fundo próprio (diferente da vizinha) — montada
 * com vários blocos sobre a arte, encostados no pé da página.
 */
export const coverPagesOf = (layout: DocumentLayout): Set<number> => {
  const covers = new Set<number>();
  const { pages } = layout;
  const art = (t: LayoutNode | null): number => {
    if (!t) return 0;
    const own = t.k === 'img' || t.k === 'raster' || (t.k !== 'text' && (t.bgImage || t.paint)) ? t.box.w * t.box.h : 0;
    return Math.max(own, ...(t.k === 'box' ? t.kids.map(art) : [0]));
  };
  const pageArea = layout.meta.pageWidthPx * layout.meta.pageHeightPx;
  for (const p of pages) {
    const blocks = layout.blocks.filter((b) => b.page === p.index && b.tree);
    // Bloco do tamanho da página só é capa se tiver arte grande (imagem de
    // fundo/ilustração): um Sumário feito de textos e linhas vira tabela.
    if (blocks.length === 1 && blocks[0].box.h >= layout.meta.pageHeightPx * 0.9 && art(blocks[0].tree) >= 0.3 * pageArea) covers.add(p.index);
  }
  const ownBackground = (i: number, neighbour: number) =>
    pages[i]?.bgImage && pages[neighbour] && pages[i].bgImage !== pages[neighbour].bgImage;
  // Página com fundo próprio só vira "foto" se o conteúdo tiver arte (imagem
  // grande): caixas coloridas e texto (Quadro Resumo na capa) vão como tabela
  // por cima do fundo, acompanhando o texto.
  const hasArt = (i: number) => layout.blocks.filter((b) => b.page === i).some((b) => art(b.tree) >= 0.15 * pageArea);
  if (pages.length > 1 && ownBackground(0, 1) && hasArt(0)) covers.add(0);
  // Contracapa só com pouco texto: a última página do anexo, com fundo
  // próprio mas conteúdo corrido, não é capa (ia inteira pro cabeçalho).
  const chars = (i: number) => {
    let n = 0;
    const walk = (t: LayoutNode | null): void => {
      if (!t) return;
      if (t.k === 'text') n += t.paras.reduce((a, p) => a + p.runs.reduce((b, r) => b + (r.t?.length ?? 0), 0), 0);
      else if (t.k === 'box') t.kids.forEach(walk);
    };
    layout.blocks.filter((b) => b.page === i).forEach((b) => walk(b.tree));
    return n;
  };
  const last = pages.length - 1;
  if (pages.length > 2 && ownBackground(last, last - 1) && chars(last) <= 150 && hasArt(last)) covers.add(last);
  return covers;
};

/**
 * Fundo de cada página. Capa: a arte da página inteira, capturada sem os
 * textos (eles vão editáveis por cima — os campos "CONTRATANTE: Dados"
 * precisam ser preenchidos). Demais páginas: a imagem de fundo do Studio,
 * desenhada pelo Chromium (o Office não desenha certo todo SVG e fundo não se
 * edita).
 */
export const pageBackgrounds = (page: Page, layout: DocumentLayout, covers: Set<number>, assets: Assets) =>
  async (index: number): Promise<ImageData | null> => {
    const p = layout.pages[index];
    if (!p) return null;
    if (covers.has(index)) {
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
    const img = await assets.pageImage(p.bgImage, p.box.w, p.box.h);
    return img?.svg ? { data: img.data, type: img.type } : img;
  };

/** Diferença entre o número do Studio e a posição da página ("02/28" na 3ª página: -1). */
export type PageOffsets = { current: number; total: number; format: FooterNumberFormat };

/** Numeração do rodapé da página: "01/12", "1" ou "Página 1 de 12" — o 1º número é a página, o 2º o total. */
export const pageNumbering = (layout: DocumentLayout, pageIndex: number): PageOffsets | null => {
  const num = layout.pages[pageIndex]?.footer?.number;
  if (!num) return null;
  const [current = pageIndex + 1, total = layout.pages.length] = (num.text.match(/\d+/g) ?? []).map(Number);
  return { current: current - (pageIndex + 1), total: total - layout.pages.length, format: num.format ?? 'page-total' };
};
