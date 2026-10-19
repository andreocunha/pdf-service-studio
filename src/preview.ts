import { withRenderPage } from './render.js';
import { logger } from './logger.js';

// ---------------------------------------------------------------------------
// Prévia das páginas — o "olho" das IAs conectadas pelo MCP do Studio.
//
// Mesmo render do PDF (withRenderPage), mas em vez de imprimir devolve uma
// foto JPEG de cada página pedida e um raio-x de TODAS: quais blocos caíram
// nela, quanto espaço sobrou embaixo, se estourou e qual bloco fecha a
// página. É com isso que a IA acerta a paginação (blocos de texto vazios),
// acha título sozinho no pé da página e confere o visual.
// ---------------------------------------------------------------------------

export type PageAudit = {
  index: number;
  blockIds: string[];
  lastBlock: { id: string; type: string } | null;
  /** Espaço livre entre o último bloco e a margem inferior (px, 96 dpi). */
  emptyBottomPx: number;
  /** Conteúdo passou da margem inferior. */
  overflow: boolean;
  /** Onde o conteúdo termina, medido do topo da página (px). Bloco
   *  `fullWidth` no fim da página pode ir até a borda (passa da margem) sem
   *  estar estourado — quem chama compara isto com a altura da página. */
  contentBottomPx: number;
};

export type PreviewResult = {
  title: string;
  pageWidthPx: number;
  pageHeightPx: number;
  pages: PageAudit[];
  images: { index: number; jpegBase64: string }[];
};

const MAX_IMAGES = 12;

export const renderDocumentPreview = async (args: {
  documentId: string;
  workspaceId: string;
  /** Índices (0-based) das páginas a fotografar. Vazio = as primeiras. */
  pages: number[];
  /** Escala da imagem: 1 = 96 dpi. */
  scale: number;
}): Promise<PreviewResult> => {
  const start = Date.now();
  return withRenderPage(args, async ({ page, meta }) => {
    const pages = await page.evaluate((paddingBottom: number) => {
      const pageEls = Array.from(document.querySelectorAll<HTMLElement>('[data-page-index]'));
      // Documento contínuo não tem páginas: o corpo inteiro vale como uma.
      const roots: HTMLElement[] = pageEls.length ? pageEls : [document.body];
      return roots.map((root, index) => {
        const box = root.getBoundingClientRect();
        const blocks = Array.from(root.querySelectorAll<HTMLElement>('[data-block-id]'))
          .filter((el) => !el.parentElement?.closest('[data-block-id]'));
        const last = blocks.at(-1);
        // Bloco fullWidth no pé da página sangra a margem de baixo de propósito
        // (margin-bottom negativo, ver Block.tsx): a capa de
        // página inteira não "estourou". Só ele ganha essa folga; cada bloco é
        // medido contra o próprio limite.
        const baseLimit = box.top + box.height - paddingBottom;
        const lastCs = last ? getComputedStyle(last) : null;
        // fullWidth = ocupa a largura inteira da página (margem lateral negativa
        // sozinha não basta: blocos comuns têm -ml-12 no editor).
        const lastBleeds = !!lastCs && !!last && parseFloat(lastCs.marginBottom) < 0
          && last.getBoundingClientRect().width >= box.width - 2;
        const bleed = lastBleeds ? Math.min(paddingBottom, -parseFloat(lastCs!.marginBottom)) : 0;
        let bottom = box.top;
        let overflow = false;
        for (const el of blocks) {
          const elBottom = el.getBoundingClientRect().bottom;
          bottom = Math.max(bottom, elBottom);
          if (elBottom > baseLimit + (el === last ? bleed : 0) + 1) overflow = true;
        }
        const limit = baseLimit + (lastBleeds && last && last.getBoundingClientRect().bottom >= bottom - 1 ? bleed : 0);
        return {
          index,
          blockIds: blocks.map((el) => el.dataset.blockId ?? ''),
          lastBlock: last ? { id: last.dataset.blockId ?? '', type: last.dataset.blockType ?? '' } : null,
          emptyBottomPx: blocks.length ? Math.max(0, Math.round(limit - bottom)) : Math.round(box.height),
          overflow,
          contentBottomPx: Math.round(bottom - box.top),
        };
      });
    }, meta.paddingBottomPx ?? 0);

    const wanted = (args.pages.length ? args.pages : pages.map((p) => p.index))
      .filter((i) => i >= 0 && i < pages.length)
      .slice(0, MAX_IMAGES);

    const handles = await page.$$('[data-page-index]');
    const images: PreviewResult['images'] = [];
    for (const index of wanted) {
      const box = handles.length ? await handles[index]?.boundingBox() : { x: 0, y: 0, width: meta.pageWidthPx, height: meta.pageHeightPx };
      if (!box) continue;
      const shot = await page.screenshot({
        type: 'jpeg',
        quality: 72,
        captureBeyondViewport: true,
        clip: { x: box.x, y: box.y, width: box.width, height: box.height, scale: args.scale },
      });
      images.push({ index, jpegBase64: Buffer.from(shot).toString('base64') });
    }

    logger.info(
      { documentId: args.documentId, elapsedMs: Date.now() - start, pages: pages.length, images: images.length },
      'Preview rendered',
    );
    return {
      title: meta.title ?? 'document',
      pageWidthPx: meta.pageWidthPx,
      pageHeightPx: meta.pageHeightPx,
      pages,
      images,
    };
  });
};
