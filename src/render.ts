import type { Page } from 'puppeteer-core';

import { getBrowser } from './browser.js';
import { config } from './config.js';
import { renderFailed } from './errors.js';
import { logger } from './logger.js';
import { signRenderToken } from './signing.js';

type RenderMeta = {
  pageWidthPx: number;
  pageHeightPx: number;
  paddingTopPx?: number;
  paddingRightPx?: number;
  paddingBottomPx?: number;
  paddingLeftPx?: number;
  title?: string;
};

const PX_TO_IN = 1 / 96;

/**
 * O render page expõe `window.__PDF_READY = true` quando o documento
 * terminou de montar e as fontes carregaram. Expõe também `window.__PDF_META`
 * com as dimensões reais da página do documento.
 */
export type UsedFont = {
  family: string;
  category: 'sans' | 'serif' | 'mono' | null;
};

const waitForReady = async (
  page: Page,
  timeoutMs: number,
): Promise<{ meta: RenderMeta; fonts: UsedFont[] }> => {
  await page.waitForFunction(
    () =>
      (window as unknown as { __PDF_READY?: boolean }).__PDF_READY === true &&
      typeof (window as unknown as { __PDF_META?: unknown }).__PDF_META === 'object',
    { timeout: timeoutMs, polling: 100 },
  );
  // Belt-and-suspenders: esperar todas fontes resolverem.
  await page.evaluate(async () => {
    if ('fonts' in document) {
      await (document as Document & { fonts: FontFaceSet }).fonts.ready;
    }
  });
  const { meta, fonts } = await page.evaluate(() => {
    const w = window as unknown as { __PDF_META: RenderMeta; __PDF_FONTS?: UsedFont[] };
    return { meta: w.__PDF_META, fonts: Array.isArray(w.__PDF_FONTS) ? w.__PDF_FONTS : [] };
  });
  return { meta, fonts };
};

export const renderDocumentPdf = async (args: {
  documentId: string;
  workspaceId: string;
}): Promise<{ buffer: Buffer; title: string; fonts: UsedFont[] }> => {
  const start = Date.now();
  const token = signRenderToken(args.documentId, args.workspaceId);
  const url = `${config.renderBaseUrl}/render-pdf/${encodeURIComponent(args.documentId)}?t=${encodeURIComponent(token)}`;

  const browser = await getBrowser();
  const page = await browser.newPage();

  try {
    // Começa com viewport A4 — será ajustada pra largura real da página
    // após o render page expor __PDF_META.
    await page.setViewport({ width: 794, height: 1123, deviceScaleFactor: 1 });

    // Usa CSS `screen` em vez de `print` (default do page.pdf). O editor
    // é pensado pra tela, e a gente controla page breaks via CSS no pdf-render-client.
    await page.emulateMediaType('screen');

    // Bloqueia recursos não-essenciais pra reduzir tempo/tamanho.
    await page.setRequestInterception(true);
    page.on('request', (req) => {
      const type = req.resourceType();
      if (type === 'media' || type === 'websocket' || type === 'eventsource') {
        req.abort().catch(() => undefined);
        return;
      }
      req.continue().catch(() => undefined);
    });

    page.on('pageerror', (err: unknown) => {
      const message = err instanceof Error ? err.message : String(err);
      logger.warn({ err: message, documentId: args.documentId }, 'Render page error');
    });
    page.on('console', (msg) => {
      const text = msg.text();
      if (msg.type() === 'error') {
        logger.warn({ text, documentId: args.documentId }, 'Render console error');
        return;
      }
      // Encaminha logs do client com prefixo [pdf-render] — úteis pra debug
      // em produção sem precisar abrir devtools remoto.
      if (text.startsWith('[pdf-render]')) {
        logger.info({ text, documentId: args.documentId }, 'Render client log');
      }
    });

    const response = await page.goto(url, {
      waitUntil: 'domcontentloaded',
      timeout: config.requestTimeoutMs,
    });
    if (!response || !response.ok()) {
      const status = response?.status() ?? 0;
      throw renderFailed(`Render page responded with status ${status}`);
    }

    const { meta, fonts } = await waitForReady(page, config.requestTimeoutMs);

    // Ajusta a viewport pra bater exatamente com a largura da página.
    // Isso garante que `mx-auto` não cause offset horizontal e que cada
    // página do doc ocupe exatamente 100% do width do PDF.
    await page.setViewport({
      width: meta.pageWidthPx,
      height: meta.pageHeightPx,
      deviceScaleFactor: 1,
    });
    // O resize acima muda o tamanho de LAYOUT das imagens, o que invalida o
    // raster que o Chromium já tinha decodificado. Capturar direto depois disso
    // é o que fazia imagem grande sair cortada no PDF (às vezes — depende de o
    // decode ganhar ou perder a corrida). __PDF_SETTLE re-decodifica tudo e
    // espera dois frames; o fallback de um rAF cobre páginas antigas em cache.
    await page
      .evaluate(async () => {
        const settle = (window as unknown as { __PDF_SETTLE?: () => Promise<void> })
          .__PDF_SETTLE;
        if (settle) return settle();
        return new Promise((r) => requestAnimationFrame(() => r(null)));
      })
      .catch((err: unknown) => {
        logger.warn({ err, documentId: args.documentId }, 'PDF settle failed — capturing anyway');
      });

    const widthIn = meta.pageWidthPx * PX_TO_IN;
    const heightIn = meta.pageHeightPx * PX_TO_IN;

    const pdf = await page.pdf({
      width: `${widthIn}in`,
      height: `${heightIn}in`,
      printBackground: true,
      preferCSSPageSize: false,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
      displayHeaderFooter: false,
      tagged: true,
    });

    const elapsed = Date.now() - start;
    logger.info(
      { documentId: args.documentId, elapsedMs: elapsed, pdfBytes: pdf.length, fonts },
      'PDF rendered',
    );

    return {
      buffer: Buffer.from(pdf),
      title: meta.title ?? 'document',
      fonts,
    };
  } finally {
    await page.close().catch(() => undefined);
  }
};
