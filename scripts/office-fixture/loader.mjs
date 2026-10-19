// Troca a página de render (/render-pdf, que lê o banco) por um HTML local
// (OFFICE_FIXTURE): o mesmo caminho do Word/PowerPoint nativo, sem banco.
const RENDER = `
import { readFileSync } from 'node:fs';
import { getBrowser } from './browser.js';
export const withRenderPage = async (args, fn) => {
  const browser = await getBrowser();
  const page = await browser.newPage();
  try {
    await page.setViewport({ width: 794, height: 1123, deviceScaleFactor: 1 });
    await page.emulateMediaType('screen');
    // Mesma origem do app: as fontes de localhost:3000 só carregam assim (CORS).
    await page.goto(process.env.RENDER_BASE_URL + '/word-fonts/manifest.json', { waitUntil: 'domcontentloaded' }).catch(() => undefined);
    await page.setContent(readFileSync(process.env.OFFICE_FIXTURE, 'utf8'), { waitUntil: 'networkidle0' });
    await page.evaluate(() => document.fonts.ready);
    const meta = await page.evaluate(() => window.__PDF_META);
    await page.setViewport({ width: meta.pageWidthPx, height: meta.pageHeightPx, deviceScaleFactor: 1 });
    return await fn({ page, meta, fonts: [] });
  } finally {
    await page.close().catch(() => undefined);
  }
};
export const renderDocumentPdf = async () => { throw Error('fixture: sem PDF'); };
`;
export async function load(url, context, nextLoad) {
  if (url.endsWith('/dist/render.js')) return { format: 'module', source: RENDER, shortCircuit: true };
  return nextLoad(url, context);
}
