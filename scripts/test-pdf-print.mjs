import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { setTimeout as delay } from 'node:timers/promises';
import { PDFDocument } from 'pdf-lib';

const expectTimeout = process.argv.includes('--expect-timeout');
const fixture = createServer((_req, res) => {
  res.setHeader('Content-Type', 'text/html');
  res.end(`<html><body><h1>Print timeout regression</h1><script>
    window.__PDF_META = {pageWidthPx:794,pageHeightPx:1123,title:'Timeout test'};
    window.__PDF_FONTS = []; window.__PDF_READY = true;
  </script></body></html>`);
});
await new Promise(resolve => fixture.listen(0, '127.0.0.1', resolve));
process.env.RENDER_BASE_URL = `http://127.0.0.1:${fixture.address().port}`;
process.env.PDF_SERVICE_SECRET = 'test';
process.env.SUPABASE_URL = 'http://127.0.0.1';
process.env.SUPABASE_PUBLISHABLE_KEY = 'test';
process.env.SUPABASE_SECRET_KEY = 'test';
process.env.PDF_PRINT_TIMEOUT_MS = expectTimeout ? '1000' : '120000';
if (process.platform === 'darwin') process.env.CHROMIUM_EXECUTABLE_PATH ||= '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const { renderDocumentPdf } = await import('../dist/render.js');
const { getBrowser, closeBrowser } = await import('../dist/browser.js');
let printedPage;
let printRequested = false;
try {
  const browser = await getBrowser();
  const newPage = browser.newPage.bind(browser);
  browser.newPage = async () => {
    const page = await newPage();
    printedPage = page;
    // Delay the real Chromium response, not page.pdf itself: Puppeteer's real
    // timeout race must stay active, and renderDocumentPdf supplies the limit.
    const client = page._client();
    const send = client.send.bind(client);
    client.send = async (method, ...args) => {
      if (method === 'Page.printToPDF') printRequested = true;
      const result = await send(method, ...args);
      if (method === 'Page.printToPDF') {
        await delay(expectTimeout ? 2000 : 31000);
      }
      return result;
    };
    return page;
  };
  const start = Date.now();
  const render = () => renderDocumentPdf({documentId:'print-timeout-test',workspaceId:'test'});
  if (expectTimeout) {
    await assert.rejects(render(), err => err.statusCode === 504 && err.code === 'pdf_print_timeout');
  } else {
    const {buffer} = await render();
    assert.equal((await PDFDocument.load(buffer)).getPageCount(), 1);
    assert.ok(Date.now() - start > 30000, 'must exceed the previous 30s default');
  }
  assert.ok(printRequested, 'exercise the actual print command');
  assert.ok(printedPage.isClosed(), 'render page must close on both success and timeout');
  console.log(`PASS: ${expectTimeout ? 'configured print timeout is bounded, returns 504 and closes the page' : 'compiled render finishes after more than 30s with real Chromium/Puppeteer'}.`);
} finally {
  await closeBrowser();
  fixture.closeAllConnections();
  await new Promise(resolve => fixture.close(resolve));
}
