// Isolate upstream services while exercising the compiled HTTP route and finalizer.
import { readFile } from 'node:fs/promises';
export async function load(url, context, nextLoad) {
  const stubs = {
    'auth.js': 'export const authorize = async () => ({documentId:"download-test", workspaceId:"test"});',
    'browser.js': 'export const closeBrowser = async () => {}; export const warmUpBrowser = async () => {};',
    'render.js': `import {readFileSync} from 'node:fs'; export const renderDocumentPdf = async () => ({buffer:readFileSync(process.env.PDF_TEST_FIXTURE), title:'Menu test', fonts:[]});`,
    'ilovepdf.js': 'export const startIlovepdfTask = async () => ({}); export const compressPdf = async (pdf) => pdf; export const convertPdfToOffice = async () => {throw Error("unexpected Office export")};',
    'supabase.js': 'export const serviceClient = () => {throw Error("unexpected database access")};',
  };
  for (const [file, source] of Object.entries(stubs)) {
    if (url.endsWith(`/dist/${file}`)) return { format: 'module', source, shortCircuit: true };
  }
  if (url.endsWith('/dist/pdf-links.js') && process.env.PDF_TEST_FAILURE === '1') {
    // Keep the real parser/normalizer: fail after it has processed a valid PDF.
    const original = await readFile(new URL(url), 'utf8');
    const source = original.replace('export async function normalizePdfLinks(', 'async function originalNormalize(')
      + '\nexport async function normalizePdfLinks(pdf) { await originalNormalize(pdf); throw Error("simulated PDF serialization failure"); }';
    return { format: 'module', source, shortCircuit: true };
  }
  return nextLoad(url, context);
}
