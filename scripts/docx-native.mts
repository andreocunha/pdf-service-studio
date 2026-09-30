/**
 * npm run docx:native -- <documentId> <saida.docx>
 * Gera o Word nativo (padrão Lex) de um documento, com o app em RENDER_BASE_URL.
 */
import { writeFileSync } from 'node:fs';
import { closeBrowser } from '../src/browser.js';
import { config } from '../src/config.js';
import { renderNativeDocx } from '../src/docx-native/index.js';
import { serviceClient } from '../src/supabase.js';

const [id, out] = process.argv.slice(2);
const { data, error } = await serviceClient().from('documents').select('workspace_id').eq('id', id).single();
if (error) throw error;
const t0 = Date.now();
try {
  const { buffer: docx } = await renderNativeDocx({ documentId: id, workspaceId: data.workspace_id, appUrl: config.renderBaseUrl });
  writeFileSync(out, docx);
  console.log(`ok ${docx.length} bytes em ${Date.now() - t0} ms -> ${out}`);
} finally {
  await closeBrowser();
}
