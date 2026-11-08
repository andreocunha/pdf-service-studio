/**
 * npm run html:export -- <documentId> <saida.html>
 * Gera o HTML autocontido de um documento (o mesmo do /html), com o app em RENDER_BASE_URL.
 */
import { writeFileSync } from 'node:fs';
import { closeBrowser } from '../src/browser.js';
import { renderDocumentHtml } from '../src/render.js';
import { serviceClient } from '../src/supabase.js';

const [id, out] = process.argv.slice(2);
const { data, error } = await serviceClient().from('documents').select('workspace_id').eq('id', id).single();
if (error) throw error;
const t0 = Date.now();
try {
  const { html } = await renderDocumentHtml({ documentId: id, workspaceId: data.workspace_id });
  writeFileSync(out, html);
  console.log(`ok ${Buffer.byteLength(html)} bytes em ${Date.now() - t0} ms -> ${out}`);
} finally {
  await closeBrowser();
}
