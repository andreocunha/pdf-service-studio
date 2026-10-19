/**
 * npm run pptx:native -- <documentId> <saida.pptx>
 * Gera o PowerPoint nativo de um documento, com o app em RENDER_BASE_URL.
 */
import { writeFileSync } from 'node:fs';
import { closeBrowser } from '../src/browser.js';
import { config } from '../src/config.js';
import { renderNativePptx } from '../src/pptx-native/index.js';
import { serviceClient } from '../src/supabase.js';

const [id, out] = process.argv.slice(2);
const { data, error } = await serviceClient().from('documents').select('workspace_id').eq('id', id).single();
if (error) throw error;
const t0 = Date.now();
try {
  const { buffer } = await renderNativePptx({ documentId: id, workspaceId: data.workspace_id, appUrl: config.renderBaseUrl });
  writeFileSync(out, buffer);
  console.log(`ok ${buffer.length} bytes em ${Date.now() - t0} ms -> ${out}`);
} finally {
  await closeBrowser();
}
