/**
 * node scripts/office-fixture/run.mjs <saida-sem-extensão> [fixture.html]
 * Gera Word e PowerPoint nativos (dist/) a partir de um HTML local com a
 * estrutura do /render-pdf — sem banco. Precisa do app (fontes) em RENDER_BASE_URL.
 */
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const [out, fixture = fileURLToPath(new URL('./sample.html', import.meta.url))] = process.argv.slice(2);
const root = fileURLToPath(new URL('../../', import.meta.url));
const register = `import {register} from 'node:module'; register(${JSON.stringify(new URL('./loader.mjs', import.meta.url).href)});`;
const main = `
import { writeFileSync, existsSync } from 'node:fs';
const { config } = await import('./dist/config.js');
const { closeBrowser } = await import('./dist/browser.js');
const args = { documentId: 'fixture', workspaceId: 'fixture', appUrl: config.renderBaseUrl };
try {
  const { renderNativeDocx } = await import('./dist/docx-native/index.js');
  writeFileSync(process.env.OUT + '.docx', (await renderNativeDocx(args)).buffer);
  if (existsSync('./dist/pptx-native/index.js')) {
    const { renderNativePptx } = await import('./dist/pptx-native/index.js');
    writeFileSync(process.env.OUT + '.pptx', (await renderNativePptx(args)).buffer);
  }
} finally { await closeBrowser(); }
`;
const r = spawnSync(process.execPath, ['--env-file=.env', '--import', `data:text/javascript,${encodeURIComponent(register)}`, '--input-type=module', '-e', main], {
  cwd: process.env.OFFICE_ROOT ?? root,
  env: { ...process.env, OFFICE_FIXTURE: fixture, OUT: out },
  stdio: 'inherit',
});
process.exit(r.status ?? 1);
