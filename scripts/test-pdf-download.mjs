import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PDFArray, PDFDict, PDFDocument, PDFName } from 'pdf-lib';

const root = fileURLToPath(new URL('../', import.meta.url));
const dir = await mkdtemp(join(tmpdir(), 'lex-pdf-download-'));
const fixturePath = join(dir, 'input.pdf');
const pdf = await PDFDocument.create();
const menu = pdf.addPage();
const section = pdf.addPage();
menu.drawText('Menu');
section.drawText('Section');
pdf.catalog.set(PDFName.of('Dests'), pdf.context.obj({ section: [section.ref, 'XYZ', 6, 700, 0] }));
menu.node.set(PDFName.of('Annots'), pdf.context.obj([{
  Type: 'Annot', Subtype: 'Link', Rect: [0, 0, 100, 30], Dest: 'section',
}]));
const input = Buffer.from(await pdf.save());
await writeFile(fixturePath, input);

async function check(failure) {
  const register = `import {register} from 'node:module'; register(${JSON.stringify(new URL('./test-pdf-download-loader.mjs', import.meta.url).href)});`;
  const child = spawn(process.execPath, ['--import', `data:text/javascript,${encodeURIComponent(register)}`, 'dist/server.js'], {
    cwd: root,
    env: {
      ...process.env, PORT: '0', RENDER_BASE_URL: 'http://127.0.0.1',
      PDF_SERVICE_SECRET: 'test', SUPABASE_URL: 'http://127.0.0.1',
      SUPABASE_PUBLISHABLE_KEY: 'test', SUPABASE_SECRET_KEY: 'test',
      PDF_TEST_FIXTURE: fixturePath, PDF_TEST_FAILURE: failure ? '1' : '0', LOG_LEVEL: 'info',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  let logs = '';
  const exited = new Promise(resolve => child.once('close', resolve));
  try {
    const address = await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(Error(`Server startup timed out: ${logs}`)), 15000);
      const receive = chunk => {
        logs += chunk.toString();
        const match = logs.match(/Server listening at http:\/\/[^:]+:(\d+)/);
        if (match) { clearTimeout(timeout); resolve(`http://127.0.0.1:${match[1]}`); }
      };
      child.stdout.on('data', receive);
      child.stderr.on('data', receive);
      child.once('error', err => { clearTimeout(timeout); reject(err); });
      child.once('exit', code => { clearTimeout(timeout); reject(Error(`Server exited ${code}: ${logs}`)); });
    });
    const response = await fetch(`${address}/pdf`, {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({documentId: 'download-test'}), signal: AbortSignal.timeout(15000),
    });
    assert.equal(response.status, 200);
    assert.equal(response.headers.get('content-type'), 'application/pdf');
    assert.match(response.headers.get('content-disposition'), /attachment/);
    const bytes = Buffer.from(await response.arrayBuffer());
    if (failure) {
      assert.deepEqual(bytes, input, 'normalization failure must deliver the original PDF byte for byte');
    } else {
      const output = await PDFDocument.load(bytes);
      const link = output.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray).lookup(0, PDFDict);
      const action = link.lookup(PDFName.of('A'), PDFDict);
      assert.equal(action.get(PDFName.of('S')).toString(), '/GoTo');
      const destination = action.lookup(PDFName.of('D'), PDFArray);
      assert.equal(destination.get(1).toString(), '/XYZ');
      assert.equal(destination.get(2).toString(), 'null');
      assert.equal(destination.get(3).asNumber(), 700);
      assert.equal(destination.get(4).toString(), 'null');
    }
    assert.equal((await fetch(`${address}/health`)).status, 200, 'server stays healthy');
    console.log(`PASS: compiled /pdf HTTP route ${failure ? 'returns original PDF after normalization failure' : 'delivers mobile-compatible menu links'}.`);
  } finally {
    child.kill('SIGTERM');
    await exited;
  }
  if (failure) assert.match(logs, /PDF link normalization failed/);
}
try {
  await check(false);
  await check(true);
} finally {
  await rm(dir, {recursive: true, force: true});
}
