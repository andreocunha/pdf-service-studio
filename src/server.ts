import cors from '@fastify/cors';
import rateLimit from '@fastify/rate-limit';
import Fastify from 'fastify';

import { authorize } from './auth.js';
import { closeBrowser, warmUpBrowser } from './browser.js';
import { config } from './config.js';
import { compressPdf, convertPdfToOffice, startIlovepdfTask } from './ilovepdf.js';
import type { IlovepdfTask, OfficeFormat } from './ilovepdf.js';
import { HttpError } from './errors.js';
import { logger } from './logger.js';
import { renderDocumentPdf } from './render.js';
import { serviceClient } from './supabase.js';

const app = Fastify({
  loggerInstance: logger,
  bodyLimit: 1 * 1024 * 1024,
  trustProxy: true,
});

await app.register(cors, {
  origin: true,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'x-api-key'],
  exposedHeaders: ['Content-Disposition'],
});

await app.register(rateLimit, {
  max: 30,
  timeWindow: '1 minute',
  keyGenerator: (req) => {
    const apiKey = req.headers['x-api-key'];
    if (typeof apiKey === 'string' && apiKey.length > 0) return `key:${apiKey}`;
    const auth = req.headers['authorization'];
    if (typeof auth === 'string' && auth.length > 0) return `jwt:${auth.slice(-24)}`;
    return `ip:${req.ip}`;
  },
});

app.setErrorHandler((err, req, reply) => {
  if (err instanceof HttpError) {
    reply.code(err.statusCode).send({ error: err.code, message: err.message });
    return;
  }
  req.log.error({ err }, 'Unhandled error');
  reply.code(500).send({ error: 'internal', message: 'Internal server error' });
});

app.get('/health', async () => ({ ok: true, version: '0.1.0' }));

type PdfBody = { documentId?: unknown };

// Sanitiza só caracteres reservados de filesystem (preserva acentos, parênteses, etc).
const cleanTitle = (raw: string | undefined): string => {
  return (
    (raw || 'documento')
      // eslint-disable-next-line no-control-regex
      .replace(/[/\\:*?"<>|\x00-\x1F]/g, '')
      .replace(/\s+/g, ' ')
      .trim() || 'documento'
  );
};

const dispositionFor = (clean: string, ext: 'pdf' | 'docx' | 'pptx'): string => {
  // ASCII fallback pra clientes que não entendem RFC 5987 (é raro hoje, mas seguro).
  const asciiFallback = clean.replace(/[^\x20-\x7E]/g, '_');
  const encodedUtf8 = encodeURIComponent(clean);
  return `attachment; filename="${asciiFallback}.${ext}"; filename*=UTF-8''${encodedUtf8}.${ext}`;
};

/**
 * Comprime o PDF, mas NUNCA piora e NUNCA quebra o download: se o ilovepdf
 * falhar, expirar, ou devolver um arquivo maior/vazio, entrega o original.
 * (Compressão que aumenta o arquivo é real — acontece em PDF só-texto, onde
 * não há imagem pra reamostrar e o reencode adiciona overhead.)
 */
const compressOrOriginal = async (
  pdf: Buffer,
  documentId: string,
  title: string,
  preCreatedTask?: IlovepdfTask,
): Promise<Buffer> => {
  const t0 = Date.now();
  try {
    const compressed = await compressPdf(pdf, documentId, title, preCreatedTask);
    const elapsedMs = Date.now() - t0;
    if (compressed.length === 0 || compressed.length >= pdf.length) {
      logger.info(
        { documentId, elapsedMs, originalBytes: pdf.length, compressedBytes: compressed.length },
        'PDF compression not worth it — sending original',
      );
      return pdf;
    }
    logger.info(
      {
        documentId,
        elapsedMs,
        originalBytes: pdf.length,
        compressedBytes: compressed.length,
        savedPct: Math.round((1 - compressed.length / pdf.length) * 100),
      },
      'PDF compressed',
    );
    return compressed;
  } catch (err) {
    logger.warn({ documentId, elapsedMs: Date.now() - t0, err }, 'PDF compression failed — sending original');
    return pdf;
  }
};

app.post<{ Body: PdfBody }>('/pdf', async (req, reply) => {
  const { documentId } = req.body ?? {};
  if (typeof documentId !== 'string' || documentId.length < 8) {
    reply.code(400).send({ error: 'bad_request', message: 'documentId is required' });
    return;
  }
  const authorization = req.headers['authorization'];
  const apiKeyHeader = req.headers['x-api-key'];
  const apiKey = Array.isArray(apiKeyHeader) ? apiKeyHeader[0] : apiKeyHeader;

  const authed = await authorize({
    documentId,
    authorization: typeof authorization === 'string' ? authorization : undefined,
    apiKey: typeof apiKey === 'string' ? apiKey : undefined,
  });

  // Abre a task de compressão em paralelo com o render — o start do ilovepdf
  // custa ~800ms e não depende do PDF. Se falhar, o compress cria a dele.
  const [{ buffer, title }, compressTask] = await Promise.all([
    renderDocumentPdf({ documentId: authed.documentId, workspaceId: authed.workspaceId }),
    startIlovepdfTask('compress').catch(() => undefined),
  ]);

  const clean = cleanTitle(title);
  const delivered = await compressOrOriginal(buffer, authed.documentId, clean, compressTask);

  reply
    .code(200)
    .header('Content-Type', 'application/pdf')
    .header('Content-Disposition', dispositionFor(clean, 'pdf'))
    .header('Cache-Control', 'no-store')
    .send(delivered);
});

// docx (Word) e pptx (PowerPoint) compartilham o mesmo fluxo via ilovepdf/pdfoffice.
const OFFICE_CONTENT_TYPE: Record<OfficeFormat, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

const handleOfficeExport = async (
  format: OfficeFormat,
  req: import('fastify').FastifyRequest<{ Body: PdfBody }>,
  reply: import('fastify').FastifyReply,
): Promise<void> => {
  const { documentId } = req.body ?? {};
  if (typeof documentId !== 'string' || documentId.length < 8) {
    reply.code(400).send({ error: 'bad_request', message: 'documentId is required' });
    return;
  }
  const authorization = req.headers['authorization'];
  const apiKeyHeader = req.headers['x-api-key'];
  const apiKey = Array.isArray(apiKeyHeader) ? apiKeyHeader[0] : apiKeyHeader;

  const authed = await authorize({
    documentId,
    authorization: typeof authorization === 'string' ? authorization : undefined,
    apiKey: typeof apiKey === 'string' ? apiKey : undefined,
  });

  // Start the ilovepdf task concurrently with PDF rendering to save ~800ms.
  const [{ buffer: pdfBuffer, title }, ilovepdfTask] = await Promise.all([
    renderDocumentPdf({ documentId: authed.documentId, workspaceId: authed.workspaceId }),
    startIlovepdfTask('pdfoffice'),
  ]);
  const clean = cleanTitle(title);
  const officeBuffer = await convertPdfToOffice(pdfBuffer, authed.documentId, clean, format, ilovepdfTask);

  reply
    .code(200)
    .header('Content-Type', OFFICE_CONTENT_TYPE[format])
    .header('Content-Disposition', dispositionFor(clean, format))
    .header('Cache-Control', 'no-store')
    .send(officeBuffer);
};

app.post<{ Body: PdfBody }>('/docx', (req, reply) => handleOfficeExport('docx', req, reply));
app.post<{ Body: PdfBody }>('/pptx', (req, reply) => handleOfficeExport('pptx', req, reply));

// ---------------------------------------------------------------------------
// /compress-attachment — comprime um PDF que o chat da IA já subiu no bucket
// ---------------------------------------------------------------------------
// Motivação: PDF até 16 MB vai NATIVO pro modelo (ele VÊ o layout das páginas);
// acima disso a rota de chat cai pra extração de texto e o resultado piora
// muito em deck exportado do Canva. Comprimir traz o arquivo de volta pra
// dentro da faixa nativa.
//
// O arquivo trafega service ↔ Supabase (o body aqui é só JSON), então o
// bodyLimit de 1 MB do Fastify não atrapalha nem em anexo de 50 MB.
//
// O client é quem apaga o original depois de trocar a referência — assim uma
// resposta perdida nunca deixa o chat apontando pra um arquivo que já sumiu.
// ---------------------------------------------------------------------------

const SOURCES_BUCKET = 'ai-design-sources';

type CompressAttachmentBody = { documentId?: unknown; path?: unknown };

app.post<{ Body: CompressAttachmentBody }>('/compress-attachment', async (req, reply) => {
  const { documentId, path } = req.body ?? {};
  if (typeof documentId !== 'string' || documentId.length < 8) {
    reply.code(400).send({ error: 'bad_request', message: 'documentId is required' });
    return;
  }
  if (typeof path !== 'string' || !path.endsWith('.pdf') || path.includes('..')) {
    reply.code(400).send({ error: 'bad_request', message: 'path must be a .pdf inside the bucket' });
    return;
  }
  const authorization = req.headers['authorization'];
  const apiKeyHeader = req.headers['x-api-key'];
  const apiKey = Array.isArray(apiKeyHeader) ? apiKeyHeader[0] : apiKeyHeader;

  const authed = await authorize({
    documentId,
    authorization: typeof authorization === 'string' ? authorization : undefined,
    apiKey: typeof apiKey === 'string' ? apiKey : undefined,
  });

  // O download abaixo usa a SECRET_KEY (bypassa RLS), então a checagem de
  // workspace tem que ser explícita: o path do bucket começa com o workspace
  // id, e ele precisa bater com o workspace do documento autorizado.
  if (!path.startsWith(`${authed.workspaceId}/`)) {
    reply.code(403).send({ error: 'forbidden', message: 'path does not belong to this workspace' });
    return;
  }

  const storage = serviceClient().storage.from(SOURCES_BUCKET);
  const { data: file, error: downloadErr } = await storage.download(path);
  if (downloadErr || !file) {
    reply.code(404).send({ error: 'not_found', message: 'attachment not found' });
    return;
  }
  const original = Buffer.from(await file.arrayBuffer());

  let compressed: Buffer;
  try {
    compressed = await compressPdf(original, authed.documentId, 'anexo');
  } catch (err) {
    req.log.warn({ documentId, err }, 'attachment compression failed');
    reply.code(200).send({ path, bytes: original.length, originalBytes: original.length, compressed: false });
    return;
  }
  if (compressed.length === 0 || compressed.length >= original.length) {
    reply.code(200).send({ path, bytes: original.length, originalBytes: original.length, compressed: false });
    return;
  }

  const compressedPath = path.replace(/\.pdf$/, '-compressed.pdf');
  const { error: uploadErr } = await storage.upload(compressedPath, compressed, {
    contentType: 'application/pdf',
    upsert: true,
  });
  if (uploadErr) {
    req.log.warn({ documentId, msg: uploadErr.message }, 'compressed attachment upload failed');
    reply.code(200).send({ path, bytes: original.length, originalBytes: original.length, compressed: false });
    return;
  }

  req.log.info(
    { documentId, originalBytes: original.length, compressedBytes: compressed.length },
    'attachment compressed',
  );
  reply.code(200).send({
    path: compressedPath,
    bytes: compressed.length,
    originalBytes: original.length,
    compressed: true,
  });
});

const start = async (): Promise<void> => {
  try {
    await app.listen({ port: config.port, host: '0.0.0.0' });
    void warmUpBrowser();
  } catch (err) {
    logger.error({ err }, 'Failed to start server');
    process.exit(1);
  }
};

const shutdown = async (signal: string): Promise<void> => {
  logger.info({ signal }, 'Shutting down');
  try {
    await app.close();
    await closeBrowser();
  } catch (err) {
    logger.error({ err }, 'Error during shutdown');
  } finally {
    process.exit(0);
  }
};

process.on('SIGINT', () => void shutdown('SIGINT'));
process.on('SIGTERM', () => void shutdown('SIGTERM'));

void start();
