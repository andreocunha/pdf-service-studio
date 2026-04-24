import cors from '@fastify/cors';
import rateLimit from '@fastify/rate-limit';
import Fastify from 'fastify';

import { authorize } from './auth.js';
import { closeBrowser, warmUpBrowser } from './browser.js';
import { config } from './config.js';
import { HttpError } from './errors.js';
import { logger } from './logger.js';
import { renderDocumentPdf } from './render.js';

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

  const { buffer, title } = await renderDocumentPdf({
    documentId: authed.documentId,
    workspaceId: authed.workspaceId,
  });

  // Sanitiza só caracteres reservados de filesystem (preserva acentos, parênteses, etc).
  const clean = (title || 'documento')
    // eslint-disable-next-line no-control-regex
    .replace(/[/\\:*?"<>|\x00-\x1F]/g, '')
    .replace(/\s+/g, ' ')
    .trim() || 'documento';
  // ASCII fallback pra clientes que não entendem RFC 5987 (é raro hoje, mas seguro).
  const asciiFallback = clean.replace(/[^\x20-\x7E]/g, '_');
  const encodedUtf8 = encodeURIComponent(clean);

  reply
    .code(200)
    .header('Content-Type', 'application/pdf')
    .header(
      'Content-Disposition',
      `attachment; filename="${asciiFallback}.pdf"; filename*=UTF-8''${encodedUtf8}.pdf`,
    )
    .header('Cache-Control', 'no-store')
    .send(buffer);
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
