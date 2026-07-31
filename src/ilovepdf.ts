import { logger } from './logger.js';

const ILOVEPDF_TOKEN =
  'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiIiLCJhdWQiOiIiLCJpYXQiOjE1MjMzNjQ4MjQsIm5iZiI6MTUyMzM2NDgyNCwianRpIjoicHJvamVjdF9wdWJsaWNfYzkwNWRkMWMwMWU5ZmQ3NzY5ODNjYTQwZDBhOWQyZjNfT1Vzd2EwODA0MGI4ZDJjN2NhM2NjZGE2MGQ2MTBhMmRkY2U3NyJ9.qvHSXgCJgqpC4gd6-paUlDLFmg0o2DsOvb1EUYPYx_E';
const SERVICE_URL = 'https://api.ilovepdf.com';
const API_VERSION = 'v1';
const APP_VERSION = 'web.0';

// Headers that mimic the browser so ilovepdf doesn't throttle server-side requests.
const BASE_HEADERS: Record<string, string> = {
  Authorization: `Bearer ${ILOVEPDF_TOKEN}`,
  Accept: 'application/json',
  'User-Agent':
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  Origin: 'https://www.ilovepdf.com',
  Referer: 'https://www.ilovepdf.com/',
};

type StartResponse = { server: string; task: string };
type UploadResponse = { server_filename: string };
type TaskStatus = { status: string; download_filename?: string; tool?: string };

const ts = () => Date.now();

/** `pdfoffice` converte pra docx/pptx; `compress` reduz o tamanho do PDF.
 *  Os dois falam o mesmo protocolo — start → upload → process → download. */
export type IlovepdfTool = 'pdfoffice' | 'compress';

// pdfoffice converte para docx (Word) ou pptx (PowerPoint) — mesmo fluxo, só muda convert_to.
export type OfficeFormat = 'docx' | 'pptx';

export type IlovepdfTask = { tool: IlovepdfTool; workerServer: string; taskId: string };

export const startIlovepdfTask = async (tool: IlovepdfTool): Promise<IlovepdfTask> => {
  const t0 = ts();
  const res = await fetch(`${SERVICE_URL}/${API_VERSION}/start/${tool}`, {
    headers: BASE_HEADERS,
  });
  if (!res.ok) throw new Error(`ilovepdf start failed: ${res.status} ${await res.text()}`);
  const data = (await res.json()) as StartResponse;
  logger.info({ elapsedMs: ts() - t0, tool }, 'ilovepdf: start');
  return {
    tool,
    workerServer: `https://${data.server}`,
    taskId: data.task,
  };
};

const uploadPdf = async (
  workerServer: string,
  taskId: string,
  pdfBuffer: Buffer,
  filename: string,
): Promise<string> => {
  const t0 = ts();
  const form = new FormData();
  form.append('task', taskId);
  form.append('name', filename);
  form.append('chunk', '0');
  form.append('chunks', '1');
  form.append('preview', '0');
  form.append('pdfinfo', '0');
  form.append('pdfforms', '0');
  form.append('pdfresetforms', '0');
  form.append('v', APP_VERSION);
  const pdfArrayBuffer = pdfBuffer.buffer.slice(
    pdfBuffer.byteOffset,
    pdfBuffer.byteOffset + pdfBuffer.byteLength,
  ) as ArrayBuffer;
  form.append('file', new Blob([pdfArrayBuffer], { type: 'application/pdf' }), filename);

  const res = await fetch(`${workerServer}/${API_VERSION}/upload`, {
    method: 'POST',
    headers: BASE_HEADERS,
    body: form,
  });
  if (!res.ok) throw new Error(`ilovepdf upload failed: ${res.status} ${await res.text()}`);
  const data = (await res.json()) as UploadResponse;
  logger.info({ elapsedMs: ts() - t0, bytes: pdfBuffer.length }, 'ilovepdf: upload');
  return data.server_filename;
};

const processTask = async (args: {
  workerServer: string;
  taskId: string;
  tool: IlovepdfTool;
  serverFilename: string;
  filename: string;
  /** Campos específicos do tool (convert_to, compression_level, ...). */
  fields: Record<string, string>;
  pollDeadlineMs: number;
}): Promise<void> => {
  const t0 = ts();
  const form = new FormData();
  form.append('task', args.taskId);
  form.append('tool', args.tool);
  for (const [key, value] of Object.entries(args.fields)) form.append(key, value);
  form.append('files[0][server_filename]', args.serverFilename);
  form.append('files[0][filename]', args.filename);

  const res = await fetch(`${args.workerServer}/${API_VERSION}/process`, {
    method: 'POST',
    headers: BASE_HEADERS,
    body: form,
  });
  const body = await res.text();
  logger.info({ elapsedMs: ts() - t0, status: res.status, body, tool: args.tool }, 'ilovepdf: process');
  if (!res.ok) throw new Error(`ilovepdf process failed: ${res.status} ${body}`);

  let parsed: TaskStatus | null = null;
  try { parsed = JSON.parse(body) as TaskStatus; } catch { /* empty body — fall through to poll */ }
  if (parsed?.status === 'TaskSuccess') return;
  if (parsed?.status === 'TaskError' || parsed?.status === 'TaskWrongPassword') {
    throw new Error(`ilovepdf task failed: ${parsed?.status}`);
  }

  // Poll when body is empty/unparseable OR returned a non-final status.
  logger.info({ taskStatus: parsed?.status ?? '(empty body)' }, 'ilovepdf: polling for task completion');
  const deadline = Date.now() + args.pollDeadlineMs;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, 1000));
    const pr = await fetch(`${args.workerServer}/${API_VERSION}/task/${args.taskId}`, { headers: BASE_HEADERS });
    if (!pr.ok) continue;
    const pd = (await pr.json()) as TaskStatus;
    logger.info({ taskStatus: pd.status }, 'ilovepdf: poll');
    if (pd.status === 'TaskSuccess') return;
    if (pd.status === 'TaskError' || pd.status === 'TaskWrongPassword') {
      throw new Error(`ilovepdf task failed: ${pd.status}`);
    }
  }
  throw new Error('ilovepdf task timed out');
};

const downloadResult = async (
  workerServer: string,
  taskId: string,
  timeoutMs: number,
): Promise<Buffer> => {
  const t0 = ts();
  const downloadHeaders = { ...BASE_HEADERS, Accept: '*/*' };
  const res = await fetch(`${workerServer}/${API_VERSION}/download/${taskId}`, {
    headers: downloadHeaders,
    // 10s: good servers deliver 4MB in ~2s; stalled ones send headers and stop — fail fast.
    signal: AbortSignal.timeout(timeoutMs),
  });
  const contentType = res.headers.get('content-type') ?? '';
  logger.info({ status: res.status, contentType }, 'ilovepdf: download response');
  if (!res.ok) throw new Error(`ilovepdf download failed: ${res.status} ${await res.text().catch(() => '')}`);
  const arrayBuffer = await res.arrayBuffer();
  logger.info({ elapsedMs: ts() - t0, bytes: arrayBuffer.byteLength }, 'ilovepdf: download');
  return Buffer.from(arrayBuffer);
};

/** Roda o pipeline completo de um tool, com retry opcional em outro servidor. */
const runTask = async (args: {
  tool: IlovepdfTool;
  pdfBuffer: Buffer;
  documentId: string;
  filename: string;
  fields: Record<string, string>;
  attempts: number;
  pollDeadlineMs: number;
  downloadTimeoutMs: number;
  preCreatedTask?: IlovepdfTask;
}): Promise<Buffer> => {
  const globalStart = ts();

  for (let attempt = 1; attempt <= args.attempts; attempt++) {
    try {
      // Re-create the task on retry — lands on a different (hopefully healthy) server.
      // Uma task pré-criada só serve se for do MESMO tool.
      const reusable =
        attempt === 1 && args.preCreatedTask?.tool === args.tool ? args.preCreatedTask : null;
      const { workerServer, taskId } = reusable ?? (await startIlovepdfTask(args.tool));
      logger.info({ documentId: args.documentId, workerServer, attempt, tool: args.tool }, 'ilovepdf task started');

      const serverFilename = await uploadPdf(workerServer, taskId, args.pdfBuffer, args.filename);
      await processTask({
        workerServer,
        taskId,
        tool: args.tool,
        serverFilename,
        filename: args.filename,
        fields: args.fields,
        pollDeadlineMs: args.pollDeadlineMs,
      });
      const out = await downloadResult(workerServer, taskId, args.downloadTimeoutMs);

      logger.info(
        { documentId: args.documentId, elapsedMs: ts() - globalStart, bytes: out.length, tool: args.tool },
        'ilovepdf result ready',
      );
      return out;
    } catch (err) {
      logger.warn({ documentId: args.documentId, attempt, tool: args.tool, err }, 'ilovepdf: attempt failed');
      if (attempt < args.attempts) continue;
      throw err;
    }
  }
  throw new Error('unreachable');
};

export const convertPdfToOffice = async (
  pdfBuffer: Buffer,
  documentId: string,
  title: string,
  format: OfficeFormat,
  preCreatedTask?: IlovepdfTask,
): Promise<Buffer> =>
  runTask({
    tool: 'pdfoffice',
    pdfBuffer,
    documentId,
    filename: `${title}.pdf`,
    fields: {
      convert_to: format,
      output_filename: title,
      packaged_filename: 'ilovepdf_converted',
      ocr: '0',
    },
    attempts: 2,
    pollDeadlineMs: 120_000,
    downloadTimeoutMs: 10_000,
    preCreatedTask,
  });

/**
 * Comprime o PDF. Best-effort por design: quem chama SEMPRE deve cair de volta
 * pro PDF original em qualquer falha — o resultado do usuário não pode depender
 * de um serviço externo. Por isso uma tentativa só, em vez do retry de 2× da
 * conversão pra Office (onde a espera é esperada e não há alternativa).
 *
 * ORÇAMENTO DE ~1 MINUTO no total (upload + poll + download). O ilovepdf entrega
 * em segundos no caso normal e em ~1 min no pior caso observado; esperar além
 * disso não recupera nada, só segura o usuário. Estourou, entrega o original —
 * que é um resultado perfeitamente bom, só maior.
 *
 * Note que a conversão pra Office (convertPdfToOffice) usa deadline bem maior
 * de propósito: lá não existe fallback, o usuário PRECISA do .docx/.pptx.
 */
const COMPRESS_POLL_MS = 35_000;
const COMPRESS_DOWNLOAD_MS = 12_000;

export const compressPdf = async (
  pdfBuffer: Buffer,
  documentId: string,
  title: string,
  preCreatedTask?: IlovepdfTask,
): Promise<Buffer> =>
  runTask({
    tool: 'compress',
    pdfBuffer,
    documentId,
    filename: `${title}.pdf`,
    fields: {
      compression_level: 'recommended',
      output_filename: title,
      packaged_filename: 'ilovepdf_compressed',
    },
    attempts: 1,
    pollDeadlineMs: COMPRESS_POLL_MS,
    downloadTimeoutMs: COMPRESS_DOWNLOAD_MS,
    preCreatedTask,
  });
