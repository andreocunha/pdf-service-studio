import { logger } from './logger.js';

const ILOVEPDF_TOKEN =
  'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiIiLCJhdWQiOiIiLCJpYXQiOjE1MjMzNjQ4MjQsIm5iZiI6MTUyMzM2NDgyNCwianRpIjoicHJvamVjdF9wdWJsaWNfYzkwNWRkMWMwMWU5ZmQ3NzY5ODNjYTQwZDBhOWQyZjNfT1Vzd2EwODA0MGI4ZDJjN2NhM2NjZGE2MGQ2MTBhMmRkY2U3NyJ9.qvHSXgCJgqpC4gd6-paUlDLFmg0o2DsOvb1EUYPYx_E';
const SERVICE_URL = 'https://api.ilovepdf.com';
const API_VERSION = 'v1';
const APP_VERSION = 'web.0';
const TOOL = 'pdfoffice';

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

// pdfoffice converte para docx (Word) ou pptx (PowerPoint) — mesmo fluxo, só muda convert_to.
export type OfficeFormat = 'docx' | 'pptx';

export type IlovepdfTask = { workerServer: string; taskId: string };

export const startIlovepdfTask = async (): Promise<IlovepdfTask> => startTask();

const startTask = async (): Promise<IlovepdfTask> => {
  const t0 = ts();
  const res = await fetch(`${SERVICE_URL}/${API_VERSION}/start/${TOOL}`, {
    headers: BASE_HEADERS,
  });
  if (!res.ok) throw new Error(`ilovepdf start failed: ${res.status} ${await res.text()}`);
  const data = (await res.json()) as StartResponse;
  logger.info({ elapsedMs: ts() - t0 }, 'ilovepdf: start');
  return {
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

const processTask = async (
  workerServer: string,
  taskId: string,
  serverFilename: string,
  filename: string,
  outputFilename: string,
  convertTo: OfficeFormat,
): Promise<void> => {
  const t0 = ts();
  const form = new FormData();
  form.append('task', taskId);
  form.append('tool', TOOL);
  form.append('convert_to', convertTo);
  form.append('output_filename', outputFilename);
  form.append('packaged_filename', 'ilovepdf_converted');
  form.append('ocr', '0');
  form.append('files[0][server_filename]', serverFilename);
  form.append('files[0][filename]', filename);

  const res = await fetch(`${workerServer}/${API_VERSION}/process`, {
    method: 'POST',
    headers: BASE_HEADERS,
    body: form,
  });
  const body = await res.text();
  logger.info({ elapsedMs: ts() - t0, status: res.status, body }, 'ilovepdf: process');
  if (!res.ok) throw new Error(`ilovepdf process failed: ${res.status} ${body}`);

  let parsed: TaskStatus | null = null;
  try { parsed = JSON.parse(body) as TaskStatus; } catch { /* empty body — fall through to poll */ }
  if (parsed?.status === 'TaskSuccess') return;
  if (parsed?.status === 'TaskError' || parsed?.status === 'TaskWrongPassword') {
    throw new Error(`ilovepdf task failed: ${parsed?.status}`);
  }

  // Poll when body is empty/unparseable OR returned a non-final status.
  logger.info({ taskStatus: parsed?.status ?? '(empty body)' }, 'ilovepdf: polling for task completion');
  const deadline = Date.now() + 120_000;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, 1000));
    const pr = await fetch(`${workerServer}/${API_VERSION}/task/${taskId}`, { headers: BASE_HEADERS });
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

const downloadOffice = async (workerServer: string, taskId: string): Promise<Buffer> => {
  const t0 = ts();
  const downloadHeaders = { ...BASE_HEADERS, Accept: '*/*' };
  const res = await fetch(`${workerServer}/${API_VERSION}/download/${taskId}`, {
    headers: downloadHeaders,
    // 10s: good servers deliver 4MB in ~2s; stalled ones send headers and stop — fail fast.
    signal: AbortSignal.timeout(10_000),
  });
  const contentType = res.headers.get('content-type') ?? '';
  logger.info({ status: res.status, contentType }, 'ilovepdf: download response');
  if (!res.ok) throw new Error(`ilovepdf download failed: ${res.status} ${await res.text().catch(() => '')}`);
  const arrayBuffer = await res.arrayBuffer();
  logger.info({ elapsedMs: ts() - t0, bytes: arrayBuffer.byteLength }, 'ilovepdf: download');
  return Buffer.from(arrayBuffer);
};

export const convertPdfToOffice = async (
  pdfBuffer: Buffer,
  documentId: string,
  title: string,
  format: OfficeFormat,
  preCreatedTask?: IlovepdfTask,
): Promise<Buffer> => {
  const globalStart = ts();
  const filename = `${title}.pdf`;
  const outputFilename = title;

  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      // Re-create the task on retry — lands on a different (hopefully healthy) server.
      const { workerServer, taskId } =
        attempt === 1 && preCreatedTask ? preCreatedTask : await startTask();
      logger.info({ documentId, workerServer, attempt, format }, 'ilovepdf task started');

      const serverFilename = await uploadPdf(workerServer, taskId, pdfBuffer, filename);
      await processTask(workerServer, taskId, serverFilename, filename, outputFilename, format);
      const officeBuffer = await downloadOffice(workerServer, taskId);

      logger.info({ documentId, elapsedMs: ts() - globalStart, bytes: officeBuffer.length, format }, 'ilovepdf office ready');
      return officeBuffer;
    } catch (err) {
      logger.warn({ documentId, attempt, format, err }, 'ilovepdf: attempt failed');
      if (attempt < 2) continue;
      throw err;
    }
  }
  throw new Error('unreachable');
};
