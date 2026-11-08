import { timingSafeEqual } from 'node:crypto';

import type { AuthedRequest } from './auth.js';
import { config } from './config.js';
import { HttpError } from './errors.js';
import { logger } from './logger.js';
import { serviceClient } from './supabase.js';

// ---------------------------------------------------------------------------
// Registro e teto de downloads (supabase/sql/205_document_downloads.sql)
// ---------------------------------------------------------------------------
// Todo /pdf, /docx, /pptx e /html é um download — do editor, do MCP ou da API — e
// passa pela fn_register_download ANTES de gerar o arquivo: ela confere o teto
// do workspace e grava quem baixou. Se a geração falhar, a linha é apagada.
//
// Exceção: o próprio Studio gerando PDF para uso interno (envio para
// assinatura, PDF final do DLM). Essas chamadas vêm do servidor do app com o
// cabeçalho x-lex-internal = PDF_SERVICE_SECRET e não contam nem registram.
// ---------------------------------------------------------------------------

export type DownloadFormat = 'pdf' | 'docx' | 'pptx' | 'html';
type DownloadSource = 'editor' | 'mcp' | 'api';

const sameSecret = (given: string): boolean => {
  const a = Buffer.from(given);
  const b = Buffer.from(config.pdfServiceSecret);
  return a.length === b.length && timingSafeEqual(a, b);
};

export const isInternalCall = (header: string | string[] | undefined): boolean => {
  const value = Array.isArray(header) ? header[0] : header;
  return typeof value === 'string' && value.length > 0 && sameSecret(value);
};

/** O corpo diz de onde veio ('mcp'); é só um rótulo — não muda se conta. */
const sourceOf = (authed: AuthedRequest, label: unknown): DownloadSource => {
  if (authed.apiKeyId) return 'api';
  return label === 'mcp' ? 'mcp' : 'editor';
};

/**
 * Registra o download (ou recusa com 402 download_limit_reached).
 * Devolve o id da linha para `cancelDownload` em caso de falha, ou null
 * quando não registrou (chamada interna, ou banco fora — aí o download segue:
 * o arquivo nunca deixa de sair por causa do registro).
 */
export const registerDownload = async (
  authed: AuthedRequest,
  format: DownloadFormat,
  sourceLabel: unknown,
): Promise<string | null> => {
  let result: { ok: boolean; id?: string; used?: number; limit?: number };
  try {
    // O supabase-js costuma devolver { error }, mas também pode lançar (rede,
    // cliente mal configurado): as duas formas liberam o download.
    const { data, error } = await serviceClient().rpc('fn_register_download', {
      p_document_id: authed.documentId,
      p_format: format,
      p_source: sourceOf(authed, sourceLabel),
      p_user_id: authed.userId ?? null,
      p_api_key_id: authed.apiKeyId ?? null,
    });
    if (error) throw new Error(`${error.code ?? ''} ${error.message}`.trim());
    result = data as typeof result;
  } catch (err) {
    logger.error({ documentId: authed.documentId, err }, 'download register failed — allowing');
    return null;
  }
  if (!result.ok) {
    logger.info({ documentId: authed.documentId, workspaceId: authed.workspaceId, used: result.used, limit: result.limit }, 'download limit reached');
    throw new HttpError(
      402,
      'download_limit_reached',
      `Este workspace atingiu o limite de ${result.limit} downloads. Fale com a Lex para liberar mais.`,
    );
  }
  return result.id ?? null;
};

export const cancelDownload = async (id: string | null): Promise<void> => {
  if (!id) return;
  try {
    const { error } = await serviceClient().from('document_downloads').delete().eq('id', id);
    if (error) throw new Error(error.message);
  } catch (err) {
    // Nunca troca o erro original da geração pelo do cancelamento.
    logger.warn({ id, err }, 'download cancel failed');
  }
};
