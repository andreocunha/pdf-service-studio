import { createHash } from 'node:crypto';

import { badRequest, forbidden, unauthorized } from './errors.js';
import { logger } from './logger.js';
import { serviceClient, userClient } from './supabase.js';

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export type AuthedRequest = {
  documentId: string;
  workspaceId: string;
  /** Presente quando autenticou por JWT de usuário */
  userId?: string;
  /** Presente quando autenticou por API key */
  apiKeyId?: string;
};

export const hashApiKey = (rawKey: string): string =>
  createHash('sha256').update(rawKey).digest('hex');

const PUBLIC_KEY_PREFIX = 'lex_live_';

const extractBearer = (header: string | undefined): string | null => {
  if (!header) return null;
  const match = /^Bearer\s+(.+)$/i.exec(header);
  return match ? match[1].trim() : null;
};

const authorizeByJwt = async (
  accessToken: string,
  documentId: string,
): Promise<AuthedRequest> => {
  const anon = userClient(accessToken);
  const { data: userData, error: userErr } = await anon.auth.getUser(accessToken);
  if (userErr || !userData.user) {
    throw unauthorized('Invalid or expired session token');
  }
  const { data: doc, error: docErr } = await anon
    .from('documents')
    .select('id, workspace_id')
    .eq('id', documentId)
    .maybeSingle();
  if (docErr) {
    logger.warn({ documentId, code: docErr.code, msg: docErr.message }, 'Supabase error on access check');
    throw forbidden('Document access check failed');
  }
  if (!doc) {
    throw forbidden('No access to this document');
  }
  return {
    documentId: doc.id,
    workspaceId: doc.workspace_id,
    userId: userData.user.id,
  };
};

const authorizeByApiKey = async (
  rawKey: string,
  documentId: string,
): Promise<AuthedRequest> => {
  if (!rawKey.startsWith(PUBLIC_KEY_PREFIX)) {
    throw unauthorized('Invalid API key format');
  }
  const keyHash = hashApiKey(rawKey);
  const svc = serviceClient();
  const { data: key, error: keyErr } = await svc
    .from('workspace_api_keys')
    .select('id, workspace_id, revoked_at')
    .eq('key_hash', keyHash)
    .maybeSingle();
  if (keyErr || !key) {
    throw unauthorized('API key not recognized');
  }
  if (key.revoked_at) {
    throw unauthorized('API key revoked');
  }
  const { data: doc, error: docErr } = await svc
    .from('documents')
    .select('id, workspace_id')
    .eq('id', documentId)
    .maybeSingle();
  if (docErr || !doc) {
    throw forbidden('Document not found');
  }
  if (doc.workspace_id !== key.workspace_id) {
    throw forbidden('API key does not belong to this document workspace');
  }
  // fire-and-forget: last_used_at
  void svc
    .from('workspace_api_keys')
    .update({ last_used_at: new Date().toISOString() })
    .eq('id', key.id);

  return {
    documentId: doc.id,
    workspaceId: doc.workspace_id,
    apiKeyId: key.id,
  };
};

export const authorize = async (args: {
  documentId: string;
  authorization?: string;
  apiKey?: string;
}): Promise<AuthedRequest> => {
  const { documentId } = args;
  if (!documentId) throw badRequest('documentId is required');
  if (!UUID_REGEX.test(documentId)) {
    throw badRequest(
      'documentId must be a UUID. Got: ' + documentId.slice(0, 64),
    );
  }
  const bearer = extractBearer(args.authorization);
  if (bearer) {
    return authorizeByJwt(bearer, documentId);
  }
  if (args.apiKey) {
    return authorizeByApiKey(args.apiKey, documentId);
  }
  throw unauthorized('Missing credentials (Authorization Bearer or x-api-key)');
};
