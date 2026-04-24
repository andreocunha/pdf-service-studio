import { createHmac, timingSafeEqual } from 'node:crypto';

import { config } from './config.js';

type RenderTokenPayload = {
  documentId: string;
  workspaceId: string;
  /** Seconds since epoch — token expiration */
  exp: number;
};

const b64url = (buf: Buffer): string =>
  buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

const b64urlDecode = (input: string): Buffer => {
  const pad = 4 - (input.length % 4);
  const padded = pad < 4 ? input + '='.repeat(pad) : input;
  return Buffer.from(padded.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
};

const sign = (input: string): string => {
  const mac = createHmac('sha256', config.pdfServiceSecret).update(input).digest();
  return b64url(mac);
};

/**
 * Produz um token curto (padrão: 60s) assinado com PDF_SERVICE_SECRET.
 * Usado pela rota /render-pdf/[id] no app principal pra autorizar o bypass de RLS.
 */
export const signRenderToken = (
  documentId: string,
  workspaceId: string,
  ttlSeconds = 60,
): string => {
  const payload: RenderTokenPayload = {
    documentId,
    workspaceId,
    exp: Math.floor(Date.now() / 1000) + ttlSeconds,
  };
  const body = b64url(Buffer.from(JSON.stringify(payload)));
  const sig = sign(body);
  return `${body}.${sig}`;
};

export const verifyRenderToken = (token: string): RenderTokenPayload => {
  const parts = token.split('.');
  if (parts.length !== 2) {
    throw new Error('Malformed render token');
  }
  const [body, sig] = parts;
  const expected = sign(body);
  const a = Buffer.from(sig);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !timingSafeEqual(a, b)) {
    throw new Error('Invalid render token signature');
  }
  const payload = JSON.parse(b64urlDecode(body).toString('utf8')) as RenderTokenPayload;
  if (typeof payload.exp !== 'number' || payload.exp < Math.floor(Date.now() / 1000)) {
    throw new Error('Render token expired');
  }
  return payload;
};
