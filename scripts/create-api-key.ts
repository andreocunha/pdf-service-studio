#!/usr/bin/env tsx
// Gera uma nova API key para um workspace e persiste no Supabase.
// Uso:
//   tsx scripts/create-api-key.ts --workspace <uuid> --name "Minha integração"
//
// Precisa das env vars SUPABASE_URL e SUPABASE_SECRET_KEY configuradas
// (use o mesmo .env que o pdf-service carrega).

import { createHash, randomBytes } from 'node:crypto';

import { createClient } from '@supabase/supabase-js';

const parseArgs = () => {
  const out: Record<string, string> = {};
  for (let i = 2; i < process.argv.length; i++) {
    const arg = process.argv[i];
    if (arg.startsWith('--')) {
      const key = arg.slice(2);
      const value = process.argv[i + 1];
      if (value && !value.startsWith('--')) {
        out[key] = value;
        i++;
      } else {
        out[key] = 'true';
      }
    }
  }
  return out;
};

const main = async () => {
  const args = parseArgs();
  const workspaceId = args.workspace;
  const name = args.name ?? 'API key';
  if (!workspaceId) {
    console.error('Missing --workspace <uuid>');
    process.exit(1);
  }
  const url = process.env.SUPABASE_URL;
  const secret = process.env.SUPABASE_SECRET_KEY;
  if (!url || !secret) {
    console.error('Missing SUPABASE_URL or SUPABASE_SECRET_KEY in env');
    process.exit(1);
  }
  // 32 bytes de entropia em base64url
  const raw = randomBytes(32).toString('base64url');
  const fullKey = `lex_live_${raw}`;
  const keyHash = createHash('sha256').update(fullKey).digest('hex');
  const keyPrefix = fullKey.slice(0, 16);

  const supabase = createClient(url, secret, {
    auth: { persistSession: false, autoRefreshToken: false },
  });
  const { data, error } = await supabase
    .from('workspace_api_keys')
    .insert({
      workspace_id: workspaceId,
      name,
      key_hash: keyHash,
      key_prefix: keyPrefix,
    })
    .select('id, key_prefix')
    .single();

  if (error) {
    console.error('Insert failed:', error.message);
    process.exit(1);
  }

  console.log('\n✓ API key criada com sucesso');
  console.log('   id:      ', data.id);
  console.log('   prefix:  ', data.key_prefix);
  console.log('\n⚠️  Copie a chave agora. Não é possível recuperá-la depois:\n');
  console.log('   ' + fullKey + '\n');
};

void main();
