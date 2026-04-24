import { createClient, type SupabaseClient } from '@supabase/supabase-js';

import { config } from './config.js';

/**
 * Client com SECRET_KEY — bypassa RLS. Use apenas pra operações de confiança:
 * - validar hash de API key em workspace_api_keys
 * - cruzar documento x workspace
 * - atualizar last_used_at
 */
export const serviceClient = (): SupabaseClient =>
  createClient(config.supabaseUrl, config.supabaseSecretKey, {
    auth: { persistSession: false, autoRefreshToken: false },
  });

/**
 * Client com PUBLISHABLE_KEY + token de sessão do usuário — RLS aplicada
 * como se fosse o próprio user. Usado pra: "este user tem acesso a este doc?"
 */
export const userClient = (accessToken: string): SupabaseClient =>
  createClient(config.supabaseUrl, config.supabasePublishableKey, {
    auth: { persistSession: false, autoRefreshToken: false },
    global: { headers: { Authorization: `Bearer ${accessToken}` } },
  });
