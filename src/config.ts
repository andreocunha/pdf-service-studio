const required = (name: string): string => {
  const value = process.env[name];
  if (!value || value.trim() === '') {
    throw new Error(`Missing required env var: ${name}`);
  }
  return value;
};

const optional = (name: string, fallback = ''): string => {
  const value = process.env[name];
  return value && value.trim() !== '' ? value : fallback;
};

export const config = {
  port: Number(optional('PORT', '8080')),
  renderBaseUrl: required('RENDER_BASE_URL').replace(/\/$/, ''),
  pdfServiceSecret: required('PDF_SERVICE_SECRET'),
  supabaseUrl: required('SUPABASE_URL').replace(/\/$/, ''),
  supabasePublishableKey: required('SUPABASE_PUBLISHABLE_KEY'),
  supabaseSecretKey: required('SUPABASE_SECRET_KEY'),
  chromiumExecutablePath: optional('CHROMIUM_EXECUTABLE_PATH'),
  requestTimeoutMs: Number(optional('REQUEST_TIMEOUT_MS', '30000')),
  isProduction: process.env.NODE_ENV === 'production',
} as const;

export type Config = typeof config;
