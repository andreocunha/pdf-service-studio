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
  pdfPrintTimeoutMs: Number(optional('PDF_PRINT_TIMEOUT_MS', '120000')),
  isProduction: process.env.NODE_ENV === 'production',
  /**
   * Motor do Word (/docx). 'native' (padrão): montado do layout do Studio no
   * padrão da Lex (src/docx-native), com o iLovePDF de reserva se falhar.
   * 'ilovepdf': a conversão antiga do PDF. Voltar é só trocar a env.
   */
  docxEngine: optional('DOCX_ENGINE', 'native') as 'native' | 'ilovepdf',
} as const;

export type Config = typeof config;
