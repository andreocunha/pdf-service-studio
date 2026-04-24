# Lex PDF Service

Serviço externo que gera PDFs a partir de documentos do Lex Studio. Recebe um
`documentId` + credencial (JWT de usuário ou API key de workspace), aciona o
Chromium headless contra a rota `/render-pdf/[id]` do app principal e devolve
o PDF.

## Arquitetura

```
  ┌──────────────┐   POST /pdf         ┌────────────────┐
  │   Cliente    │ ──────────────────▶ │   pdf-service  │
  │ (frontend /  │  { documentId }     │  (Node +       │
  │  integração) │  Auth: JWT ou       │   Puppeteer)   │
  │              │     x-api-key       │                │
  └──────────────┘                     └───────┬────────┘
                                               │ validar credencial
                                               │ assinar HMAC (60s)
                                               │ abrir Chromium
                                               ▼
                                      ┌────────────────────┐
                                      │ lex-studio-v2      │
                                      │ /render-pdf/[id]?t │
                                      │ (server component) │
                                      │  ↓                 │
                                      │  fetch yjs_state   │
                                      │  render readOnly   │
                                      │  window.__PDF_READY│
                                      └────────┬───────────┘
                                               │ page.pdf()
                                               ▼
                                            binário PDF
```

**Por que Puppeteer?** O editor renderiza documentos com tipografia custom
(Metropolis, Graphik, etc.), design blocks com HTML arbitrário, menus de
navegação e layouts paginados. Replicar esse rendering em Node seria
reescrever meio editor. O Chromium usa o mesmo pipeline de renderização do
navegador, então o PDF fica pixel-perfect. Links inline `<a href>` e botões
de menu (SectionNav) viram links clicáveis dentro do PDF.

## Deploy (Railway, plano Hobby)

1. `railway init` dentro deste diretório (ou via dashboard apontando pro
   subpath `pdf-service/`).
2. Variáveis de ambiente (copiar de `.env.example`):
   - `RENDER_BASE_URL` — URL pública do Next.js (`https://lexstudio.ai`)
   - `PDF_SERVICE_SECRET` — HMAC compartilhado com o Next.js (`openssl rand -hex 32`)
   - `SUPABASE_URL`, `SUPABASE_PUBLISHABLE_KEY`, `SUPABASE_SECRET_KEY`
3. `railway up`. O Dockerfile provisiona o Chromium do sistema em
   `/usr/bin/chromium` (200 MB, cabe de sobra nos 8 GB do Hobby).
4. No app Next.js, adicionar a mesma `PDF_SERVICE_SECRET` às env vars do
   Vercel/Railway. Sem isso a rota `/render-pdf/[id]` retorna 404.

## Dev local

```bash
cp .env.example .env
# Edite .env com os valores corretos.
# Para dev: CHROMIUM_EXECUTABLE_PATH=/Applications/Google Chrome.app/Contents/MacOS/Google Chrome
npm install
npm run dev       # tsx watch + auto-reload
```

Lado do Next.js (`lex-studio-v2`):
```bash
echo "PDF_SERVICE_SECRET=<mesmo_hex>" >> .env.local
npm run dev
```

## Endpoints

### `GET /health`

```json
{ "ok": true, "version": "0.1.0" }
```

### `POST /pdf`

**Auth** (uma das duas):
- `Authorization: Bearer <supabase_jwt>` — JWT do usuário logado (seu frontend).
- `x-api-key: lex_live_...` — API key do workspace (integrações externas).

**Body**:
```json
{ "documentId": "<uuid>" }
```

**Resposta**: `application/pdf` binário.

### Exemplo com JWT (frontend)

```ts
const session = await supabase.auth.getSession();
const res = await fetch(`${PDF_SERVICE_URL}/pdf`, {
  method: 'POST',
  headers: {
    'Authorization': `Bearer ${session.data.session!.access_token}`,
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({ documentId }),
});
const blob = await res.blob();
// triggering download...
```

### Exemplo com API key (servidor-a-servidor)

```bash
curl -X POST https://pdf.lexstudio.ai/pdf \
  -H "x-api-key: lex_live_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" \
  -H "Content-Type: application/json" \
  -d '{"documentId":"<uuid>"}' \
  -o documento.pdf
```

## Criar API key para um workspace

Enquanto não houver UI no `lex-studio-v2`, use o script:

```bash
npm install
# Exporta SUPABASE_URL e SUPABASE_SECRET_KEY (mesmos do .env)
npx tsx scripts/create-api-key.ts \
  --workspace <uuid_do_workspace> \
  --name "Integração ACME"
```

O script imprime a key pleno uma única vez — copie no ato. O banco guarda
apenas o SHA-256. Para revogar, basta `UPDATE workspace_api_keys SET
revoked_at = now() WHERE id = '<uuid>'`.

## Performance

- Chromium fica vivo entre requisições (warm pool). Cold-start apenas na
  primeira request pós-deploy (~2 s). Requisições subsequentes: 400-1200 ms
  de overhead + tempo de render do documento.
- Documentos de até ~30 páginas: tipicamente < 3 s total.
- Documentos gigantes (100+ páginas com muitas imagens): pode passar de 10 s.

## Observações de segurança

- **HMAC**: o token assinado pelo serviço tem TTL de 60 s e só é aceito pela
  rota `/render-pdf/[id]` se a assinatura bater. O secret nunca sai do lado
  servidor.
- **API keys**: armazenadas como SHA-256. Prefixadas com `lex_live_` para
  facilitar detecção acidental em logs/diffs. Scope: workspace — uma key só
  pode baixar documentos do workspace dela.
- **Rate limit**: 30 req/minuto por chave (JWT ou API key). Configurável.
- **Bypass de RLS**: a rota `/render-pdf/[id]` usa SECRET_KEY, mas só depois
  de validar HMAC — não é um endpoint público de leitura de documentos.
