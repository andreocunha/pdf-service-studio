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

## Compatibilidade com Preview e iOS

Antes de imprimir, `preparePdfMasks` converte máscaras CSS de gradiente com
bordas duras (por exemplo, cartões com recorte de selo) em `clip-path` vetorial.
Isso evita os soft masks/padrões que o Quartz interpreta incorretamente. A
conversão acontece apenas na página de exportação, após o layout final; não
modifica templates salvos, texto, links ou dimensões. O contorno é amostrado a
288 dpi, com limites de memória e de complexidade. Gradientes suaves, máscaras
URL/SVG, pseudo-elementos e elementos que já têm `clip-path` ficam fora desta
conversão.

Validação:

```bash
npm run typecheck
npm run test:pdf-masks
```

O teste usa `CHROMIUM_EXECUTABLE_PATH` quando informado; no Mac usa o Chrome
instalado e em Linux o Chromium do pacote. No macOS, requer também `swiftc` e
`pdftotext` (Poppler) para renderizar o PDF com CoreGraphics/Quartz e verificar
as quatro bordas e a preservação de texto. PDFs de diagnóstico são gravados em
um diretório temporário informado pelo teste. Nenhum documento de produção é
necessário para esse teste. O workflow `PDF regression` compila a imagem do
Dockerfile e roda a regressão no Chromium Linux empacotado, usando os mesmos
argumentos de inicialização do serviço; não depende de credenciais ou documentos.

### Critério de aceitação do PDF final

O teste de máscaras acima verifica o recorte no mesmo navegador; ele **não**
valida equivalência entre Mac e Linux nem substitui o download completo.
`renderDocumentPdf()` devolve o PDF bruto. A rota `/pdf` ainda passa esse arquivo
por `compressOrOriginal()`. Não entregar o retorno bruto como amostra final.
O editor e a rota de PDF agora usam as mesmas métricas canônicas de fonte.
Ainda assim, a validação local no Mac não substitui a execução no Chromium
Linux da imagem de produção. Compare o mesmo documento/estado nos dois
ambientes e após a compressão normal. A bateria completa e reproduzível de
editor versus PDF está em [scripts/fidelity](scripts/fidelity/README.md).
Fontes ou imagens que falham impedem a exportação, inclusive na espera final
após o ajuste de viewport; um carregamento parcial não autoriza a captura.

O verificador abaixo rejeita mudanças de páginas, coordenadas de palavras,
bytes das fontes incorporadas, links e aumento de tamanho acima de 5%:

```bash
python3 -m pip install -r scripts/requirements-pdf-validation.txt
python3 scripts/check-pdf-regression.py referencia.pdf candidato.pdf
```

Os scripts Python são ferramentas locais de diagnóstico, não dependências do
servidor. Para recuperar um PDF já exportado sem recalcular seu layout, o
utilitário abaixo reconhece apenas a estrutura de máscaras Skia suportada e
recusa gradientes suaves. Requer `pdftocairo` (Poppler); revisar o resultado no
Quartz e no Poppler e executar o verificador antes de entregar:

```bash
python3 scripts/repair-skia-mask-pdf.py referencia.pdf corrigido.pdf
python3 scripts/check-pdf-regression.py referencia.pdf corrigido.pdf
```

### Ícones Phosphor durante a exportação

O frontend serve os ícones Phosphor empacotados por `/api/icons/ph/[icon]`.
Os templates mantêm as URLs portáveis originais; a renderização no editor e
no PDF converte URLs Iconify `ph:nome.svg` e `ph/nome.svg` para esse endpoint.
Isso elimina a dependência do limite de requisições do Iconify para os ícones
locais suportados. Vetores, pesos, cores e dimensões intrínsecas são preservados.
A rota não consulta a rede, e seus assets entram no file tracing de produção.

Teste no frontend: `npm run test:local-icons`. A regressão foi reproduzida nos
documentos `2f182d93-3431-4a29-bd2e-551b39ce57cc` e
`98db49b5-f35b-43b8-b89e-d0b621352419`; ambos ficaram prontos com todas as
imagens carregadas e sem chamadas ao Iconify, inclusive com o domínio bloqueado.
Os PDFs finais passaram pela renderização e compressão normais do serviço.

### Links internos em leitores de PDF

A rota `/pdf` resolve destinos nomeados para ações `/GoTo` com referência direta
à página, após a compressão. Isso remove a necessidade de o leitor consultar o
dicionário de destinos do Chromium. Destinos `/XYZ` com zoom herdado usam `/FitH`
para ajustar à largura e solicitar o alinhamento da seção ao topo da tela. A página
e a coordenada vertical são preservadas, inclusive para saltos dentro da mesma
página; destinos com zoom explícito permanecem intactos. Links externos
e destinos não resolvidos permanecem intactos; os contadores ficam no log.

`npm run test:pdf-links` cobre destinos legados, árvores de nomes, links externos,
destinos ausentes/cíclicos e reprocessamento. `check-pdf-regression.py` compara o
destino efetivo, além de textos, fontes, geometria e regiões clicáveis. A opção
`--fit-width-links` permite somente essa mudança de enquadramento, exigindo a
mesma página e coordenada vertical.

A normalização amplia a compatibilidade do formato, mas não acrescenta suporte
à navegação interna em aplicativos que não a implementam. Validar o toque nos
leitores nativos de iOS/Android antes de anunciar a correção para esses leitores;
a resolução de destinos pelo PDFKit no macOS não substitui esse teste.
