# syntax=docker/dockerfile:1.7
FROM node:20-bookworm-slim AS builder

WORKDIR /app

# tsc + tsx precisam apenas de node, não do chromium
COPY package.json package-lock.json* ./
RUN npm ci --no-audit --no-fund --ignore-scripts

COPY tsconfig.json ./
COPY src ./src

RUN npm run build


FROM node:20-bookworm-slim AS runtime

# Chromium vem do pacote @sparticuz/chromium (bundled no node_modules) — não do apt.
# O binário do apt trava com SIGTRAP em sandboxes restritos tipo gVisor (usado por
# algumas plataformas de hosting) ao fazer qualquer request de rede.
# Fontes base do sistema + libs necessárias pro headless seguem via apt.
# Fontes custom do projeto são servidas pelo app Next.js via /api/fonts,
# carregadas em tempo de render como @font-face dos arquivos .woff2.
RUN apt-get update && apt-get install -y --no-install-recommends \
      ca-certificates \
      fonts-liberation \
      fonts-noto-core \
      fonts-noto-cjk \
      fonts-noto-color-emoji \
      fontconfig \
      libnss3 \
      libatk-bridge2.0-0 \
      libxkbcommon0 \
      libx11-xcb1 \
      libxcomposite1 \
      libxdamage1 \
      libxrandr2 \
      libgbm1 \
      libpango-1.0-0 \
      libpangocairo-1.0-0 \
      libcups2 \
      libdrm2 \
      libasound2 \
      dumb-init \
  && rm -rf /var/lib/apt/lists/*

ENV NODE_ENV=production \
    PORT=8080 \
    # Puppeteer já baixaria um Chrome no install — evita (usamos @sparticuz/chromium).
    PUPPETEER_SKIP_DOWNLOAD=true

WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci --omit=dev --no-audit --no-fund --ignore-scripts \
  && npm cache clean --force

COPY --from=builder /app/dist ./dist

# Usuário não-root
RUN groupadd -r app && useradd -r -g app -G audio,video app \
  && mkdir -p /home/app && chown -R app:app /home/app /app
USER app

EXPOSE 8080

ENTRYPOINT ["/usr/bin/dumb-init", "--"]
CMD ["node", "dist/server.js"]
