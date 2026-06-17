# Base oficial Apify com Node + Playwright (Chromium já instalado)
FROM apify/actor-node-playwright-chrome:20

# Instala Python 3 + pip e as libs usadas pelos scripts de análise.
# (precisa de root para apt-get; a imagem volta para o usuário padrão depois)
USER root
RUN apt-get update \
 && apt-get install -y --no-install-recommends python3 python3-pip python3-venv \
 && rm -rf /var/lib/apt/lists/*
RUN pip3 install --no-cache-dir --break-system-packages pandas openpyxl
USER myuser

# Dependências Node
COPY --chown=myuser:myuser package.json ./
RUN npm install --omit=dev --no-audit --no-fund \
 && npm ls || true

# Código do actor + scripts Python
COPY --chown=myuser:myuser . ./

CMD ["node", "main.js"]
