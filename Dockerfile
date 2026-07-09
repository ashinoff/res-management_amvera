# ── Stage 1: сборка фронтенда ────────────────────────────────────
FROM node:20-alpine AS build
WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm install
COPY frontend/ ./
RUN npm run build

# ── Stage 2: runtime — Node + Python (для анализаторов) ──────────
FROM node:20-slim
ENV TZ=Europe/Moscow
ENV NODE_ENV=production
WORKDIR /app

# Python 3 и зависимости анализаторов (pandas/openpyxl/xlrd/xlwt).
# --break-system-packages — Debian bookworm помечает окружение externally-managed.
RUN apt-get update \
 && apt-get install -y --no-install-recommends python3 python3-pip \
 && rm -rf /var/lib/apt/lists/*
COPY backend/requirements.txt ./backend/requirements.txt
RUN pip3 install --no-cache-dir --break-system-packages -r backend/requirements.txt

# Node-зависимости бэкенда
COPY backend/package*.json ./backend/
RUN cd backend && npm install --omit=dev

# Код бэкенда и собранный фронт из stage 1
COPY backend/ ./backend/
COPY --from=build /app/frontend/dist ./frontend/dist

ENV PORT=8000
EXPOSE 8000
WORKDIR /app/backend
CMD ["node", "server.js"]
