FROM node:20-bookworm-slim

RUN apt-get update && apt-get install -y \
    python3 python3-pip \
    chromium fonts-noto fonts-liberation \
    libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 libdrm2 \
    libxkbcommon0 libxcomposite1 libxdamage1 libxfixes3 libxrandr2 \
    libgbm1 libpango-1.0-0 libcairo2 libasound2 \
    ca-certificates curl git \
    && rm -rf /var/lib/apt/lists/*

RUN pip3 install --no-cache-dir --break-system-packages \
    ddddocr==1.5.6 onnxruntime pillow numpy

WORKDIR /app

# تثبيت إصدارات محددة ومتوافقة من الحزم الأساسية
RUN npm init -y && \
    npm install vite@^6.2.3 @vitejs/plugin-react@^5.0.4 @tailwindcss/vite@^4.1.14 typescript@~5.8.2 tsx@^4.21.0 esbuild@^0.25.0 --save-dev --no-audit --no-fund --legacy-peer-deps

# تثبيت باقي الاعتماديات من package.json
COPY package*.json ./
RUN npm install --no-audit --no-fund --legacy-peer-deps

COPY . .

# بناء الواجهة والخادم
RUN npx vite build && \
    npx esbuild server.ts --bundle --platform=node --format=cjs --packages=external --sourcemap --outfile=dist/server.cjs

ENV NODE_ENV=production
ENV PORT=5000
ENV CHROMIUM_PATH=/usr/bin/chromium
ENV PUPPETEER_SKIP_DOWNLOAD=true

EXPOSE 5000
CMD ["node", "dist/server.cjs"]
