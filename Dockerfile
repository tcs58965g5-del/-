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

# 1. تثبيت أدوات البناء عالمياً (لضمان وجودها في PATH)
RUN npm install -g vite@6.2.3 esbuild@0.25.0 typescript@5.8.2 tsx@4.21.0

# 2. نسخ وتثبيت اعتماديات المشروع
COPY package*.json ./
RUN npm install --no-audit --no-fund --legacy-peer-deps

# 3. ضمان وجود vite محلياً أيضاً
RUN npm install --save-dev --no-audit --no-fund --legacy-peer-deps \
    vite@6.2.3 \
    @vitejs/plugin-react@5.0.4 \
    @tailwindcss/vite@4.1.14

COPY . .

# 4. البناء باستخدام المسار المباشر (يتجاوز مشاكل PATH)
RUN ./node_modules/.bin/vite build && \
    ./node_modules/.bin/esbuild server.ts --bundle --platform=node --format=cjs --packages=external --sourcemap --outfile=dist/server.cjs

ENV NODE_ENV=production
ENV PORT=5000
ENV CHROMIUM_PATH=/usr/bin/chromium
ENV PUPPETEER_SKIP_DOWNLOAD=true

EXPOSE 5000
CMD ["node", "dist/server.cjs"]
