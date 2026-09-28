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
COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build

ENV NODE_ENV=production
ENV PORT=5000
ENV CHROMIUM_PATH=/usr/bin/chromium

EXPOSE 5000
CMD ["node", "dist/server.cjs"]
