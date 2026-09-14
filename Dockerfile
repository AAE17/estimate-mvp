FROM node:20-bookworm-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-noto-core \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev
COPY . .
RUN mkdir -p /usr/share/fonts/truetype/custom \
    && (cp -f NotoSansGujarati-Regular.ttf /usr/share/fonts/truetype/custom/ || true) \
    && (cp -f fonts/NotoSansGujarati-Regular.ttf /usr/share/fonts/truetype/custom/ || true) \
    && fc-cache -f || true
ENV PORT=10000
EXPOSE 10000
CMD ["node", "server.js"]
