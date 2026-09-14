FROM node:20-bookworm-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-noto-core \
    fonts-lohit-gujarati \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev
COPY . .
ENV PORT=10000
EXPOSE 10000
CMD ["node", "server.js"]
