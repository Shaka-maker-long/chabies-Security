FROM node:22-bookworm-slim
WORKDIR /app

RUN apt-get update \
  && apt-get install -y --no-install-recommends ca-certificates \
  && rm -rf /var/lib/apt/lists/*

COPY package.json package-lock.json* ./
RUN if [ -f package-lock.json ]; then npm ci --omit=dev; else npm install --omit=dev; fi

COPY . .
RUN mkdir -p /app/data && chmod 777 /app/data

ENV TZ=Africa/Johannesburg
ENV NODE_ENV=production
ENV DATA_DIR=/app/data
EXPOSE 8080
CMD ["node", "server/index.js"]
