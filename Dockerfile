FROM node:20-bookworm-slim
RUN apt-get update && apt-get install -y python3 make g++ \
  && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY studio-delta-production/package.json studio-delta-production/package-lock.json* ./
RUN npm install --omit=dev
COPY studio-delta-production/ ./
RUN mkdir -p /app/data && chmod 777 /app/data
ENV TZ=Africa/Johannesburg
ENV NODE_ENV=production
ENV DATA_DIR=/app/data
EXPOSE 8080
CMD ["node", "server/index.js"]
