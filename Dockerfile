FROM node:20-slim
WORKDIR /app
COPY studio-delta-production/package.json studio-delta-production/package-lock.json* ./
RUN npm install --omit=dev
COPY studio-delta-production/ ./
ENV TZ=Africa/Johannesburg
ENV NODE_ENV=production
EXPOSE 8080
CMD ["node", "server/index.js"]
