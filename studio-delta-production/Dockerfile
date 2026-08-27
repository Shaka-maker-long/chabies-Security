FROM node:20-slim
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --omit=dev
COPY . .
ENV TZ=Africa/Johannesburg
ENV NODE_ENV=production
EXPOSE 8080
CMD ["node", "server/index.js"]
