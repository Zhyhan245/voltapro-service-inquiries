FROM node:24-alpine

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY migrations ./migrations
COPY src ./src
COPY public ./public

RUN mkdir -p /app/data

ENV NODE_ENV=production
ENV PORT=3000
ENV DB_PATH=/app/data/voltapro.db

EXPOSE 3000

CMD ["sh", "-c", "node src/db/migrate.js && node src/db/seed.js && node src/server.js"]
