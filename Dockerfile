FROM node:20-bullseye

# LibreOffice + basic fonts
RUN apt-get update && \
    apt-get install -y libreoffice fonts-dejavu-core && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy package files first (better caching)
COPY package*.json ./
RUN npm ci --omit=dev

# Copy the rest
COPY . .

ENV NODE_ENV=production
ENV PORT=10000

EXPOSE 10000

CMD ["node", "server.js"]
