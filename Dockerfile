FROM node:20-bullseye

RUN apt-get update && apt-get install -y libreoffice && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .

ENV NODE_ENV=production
EXPOSE 10000
CMD ["node", "server.js"]
