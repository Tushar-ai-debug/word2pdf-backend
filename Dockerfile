FROM node:18-slim

# Install LibreOffice for PDF conversion
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-writer \
      fonts-liberation \
      fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --production

COPY . .

# Create required directories
RUN mkdir -p uploads converted

EXPOSE 5000

CMD ["node", "server.js"]
