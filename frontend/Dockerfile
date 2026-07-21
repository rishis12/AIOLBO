# AIO LBO Frontend Docker Image
FROM node:20-alpine

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm ci

# Copy source files
COPY . .

# Expose dev server port
EXPOSE 5173

# Run dev server (for development)
CMD ["npm", "run", "dev", "--", "--host"]
