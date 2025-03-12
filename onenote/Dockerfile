FROM node:20-alpine

WORKDIR /app

# Install dependencies
COPY package*.json ./
RUN npm install

# Copy source code
COPY . .

# Build TypeScript code
RUN npm run build

# Create a directory for token cache and device code
RUN mkdir -p /app/dist


# Set environment variables
ENV NODE_ENV=production

# Run the application
CMD ["node", "dist/index.js"] 