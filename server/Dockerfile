# A small image for the DuckWriter relay: docker build -t duckwriter-relay . && docker run -p 8787:8787 duckwriter-relay
FROM node:22-alpine
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --omit=dev
COPY relay.mjs ./
ENV PORT=8787
EXPOSE 8787
USER node
CMD ["node", "relay.mjs"]
