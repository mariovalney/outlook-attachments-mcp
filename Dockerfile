FROM node:22-alpine

LABEL org.opencontainers.image.title="Outlook Attachments MCP"
LABEL org.opencontainers.image.description="Multi-user remote MCP server for Microsoft Outlook — OAuth with Microsoft Entra ID, email, attachments, calendar, and contacts via Graph API"
LABEL org.opencontainers.image.source="https://github.com/mariovalney/outlook-attachments-mcp"
LABEL org.opencontainers.image.licenses="MIT"

ENV NODE_ENV=production
WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci --omit=dev --ignore-scripts

COPY http-server.js index.js server-factory.js config.js outlook-auth-server.js ./
COPY auth/ auth/
COPY oauth/ oauth/
COPY calendar/ calendar/
COPY categories/ categories/
COPY contacts/ contacts/
COPY email/ email/
COPY folder/ folder/
COPY rules/ rules/
COPY settings/ settings/
COPY advanced/ advanced/
COPY utils/ utils/

# Run as the unprivileged user shipped with the Node image
USER node

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD wget -qO- "http://127.0.0.1:${PORT:-3000}/healthz" || exit 1

ENTRYPOINT ["node", "http-server.js"]
