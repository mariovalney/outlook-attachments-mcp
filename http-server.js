#!/usr/bin/env node
/**
 * Outlook Attachments MCP — HTTP entry point (remote, multi-user)
 *
 * Exposes the MCP server over Streamable HTTP so it can be added to Claude
 * (or any MCP client) as a remote custom connector:
 *
 *   GET  /healthz          — unauthenticated health check
 *   *    /.well-known/...  — OAuth discovery metadata
 *   GET  /authorize, /callback, POST /register, /token — OAuth proxy to Entra ID
 *   POST/GET/DELETE /mcp   — MCP endpoint (requires a Graph bearer token)
 *
 * Every MCP request carries the calling user's own Microsoft Graph access
 * token as the OAuth bearer; the token is bound to the request context so
 * all tools operate on that user's mailbox. Nothing is persisted server-side.
 */
require('dotenv').config();

const { randomUUID } = require('crypto');
const express = require('express');
const {
  StreamableHTTPServerTransport,
} = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const { isInitializeRequest } = require('@modelcontextprotocol/sdk/types.js');

const config = require('./config');
const { buildTools, createServer } = require('./server-factory');
const { runWithAccessToken } = require('./auth/request-context');
const { oauthRouter, requireGraphBearer } = require('./oauth');

const HTTP = config.HTTP_CONFIG;

// --- Startup validation ----------------------------------------------------

if (!config.AUTH_CONFIG.clientId && !config.USE_TEST_MODE) {
  console.error(
    'FATAL: OUTLOOK_CLIENT_ID is not set. Create an Entra ID app registration and set OUTLOOK_CLIENT_ID (see README).'
  );
  process.exit(1);
}

if (!HTTP.baseUrl) {
  console.error(
    'FATAL: BASE_URL is not set. Set it to the public HTTPS URL of this server (e.g. https://outlook-mcp.example.com).'
  );
  process.exit(1);
}

if (!HTTP.stateSecret) {
  // Random fallback keeps the server usable, but in-flight OAuth flows break
  // on restart and multi-replica deployments won't share state signatures.
  HTTP.stateSecret = randomUUID() + randomUUID();
  console.error(
    '⚠ STATE_SECRET is not set — generated an ephemeral one. Set STATE_SECRET to a fixed random value in production.'
  );
}

// --- MCP wiring ------------------------------------------------------------

const TOOLS = buildTools();

// Active Streamable HTTP sessions (in-memory; use sticky sessions if scaling
// horizontally).
const transports = new Map(); // sessionId -> StreamableHTTPServerTransport

async function handleMcpRequest(req, res) {
  const sessionId = req.headers['mcp-session-id'];
  let transport = sessionId ? transports.get(sessionId) : undefined;

  if (!transport) {
    if (!sessionId && req.method === 'POST' && isInitializeRequest(req.body)) {
      transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (sid) => {
          transports.set(sid, transport);
        },
      });
      transport.onclose = () => {
        if (transport.sessionId) transports.delete(transport.sessionId);
      };
      const server = createServer(TOOLS);
      await server.connect(transport);
    } else {
      res.status(400).json({
        jsonrpc: '2.0',
        error: {
          code: -32000,
          message: 'Bad Request: no valid MCP session. Send an initialize request first.',
        },
        id: null,
      });
      return;
    }
  }

  // Bind the caller's Graph token to this request's async context so every
  // tool handler resolves it via ensureAuthenticated().
  await runWithAccessToken(req.auth.accessToken, () =>
    transport.handleRequest(req, res, req.body)
  );
}

// --- HTTP app ---------------------------------------------------------------

const app = express();
app.set('trust proxy', true);
app.use(express.json({ limit: '4mb' }));
app.use(express.urlencoded({ extended: false }));

app.get('/healthz', (_req, res) => {
  res.json({
    status: 'ok',
    name: config.SERVER_NAME,
    version: config.SERVER_VERSION,
    tools: TOOLS.length,
  });
});

app.use(oauthRouter());

app.post('/mcp', requireGraphBearer, (req, res, next) => {
  handleMcpRequest(req, res).catch(next);
});
app.get('/mcp', requireGraphBearer, (req, res, next) => {
  handleMcpRequest(req, res).catch(next);
});
app.delete('/mcp', requireGraphBearer, (req, res, next) => {
  handleMcpRequest(req, res).catch(next);
});

// eslint-disable-next-line no-unused-vars
app.use((error, _req, res, _next) => {
  console.error(`[HTTP] Unhandled error: ${error.stack || error.message}`);
  if (!res.headersSent) {
    res.status(500).json({
      jsonrpc: '2.0',
      error: { code: -32603, message: 'Internal server error' },
      id: null,
    });
  }
});

const httpServer = app.listen(HTTP.port, '0.0.0.0', () => {
  console.log(
    `${config.SERVER_NAME} v${config.SERVER_VERSION} listening on :${HTTP.port} (MCP endpoint: ${HTTP.baseUrl}/mcp)`
  );
  console.log(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);
  if (
    !process.env.OUTLOOK_MAX_EMAILS_PER_SESSION &&
    !process.env.OUTLOOK_ALLOWED_RECIPIENTS &&
    !config.USE_TEST_MODE
  ) {
    console.error(
      '⚠ Safety belts not configured. Consider setting OUTLOOK_MAX_EMAILS_PER_SESSION and OUTLOOK_ALLOWED_RECIPIENTS for safer AI-assisted sending.'
    );
  }
});

function shutdown(signal) {
  console.log(`${signal} received, shutting down`);
  for (const transport of transports.values()) {
    try {
      transport.close();
    } catch (_e) {
      // best effort
    }
  }
  httpServer.close(() => process.exit(0));
  // Force-exit if connections linger past the platform grace period.
  setTimeout(() => process.exit(0), 10_000).unref();
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
