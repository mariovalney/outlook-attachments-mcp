/**
 * Stateless OAuth 2.1 proxy in front of Microsoft Entra ID.
 *
 * Implements the endpoints a remote MCP client (e.g. Claude custom
 * connectors) needs to authenticate users with Microsoft and obtain a
 * Microsoft Graph access token, which the client then sends as the bearer
 * on every MCP request:
 *
 *   GET  /.well-known/oauth-protected-resource   (RFC 9728)
 *   GET  /.well-known/oauth-authorization-server (RFC 8414)
 *   POST /register                               (RFC 7591 stub)
 *   GET  /authorize   -> redirects to Microsoft login (PKCE passed through)
 *   GET  /callback    -> relays the authorization code back to the client
 *   POST /token       -> proxies code/refresh grants to the Entra endpoint
 *
 * No token or session is persisted server-side: the MCP client stores the
 * access/refresh tokens and refreshes them through /token. The only state
 * crossing the Microsoft round-trip travels inside an HMAC-signed `state`
 * parameter. PKCE (bound by Microsoft to the authorization code) protects
 * the code relay.
 */
const crypto = require('crypto');
const express = require('express');
const config = require('../config');

const AUTH = config.AUTH_CONFIG;
const HTTP = config.HTTP_CONFIG;

const STATE_TTL_MS = 10 * 60 * 1000; // authorize -> callback window
const TOKEN_CACHE_TTL_MS = 5 * 60 * 1000; // bearer validation cache

// ---------------------------------------------------------------------------
// Signed state helpers (stateless relay of the client's redirect_uri/state)
// ---------------------------------------------------------------------------

function b64url(buf) {
  return Buffer.from(buf).toString('base64url');
}

function signState(payload, secret) {
  const body = b64url(JSON.stringify(payload));
  const sig = crypto.createHmac('sha256', secret).update(body).digest('base64url');
  return `${body}.${sig}`;
}

function verifyState(state, secret) {
  if (typeof state !== 'string') return null;
  const dot = state.lastIndexOf('.');
  if (dot <= 0) return null;
  const body = state.slice(0, dot);
  const sig = state.slice(dot + 1);
  const expected = crypto
    .createHmac('sha256', secret)
    .update(body)
    .digest('base64url');
  const a = Buffer.from(sig);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try {
    const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    if (!payload.exp || Date.now() > payload.exp) return null;
    return payload;
  } catch (_e) {
    return null;
  }
}

// ---------------------------------------------------------------------------
// redirect_uri validation
// ---------------------------------------------------------------------------

function isAllowedRedirectUri(redirectUri) {
  let url;
  try {
    url = new URL(redirectUri);
  } catch (_e) {
    return false;
  }
  const isLocalhost =
    url.hostname === 'localhost' || url.hostname === '127.0.0.1';
  if (url.protocol !== 'https:' && !(url.protocol === 'http:' && isLocalhost)) {
    return false;
  }
  if (HTTP.allowedRedirectHosts.length > 0) {
    const host = url.hostname.toLowerCase();
    return HTTP.allowedRedirectHosts.some(
      (allowed) => host === allowed || host.endsWith(`.${allowed}`)
    );
  }
  return true;
}

// ---------------------------------------------------------------------------
// Bearer validation — the bearer IS a Microsoft Graph access token.
// Validated with a GET /me round-trip, cached in memory per token hash.
// ---------------------------------------------------------------------------

const tokenCache = new Map(); // sha256(token) -> { user, expiresAt }

function cacheKey(token) {
  return crypto.createHash('sha256').update(token).digest('hex');
}

function pruneTokenCache() {
  const now = Date.now();
  for (const [key, entry] of tokenCache) {
    if (entry.expiresAt <= now) tokenCache.delete(key);
  }
}

async function validateGraphToken(token) {
  const key = cacheKey(token);
  const cached = tokenCache.get(key);
  if (cached && cached.expiresAt > Date.now()) {
    return cached.user;
  }

  const response = await fetch(
    `${config.GRAPH_API_ENDPOINT}me?$select=id,displayName,userPrincipalName,mail`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (response.status === 401 || response.status === 403) {
    return null;
  }
  if (!response.ok) {
    throw new Error(`Graph token validation failed with ${response.status}`);
  }

  const me = await response.json();
  const user = {
    id: me.id,
    displayName: me.displayName,
    email: me.mail || me.userPrincipalName,
  };
  pruneTokenCache();
  tokenCache.set(key, { user, expiresAt: Date.now() + TOKEN_CACHE_TTL_MS });
  return user;
}

function unauthorized(res, description) {
  const resourceMetadata = `${HTTP.baseUrl}/.well-known/oauth-protected-resource`;
  res
    .status(401)
    .set(
      'WWW-Authenticate',
      `Bearer realm="mcp", error="invalid_token", error_description="${description}", resource_metadata="${resourceMetadata}"`
    )
    .json({
      jsonrpc: '2.0',
      error: { code: -32001, message: `Unauthorized: ${description}` },
      id: null,
    });
}

/**
 * Express middleware protecting /mcp. On success sets req.auth =
 * { accessToken, user } for the request handler.
 */
async function requireGraphBearer(req, res, next) {
  const header = req.headers.authorization || '';
  const match = header.match(/^Bearer\s+(.+)$/i);
  if (!match) {
    unauthorized(res, 'Missing bearer token');
    return;
  }
  const token = match[1].trim();

  // Test mode bypass mirrors the stdio behaviour (mock Graph data).
  if (config.USE_TEST_MODE && token.startsWith('test_access_token_')) {
    req.auth = { accessToken: token, user: { id: 'test', email: 'test@example.com' } };
    next();
    return;
  }

  try {
    const user = await validateGraphToken(token);
    if (!user) {
      unauthorized(res, 'Token rejected by Microsoft Graph');
      return;
    }
    req.auth = { accessToken: token, user };
    next();
  } catch (error) {
    console.error(`[OAUTH] Bearer validation error: ${error.message}`);
    res.status(503).json({
      jsonrpc: '2.0',
      error: { code: -32000, message: 'Token validation unavailable' },
      id: null,
    });
  }
}

// ---------------------------------------------------------------------------
// Router
// ---------------------------------------------------------------------------

function oauthRouter() {
  const router = express.Router();
  const scopes = AUTH.scopes.join(' ');

  const authServerMetadata = () => ({
    issuer: HTTP.baseUrl,
    authorization_endpoint: `${HTTP.baseUrl}/authorize`,
    token_endpoint: `${HTTP.baseUrl}/token`,
    registration_endpoint: `${HTTP.baseUrl}/register`,
    response_types_supported: ['code'],
    response_modes_supported: ['query'],
    grant_types_supported: ['authorization_code', 'refresh_token'],
    code_challenge_methods_supported: ['S256'],
    token_endpoint_auth_methods_supported: ['none', 'client_secret_post'],
    scopes_supported: AUTH.scopes,
  });

  const protectedResourceMetadata = () => ({
    resource: `${HTTP.baseUrl}/mcp`,
    authorization_servers: [HTTP.baseUrl],
    bearer_methods_supported: ['header'],
    scopes_supported: AUTH.scopes,
  });

  // RFC 9728 — served both at the root and with the /mcp path suffix,
  // since clients may probe either form.
  router.get(
    ['/.well-known/oauth-protected-resource', '/.well-known/oauth-protected-resource/mcp'],
    (_req, res) => res.json(protectedResourceMetadata())
  );

  // RFC 8414 (some clients also probe the OIDC discovery path)
  router.get(
    ['/.well-known/oauth-authorization-server', '/.well-known/openid-configuration'],
    (_req, res) => res.json(authServerMetadata())
  );

  // RFC 7591 stub — Entra ID has no dynamic client registration, so every
  // MCP client shares the single Entra app registration configured via env.
  // The synthetic client_id keeps clients happy; it is never checked later.
  router.post('/register', (req, res) => {
    const body = req.body || {};
    res.status(201).json({
      client_id: crypto.randomUUID(),
      client_id_issued_at: Math.floor(Date.now() / 1000),
      redirect_uris: body.redirect_uris || [],
      client_name: body.client_name || 'mcp-client',
      grant_types: ['authorization_code', 'refresh_token'],
      response_types: ['code'],
      token_endpoint_auth_method: 'none',
    });
  });

  router.get('/authorize', (req, res) => {
    const {
      redirect_uri: redirectUri,
      state: clientState,
      code_challenge: codeChallenge,
      code_challenge_method: codeChallengeMethod,
      response_type: responseType,
    } = req.query;

    if (!redirectUri || !isAllowedRedirectUri(redirectUri)) {
      res.status(400).json({
        error: 'invalid_request',
        error_description: 'redirect_uri missing or not allowed',
      });
      return;
    }
    if (responseType && responseType !== 'code') {
      res.status(400).json({
        error: 'unsupported_response_type',
        error_description: 'Only response_type=code is supported',
      });
      return;
    }
    if (!codeChallenge) {
      res.status(400).json({
        error: 'invalid_request',
        error_description: 'PKCE code_challenge is required',
      });
      return;
    }

    const state = signState(
      {
        ru: redirectUri,
        st: typeof clientState === 'string' ? clientState : '',
        exp: Date.now() + STATE_TTL_MS,
      },
      HTTP.stateSecret
    );

    const authorizeUrl = new URL(AUTH.authorizeEndpoint);
    authorizeUrl.searchParams.set('client_id', AUTH.clientId);
    authorizeUrl.searchParams.set('response_type', 'code');
    authorizeUrl.searchParams.set('redirect_uri', `${HTTP.baseUrl}/callback`);
    authorizeUrl.searchParams.set('response_mode', 'query');
    authorizeUrl.searchParams.set('scope', scopes);
    authorizeUrl.searchParams.set('state', state);
    authorizeUrl.searchParams.set('code_challenge', codeChallenge);
    authorizeUrl.searchParams.set(
      'code_challenge_method',
      codeChallengeMethod || 'S256'
    );
    // Multiple people connect their own mailboxes through the same server —
    // always let the user pick/confirm the account.
    authorizeUrl.searchParams.set('prompt', 'select_account');

    res.redirect(authorizeUrl.toString());
  });

  router.get('/callback', (req, res) => {
    const { code, state, error, error_description: errorDescription } =
      req.query;

    const payload = verifyState(state, HTTP.stateSecret);
    if (!payload) {
      res
        .status(400)
        .send('Invalid or expired OAuth state. Please restart the connection flow.');
      return;
    }

    const target = new URL(payload.ru);
    if (error) {
      target.searchParams.set('error', String(error));
      if (errorDescription) {
        target.searchParams.set('error_description', String(errorDescription));
      }
    } else if (code) {
      target.searchParams.set('code', String(code));
    } else {
      res.status(400).send('Missing authorization code.');
      return;
    }
    if (payload.st) {
      target.searchParams.set('state', payload.st);
    }

    res.redirect(target.toString());
  });

  router.post('/token', async (req, res) => {
    const body = req.body || {};
    const grantType = body.grant_type;

    const upstream = new URLSearchParams();
    upstream.set('client_id', AUTH.clientId);
    if (AUTH.clientSecret) {
      upstream.set('client_secret', AUTH.clientSecret);
    }
    upstream.set('scope', scopes);

    if (grantType === 'authorization_code') {
      if (!body.code) {
        res.status(400).json({
          error: 'invalid_request',
          error_description: 'code is required',
        });
        return;
      }
      upstream.set('grant_type', 'authorization_code');
      upstream.set('code', body.code);
      // Microsoft bound the code to OUR /callback, not to the MCP client's
      // redirect_uri — swap it on the upstream exchange.
      upstream.set('redirect_uri', `${HTTP.baseUrl}/callback`);
      if (body.code_verifier) {
        upstream.set('code_verifier', body.code_verifier);
      }
    } else if (grantType === 'refresh_token') {
      if (!body.refresh_token) {
        res.status(400).json({
          error: 'invalid_request',
          error_description: 'refresh_token is required',
        });
        return;
      }
      upstream.set('grant_type', 'refresh_token');
      upstream.set('refresh_token', body.refresh_token);
    } else {
      res.status(400).json({
        error: 'unsupported_grant_type',
        error_description: `Unsupported grant_type: ${grantType}`,
      });
      return;
    }

    try {
      const response = await fetch(AUTH.tokenEndpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: upstream.toString(),
      });
      const json = await response.json();
      res.status(response.status).set('Cache-Control', 'no-store').json(json);
    } catch (fetchError) {
      console.error(`[OAUTH] Token proxy error: ${fetchError.message}`);
      res.status(502).json({
        error: 'server_error',
        error_description: 'Failed to reach the Microsoft token endpoint',
      });
    }
  });

  return router;
}

module.exports = { oauthRouter, requireGraphBearer };
