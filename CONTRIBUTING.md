# Contributing

Thanks for considering a contribution to Outlook Attachments MCP.

## Project structure

The server is organized by Outlook domain, one directory per area of the Microsoft Graph API:

| Path | What it holds |
|---|---|
| `auth/` | Per-request auth context (`request-context.js`) — resolves the caller's Graph bearer token |
| `oauth/` | The OAuth 2.1 proxy in front of Microsoft Entra ID (discovery metadata, `/register`, `/authorize`, `/callback`, `/token`, bearer validation) |
| `email/`, `calendar/`, `contacts/`, `categories/`, `folder/`, `rules/`, `settings/`, `advanced/` | Tool implementations, one module per Outlook feature area |
| `utils/` | Shared helpers: the raw Graph API client (`graph-api.js`), field presets, response formatting, schema coercion |
| `config.js` | All configuration and environment variable resolution |
| `server-factory.js` | Builds the MCP `Server` instance and wires up the tool list |
| `http-server.js` | The HTTP entry point: Express app, Streamable HTTP transport, session handling, mounts the OAuth router |

Most of `email/`, `calendar/`, `contacts/`, `categories/`, `folder/`, `rules/`, `settings/`, `advanced/`, and `utils/` are vendored from [outlook-assistant](https://github.com/littlebearapps/outlook-assistant) (MIT, © Little Bear Apps) — see [LICENSE](LICENSE). If you're changing behavior in those directories, please keep changes focused so they stay easy to compare against upstream.

## Local setup

```bash
npm install
npm run check   # syntax-checks the core entry points
```

Run the server with mock Graph data — no Azure app registration or real Microsoft account needed:

```bash
BASE_URL=http://localhost:3000 npm run test-mode   # USE_TEST_MODE=true node http-server.js, listens on :3000
```

`BASE_URL` is always required (it's used to build the OAuth metadata URLs), even in test mode. `USE_TEST_MODE=true` routes every Graph API call through the mock data in `utils/mock-data.js` instead of hitting `graph.microsoft.com`, and the bearer check accepts any token starting with `test_access_token_`. This is the fastest way to exercise a change end-to-end.

To test against your own real mailbox instead, follow the [Entra ID app registration steps in the README](README.md#1-register-an-app-in-microsoft-entra-id) and run without `USE_TEST_MODE`.

## Testing a change end to end

With the server running in test mode:

```bash
# Health check
curl -s http://localhost:3000/healthz

# Open an MCP session and call a tool
curl -s http://localhost:3000/mcp -X POST \
  -H 'Content-Type: application/json' -H 'Accept: application/json, text/event-stream' \
  -H 'Authorization: Bearer test_access_token_123' \
  -d '{"jsonrpc":"2.0","method":"initialize","id":1,"params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"dev","version":"1"}}}' -i

# Grab the mcp-session-id from the response headers above, then:
curl -s http://localhost:3000/mcp -X POST \
  -H 'Content-Type: application/json' -H 'Accept: application/json, text/event-stream' \
  -H 'Authorization: Bearer test_access_token_123' -H 'mcp-session-id: <id-from-above>' \
  -d '{"jsonrpc":"2.0","method":"tools/call","id":2,"params":{"name":"search-emails","arguments":{"maxResults":1}}}'
```

For Docker changes, rebuild and smoke-test the image before opening a PR:

```bash
docker build -t outlook-attachments-mcp:dev .
docker run --rm -p 3000:3000 \
  -e USE_TEST_MODE=true -e BASE_URL=http://localhost:3000 \
  -e STATE_SECRET=dev-secret -e OUTLOOK_CLIENT_ID=dev-client-id \
  outlook-attachments-mcp:dev
```

## Submitting a change

- Keep pull requests small and focused — one change, one PR.
- Run `npm run check` before opening a PR (there's no CI yet, so this is the only automated safety net).
- If you touch the OAuth proxy (`oauth/`) or the per-request auth context (`auth/request-context.js`), re-read [SECURITY.md](SECURITY.md) — that code path handles live Microsoft Graph tokens.
- Explain the *why* in the PR description, not just the *what*.

## Reporting bugs / requesting features

Open a [GitHub issue](https://github.com/mariovalney/outlook-attachments-mcp/issues). For security issues, see [SECURITY.md](SECURITY.md) instead — please don't file those as public issues.
