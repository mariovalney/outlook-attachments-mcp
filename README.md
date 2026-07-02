# Outlook Attachments MCP

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Node.js >= 18.18.0](https://img.shields.io/badge/node-%3E%3D18.18.0-brightgreen.svg)](package.json)
[![Docker](https://img.shields.io/badge/deploy-docker-2496ED.svg?logo=docker&logoColor=white)](Dockerfile)

A **multi-user remote MCP server** for Microsoft Outlook / Microsoft 365. Add it to Claude (claude.ai) as a custom connector, sign in with your Microsoft account, and let Claude work with your email, attachments, calendar, contacts, folders, rules, and mailbox settings — 21 tools backed by the Microsoft Graph API.

Runs anywhere a Dockerfile runs: any cloud platform or PaaS that builds container images, or plain `docker run` on your own server.

Based on the excellent [outlook-assistant](https://github.com/littlebearapps/outlook-assistant) by Little Bear Apps (MIT), with inspiration from [Claude-MCP-Read-Email-Attachments](https://github.com/Zacccck/Claude-MCP-Read-Email-Attachments). This fork replaces the original's local, single-user stdio server and device-code login with a cloud-ready, multi-user remote MCP server authenticated via OAuth.

## How it works

```
Claude (per user) ──OAuth login──▶ Microsoft Entra ID
      │                                   │
      │  bearer = user's Graph token      ▼
      └────────▶  this server  ──────▶ Microsoft Graph (user's mailbox)
```

- **OAuth with Microsoft**: the server exposes standard OAuth discovery endpoints (RFC 8414/9728) and proxies the authorization flow to Microsoft Entra ID. When a user connects the Claude connector, they log in with their own Microsoft account.
- **The login token *is* the Graph token**: the access token issued by Microsoft is what Claude sends as the bearer on every MCP request, and it is used directly for Graph calls on that user's mailbox.
- **Claude manages tokens**: access/refresh tokens are held by the MCP client. Refresh happens through the server's `/token` endpoint, which simply proxies to Microsoft. **Nothing is stored server-side** — no database, no volume; the server is stateless and each request is isolated to the calling user.
- **PKCE end-to-end**: the client's PKCE challenge is passed through to Microsoft, which binds it to the authorization code.

> **Security note**: this design intentionally passes the user's Microsoft Graph token through the MCP client (the "token pass-through" pattern). It keeps the server stateless and dependency-free, at the cost of the MCP client holding a Graph-scoped token. Deploy behind HTTPS only, and consider `OAUTH_ALLOWED_REDIRECT_HOSTS` to pin redirect targets.

## Setup

### 1. Register an app in Microsoft Entra ID

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → **New registration**.
2. **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts" (or restrict as you prefer — match `OUTLOOK_AUTH_AUDIENCE`).
3. **Redirect URI**: platform **Web**, value `https://<your-host>/callback`.
4. Under **API permissions**, add these **delegated** Microsoft Graph permissions:
   `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read`, `Contacts.ReadWrite`, `People.Read`, `MailboxSettings.ReadWrite`, `offline_access`.
5. Copy the **Application (client) ID** → `OUTLOOK_CLIENT_ID`.
6. Either create a **client secret** (Certificates & secrets) → `OUTLOOK_CLIENT_SECRET`, **or** enable **Allow public client flows** and leave the secret empty (PKCE-only).

### 2. Deploy with Docker

```bash
docker build -t outlook-attachments-mcp .

docker run -d --name outlook-mcp -p 3000:3000 \
  -e BASE_URL=https://outlook-mcp.example.com \
  -e OUTLOOK_CLIENT_ID=<client-id> \
  -e OUTLOOK_CLIENT_SECRET=<secret-or-empty> \
  -e STATE_SECRET=$(openssl rand -hex 32) \
  outlook-attachments-mcp
```

**On any cloud container platform**: point it at this repository (it builds the `Dockerfile`), set the environment variables below, expose port `3000`, and put HTTPS in front (most platforms do this automatically). The container is stateless — no volumes needed — and `GET /healthz` is available for health checks.

| Variable | Required | Description |
|---|---|---|
| `BASE_URL` | ✅ | Public HTTPS URL of the server (e.g. `https://outlook-mcp.example.com`) |
| `OUTLOOK_CLIENT_ID` | ✅ | Entra ID application (client) ID |
| `STATE_SECRET` | ✅ | Long random string (signs the OAuth state) |
| `OUTLOOK_CLIENT_SECRET` | — | Only for confidential (Web) app registrations |
| `OUTLOOK_AUTH_AUDIENCE` | — | `common` (default), `consumers`, `organizations`, or a tenant GUID |
| `OAUTH_ALLOWED_REDIRECT_HOSTS` | — | Comma-separated allowlist of OAuth redirect hosts (e.g. `claude.ai,claude.com`) |
| `PORT` | — | Listen port (default `3000`) |
| `OUTLOOK_MAX_EMAILS_PER_SESSION` | — | Safety belt: max emails sent per session |
| `OUTLOOK_ALLOWED_RECIPIENTS` | — | Safety belt: allowed recipient domains/addresses |
| `OUTLOOK_DEFAULT_TIMEZONE` | — | IANA timezone for calendar tools |

### 3. Connect from Claude

1. In Claude, go to **Settings → Connectors → Add custom connector**.
2. Enter the URL: `https://<your-host>/mcp`.
3. Claude discovers the OAuth endpoints automatically. Click **Connect** and sign in with your Microsoft account.
4. Done — every user who adds the connector signs in with their own account and only sees their own mailbox.

## Tools (21)

| Area | Tools |
|---|---|
| Email | `search-emails`, `read-email`, `send-email`, `draft`, `update-email`, `attachments`, `export`, `get-mail-tips` |
| Calendar | `list-events`, `create-event`, `manage-event` |
| Contacts | `manage-contact`, `search-people` |
| Categories | `manage-category`, `apply-category`, `manage-focused-inbox` |
| Folders & rules | `folders`, `manage-rules` |
| Settings | `mailbox-settings` |
| Advanced (work/school) | `access-shared-mailbox`, `find-meeting-rooms` |

## Endpoints

| Endpoint | Description |
|---|---|
| `POST/GET/DELETE /mcp` | MCP Streamable HTTP endpoint (bearer required) |
| `GET /healthz` | Health check (public) |
| `GET /.well-known/oauth-protected-resource` | RFC 9728 resource metadata |
| `GET /.well-known/oauth-authorization-server` | RFC 8414 AS metadata |
| `POST /register` | Dynamic client registration (RFC 7591 stub) |
| `GET /authorize` → `GET /callback` → `POST /token` | OAuth proxy to Microsoft Entra ID |

## Limitations

- MCP sessions are kept in memory — when scaling to multiple replicas, use sticky sessions.
- Shared mailboxes and meeting rooms require work/school accounts and extra Graph scopes (see `config.js`).
- Attachments are listed/downloaded via Graph; there is no server-side document parsing (PDF/OCR etc.).

## Contributing

Bug reports, feature requests, and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) for local setup, the test-mode workflow (no Azure account needed), and how to submit a change.

## Security

Found a vulnerability? Please **don't** open a public issue — see [SECURITY.md](SECURITY.md) for how to report it privately, and for the security-relevant design notes (token pass-through, `STATE_SECRET`, redirect allowlisting).

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

## License

[MIT](LICENSE). Derived from [outlook-assistant](https://github.com/littlebearapps/outlook-assistant) © Little Bear Apps, also MIT.
