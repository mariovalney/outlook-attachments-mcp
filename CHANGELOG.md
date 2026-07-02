# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-07-02

### Added

- Initial release: a multi-user, remote MCP server for Microsoft Outlook / Microsoft 365, deployable anywhere a Dockerfile can run.
- OAuth 2.1 proxy in front of Microsoft Entra ID (`oauth/`) implementing protected-resource metadata (RFC 9728), authorization-server metadata (RFC 8414), a dynamic-client-registration stub (RFC 7591), and `/authorize` → `/callback` → `/token` relaying to Microsoft, with PKCE passed through end-to-end.
- Stateless, per-user authentication: the bearer sent on every MCP request is the caller's own Microsoft Graph access token, bound to that request's context (`auth/request-context.js`) — no server-side token storage.
- Streamable HTTP MCP transport (`http-server.js`) with in-memory session handling and a public `GET /healthz` for platform health checks.
- 21 Outlook tools across email, calendar, contacts, categories, folders, rules, mailbox settings, and (work/school) shared mailboxes / meeting rooms, backed by the Microsoft Graph API.
- `Dockerfile` for cloud-agnostic deployment.
- `USE_TEST_MODE` for exercising the full tool surface against mock data without a real Microsoft account.

### Changed from upstream

Built on top of [outlook-assistant](https://github.com/littlebearapps/outlook-assistant) v3.8.1 (MIT, © Little Bear Apps), with inspiration from [Claude-MCP-Read-Email-Attachments](https://github.com/Zacccck/Claude-MCP-Read-Email-Attachments). The original's local, single-user stdio server and device-code login were replaced with the remote, multi-user, OAuth-authenticated HTTP server described above.
