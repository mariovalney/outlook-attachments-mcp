# Security Policy

## Reporting a vulnerability

Please **do not** open a public GitHub issue for security vulnerabilities.

Instead, report it privately using one of:

- [GitHub Security Advisories](https://github.com/mariovalney/outlook-attachments-mcp/security/advisories/new) for this repository, or
- Email **mariovalney@gmail.com** with details and, if possible, steps to reproduce.

You should get an acknowledgement within a few days. Please give us a reasonable amount of time to investigate and ship a fix before any public disclosure.

## Supported versions

This project is early-stage and released as a rolling `main`/latest build — there are no maintained older versions. Security fixes land on the default branch; please run the latest build.

## Security model — things worth knowing

This server is a stateless OAuth proxy in front of Microsoft Entra ID (see the [README's "How it works" section](README.md#how-it-works)). A few design points matter for anyone deploying or auditing it:

- **Token pass-through by design.** The bearer token an MCP client (e.g. Claude) sends on `/mcp` *is* the user's Microsoft Graph access token — the same token is used directly for Graph API calls. This is an intentional trade-off to keep the server stateless (no database, no token storage, no volume). It also means anyone who can read that bearer can act on the user's mailbox with the granted scopes, so:
  - **Always deploy behind HTTPS.** Never run this with a plaintext HTTP listener reachable from outside your own machine.
  - Treat the deployment's logs and any request-tracing/proxy layer as sensitive — avoid logging the `Authorization` header.
- **`STATE_SECRET`** signs the OAuth `state` parameter relayed through the Microsoft login round-trip. Set it to a long, random, fixed value in production (`openssl rand -hex 32`). A weak or absent value lets an attacker forge state and potentially redirect the OAuth callback.
- **`OAUTH_ALLOWED_REDIRECT_HOSTS`** restricts which hosts `/authorize` will redirect back to after login. Leaving it unset allows any HTTPS `redirect_uri` (PKCE still protects the authorization code itself, but pinning this in production narrows the attack surface for open-redirect-style abuse).
- **No server-side token storage.** Access/refresh tokens are held by the MCP client, not this server — there's no token database to leak. `/token` is a pure proxy to Microsoft's token endpoint.
- **Per-request isolation.** Each MCP request's Graph token is bound to that request's async context (`auth/request-context.js`) and never persisted or shared across requests/sessions.

If you find a way to bypass bearer validation, leak another user's token, or forge the OAuth state, that's a high-priority report — please use the private channels above.
