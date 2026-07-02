#!/usr/bin/env node
/**
 * Outlook Attachments MCP — stdio entry point (local, single-user)
 *
 * A Model Context Protocol server that provides access to
 * Microsoft Outlook through the Microsoft Graph API.
 *
 * This mode preserves the original outlook-assistant behaviour: device-code
 * authentication via the `auth` tool with tokens persisted on disk.
 * For the multi-user HTTP mode (Claude custom connector), see http-server.js.
 */
const {
  StdioServerTransport,
} = require('@modelcontextprotocol/sdk/server/stdio.js');
const config = require('./config');
const { buildTools, createServer } = require('./server-factory');

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// F-1 / F-48: warn at startup when safety belts are unset. Mirrors the
// warning surfaced by `auth action=about`. Visible to operators reading
// stderr; AI clients reading the JSON-RPC stream are unaffected.
if (
  !process.env.OUTLOOK_MAX_EMAILS_PER_SESSION &&
  !process.env.OUTLOOK_ALLOWED_RECIPIENTS &&
  !config.USE_TEST_MODE
) {
  console.error(
    '⚠ Safety belts not configured. Consider setting OUTLOOK_MAX_EMAILS_PER_SESSION and OUTLOOK_ALLOWED_RECIPIENTS in your .mcp.json env block for safer AI-assisted sending. See `auth action=about` for details.'
  );
}

const TOOLS = buildTools({ includeAuthTools: true });
const server = createServer(TOOLS);

// Make the script executable
process.on('SIGTERM', () => {
  console.error('SIGTERM received but staying alive');
});

// Start the server
const transport = new StdioServerTransport();
server
  .connect(transport)
  .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
  .catch((error) => {
    console.error(`Connection error: ${error.message}`);
    process.exit(1);
  });
