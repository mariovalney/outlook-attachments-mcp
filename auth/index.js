/**
 * Authentication module for the Outlook Attachments MCP server.
 *
 * This server is HTTP-only and multi-user: every MCP request carries the
 * calling user's own Microsoft Graph access token as the OAuth bearer (see
 * oauth/index.js and http-server.js). ensureAuthenticated() simply resolves
 * that per-request token — there is no server-side token storage or
 * device-code flow.
 */
const { getRequestAccessToken } = require('./request-context');

/**
 * Returns the Graph access token bound to the current request.
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If called outside a request context (no bearer bound)
 */
async function ensureAuthenticated() {
  const requestToken = getRequestAccessToken();
  if (!requestToken) {
    throw new Error('Authentication required');
  }
  return requestToken;
}

module.exports = {
  ensureAuthenticated,
};
