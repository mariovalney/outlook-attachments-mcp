/**
 * Per-request authentication context for HTTP (multi-user) mode.
 *
 * In HTTP mode every MCP request arrives with the caller's own Microsoft
 * Graph access token as the OAuth bearer. The HTTP layer wraps request
 * handling in runWithAccessToken() so that ensureAuthenticated() — the
 * single choke point every tool goes through — resolves the token of the
 * user making the request instead of the single-user token store used in
 * stdio mode.
 */
const { AsyncLocalStorage } = require('async_hooks');

const storage = new AsyncLocalStorage();

/**
 * Runs fn with the given access token bound to the async context.
 * @param {string} accessToken - Microsoft Graph access token for this request
 * @param {Function} fn - Function (sync or async) to run within the context
 * @returns {*} - Return value of fn
 */
function runWithAccessToken(accessToken, fn) {
  return storage.run({ accessToken }, fn);
}

/**
 * Returns the access token bound to the current async context, if any.
 * @returns {string|null}
 */
function getRequestAccessToken() {
  const store = storage.getStore();
  return store ? store.accessToken : null;
}

module.exports = { runWithAccessToken, getRequestAccessToken };
