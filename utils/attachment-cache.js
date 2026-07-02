/**
 * Ephemeral, in-memory store for attachment bytes pending a one-time HTTP
 * download.
 *
 * Some MCP clients only support a fixed allowlist of mimeTypes for
 * embedded `resource` content blocks (confirmed in production: a .docx
 * attachment returned as a resource/blob was rejected with "Resources of
 * type '...wordprocessingml.document' are not currently supported").
 * Rather than embedding bytes in the MCP response, `download` stashes them
 * here and hands back a plain HTTPS link — sidestepping MCP resource-type
 * support entirely, since the client just opens it as a normal file
 * download.
 *
 * In-memory and short-lived by design, matching the server's existing
 * "nothing persisted" posture. Single-instance only: on a multi-replica
 * deployment without sticky sessions, a download link may be served by a
 * different instance than the one that created it and 404. Same caveat as
 * MCP session state — see the README.
 */
const crypto = require('crypto');

const TTL_MS = 5 * 60 * 1000; // 5 minutes

const store = new Map(); // token -> { buffer, filename, mimeType, expiresAt }

/**
 * Registers a buffer for one-time HTTP download.
 * @param {Buffer} buffer - File content
 * @param {string} filename - Original filename
 * @param {string} mimeType - Content type
 * @returns {{ token: string, expiresAt: number }}
 */
function put(buffer, filename, mimeType) {
  const token = crypto.randomBytes(32).toString('base64url');
  const expiresAt = Date.now() + TTL_MS;
  store.set(token, { buffer, filename, mimeType, expiresAt });
  return { token, expiresAt };
}

/**
 * Retrieves a previously stored entry, or null if missing/expired.
 * @param {string} token
 * @returns {{ buffer: Buffer, filename: string, mimeType: string } | null}
 */
function get(token) {
  const entry = store.get(token);
  if (!entry) return null;
  if (entry.expiresAt <= Date.now()) {
    store.delete(token);
    return null;
  }
  return entry;
}

/** Removes all expired entries. Call periodically to bound memory use. */
function pruneExpired() {
  const now = Date.now();
  for (const [token, entry] of store) {
    if (entry.expiresAt <= now) store.delete(token);
  }
}

module.exports = { put, get, pruneExpired, TTL_MS };
