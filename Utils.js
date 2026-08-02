/**
 * Utils.gs
 * -----------------------------------------------------------------------
 * Shared helper functions used across the entire project:
 *   - fetchWithRetry()  — HTTP fetch with automatic retry on 5xx/network errors
 *   - normalizePhone()  — digits-only phone for dedup comparison
 *   - normalizeAddress() — lowercase, alphanumeric-only address for dedup
 *   - normalizeName()   — lowercase, alphanumeric-only business name for dedup
 *   - normalizeUrl()    — ensures a URL has an https:// scheme
 *   - getOrigin()       — extracts scheme + host from a URL
 *   - now()             — consistent timestamp for logging
 *
 * Keeping these in one place avoids duplicated logic across files.
 */

/**
 * Wraps UrlFetchApp.fetch with retry-on-failure logic.
 * Retries on network errors and 5xx responses; does NOT retry on 4xx,
 * since those are unlikely to succeed on retry. Uses simple linear backoff.
 *
 * @param {string} url - the URL to fetch
 * @param {Object} [options] - UrlFetchApp.fetch options (method, headers, payload, etc.)
 * @param {number} [maxRetries] - override retry count (defaults to Settings sheet value, then 2)
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse} the HTTP response
 * @throws {Error} if all retry attempts fail with network errors
 */
function fetchWithRetry(url, options, maxRetries) {
  options = options || {};
  options.muteHttpExceptions = true;
  const retries = typeof maxRetries === 'number' ? maxRetries : (getSettings()['Request Timeout Retries'] || 2);

  let lastError = null;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      if (code >= 500 && attempt < retries) {
        Utilities.sleep(500 * (attempt + 1)); // simple backoff
        continue;
      }
      return response;
    } catch (e) {
      lastError = e;
      if (attempt < retries) {
        Utilities.sleep(500 * (attempt + 1));
        continue;
      }
    }
  }
  if (lastError) throw lastError;
  return null;
}

/**
 * Normalizes a phone number to digits-only for dedup comparison.
 * @param {string|number} phone - raw phone value
 * @returns {string} digits only (e.g. "(555) 123-4567" → "5551234567")
 */
function normalizePhone(phone) {
  return (phone || '').toString().replace(/\D/g, '');
}

/**
 * Normalizes an address for loose dedup comparison (lowercase, alphanumeric only).
 * @param {string} address - raw address string
 * @returns {string} normalized address
 */
function normalizeAddress(address) {
  return (address || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Normalizes a business name for loose dedup comparison (lowercase, alphanumeric only).
 * @param {string} name - raw business name
 * @returns {string} normalized name
 */
function normalizeName(name) {
  return (name || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Ensures a URL has a scheme; defaults to https:// if none present.
 * @param {string} url - raw URL (may lack scheme)
 * @returns {string} URL guaranteed to start with https:// or http://
 */
function normalizeUrl(url) {
  if (!url) return url;
  if (!/^https?:\/\//i.test(url)) return 'https://' + url;
  return url;
}


/**
 * Safely extracts the origin (scheme + host) from a URL string.
 * Example: "https://example.com/about" → "https://example.com"
 * @param {string} url - full URL
 * @returns {string} origin portion, or the original input if parsing fails
 */
function getOrigin(url) {
  try {
    const match = url.match(/^(https?:\/\/[^\/]+)/i);
    return match ? match[1] : url;
  } catch (e) {
    return url;
  }
}

/** Returns the current timestamp as a Date object (for consistent logging across files). */
function now() {
  return new Date();
}
