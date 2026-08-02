/**
 * EmailFinder.gs
 * -----------------------------------------------------------------------
 * Attempts to discover a publicly-listed email address for a business by
 * checking the homepage (already fetched by WebsiteAnalyzer, reused here
 * to avoid a redundant HTTP request) and a handful of common sub-pages.
 *
 * Strategy:
 *   1. Check homepage HTML for mailto: links, then plain-text email matches
 *   2. Fetch common sub-pages (/contact, /about, /privacy, etc.) in parallel
 *      via UrlFetchApp.fetchAll() for speed
 *   3. Filter out known placeholder/template emails (GoDaddy fillers, noreply, etc.)
 *
 * Never invents an email — if nothing is found, returns blank.
 */

const EMAIL_CANDIDATE_PATHS = [
  '', // homepage itself (reuses html passed in)
  '/contact', '/contact-us', '/contactus',
  '/about', '/about-us', '/aboutus',
  '/privacy', '/privacy-policy',
  '/terms', '/terms-of-service', '/terms-and-conditions'
];

const EMAIL_REGEX = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;

/**
 * Known placeholder/template emails left behind by website builders when a
 * business never replaces the demo contact info. Matching any of these
 * patterns means "not a real contact" — never accept them as a lead's email,
 * even though they're technically valid, real-looking email addresses.
 */
const PLACEHOLDER_EMAIL_PATTERNS = [
  /^filler@/i,                    // GoDaddy Website Builder demo contact
  /^test@/i,
  /^example@/i,
  /^demo@/i,
  /^sample@/i,
  /^placeholder@/i,
  /^youremail@/i,
  /^email@example\./i,
  /@example\.(com|org|net)$/i,
  /@domain\.(com|org|net)$/i,
  /@yourdomain\.(com|org|net)$/i,
  /@yoursite\.(com|org|net)$/i,
  /^name@/i,
  /^info@example/i,
  /^admin@localhost/i,
  /@godaddy\.com$/i,               // GoDaddy's own domain showing up as a "contact" is always a template leftover
  /@wixpress\.com$/i,              // Wix internal/demo addresses
  /@sentry\./i,                    // error-tracking service addresses (e.g., sentry.io, sentry-next.wixpress.com)
  /noreply@/i,
  /no-reply@/i,
  /donotreply@/i,
  
  // --- Invalid / Malformed matches (e.g., image files that look like emails) ---
  /\.(jpg|jpeg|png|gif|svg|webp)$/i,  // Image filenames
  /\b\d{2,4}x\d{2,4}\b/i,             // Image dimensions (e.g., 300x300)
  /@2x/i,                             // Retina image suffixes
  /\//,                               // File paths
  /\\/                                // Windows file paths
];

/**
 * Returns true if an email matches a known placeholder/template pattern.
 * @param {string} email - email address to check
 * @returns {boolean} true if the email is a placeholder
 */
function isPlaceholderEmail(email) {
  return PLACEHOLDER_EMAIL_PATTERNS.some(pattern => pattern.test(email));
}

/**
 * Tries to find a public email for a business.
 * @param {string} websiteUrl - normalized site URL (may be blank if no site)
 * @param {string} homepageHtml - HTML already fetched for the homepage, if any
 * @returns {{email: string, sourceUrl: string}}
 */
function findPublicEmail(websiteUrl, homepageHtml) {
  if (!websiteUrl) return { email: '', sourceUrl: '' };

  const origin = getOrigin(normalizeUrl(websiteUrl));

  // Check homepage HTML we already have before spending more fetches.
  const homepageEmail = extractEmail(homepageHtml);
  if (homepageEmail) return { email: homepageEmail, sourceUrl: normalizeUrl(websiteUrl) };

  // Fetch all candidate sub-pages in ONE parallel batch instead of one at a
  // time — this is what was taking up to ~11 sequential round-trips (many
  // seconds) per lead. UrlFetchApp.fetchAll() dispatches every request
  // together and Google's infrastructure fetches them concurrently, so this
  // typically finishes in roughly the time of the SLOWEST single request,
  // not the sum of all of them.
  const subPaths = EMAIL_CANDIDATE_PATHS.slice(1); // skip '' (homepage, already checked above)
  const requests = subPaths.map(path => ({
    url: origin + path,
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: false
  }));

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    return { email: '', sourceUrl: '' }; // batch fetch itself failed (rare) — treat as no email found
  }

  for (let i = 0; i < responses.length; i++) {
    try {
      if (responses[i].getResponseCode() >= 400) continue;
      const email = extractEmail(responses[i].getContentText());
      if (email) return { email: email, sourceUrl: requests[i].url };
    } catch (e) {
      continue;
    }
  }

  return { email: '', sourceUrl: '' };
}

/**
 * Extracts a real (non-placeholder) email from HTML. Checks ALL mailto:
 * links first (preferring these over plain text, since mailto is a more
 * deliberate signal), skipping any that match a known placeholder pattern,
 * then falls back to scanning all plain-text email matches the same way.
 * Returns '' if every match found is a placeholder or none exist at all.
 */
function extractEmail(html) {
  if (!html) return '';

  const mailtoRegex = /mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/gi;
  let match;
  while ((match = mailtoRegex.exec(html)) !== null) {
    if (!isPlaceholderEmail(match[1])) return match[1];
  }

  const plainRegex = new RegExp(EMAIL_REGEX.source, 'gi');
  while ((match = plainRegex.exec(html)) !== null) {
    if (!isPlaceholderEmail(match[0])) return match[0];
  }

  return ''; // every match found (if any) was a placeholder — correctly treat as "no real email"
}