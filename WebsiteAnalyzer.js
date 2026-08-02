/**
 * WebsiteAnalyzer.gs
 * -----------------------------------------------------------------------
 * Fetches a business's website (if any) and classifies it into one of:
 *   Excellent / Good / Basic / Outdated / Very Outdated / Broken / No Website
 *
 * Heuristic checks performed on the HTML:
 *   - HTTPS support (does the site redirect to a secure connection?)
 *   - Mobile viewport meta tag (responsive design indicator)
 *   - Stale copyright year (4+ years old)
 *   - Obsolete HTML elements (frameset, marquee, blink)
 *   - Flash references (.swf files)
 *   - Heavy table-based layout without modern CSS classes
 *   - Very small page with no recognizable platform
 *   - Placeholder/broken internal links (href="#")
 *
 * Classification scale (based on flag count):
 *   0 flags → Good or Excellent (Excellent if rich structured data found)
 *   1 flag  → Basic
 *   2-3 flags → Outdated
 *   4+ flags → Very Outdated
 *
 * This is a heuristic, not a certainty — it's a fast, free proxy for
 * "does this site need work," using signals a search engine or a human
 * skimming the HTML would notice. It costs no API calls beyond UrlFetchApp.
 */

/**
 * Main entry point for website analysis.
 *
 * HTTPS handling: Google's stored website URL sometimes uses "http://" even
 * when the live site fully supports (and redirects to) "https://" — so we
 * always attempt the https:// version FIRST, regardless of what scheme was
 * stored, and only fall back to the original URL if that fails. This avoids
 * false "No HTTPS" flags on sites that are actually secure.
 *
 * @param {string} rawUrl - the business's website URL (may be blank)
 * @returns {{status: string, flags: string[], httpCode: number|null, html: string}}
 *          status is one of WEBSITE_STATUS.*, flags lists what was found,
 *          html is the raw page source (reused by EmailFinder to avoid a second fetch)
 */
function analyzeWebsite(rawUrl) {
  if (!rawUrl) {
    return { status: WEBSITE_STATUS.NO_WEBSITE, flags: ['No website field on Google Business Profile'], html: '' };
  }

  const httpsUrl = forceHttpsScheme(rawUrl);
  let response, finalUrl, usedHttps;

  try {
    response = fetchWithRetry(httpsUrl, { followRedirects: true, validateHttpsCertificates: false });
    if (response.getResponseCode() < 400) {
      finalUrl = httpsUrl;
      usedHttps = true;
    } else {
      throw new Error('https attempt returned ' + response.getResponseCode());
    }
  } catch (e) {
    // https:// attempt failed — fall back to whatever scheme was originally stored.
    try {
      const fallbackUrl = normalizeUrl(rawUrl);
      response = fetchWithRetry(fallbackUrl, { followRedirects: true, validateHttpsCertificates: false });
      finalUrl = fallbackUrl;
      usedHttps = /^https:\/\//i.test(fallbackUrl);
    } catch (e2) {
      return { status: WEBSITE_STATUS.BROKEN, flags: ['Site did not respond: ' + e2.message], httpCode: null, html: '' };
    }
  }

  const code = response.getResponseCode();
  if (code >= 400 || code === 0) {
    return { status: WEBSITE_STATUS.BROKEN, flags: ['HTTP ' + code], httpCode: code, html: '' };
  }

  const html = response.getContentText();
  const flags = collectOutdatedFlags(html, finalUrl, usedHttps);
  const status = classifyFromFlags(flags.length, html);

  return { status: status, flags: flags, httpCode: code, html: html };
}

/**
 * Forces a URL to use https://, regardless of its original scheme (or lack of one).
 * @param {string} url - any URL string
 * @returns {string} the URL with https:// scheme
 */
function forceHttpsScheme(url) {
  return url.replace(/^https?:\/\//i, '').replace(/^/, 'https://');
}

/**
 * Runs each individual signal check against the page HTML and returns matched flags.
 * @param {string} html - raw HTML source
 * @param {string} url - the final URL fetched (after redirects)
 * @param {boolean} usedHttps - whether the successful fetch used https://
 * @returns {string[]} array of human-readable flag descriptions
 */
function collectOutdatedFlags(html, url, usedHttps) {
  const lower = html.toLowerCase();
  const flags = [];

  if (!usedHttps) flags.push('No HTTPS (site does not redirect to a secure connection)');
  if (!checkViewportTag(lower)) flags.push('No mobile viewport tag (likely not responsive)');

  const staleYear = checkStaleCopyright(html);
  if (staleYear) flags.push('Copyright year is ' + staleYear);

  if (checkObsoleteHtml(lower)) flags.push('Uses obsolete HTML (frameset/marquee/blink)');
  if (checkFlash(lower)) flags.push('References Flash (.swf)');
  if (checkHeavyTableLayout(lower)) flags.push('Heavy table-based layout, no modern CSS layout classes');
  if (checkTinyPageNoPlatform(html, lower)) flags.push('Very small page with no recognizable modern platform');
  if (checkBrokenInternalLinksHint(lower)) flags.push('Contains placeholder/broken-looking links (href="#")');

  return flags;
}

function checkViewportTag(lowerHtml) {
  return /name=["']viewport["']/.test(lowerHtml);
}

/**
 * Returns the stale year if a copyright year 4+ years old is found, else null.
 * Handles year RANGES like "© 2022 - 2026" or "© 2019-2026" by taking the
 * LAST year in the range (the most recent one), since that's what actually
 * indicates whether the site is being maintained — a range ending in the
 * current year is NOT stale, even if it started years ago.
 */
function checkStaleCopyright(html) {
  // Look for one or two 4-digit years near a © or "copyright" mention.
  const match = html.match(/(?:©|copyright)\D{0,10}(\d{4})(?:\s*[-–—]\s*(\d{4}))?/i);
  if (!match) return null;

  // If a second year exists (a range), use it; otherwise use the single year found.
  const year = match[2] ? parseInt(match[2], 10) : parseInt(match[1], 10);
  const currentYear = new Date().getFullYear();
  return (currentYear - year >= 4) ? year : null;
}

function checkObsoleteHtml(lowerHtml) {
  return /<frameset|<marquee|<blink/.test(lowerHtml);
}

function checkFlash(lowerHtml) {
  return /\.swf["']/.test(lowerHtml);
}

function checkHeavyTableLayout(lowerHtml) {
  const tableCount = (lowerHtml.match(/<table/g) || []).length;
  const hasModernContainer = /<div[^>]*class=["'][^"']*(container|grid|flex)/.test(lowerHtml);
  return tableCount >= 3 && !hasModernContainer;
}

function checkTinyPageNoPlatform(html, lowerHtml) {
  const modernSignals = /wp-content|squarespace|wixsite|shopify|webflow|react|next\.js|__next|tailwind|godaddy/i.test(html);
  return !modernSignals && html.length < 3000;
}

/** Cheap heuristic: many "#" placeholder hrefs suggests an unfinished/neglected site. */
function checkBrokenInternalLinksHint(lowerHtml) {
  const hashLinks = (lowerHtml.match(/href=["']#["']/g) || []).length;
  return hashLinks >= 3;
}

/**
 * Turns a flag count into a final website status label.
 * Also checks for rich structured data signals (schema.org, og:title, JSON-LD)
 * to distinguish Good from Excellent when no flags are present.
 *
 * @param {number} flagCount - number of outdated-signal flags found
 * @param {string} html - raw HTML (checked for rich signals at 0 flags)
 * @returns {string} one of WEBSITE_STATUS.*
 */
function classifyFromFlags(flagCount, html) {
  if (flagCount >= 4) return WEBSITE_STATUS.VERY_OUTDATED;
  if (flagCount >= 2) return WEBSITE_STATUS.OUTDATED;
  if (flagCount === 1) return WEBSITE_STATUS.BASIC;

  // 0 flags: distinguish Good vs Excellent by looking for richer modern signals.
  const richSignals = /schema\.org|og:title|application\/ld\+json/i.test(html);
  return richSignals ? WEBSITE_STATUS.EXCELLENT : WEBSITE_STATUS.GOOD;
}