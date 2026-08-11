/**
 * EmailFinder.gs
 * -----------------------------------------------------------------------
 * Discovers publicly-listed business emails from the website homepage and
 * common contact pages. Never invents an email.
 *
 * Important outreach rule:
 * Technical email syntax is NOT enough. We also reject known placeholders,
 * system addresses, file/path artifacts, and suspicious phone-number-style
 * addresses so they cannot enter the qualified outreach pipeline.
 */

const EMAIL_CANDIDATE_PATHS = [
  '',
  '/contact', '/contact-us', '/contactus',
  '/about', '/about-us', '/aboutus',
  '/privacy', '/privacy-policy',
  '/terms', '/terms-of-service', '/terms-and-conditions'
];

const EMAIL_REGEX = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
const GENERIC_EMAIL_PATTERN = /^(info|contact|office|hello|support|sales|admin|team|hi|mail|reception|bookings|service|billing|help|hr)@/i;

function classifyEmail(email) {
  if (!email) return '';
  return GENERIC_EMAIL_PATTERN.test(email) ? 'Generic Inbox' : 'Named Person';
}

/**
 * Returns true only when an address is acceptable for outreach.
 * This is intentionally stricter than RFC-style syntax validation.
 */
function isValidOutreachEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const trimmed = email.trim();
  if (!trimmed) return false;
  if (!EMAIL_REGEX.test(trimmed)) return false;
  if (isPlaceholderEmail(trimmed)) return false;
  if (isSuspiciousOutreachEmail(trimmed)) return false;
  return true;
}

const PLACEHOLDER_EMAIL_PATTERNS = [
  /^filler@/i,
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
  /@godaddy\.com$/i,
  /@wixpress\.com$/i,
  /@sentry\./i,
  /noreply@/i,
  /no-reply@/i,
  /donotreply@/i,
  /\.(jpg|jpeg|png|gif|svg|webp)$/i,
  /\b\d{2,4}x\d{2,4}\b/i,
  /@2x/i,
  /\//,
  /\\/
];

/**
 * Suspicious addresses that are technically valid but look like scraped
 * phone numbers or machine-generated contact strings rather than a useful
 * business mailbox. We deliberately keep this narrow to avoid rejecting
 * legitimate addresses containing normal digits.
 */
const SUSPICIOUS_OUTREACH_EMAIL_PATTERNS = [
  /^\d{10}$/i,
  /^\d{10}[a-z][a-z0-9._-]*$/i
];

function isPlaceholderEmail(email) {
  const value = String(email || '').trim();
  return PLACEHOLDER_EMAIL_PATTERNS.some(pattern => pattern.test(value)) || isSuspiciousOutreachEmail(value);
}

function isSuspiciousOutreachEmail(email) {
  const value = String(email || '').trim();
  const at = value.lastIndexOf('@');
  if (at <= 0) return true;
  const localPart = value.slice(0, at);
  return SUSPICIOUS_OUTREACH_EMAIL_PATTERNS.some(pattern => pattern.test(localPart));
}

function findPublicEmail(websiteUrl, homepageHtml) {
  if (!websiteUrl) return { email: '', sourceUrl: '', type: '' };

  const origin = getOrigin(normalizeUrl(websiteUrl));
  const homepageEmail = extractEmail(homepageHtml);
  if (homepageEmail) {
    return {
      email: homepageEmail,
      sourceUrl: normalizeUrl(websiteUrl),
      type: classifyEmail(homepageEmail)
    };
  }

  const subPaths = EMAIL_CANDIDATE_PATHS.slice(1);
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
    return { email: '', sourceUrl: '', type: '' };
  }

  for (let i = 0; i < responses.length; i++) {
    try {
      if (responses[i].getResponseCode() >= 400) continue;
      const email = extractEmail(responses[i].getContentText());
      if (email) {
        return {
          email: email,
          sourceUrl: requests[i].url,
          type: classifyEmail(email)
        };
      }
    } catch (e) {
      continue;
    }
  }

  return { email: '', sourceUrl: '', type: '' };
}

function extractEmail(html) {
  if (!html) return '';

  const mailtoRegex = /mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/gi;
  let match;
  while ((match = mailtoRegex.exec(html)) !== null) {
    const candidate = match[1].trim();
    if (isValidOutreachEmail(candidate)) return candidate;
  }

  const plainRegex = new RegExp(EMAIL_REGEX.source, 'gi');
  while ((match = plainRegex.exec(html)) !== null) {
    const candidate = match[0].trim();
    if (isValidOutreachEmail(candidate)) return candidate;
  }

  return '';
}
