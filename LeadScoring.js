/**
 * LeadScoring.gs
 * -----------------------------------------------------------------------
 * Turns website status + email presence + Google rating/review signals
 * into a single 0–100 score, and determines whether a lead qualifies.
 *
 * Point values live in Config.js (SCORING_RULES) so they can be tuned
 * without touching this logic. The qualification rule (which statuses
 * count as "qualifying") is also in Config.js (QUALIFYING_WEBSITE_STATUSES).
 */

/**
 * @param {object} params
 *   websiteStatus: one of WEBSITE_STATUS.*
 *   hasEmail: boolean
 *   rating: number|null
 *   reviewCount: number|null
 *   hasRecentReview: boolean (optional signal; defaults false if unknown)
 * @returns {number} score capped at 0-100
 */
function computeLeadScore(params) {
  let score = 0;

  switch (params.websiteStatus) {
    case WEBSITE_STATUS.NO_WEBSITE: score += SCORING_RULES.NO_WEBSITE; break;
    case WEBSITE_STATUS.BROKEN: score += SCORING_RULES.BROKEN; break;
    case WEBSITE_STATUS.VERY_OUTDATED: score += SCORING_RULES.VERY_OUTDATED; break;
    case WEBSITE_STATUS.OUTDATED: score += SCORING_RULES.OUTDATED; break;
    case WEBSITE_STATUS.BASIC: score += SCORING_RULES.BASIC; break;
    default: break; // Good / Excellent contribute 0
  }

  if (params.hasEmail) score += SCORING_RULES.HAS_EMAIL;
  if (params.rating && params.rating > 4.5) score += SCORING_RULES.RATING_ABOVE_4_5;
  if (params.reviewCount && params.reviewCount >= 50) score += SCORING_RULES.REVIEWS_50_PLUS;
  if (params.hasRecentReview) score += SCORING_RULES.RECENT_REVIEW;

  return Math.max(0, Math.min(100, score));
}

/**
 * Decides whether a lead is qualified per the mandatory rule:
 * Public Email is REQUIRED. Website quality does not disqualify a lead.
 * Returns { qualified: boolean, reason: string }
 */
function evaluateQualification(hasEmail, websiteStatus) {
  if (!hasEmail) {
    return { qualified: false, reason: 'No public email found' };
  }
  return { qualified: true, reason: '' };
}
