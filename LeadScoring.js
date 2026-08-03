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
 *   emailType: string (optional, 'Named Person' or 'Generic Inbox')
 *   readinessScore: number (optional)
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
  
  if (params.emailType === 'Named Person') score += (SCORING_RULES.NAMED_PERSON || 15);
  if (params.readinessScore >= 70) score += (SCORING_RULES.HIGH_READINESS || 20);

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

/**
 * Computes the Outreach Readiness Score.
 * @param {object} lead 
 * @param {string} emailType 
 * @param {string} recommendedChannel 
 * @returns {{score: number, notes: string}}
 */
function computeOutreachReadiness(lead, emailType, recommendedChannel) {
  let score = 0;
  let notes = [];

  // Channel & Email
  if (recommendedChannel === 'Email') {
    score += 40;
    notes.push('Good channel fit (Email)');
  } else if (recommendedChannel === 'Phone') {
    score += 20;
    notes.push('Better suited for Phone outreach');
  } else {
    notes.push('No primary outreach channel identified');
  }

  if (emailType === 'Named Person') {
    score += 20;
    notes.push('Named decision-maker found');
  } else if (emailType === 'Generic Inbox') {
    score += 5;
    notes.push('Generic inbox (lower conversion)');
  }

  // Website
  if (lead.websiteStatus === WEBSITE_STATUS.BROKEN || lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE) {
    notes.push('Website is broken or missing');
  } else {
    score += 15;
    notes.push('Website is active');
  }

  // Observations (Proxy for personalization potential)
  const rating = Number(lead.rating) || 0;
  const reviews = Number(lead.reviewCount) || 0;
  let hasGenuineDetail = false;
  if (rating >= 4.0 && reviews >= 15) {
    hasGenuineDetail = true;
    score += 25;
    notes.push('Strong rating/reviews for personalization');
  } else if (lead.notes && lead.notes.length > 5) {
    hasGenuineDetail = true;
    score += 15;
    notes.push('Specific website observations available');
  }

  if (!hasGenuineDetail) {
    notes.push('Lacks specific personalization details');
  }

  score = Math.max(0, Math.min(100, score));

  return {
    score: score,
    notes: notes.join(' | ')
  };
}
