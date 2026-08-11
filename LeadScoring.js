/**
 * LeadScoring.gs
 * -----------------------------------------------------------------------
 * Lead score + qualification + outreach readiness logic.
 *
 * Qualification is intentionally strict on email quality: a non-empty email
 * is NOT enough. isValidOutreachEmail() must approve it.
 */

function computeLeadScore(params) {
  let score = 0;

  switch (params.websiteStatus) {
    case WEBSITE_STATUS.NO_WEBSITE: score += SCORING_RULES.NO_WEBSITE; break;
    case WEBSITE_STATUS.BROKEN: score += SCORING_RULES.BROKEN; break;
    case WEBSITE_STATUS.VERY_OUTDATED: score += SCORING_RULES.VERY_OUTDATED; break;
    case WEBSITE_STATUS.OUTDATED: score += SCORING_RULES.OUTDATED; break;
    case WEBSITE_STATUS.BASIC: score += SCORING_RULES.BASIC; break;
    default: break;
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
 * Public email is required AND must pass the outreach-quality validator.
 */
function evaluateQualification(email, websiteStatus) {
  if (!isValidOutreachEmail(email)) {
    return { qualified: false, reason: email ? 'Invalid or suspicious public email' : 'No public email found' };
  }
  return { qualified: true, reason: '' };
}

function hasConcreteObservation(lead) {
  if (!lead) return false;

  if (lead.notes && typeof lead.notes === 'string' && lead.notes.trim().length > 0) {
    const flags = lead.notes.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
    const specificTechnicalRegex = /http \d|did not respond|placeholder|href="#"|no https|viewport|flash|obsolete html|table-based|very small page|copyright year/;
    for (const flag of flags) {
      if (specificTechnicalRegex.test(flag)) return true;
    }
  }

  const rating = Number(lead.rating) || 0;
  const reviews = Number(lead.reviewCount) || 0;
  if ((rating >= 4.5 && reviews >= 50) || (rating >= 4.0 && reviews >= 15)) return true;

  return false;
}

function computeOutreachReadiness(lead, emailType, recommendedChannel) {
  let score = 0;
  let notes = [];

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

  if (lead.websiteStatus === WEBSITE_STATUS.BROKEN || lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE) {
    notes.push('Website is broken or missing');
  } else {
    score += 15;
    notes.push('Website is active');
  }

  const concreteObs = hasConcreteObservation(lead);
  if (concreteObs) {
    score += 20;
    notes.push('Concrete observation available for personalization');
  } else {
    notes.push('Lacks concrete observation for personalization');
  }

  score = Math.max(0, Math.min(100, score));

  return {
    score: score,
    hasConcreteObservation: concreteObs,
    notes: notes.join(' | ')
  };
}
