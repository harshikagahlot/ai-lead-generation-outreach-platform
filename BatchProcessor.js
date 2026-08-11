/**
 * BatchProcessor.gs
 * -----------------------------------------------------------------------
 * Resumable lead generation worker.
 *
 * Important design change:
 * - Google Places Text Search is now fetched ONE PAGE at a time.
 * - The page token is saved in PropertiesService.
 * - Lead processing is also saved after every lead.
 * - A time-driven trigger continues the job automatically.
 *
 * This prevents a large Places search from timing out BEFORE the resumable
 * worker starts. Google Places Text Search (New) currently allows up to 20
 * results per page and up to 60 results across pages for one text query.
 */

const JOB_PROPERTY_KEY = 'LEAD_GEN_JOB_STATE';
const TRIGGER_PROPERTY_KEY = 'LEAD_GEN_CONTINUATION_TRIGGER_ID';
const PLACES_PAGE_SIZE = 20;
const PLACES_MAX_RESULTS_PER_QUERY = 60;

// =========================================================================
// JOB STATE
// =========================================================================

function getJobState() {
  const raw = PropertiesService.getScriptProperties().getProperty(JOB_PROPERTY_KEY);
  return raw ? JSON.parse(raw) : null;
}

function saveJobState(job) {
  PropertiesService.getScriptProperties().setProperty(JOB_PROPERTY_KEY, JSON.stringify(job));
}

function clearJobState() {
  PropertiesService.getScriptProperties().deleteProperty(JOB_PROPERTY_KEY);
}

// =========================================================================
// PLACES SEARCH — ONE PAGE AT A TIME
// =========================================================================

/**
 * Fetches exactly one Places Text Search page.
 * The nextPageToken is persisted by the job so another execution can fetch
 * the next page later instead of collecting the whole search up front.
 */
function fetchPlacesSearchPage(industry, location, pageToken, pageSize) {
  const apiKey = getPlacesApiKey();
  const url = 'https://places.googleapis.com/v1/places:searchText';
  const size = Math.min(PLACES_PAGE_SIZE, Math.max(1, Number(pageSize) || PLACES_PAGE_SIZE));

  const body = {
    textQuery: industry + ' in ' + location,
    pageSize: size
  };

  if (pageToken) body.pageToken = pageToken;

  const response = fetchWithRetry(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Goog-Api-Key': apiKey,
      'X-Goog-FieldMask': 'places.id,places.displayName,nextPageToken'
    },
    payload: JSON.stringify(body)
  });

  if (!response) throw new Error('Places API returned no response.');

  const json = JSON.parse(response.getContentText());
  if (json.error) {
    throw new Error('Places API error: ' + json.error.message);
  }

  return {
    places: json.places || [],
    nextPageToken: json.nextPageToken || null
  };
}

// =========================================================================
// STARTING A NEW JOB
// =========================================================================

function startNewLeadJob(industry, city, state, maxBusinesses) {
  const ui = SpreadsheetApp.getUi();
  const existing = getJobState();

  if (existing && existing.status === 'running') {
    ui.alert(
      'A lead generation job is already in progress: "' + existing.industry + '" in ' +
      existing.city + ', ' + existing.state + '.\n\n' +
      'Progress: ' + existing.nextIndex + ' / ' + existing.placeIds.length +
      ' currently discovered.\n\n' +
      'Use "Continue Lead Generation Job" to keep going, or "Cancel Current Job" first if you want to start something new.'
    );
    return;
  }

  const location = city + ', ' + state;
  const requested = parseInt(maxBusinesses, 10) || 20;
  const targetMax = Math.min(Math.max(requested, 1), PLACES_MAX_RESULTS_PER_QUERY);

  let firstPage;
  try {
    firstPage = fetchPlacesSearchPage(
      industry,
      location,
      null,
      Math.min(PLACES_PAGE_SIZE, targetMax)
    );
  } catch (e) {
    ui.alert('Places search failed: ' + e.message);
    return;
  }

  if (!firstPage.places.length) {
    ui.alert('No results found for "' + industry + '" in "' + location + '".');
    return;
  }

  const job = {
    industry: industry,
    city: city,
    state: state,
    location: location,
    targetMaxBusinesses: targetMax,
    requestedMaxBusinesses: requested,
    placeIds: firstPage.places.map(p => p.id).filter(Boolean),
    nextIndex: 0,
    nextPageToken: firstPage.nextPageToken,
    pagesFetched: 1,
    stats: { checked: 0, qualified: 0, rejected: 0, errors: 0 },
    status: 'running',
    startedAt: new Date().toISOString(),
    lastPageFetchedAt: new Date().getTime()
  };

  saveJobState(job);

  let message =
    'Found the first ' + job.placeIds.length + ' businesses for "' + industry + '" in ' + location + '.\n\n' +
    'Lead processing is now resumable. More Places pages will be fetched only when needed, so the initial search cannot consume the whole execution time.';

  if (requested > PLACES_MAX_RESULTS_PER_QUERY) {
    message +=
      '\n\nGoogle Places Text Search currently returns at most 60 results across pages for one text query. ' +
      'This job is therefore capped at 60 for this query. For more than 60, we will need additional search queries/locations rather than pretending one query can return 200.';
  }

  message +=
    '\n\nThe first processing batch will start now. Remaining work will continue automatically about once a minute.';

  ui.alert(message);
  processBatch();
}

// =========================================================================
// PROCESSING ONE RESUMABLE BATCH
// =========================================================================

function processBatch() {
  const startTime = new Date().getTime();
  const settings = getSettings();
  const maxSeconds = Number(settings['Max Seconds Per Batch']) || 280;
  const batchSize = Number(settings['Batch Size (leads per execution)']) || 15;

  const job = getJobState();
  if (!job || job.status !== 'running') {
    Logger.log('processBatch: no active job to process.');
    return;
  }

  const dedupData = getExistingDedupData();
  const existingIds = dedupData.placeIds;
  const existingFingerprints = dedupData.fingerprints;
  const minRating = Number(settings['Minimum Rating']) || 0;
  const minReviews = Number(settings['Minimum Reviews']) || 0;

  let processedThisBatch = 0;

  while (processedThisBatch < batchSize) {
    const elapsedSeconds = (new Date().getTime() - startTime) / 1000;
    if (elapsedSeconds > maxSeconds) break;

    // ---------------------------------------------------------------
    // If the currently discovered page is exhausted, fetch the next
    // Places page only if the job still needs more results.
    // ---------------------------------------------------------------
    if (job.nextIndex >= job.placeIds.length) {
      const reachedTarget = job.placeIds.length >= job.targetMaxBusinesses;
      const noMorePages = !job.nextPageToken;

      if (reachedTarget || noMorePages) break;

      // Leave a safety margin so the page request itself cannot push the
      // execution into the hard Apps Script timeout.
      if (elapsedSeconds > maxSeconds - 20) break;

      // Google requires a short activation delay before a newly issued
      // page token can be used. If this execution is already well separated
      // from the previous page request, no meaningful delay is needed.
      const ageMs = new Date().getTime() - Number(job.lastPageFetchedAt || 0);
      if (ageMs < 2500) Utilities.sleep(2500 - ageMs);

      const remaining = job.targetMaxBusinesses - job.placeIds.length;
      const pageSize = Math.min(PLACES_PAGE_SIZE, remaining);
      let page;

      try {
        page = fetchPlacesSearchPage(
          job.industry,
          job.location,
          job.nextPageToken,
          pageSize
        );
      } catch (e) {
        job.stats.errors++;
        Logger.log('Places page fetch failed: ' + e.message);
        saveJobState(job);
        break;
      }

      const knownIds = new Set(job.placeIds);
      let added = 0;
      page.places.forEach(place => {
        if (!place || !place.id || knownIds.has(place.id)) return;
        job.placeIds.push(place.id);
        knownIds.add(place.id);
        added++;
      });

      job.nextPageToken = page.nextPageToken;
      job.pagesFetched = Number(job.pagesFetched || 0) + 1;
      job.lastPageFetchedAt = new Date().getTime();
      saveJobState(job);

      Logger.log(
        'Fetched Places page ' + job.pagesFetched + ': added ' + added +
        ' new place(s); discovered ' + job.placeIds.length + '/' + job.targetMaxBusinesses + '.'
      );

      if (!added && !job.nextPageToken) break;
      if (!added && job.nextPageToken) continue;
    }

    const placeId = job.placeIds[job.nextIndex];
    job.nextIndex++; // advance first so a killed execution does not repeat this ID
    processedThisBatch++;

    try {
      if (existingIds.has(placeId)) {
        job.stats.rejected++;
        appendRejectedRow('(duplicate)', 'Duplicate (Place ID already in Raw_Data)', placeId);
        saveJobState(job);
        continue;
      }

      const details = getPlaceDetails(placeId);
      if (!details) {
        job.stats.errors++;
        saveJobState(job);
        continue;
      }

      const fingerprint = buildFingerprint(
        details.displayName ? details.displayName.text : '',
        details.nationalPhoneNumber,
        details.formattedAddress
      );

      if (existingFingerprints.has(fingerprint)) {
        job.stats.rejected++;
        appendRejectedRow(
          details.displayName ? details.displayName.text : '',
          'Duplicate (matching name/phone/address)',
          placeId
        );
        saveJobState(job);
        continue;
      }

      existingFingerprints.add(fingerprint);

      if ((details.rating || 0) < minRating || (details.userRatingCount || 0) < minReviews) {
        job.stats.rejected++;
        appendRejectedRow(
          details.displayName ? details.displayName.text : '',
          'Below minimum rating/review threshold',
          placeId
        );
        saveJobState(job);
        continue;
      }

      const lead = buildLeadFromPlace(details, job.industry);
      appendRawRow(lead);
      job.stats.checked++;

      const qualification = evaluateQualification(!!lead.email, lead.websiteStatus);
      if (qualification.qualified) {
        job.stats.qualified++;
        appendQualifiedRow(lead);
      } else {
        job.stats.rejected++;
        appendRejectedRow(lead.name, qualification.reason, lead.placeId);
      }

      existingIds.add(placeId);

    } catch (e) {
      job.stats.errors++;
      Logger.log('Error processing place ' + placeId + ': ' + e.message);
    }

    // Persist after EVERY lead.
    saveJobState(job);
  }

  const execSeconds = (new Date().getTime() - startTime) / 1000;
  const discoveredDone = job.nextIndex >= job.placeIds.length;
  const targetReached = job.placeIds.length >= job.targetMaxBusinesses;
  const noMorePages = !job.nextPageToken;
  const done = discoveredDone && (targetReached || noMorePages);

  if (done) {
    job.status = 'complete';
    saveJobState(job);
    clearContinuationTrigger();

    appendLogRow(
      'Generate Leads (batched)',
      job.stats.checked,
      job.stats.qualified,
      job.stats.rejected,
      job.stats.errors,
      execSeconds,
      job.industry + ' in ' + job.city + ', ' + job.state +
        ' — COMPLETE (' + job.placeIds.length + ' places discovered)'
    );

    notifyJobComplete(job);
  } else {
    scheduleContinuation();

    const remainingKnown = Math.max(0, job.placeIds.length - job.nextIndex);
    const morePages = !!job.nextPageToken && !targetReached;

    appendLogRow(
      'Generate Leads (batch)',
      processedThisBatch,
      job.stats.qualified,
      job.stats.rejected,
      job.stats.errors,
      execSeconds,
      job.industry + ' in ' + job.city + ', ' + job.state +
        ' — batch done, ' + remainingKnown + ' known remaining' +
        (morePages ? ' + more Places pages' : '')
    );
  }
}

// =========================================================================
// AUTO-CONTINUATION VIA TIME-DRIVEN TRIGGER
// =========================================================================

function scheduleContinuation() {
  clearContinuationTrigger();
  const trigger = ScriptApp.newTrigger('processBatch')
    .timeBased()
    .after(60 * 1000)
    .create();
  PropertiesService.getScriptProperties().setProperty(TRIGGER_PROPERTY_KEY, trigger.getUniqueId());
}

function clearContinuationTrigger() {
  const props = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(TRIGGER_PROPERTY_KEY);
  if (!triggerId) return;

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
  });

  props.deleteProperty(TRIGGER_PROPERTY_KEY);
}

function notifyJobComplete(job) {
  try {
    SpreadsheetApp.getUi().alert(
      'Lead generation complete: "' + job.industry + '" in ' + job.city + ', ' + job.state + '\n\n' +
      'Places discovered: ' + job.placeIds.length + '\n' +
      'Checked: ' + job.stats.checked + '\n' +
      'Qualified: ' + job.stats.qualified + '\n' +
      'Rejected: ' + job.stats.rejected + '\n' +
      'Errors: ' + job.stats.errors
    );
  } catch (e) {
    Logger.log('Job complete (ran via trigger, no UI available): ' + JSON.stringify(job.stats));
  }
}

// =========================================================================
// MENU ACTIONS: manual continue / status / cancel
// =========================================================================

function menuContinueLeadJob() {
  const job = getJobState();
  if (!job || job.status !== 'running') {
    SpreadsheetApp.getUi().alert('No lead generation job currently in progress.');
    return;
  }
  processBatch();
}

function menuJobStatus() {
  const job = getJobState();
  const ui = SpreadsheetApp.getUi();
  if (!job) {
    ui.alert('No lead generation job on record.');
    return;
  }

  ui.alert(
    'Job: ' + job.industry + ' in ' + job.city + ', ' + job.state + '\n' +
    'Status: ' + job.status + '\n' +
    'Processed/discovered: ' + job.nextIndex + ' / ' + job.placeIds.length + '\n' +
    'Target: ' + job.targetMaxBusinesses + '\n' +
    'Places pages fetched: ' + (job.pagesFetched || 1) + '\n' +
    'More Places pages: ' + (job.nextPageToken ? 'Yes' : 'No') + '\n' +
    'Checked: ' + job.stats.checked + ' | Qualified: ' + job.stats.qualified +
    ' | Rejected: ' + job.stats.rejected + ' | Errors: ' + job.stats.errors
  );
}

function menuCancelLeadJob() {
  clearContinuationTrigger();
  clearJobState();
  SpreadsheetApp.getUi().alert(
    'Current lead generation job cancelled. Any leads already saved to Raw_Data/Qualified_Leads remain untouched.'
  );
}
