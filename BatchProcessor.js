/**
 * BatchProcessor.gs
 * -----------------------------------------------------------------------
 * Makes lead generation RESUMABLE across multiple short executions instead
 * of one long one, which is what was hitting Apps Script's execution time
 * limit. The workflow now looks like:
 *
 *   1. startNewLeadJob() runs the Places search ONCE, saves the full list
 *      of Place IDs + a "next index" pointer to PropertiesService, then
 *      processes one batch immediately.
 *   2. processBatch() processes leads starting at the saved pointer, for
 *      up to (Batch Size) leads OR (Max Seconds Per Batch) — whichever
 *      comes first — saving progress after EVERY single lead so a crash
 *      or forced stop never loses more than the current lead.
 *   3. If leads remain, a one-time trigger is scheduled to call
 *      processBatch() again automatically in ~1 minute. If everything is
 *      done, the job is marked complete and a summary is shown/logged.
 *
 * This means a 200-business search that used to time out now runs as many
 * short, safe batches, continuing on its own until finished.
 *
 * Functions called from other files:
 *   - searchPlacesText()       → Code.js
 *   - getPlaceDetails()        → Code.js
 *   - buildLeadFromPlace()     → Code.js
 *   - getExistingPlaceIds()    → Sheets.js
 *   - getExistingFingerprints() → Sheets.js
 *   - appendRawRow() etc.      → Sheets.js
 *   - evaluateQualification()  → LeadScoring.js
 *
 * Menu actions defined here:
 *   - menuContinueLeadJob() — manually trigger the next batch
 *   - menuJobStatus() — show progress alert
 *   - menuCancelLeadJob() — cancel and clean up
 */

const JOB_PROPERTY_KEY = 'LEAD_GEN_JOB_STATE';
const TRIGGER_PROPERTY_KEY = 'LEAD_GEN_CONTINUATION_TRIGGER_ID';

// =========================================================================
// JOB STATE (persisted in PropertiesService so it survives between executions)
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
// STARTING A NEW JOB
// =========================================================================

/**
 * Runs the Places search ONCE (this is the only part that still happens in
 * a single shot — it's fast and cheap compared to per-lead processing),
 * saves the resulting Place ID list as a new job, then immediately runs
 * the first batch so the person sees progress right away.
 */
function startNewLeadJob(industry, city, state, maxBusinesses) {
  const ui = SpreadsheetApp.getUi();
  const existing = getJobState();
  if (existing && existing.status === 'running') {
    ui.alert(
      'A lead generation job is already in progress: "' + existing.industry + '" in ' +
      existing.city + ', ' + existing.state + ' (' + existing.nextIndex + '/' + existing.placeIds.length + ' processed).\n\n' +
      'Use "Continue Lead Generation Job" to keep going, or "Cancel Current Job" first if you want to start something new.'
    );
    return;
  }

  const location = city + ', ' + state;
  let places;
  try {
    places = searchPlacesText(industry, location, maxBusinesses);
  } catch (e) {
    ui.alert('Places search failed: ' + e.message);
    return;
  }

  if (!places.length) {
    ui.alert('No results found for "' + industry + '" in "' + location + '".');
    return;
  }

  const job = {
    industry: industry,
    city: city,
    state: state,
    placeIds: places.map(p => p.id),
    nextIndex: 0,
    stats: { checked: 0, qualified: 0, rejected: 0, errors: 0 },
    status: 'running',
    startedAt: new Date().toISOString()
  };
  saveJobState(job);

  ui.alert(
    'Found ' + job.placeIds.length + ' businesses for "' + industry + '" in ' + location + '.\n\n' +
    'Processing in small batches to avoid timeouts — running the first batch now. ' +
    'If more remain after this, it will continue automatically in the background (about once a minute) ' +
    'until done. You can also click "Continue Lead Generation Job" any time to speed it along manually.'
  );

  processBatch();
}

// =========================================================================
// PROCESSING ONE BATCH (the resumable worker — this is what a trigger calls)
// =========================================================================

/**
 * Processes leads starting from the job's saved pointer, stopping at
 * whichever limit is hit first: batch size (count) or max seconds
 * (wall-clock time budget). Saves progress after EVERY lead, so partial
 * progress is never lost even if this execution is killed mid-batch.
 */
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

  // Read both dedup structures from Raw_Data in a single sheet read
  // instead of two separate reads (getExistingPlaceIds + getExistingFingerprints).
  const dedupData = getExistingDedupData();
  const existingIds = dedupData.placeIds;
  const existingFingerprints = dedupData.fingerprints;
  const minRating = Number(settings['Minimum Rating']) || 0;
  const minReviews = Number(settings['Minimum Reviews']) || 0;

  let processedThisBatch = 0;

  while (job.nextIndex < job.placeIds.length) {
    const elapsedSeconds = (new Date().getTime() - startTime) / 1000;
    if (elapsedSeconds > maxSeconds) break;   // time budget guard — leaves safety margin under the platform limit
    if (processedThisBatch >= batchSize) break; // count guard

    const placeId = job.placeIds[job.nextIndex];
    job.nextIndex++;          // advance FIRST so a mid-processing crash never reprocesses this same id
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
        appendRejectedRow(details.displayName ? details.displayName.text : '', 'Duplicate (matching name/phone/address)', placeId);
        saveJobState(job);
        continue;
      }
      existingFingerprints.add(fingerprint);

      if ((details.rating || 0) < minRating || (details.userRatingCount || 0) < minReviews) {
        job.stats.rejected++;
        appendRejectedRow(details.displayName ? details.displayName.text : '', 'Below minimum rating/review threshold', placeId);
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
      // One failed website/API request never stops the batch — log and move on.
      job.stats.errors++;
      Logger.log('Error processing place ' + placeId + ': ' + e.message);
    }

    saveJobState(job); // persist progress after EVERY lead, per the resumability requirement
  }

  const execSeconds = (new Date().getTime() - startTime) / 1000;
  const done = job.nextIndex >= job.placeIds.length;

  if (done) {
    job.status = 'complete';
    saveJobState(job);
    clearContinuationTrigger();
    appendLogRow(
      'Generate Leads (batched)', job.stats.checked, job.stats.qualified, job.stats.rejected,
      job.stats.errors, execSeconds, job.industry + ' in ' + job.city + ', ' + job.state + ' — COMPLETE'
    );
    notifyJobComplete(job);
  } else {
    scheduleContinuation();
    appendLogRow(
      'Generate Leads (batch)', processedThisBatch, job.stats.qualified, job.stats.rejected,
      job.stats.errors, execSeconds,
      job.industry + ' in ' + job.city + ', ' + job.state + ' — batch done, ' +
      (job.placeIds.length - job.nextIndex) + ' remaining'
    );
  }
}

// =========================================================================
// AUTO-CONTINUATION VIA TIME-DRIVEN TRIGGER
// =========================================================================

/**
 * Schedules processBatch() to run again shortly, deleting any previous
 * continuation trigger first so they never pile up. Requires the
 * script.scriptapp authorization scope — Apps Script will prompt for this
 * the first time a trigger is created; this is expected.
 */
function scheduleContinuation() {
  clearContinuationTrigger();
  const trigger = ScriptApp.newTrigger('processBatch')
    .timeBased()
    .after(60 * 1000) // 1 minute is the practical minimum for one-time triggers
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

/** Shows a completion alert if run interactively; falls back to logging if run via a trigger (no UI context). */
function notifyJobComplete(job) {
  try {
    SpreadsheetApp.getUi().alert(
      'Lead generation complete: "' + job.industry + '" in ' + job.city + ', ' + job.state + '\n\n' +
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
  if (!job) { ui.alert('No lead generation job on record.'); return; }
  ui.alert(
    'Job: ' + job.industry + ' in ' + job.city + ', ' + job.state + '\n' +
    'Status: ' + job.status + '\n' +
    'Progress: ' + job.nextIndex + ' / ' + job.placeIds.length + '\n' +
    'Checked: ' + job.stats.checked + ' | Qualified: ' + job.stats.qualified +
    ' | Rejected: ' + job.stats.rejected + ' | Errors: ' + job.stats.errors
  );
}

function menuCancelLeadJob() {
  clearContinuationTrigger();
  clearJobState();
  SpreadsheetApp.getUi().alert('Current lead generation job cancelled. Any leads already saved to Raw_Data/Qualified_Leads remain untouched.');
}