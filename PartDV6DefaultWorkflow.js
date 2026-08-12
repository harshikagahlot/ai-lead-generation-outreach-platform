/**
 * PartDV6DefaultWorkflow.js
 * -----------------------------------------------------------------------
 * Safe entry point for the outreach workflow.
 *
 * It intentionally does NOT replace or alter the existing lead-generation,
 * qualification, email-validation, spam/placeholder filtering, or Gmail
 * sending architecture.
 *
 * Workflow:
 *   1. Create drafts for newly qualified leads using the existing pipeline.
 *   2. Immediately upgrade those drafts to Part D v6.
 *
 * Part D v6 is the observation-led structure already implemented in
 * EmailOutreachUpgrade.js, including the VASHA signature and inline logo.
 * Gmail drafts are created/updated as drafts only; nothing is sent.
 */
function menuDraftOutreachEmailsUsingPartD() {
  // Keep the existing lead/email gates untouched.
  menuDraftOutreachEmails();

  // Replace the generic draft content with the Part D v6 structure and
  // create/update the corresponding Gmail drafts with the existing HTML/logo.
  menuUpgradeOutreachEmails();
}
