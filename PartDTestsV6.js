/**
 * PartDTestsV6.js
 * Smoke tests for the Part D v6 outreach structure.
 * Kept separate because the older PartDTests.js still checks the v5 logo size.
 */
function testPartDV6() {
  if (isValidOutreachEmail('6399988845karki@gmail.com')) throw new Error('FAIL: phone-number-style email was accepted.');
  if (isValidOutreachEmail('filler@godaddy.com')) throw new Error('FAIL: placeholder email was accepted.');
  if (!isValidOutreachEmail('alex@exampleproperty.com')) throw new Error('FAIL: normal named email was rejected.');

  const lead = {
    name: 'Example Property Management',
    industry: 'Property Management',
    owner: 'Alex',
    email: 'alex@exampleproperty.com',
    websiteStatus: WEBSITE_STATUS.OUTDATED,
    notes: 'No HTTPS; viewport issue',
    readinessScore: 60,
    rating: 4.5,
    reviewCount: 50
  };

  if (!hasConcreteObservation(lead)) throw new Error('FAIL: concrete observation gate.');
  if (!evaluateQualification(lead.email, lead.websiteStatus).qualified) throw new Error('FAIL: valid email did not qualify.');
  if (evaluateQualification('6399988845karki@gmail.com', lead.websiteStatus).qualified) throw new Error('FAIL: suspicious email qualified.');

  const email = buildPartDEmail_(lead);
  if (!email.subject || !email.body || !email.htmlBody) throw new Error('FAIL: email fields missing.');
  if (email.body.indexOf('Harshika') === -1) throw new Error('FAIL: sender name missing.');
  if (email.body.indexOf('VASHA Technologies') === -1) throw new Error('FAIL: VASHA signature missing.');
  if (email.body.indexOf('Example Property Management') === -1) throw new Error('FAIL: business name missing.');
  if (email.hypothesis.indexOf('owner enquiry') === -1) throw new Error('FAIL: property-management hypothesis missing.');
  if (email.capability.indexOf('custom systems') === -1) throw new Error('FAIL: VASHA capability missing from generated email.');
  if (email.htmlBody.indexOf('cid:vashaLogo') === -1) throw new Error('FAIL: inline logo reference missing.');
  if (email.htmlBody.indexOf('width="110"') === -1) throw new Error('FAIL: logo width was not reduced.');

  const bodyWords = email.body.trim().split(/\s+/).length;
  if (bodyWords > 160) throw new Error('FAIL: Part D v6 email is too long (' + bodyWords + ' words).');

  const blob = getVashaLogoBlob_();
  if (blob.getContentType() !== 'image/jpeg') throw new Error('FAIL: logo blob MIME type.');

  Logger.log('PASS: Part D v6 smoke tests.');
  return 'PASS';
}
