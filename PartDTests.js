/**
 * PartDTests.js
 * Small deterministic smoke tests for the Part D email layer.
 */
function testPartD() {
  const lead = {
    name: 'Example Property Management',
    industry: 'Property Management',
    owner: 'Alex',
    email: 'alex@exampleproperty.com',
    websiteStatus: WEBSITE_STATUS.OUTDATED,
    notes: 'No HTTPS; viewport issue',
    readinessScore: 60
  };

  if (!hasConcreteObservation(lead)) throw new Error('FAIL: concrete observation gate.');
  const email = buildPartDEmail_(lead);

  if (!email.subject || !email.body || !email.htmlBody) throw new Error('FAIL: email fields missing.');
  if (email.body.indexOf('Harshika') === -1) throw new Error('FAIL: sender name missing.');
  if (email.body.indexOf('VASHA Technologies') === -1) throw new Error('FAIL: VASHA signature missing.');
  if (email.body.indexOf('Example Property Management') === -1) throw new Error('FAIL: business name missing.');
  if (email.hypothesis.indexOf('owner enquiry') === -1) throw new Error('FAIL: property-management hypothesis missing.');
  if (email.htmlBody.indexOf('cid:vashaLogo') === -1) throw new Error('FAIL: inline logo reference missing.');

  const blob = getVashaLogoBlob_();
  if (blob.getContentType() !== 'image/jpeg') throw new Error('FAIL: logo blob MIME type.');

  Logger.log('PASS: Part D smoke tests.');
  return 'PASS';
}
