/**
 *
 */
// eslint-disable-next-line no-unused-vars
const createPloDocument = () => {
  o.adminMessage.setValue(o.msgInitPlo);
  o.adminMessage.setBackground(o.bgDefault);
  let auditSs;
  let auditSheet;

  try {
    auditSs = SpreadsheetApp.openByUrl(o.adminFileUrl.getValue());
    auditSheet = auditSs.getSheetByName('Worksheet');
  } catch (e) {
    o.adminMessage.setBackground(o.bgWarning);
    o.adminMessage.setValue(o.msgInvalidAudit);
    return;
  }
  let auditData = auditSheet.getDataRange().getValues();
  auditData.splice(0, 1);
  const auditHeaders = auditData
    .splice(0, 1)
    .flat()
    .map((h) => h.toLowerCase().trim());
  auditHeaders[auditHeaders.indexOf('notes')] = ''; // don't copy notes
  const actionColNum = auditHeaders.indexOf('action');
  auditData = auditData.filter(
    (row) => row[actionColNum] === 'On Page Refresh'
  );

  if (!auditData || auditData.length < 1) {
    o.adminMessage.setBackground(o.bgWarning);
    o.adminMessage.setValue(o.msgMissingOpr);
    Logger.log(o.msgMissingUrl);
    return;
  }
  auditData = _transpose(auditData);
  const driveFile = DriveApp.getFileById(auditSs.getId());
  const parentFolder = driveFile.getParents();
  const folder = parentFolder.next();
  const ploFile = _duplicateTemplate(o.ploTemplateId, folder);
  const ploSS = SpreadsheetApp.openById(ploFile.getId());
  _populateTab(auditHeaders, auditData, ploSS.getSheetByName('PLOs'), 2);
  _populateTab(auditHeaders, auditData, ploSS.getSheetByName('Data'), 2);
  o.adminMessage.setValue(o.msgPloComplete);
  o.adminMessage.setBackground(o.bgSuccess);
};
