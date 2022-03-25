const o = {};

(function init() {
  o.invalidConfig = false;
  o.adminSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  o.adminSummary = o.adminSpreadsheet.getSheetByName('Summary');
  o.adminConfig = o.adminSpreadsheet.getSheetByName('Config');
  o.adminFolderUrl = o.adminSummary.getRange('B3');
  o.adminFileUrl = o.adminSummary.getRange('B6');
  o.adminName = o.adminSummary.getRange('B9');
  o.adminMessage = o.adminSummary.getRange('B12');
  o.msgMissingUrl =
    'Please enter a client Google Drive folder URL, and select "Merge Sources."';
  o.msgBadUrl =
    'Please provide a valid folder url: https://drive.google.com/drive/folders/<folder_id>';
  o.msgNoFolder = 'Unable to locate folder with ID: ';
  o.msgUnexpectedFile = 'Unexpected file found in client folder: ';
  o.duplicateFiles = 'Duplicate files found. Count: ';
  o.msgInitKwp =
    'Creating Keyword Prioritization document. Please wait a few moments...';
  o.msgKwpComplete = 'Keyword Prioritization document created.';
  o.msgInitMerge = 'Merging sources. Please wait a few moments...';
  o.msgInitPlo = 'Creating PLO document. Please wait a few moments...';
  o.msgMergeComplete = 'File merge completed successfully';
  o.msgInvalidAudit = 'Please enter a valid URL for the audit file.';
  o.msgMissingKwp = `No file found containing "%%"`;
  o.msgMissingOpr = 'No audit rows with "On Page Refresh" action found.';
  o.msgPloComplete = 'PLO document created successfully';
  o.bgWarning = '#ee9999';
  o.bgDefault = '#efefef';
  o.bgSuccess = '#99ee99';
  o.textDefault = '#333333';
  o.keywordPrioritizationOrganicKeywords = 'organic-keywords-subdomains';
  o.keywordPrioritizationListOverview = 'list-overview';
  o.mergeFileNames =
    'content-event-pages,content-pages,internal-urls,top-pages';
  o.noteContentNrWord = 'Add more fresh and relevant content.';
  o.noteIncomingLinks =
    'Create 1-3 more internal links on another page pointing back to this one.';
  o.noteTitlesLength =
    'Update title tag with a secondary keyword or keyword modifiers up to 70 (max 600 pixels).';
  o.noteMetaDescriptionLength =
    'Update meta description with a secondary keyword, company name, call to action, or more compelling text up to 160  (max 920 pixels).';
  o.noteDoFollow =
    'Add this to backlink target list, use variation of primary keyword target as anchor text, and aim for 1-3 more links.';
  o.adminConfig
    .getDataRange()
    .getValues()
    .forEach(([k, v]) => (o[k.trim()] = v.trim()));
})();

/**
 * Set custom menu in the Admin spreadsheet
 */
// eslint-disable-next-line no-unused-vars
function onOpen() {
  if (!o.invalidConfig) {
    o.adminMessage.setValue(o.msgMissingUrl);
    o.adminMessage.setBackground(o.bgDefault);
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🟢 Content Audit')
      .addItem('Merge sources', 'mergeSources')
      .addItem('Create PLO', 'createPloDocument')
      .addSeparator()
      .addItem('Create Keyword Prioritization', 'createKeywordPrioritization')
      .addToUi();
  }
}

/**
 *
 * @param {*} file
 * @returns
 */
// eslint-disable-next-line no-unused-vars
const _copyAsGoogleSheet = (fileId, folder) => {
  const resource = {
    parents: [folder],
    mimeType: MimeType.GOOGLE_SHEETS,
  };
  return Drive.Files.copy(resource, fileId);
};

/**
 *
 * @param {*} template
 */
// eslint-disable-next-line no-unused-vars
const _duplicateTemplate = (templateId, folder) => {
  const templateFile = DriveApp.getFileById(templateId);
  const outputFile = templateFile.makeCopy(
    templateFile.getName().replace('[TEMPLATE]', `– ${o.adminName.getValue()}`),
    folder
  );
  return outputFile;
};

/**
 * Check the user supplied url of the client's folder
 */
// eslint-disable-next-line no-unused-vars
const _getFolder = () => {
  const url = o.adminFolderUrl.getValue();
  if (!url) {
    // user didn't supply a url
    o.adminMessage.setBackground(o.bgWarning);
    o.adminMessage.setValue(o.msgMissingUrl);
    Logger.log(o.msgMissingUrl);
    return false;
  }
  const id = url.match(/(\w|-)*(?=(\?|\/$|$))/)[0];
  let folder;
  try {
    folder = DriveApp.getFolderById(id);
  } catch (e) {
    // user supplied a url that doesn't point to a folder
    o.adminMessage.setBackground(o.bgWarning);
    o.adminMessage.setValue(o.msgBadUrl);
    return false;
  }
  return folder;
};

/**
 * Populate a target sheet with source data by matching
 * column headers. Target columns with unmatched headers
 * are left empty.
 */
// eslint-disable-next-line no-unused-vars
const _populateTab = (
  sourceHeaders,
  sourceData,
  targetSheet,
  targetHeaderRow = 1
) => {
  const targetHeaders = targetSheet
    .getRange(targetHeaderRow, 1, 1, targetSheet.getLastColumn())
    .getValues()
    .flat()
    .map((h) => h.toLowerCase().trim());
  let targetData = new Array(targetHeaders.length).fill(
    new Array(sourceData[0].length).fill('')
  );
  targetHeaders.forEach((header, i) => {
    const colNum = sourceHeaders.indexOf(header);
    if (colNum > -1) {
      targetData[i] = sourceData[colNum];
    }
  });
  targetData = _transpose(targetData);
  targetSheet
    .getRange(targetHeaderRow + 1, 1, targetData.length, targetData[0].length)
    .setValues(targetData);
  SpreadsheetApp.flush();
};

/**
 * Convert rows to columns and vice versa in a nested array
 * @param {array<array>} arr nested array to transpose
 * @returns {array<array>} the transposed array
 */
// eslint-disable-next-line no-unused-vars
const _transpose = (arr) => {
  return Object.keys(arr[0]).map(function (c) {
    return arr.map(function (r) {
      return r[c];
    });
  });
};
