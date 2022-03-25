/**
 * Merge the client's input files
 * into an output spreadsheet
 */
// eslint-disable-next-line no-unused-vars
function mergeSources() {
  _removeSpreadsheetUrl();
  o.adminMessage.setValue(o.msgInitMerge);
  o.adminMessage.setBackground(o.bgDefault);
  const folder = _getFolder();
  if (!folder) {
    Logger.log('No folder. Terminating.');
    return;
  }
  const filesObj = _getFolderFilesInfo(folder);
  if (!filesObj) {
    Logger.log('No files. Terminating.');
    return;
  }
  const url = _mergeFiles(filesObj, folder);
  o.adminMessage.setValue(o.msgMergeComplete);
  o.adminMessage.setBackground(o.bgSuccess);
  // o.adminName.setValue('');
  o.adminFileUrl.setValue(url);
}

/**
 * Get names and ids of folder files
 */
const _getFolderFilesInfo = (folder) => {
  const fileNames = o.mergeFileNames.split(',');
  const files = folder.getFiles();
  const filesInfo = {};
  while (files.hasNext()) {
    const file = files.next();
    const idx = fileNames.findIndex((fileName) =>
      file.getName().includes(fileName)
    );

    if (idx > -1) {
      filesInfo[fileNames[idx]] = file.getId();
    }
  }
  return Object.keys(filesInfo).length > 0 ? filesInfo : null;
};

/**
 * Merge client files into the templated spreadsheet
 */
const _mergeFiles = (filesObj, folder) => {
  const outputFile = _duplicateTemplate(o.auditTemplateId, folder);
  const templateSS = SpreadsheetApp.openById(outputFile.getId());
  const templateSheet = templateSS.getSheetByName('Worksheet');
  const templateHeaders = templateSheet
    .getRange(2, 1, 1, templateSheet.getLastColumn())
    .getValues()
    .flat()
    .map((h) => h.toLowerCase());
  let categoryColNum = templateHeaders.indexOf('category');
  let notesColNum = templateHeaders.indexOf('notes');

  // get top-pages
  const resource = {
    parents: [folder],
    mimeType: MimeType.GOOGLE_SHEETS,
  };
  let newFile = Drive.Files.copy(resource, filesObj['top-pages']);
  let tempSS = SpreadsheetApp.openById(newFile.getId());
  let tempSheet = tempSS.getActiveSheet();
  const urlStr = tempSheet.getRange('A2').getValue();
  const re =
    /^((http[s]?|ftp):\/)?\/?([^:\/\s]+)((\/\w+)*\/)([\w\-\.]+[^#?\s]+)(.*)?(#[\w\-]+)?$/;
  let prefix = urlStr.replace(re, '$1/$3');
  if (prefix.endsWith('/')) {
    prefix = prefix.slice(0, prefix.length - 1);
  }
  const tRows = tempSheet.getDataRange().getValues();
  const tHeaders = tRows
    .splice(0, 1)
    .flat()
    .map((h) => h.toLowerCase());
  const tUrlColNum = tHeaders.indexOf('url');

  //newFile.setTrashed(true)

  // get internal-urls
  newFile = Drive.Files.copy(resource, filesObj['internal-urls']);
  tempSS = SpreadsheetApp.openById(newFile.getId());
  tempSheet = tempSS.getActiveSheet();
  const iRows = tempSheet.getDataRange().getValues();
  const iHeaders = iRows
    .splice(0, 1)
    .flat()
    .map((h) => h.toLowerCase());
  const hUrlColNum = iHeaders.indexOf('url');
  //newFile.setTrashed(true)

  // get content-pages
  tempSS = SpreadsheetApp.openById(filesObj['content-pages']);
  tempSheet = tempSS.getActiveSheet();
  let cRows = tempSheet.getDataRange().getValues();
  cRows.splice(0, 6);
  const cHeaders = cRows
    .splice(0, 1)
    .flat()
    .map((h) => h.toLowerCase());
  const cPageColNum = cHeaders.indexOf('page');
  let nonUrlIdx = cRows.findIndex((row) => !row[cPageColNum].includes('/'));
  // remove non-url rows
  cRows.splice(nonUrlIdx);
  // standardize full url
  cRows.forEach((row) => (row[0] = prefix + row[0]));

  // get content-event-pages
  tempSS = SpreadsheetApp.openById(filesObj['content-event-pages']);
  tempSheet = tempSS.getActiveSheet();
  let eRows = tempSheet.getDataRange().getValues();
  eRows.splice(0, 6);
  const eHeaders = eRows
    .splice(0, 1)
    .flat()
    .map((h) => h.toLowerCase());
  const ePageColNum = eHeaders.indexOf('page');
  nonUrlIdx = eRows.findIndex((row) => !row[ePageColNum].includes('/'));
  // remove non-url rows
  eRows.splice(nonUrlIdx);
  // standardize full url
  eRows.forEach((row) => (row[0] = prefix + row[0]));

  // prepare the output array
  const output = [];

  // map position of template header positions
  // to their respective source header positions
  const headerObj = {
    c: [{ sourcePos: cPageColNum, targetPos: 0 }],
    e: [],
    i: [],
    t: [],
  };
  for (let i = 1; i < templateHeaders.length; i++) {
    const h = templateHeaders[i];
    let idx = cHeaders.indexOf(h);
    if (idx > -1) {
      headerObj.c.push({ sourcePos: idx, targetPos: i });
      continue;
    }
    idx = eHeaders.indexOf(h);
    if (idx > -1) {
      headerObj.e.push({ sourcePos: idx, targetPos: i });
      continue;
    }
    idx = tHeaders.indexOf(h);
    if (idx > -1) {
      headerObj.t.push({ sourcePos: idx, targetPos: i });
      continue;
    }
    idx = iHeaders.indexOf(h);
    if (idx > -1) {
      headerObj.i.push({ sourcePos: idx, targetPos: i });
    }
  }

  // iterate through cRows
  // create an empty output row
  // iterate through template headers
  // look up header in object
  // get value and push to empty row
  // push row to output
  cRows.forEach((cRow) => {
    let row = new Array(templateHeaders.length).fill('');
    const cUrl = cRow[cPageColNum];
    for (v of Object.values(headerObj.c)) {
      row[v.targetPos] = cRow[v.sourcePos];
    }
    let xi = eRows.findIndex((eRow) => eRow[ePageColNum] == cUrl);
    if (xi > -1) {
      for (v of Object.values(headerObj.e)) {
        row[v.targetPos] = eRows[xi][v.sourcePos];
      }
    }
    xi = tRows.findIndex((tRow) => tRow[tUrlColNum] == cUrl);
    if (xi > -1) {
      for (v of Object.values(headerObj.t)) {
        row[v.targetPos] = tRows[xi][v.sourcePos];
      }
    }
    xi = iRows.findIndex((iRow) => iRow[hUrlColNum] == cUrl);
    if (xi > -1) {
      for (v of Object.values(headerObj.i)) {
        row[v.targetPos] = iRows[xi][v.sourcePos];
      }
    }
    row[categoryColNum] = _getCategory(cUrl);
    row[notesColNum] = _getNotes(templateHeaders, row);
    output.push(row);
  });
  templateSS
    .getSheetByName('Worksheet')
    .getRange(3, 1, output.length, output[0].length)
    .setValues(output);
  return outputFile.getUrl();
};

/**
 * Clear the message URL
 */
function _removeSpreadsheetUrl() {
  const style = SpreadsheetApp.newTextStyle()
    .setUnderline(false)
    .setForegroundColor(o.textDefault)
    .build();
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(o.msgMissingUrl)
    .setLinkUrl(null)
    .setTextStyle(style)
    .build();
  o.adminMessage.setRichTextValue(richText);
}

/**
 *
 */
function _getCategory(url) {
  const path = url
    .slice(url.indexOf('/', url.indexOf('//') + 2) + 1)
    .trim()
    .replace(/(\s|_)/, '');
  if (
    [
      'blog',
      'podcast',
      'episode',
      'how',
      'what',
      'mean',
      'with',
      'can',
      'strategies',
      'strategy',
      'tips',
      'when',
      'why',
      'blogs',
      'resource',
      'guide',
      'ebook',
      'whitepapers',
      'results',
      'casestudy',
      'casestudies',
    ].some((pattern) => path.includes(pattern))
  ) {
    return 'Resource/Guide';
  }
  if (
    ['tag', 'category', 'categories'].some((pattern) => path.includes(pattern))
  ) {
    return 'Blog Category/Tag';
  }
  if (
    [
      'product',
      'usecase',
      'demo',
      'platform',
      'solution',
      'service',
      'pricing',
      'integration',
      'industry',
      'industries',
      'feature',
    ].some((pattern) => path.includes(pattern))
  ) {
    return 'Solutions Page';
  }
  if (
    [
      'company',
      'terms',
      'privacypolicy',
      'contact',
      'career',
      'about',
      'press',
      'team',
      'whatwedo',
    ].some((pattern) => path.includes(pattern))
  ) {
    return 'About/Company';
  }
  if (path.length < 2) {
    return 'Homepage';
  }
  return '';
}

/**
 *
 * @param {*} row
 * @returns
 */
function _getNotes(templateHeaders, row) {
  const note = [];
  const contentNrWord = row[templateHeaders.indexOf('contentnrword')];
  const incomingLinks = row[templateHeaders.indexOf('incominglinks')];
  const titlesLength = row[templateHeaders.indexOf('titleslength')];
  const metaDescriptionLength =
    row[templateHeaders.indexOf('metadescriptionlength')];
  const doFollow = row[templateHeaders.indexOf('dofollow')];

  if (contentNrWord < 400) {
    note.push(o.noteContentNrWord);
  }
  if (incomingLinks < 3) {
    note.push(o.noteIncomingLinks);
  }
  if (titlesLength < 40) {
    note.push(o.noteTitlesLength);
  }
  if (metaDescriptionLength < 100) {
    note.push(o.noteMetaDescriptionLength);
  }
  if (doFollow < 3) {
    note.push(o.noteDoFollow);
  }
  return note.join(' ');
}
