/**
 *
 */
// eslint-disable-next-line no-unused-vars
const createKeywordPrioritization = () => {
  o.adminMessage.setValue(o.msgInitKwp);
  o.adminMessage.setBackground(o.bgDefault);
  // get the client folder to find the keyword input xlsx file
  const folder = _getFolder();
  // get the two input xlsx files
  let organicKeywordsId;
  let listOverviewId;

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().includes(o.keywordPrioritizationOrganicKeywords)) {
      organicKeywordsId = file.getId();
    }
    if (file.getName().includes(o.keywordPrioritizationListOverview)) {
      listOverviewId = file.getId();
    }
  }
  if (!organicKeywordsId) {
    o.adminMessage.setValue(
      o.msgMissingKwp.replace('%%', o.keywordPrioritizationOrganicKeywords)
    );
    o.adminMessage.setBackground(o.bgWarning);
    return;
  }
  if (!listOverviewId) {
    o.adminMessage.setValue(
      o.msgMissingKwp.replace('%%', o.keywordPrioritizationListOverview)
    );
    o.adminMessage.setBackground(o.bgWarning);
    return;
  }

  // duplicate the Keyword Prioritization template
  const target = _duplicateTemplate(o.keywordTemplateId, folder);
  const targetSheet = SpreadsheetApp.openById(target.getId()).getSheetByName(
    'Worksheet'
  );

  let organicKeywordsData = _prepSourceData(
    organicKeywordsId,
    folder,
    targetSheet
  );
  let listOverviewData = _prepSourceData(listOverviewId, folder, targetSheet);
  const organicKeywords = organicKeywordsData.map((row) => row[1]).flat();
  listOverviewData = listOverviewData.filter(
    (row) => !organicKeywords.includes(row[1])
  );
  let targetData = organicKeywordsData.concat(listOverviewData);

  targetSheet
    .getRange(2, 1, targetData.length, targetData[0].length)
    .setValues(targetData);
  _applyFormulas(targetSheet);
  // Sort by keyword prioritization, highest to lowest
  // since the column has a formula, sort must happen in the sheet directly
  const targetHeaders = targetSheet
    .getRange(1, 1, 1, targetSheet.getLastColumn())
    .getValues()
    .flat()
    .map((x) => x.toLowerCase().trim());
  const kwpCol = targetHeaders.findIndex((h) => h == 'priority score') + 1;
  targetSheet.sort(kwpCol, false);
  SpreadsheetApp.flush();
  o.adminMessage.setValue(o.msgKwpComplete);
  o.adminMessage.setBackground(o.bgSuccess);
};

const _prepSourceData = (xlsxId, folder, targetSheet) => {
  // copy the keyword input xlsx as Google Sheet
  // and extract its data
  const source = _copyAsGoogleSheet(xlsxId, folder);
  let sourceData = SpreadsheetApp.openById(source.getId())
    .getActiveSheet()
    .getDataRange()
    .getValues();

  const sourceHeaders = sourceData
    .splice(0, 1)
    .flat()
    .map((h) => {
      h = h.trim().toLowerCase();
      switch (h) {
        case 'current position':
          h = 'position';
          break;
        case 'kd':
          h = 'difficulty';
          break;
        case 'current url':
          h = 'url';
          break;
      }
      return h;
    });

  const targetHeaders = targetSheet
    .getRange(1, 1, 1, targetSheet.getLastColumn())
    .getValues()
    .flat()
    .map((h) => h.toLowerCase().trim());

  sourceData = _transpose(sourceData);
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
  return targetData;
};

const _applyFormulas = (sheet) => {
  // locate the three columns to fill
  // with formulas
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()
    .flat()
    .map((x) => x.trim().toLowerCase());
  let psI = -1;
  let scI = -1;
  let siI = -1;
  headers.forEach((header, i) => {
    if (header == 'priority score') {
      psI = i + 1;
    }
    if (header == 'suggested content') {
      scI = i + 1;
    }
    if (header == 'search intent') {
      siI = i + 1;
    }
  });
  // create arrays of formulas
  const psF = [...Array(sheet.getLastRow() - 1)].map((_, i) => {
    const j = i + 2;
    return [
      `=if(OR(E${j}="Transactional",E${j}="Commercial"),5,0)+(I${j}/15)+if(and(F${j}<20,F${j}>1),10*(1/F${j}),0)+2*if(and(H${j}<80,H${j}>10),
      (100-H${j})/100,-1)+if((F${j}>1),(G${j}/1000),0)-if(A${j},10,0)`,
    ];
  });
  const scF = [...Array(sheet.getLastRow() - 1)].map((_, i) => {
    const j = i + 2;
    return [
      `=if(REGEXMATCH(B${j},"advantages|benefits|can|do|lastest|example|how|ideas|learn|mean|Meaning|means|near|study|system|tactic|That|tips|trend|what|when|where|which|why|with"),"Blog Post",if(REGEXMATCH(B${j},"framework|guide|strategy|Best|cheapest|companies|comparison|firms|platforms|products|providers|review|leading|solutions|specialists|Top|vs|tutorial"),"Guide", if(REGEXMATCH(B${j},"agency|company|consultant|consulting|firm|hire|platform|product|provider|reviews|saas|service|solution|specialist|strategist|Buy|cheap|cost|download|free|near me|nearby|price|pricing|purchase|trial|software|services|enterprise
    "),"Landing Page", if(REGEXMATCH(B${j},"Template|ebook|e book|tool"),"Resource/Tool","Blog Post"))))`,
    ];
  });
  const siF = [...Array(sheet.getLastRow() - 1)].map((_, i) => {
    const j = i + 2;
    return [
      `=if(REGEXMATCH(B${j},"meaning|Meaning|means|Means|mean|Meanthat|That|with|With|how|How|what|What|why|Why|when|When|why|Why|can|Can|do"),"Informational",if(REGEXMATCH(B${j},"best|Best|vs|Vs|services|Services|service|Service|firm|Firm|consultant|consulting|agency|Consultant|Consulting|Agency|company|companies|Company|Companies|product|Product|platform|Platform|saas|SaaS|SAAS|software|Software|Tool|tool|provider|Provider|providers|Providers|solution|Solution|solutions|Solutions"),"Commercial", if(REGEXMATCH(B${j},"near|Near|where|Where"),"Navigational","Informational")))`,
    ];
  });
  // set arrays to formula columns
  sheet.getRange(2, psI, sheet.getLastRow() - 1, 1).setFormulas(psF);
  sheet.getRange(2, scI, sheet.getLastRow() - 1, 1).setFormulas(scF);
  sheet.getRange(2, siI, sheet.getLastRow() - 1, 1).setFormulas(siF);
};
