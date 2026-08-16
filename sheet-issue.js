/**
 * Issue sheet for COB-CHA
 * functions:
 * - isEditIssue
 * - dialogValueIssue
 * - generateIssueSheet
 * - openDialogIssue
 * - applyIssue
 * - setIssueList
 * - showEachIssue
 */

/**
 * is Edit or Add
 * @return Bool
 */
function isEditIssue() {
  // if not exist issue sheet then this time is not for edit
  if ( ! isSheetExist(gIssueSheetName)) return false;

  // if current sheet is not issue sheet then this time is not for edit
  if (getActiveSheet().getName() != gIssueSheetName) return false;

  // if current row has id then this time is for edit
  var sheet = getSpreadSheet().getSheetByName(gIssueSheetName);
  var activeRow = sheet.getActiveCell().getRow();
  var issueId = sheet.getRange(activeRow, 1).getValue();
  
  return (String(issueId).length > 0 && activeRow > 1);
}

/**
 * Build lookups for the displayed No. in *URLs.
 * @param Array rows
 * @return Object
 */
function buildPageNumberMap(rows) {
  var pageNumbers = {
    'bySheetAndUrl': {},
    'bySheetName': {},
    'byUrl': {}
  };
  for (var i = 0; i < rows.length; i++) {
    var rawPageNumber = rows[i][0];
    var pageNumber = String(rawPageNumber == null ? '' : rawPageNumber).trim();
    var url = String(rows[i][1] == null ? '' : rows[i][1]).trim();
    if (pageNumber == '') continue;
    pageNumbers.bySheetName[pageNumber] = pageNumber;
    if (url != '') {
      pageNumbers.bySheetAndUrl[pageNumber + '\n' + url] = pageNumber;
      if (!pageNumbers.byUrl[url]) pageNumbers.byUrl[url] = pageNumber;
    }
  }
  return pageNumbers;
}

/**
 * Find the best matching page number for a record sheet.
 * @param Object pageNumberMap
 * @param String sheetName
 * @param String url
 * @return String
 */
function findPageNumber(pageNumberMap, sheetName, url) {
  return pageNumberMap.bySheetAndUrl[sheetName + '\n' + url] ||
    pageNumberMap.byUrl[url] ||
    pageNumberMap.bySheetName[sheetName] || '';
}

/**
 * set dialog Value Issue
 * @return {Object}
 */
function dialogValueIssue() {
  var ret = {};
  ret['isEdit'] = isEditIssue();
  ret['lang'] = getProp('lang');
  ret['type'] = getProp('type');
  ret['level'] = getProp('level');
  ret['usingCriteria'] = getUsingCriteria();
  ret['usingTechs'] = getUsingTechs();

  ret['allPlaces'] = [];
  var pageNumberMap = buildPageNumberMap([]);
  var urlListSheet = getSheetIfExists(gUrlListSheetName);
  if (urlListSheet && urlListSheet.getLastRow() >= 2) {
    var urlRows = urlListSheet
      .getRange(2, 1, urlListSheet.getLastRow() - 1, 3)
      .getDisplayValues();
    pageNumberMap = buildPageNumberMap(urlRows);
  }
  var all = getAllSheets();
  for (var i = 0; i < all.length; i++) {
    var targetUrl = getUrlFromSheet(all[i]);
    ret['allPlaces'].push({
      'no' : findPageNumber(pageNumberMap, all[i].getName(), targetUrl),
      'url' : targetUrl,
      'title' : getTitleFromSheet(all[i])
    });
  }
  
  ret['vals'] = {};
  var celposes = {
    'issueId': 1,
    'issueName': 2,
    'issueVisibility': 3,
    'errorNotice': 4,
    'html': 5,
    'explanation': 6,
    'checked': 7,
    'techs': 8,
    'places': 9,
    'memo': 10
  };

  if (ret['isEdit']) {
    // issue sheet must be existed and activated
    var ss = getSpreadSheet();
    var sheet = ss.getSheetByName(gIssueSheetName);
    var activeRow = sheet.getActiveCell().getRow();
    var row = sheet.getRange(activeRow, 1, 1, 10).getValues()[0];
    for (var key in celposes) {
      var idx = celposes[key] - 1;
      var v = row[idx];
      ret['vals'][key] = (v == null ? '' : String(v));
    }
  }

  return ret;
}

/**
 * generate Issue sheet
 * @return Void
 */
function generateIssueSheet() {
  if (isSheetExist(gIssueSheetName)) return;
  
  // generate Issue sheet
  var defaults = [[
    "ID",
    getUiLang('name', 'Name'),
    getUiLang('issue-solved', 'Solved'),
    'Type',
    'HTML',
    getUiLang('explanation', 'Explanation'),
    getUiLang('criterion', 'Criteria'),
    getUiLang('tech', 'Techniques'),
    getUiLang('places', 'Places'),
    getUiLang('memo', 'Memo')
  ]];
  var sheet = generateSheetIfNotExists(gIssueSheetName, defaults, "row"); // do not return msg
  sheet.getRange("F:F").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 45);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);

  // When I tried to set conditional formatting for the entire sheet in advance,
  // it was very heavy and impractical, so I decided to set it each time.
}

/**
 * open dialog Issue
 * @return Void
 */
function openDialogIssue() {
  // to tell current page
  var activeSheet = getActiveSheet();
  var targetUrl = '';
  if (activeSheet.getName().charAt(0) != '*') {
    targetUrl = getUrlFromSheet(activeSheet);
  }

  // generate
  generateIssueSheet()  
  
  var title = isEditIssue() ? getUiLang('edit-issue', 'Edit issue') : getUiLang('add-new-issue', 'Add new issue');
  showDialog('ui-issue', 500, 400, title, '', {'targetUrl': targetUrl});
}

/**
 * Find spreadsheet rows that have the specified Issue ID.
 * @param Array dataObj
 * @param Number issueId
 * @return Array
 */
function getIssueRowsById(dataObj, issueId) {
  var rows = [];
  for (var i = 1; i < dataObj.length; i++) {
    if (Number(dataObj[i][0]) === issueId) rows.push(i + 1);
  }
  return rows;
}

/**
 * Get the next Issue ID without reusing IDs from deleted rows.
 * @param Array dataObj
 * @return Number
 */
function getNextIssueId(dataObj) {
  var maxId = 0;
  for (var i = 1; i < dataObj.length; i++) {
    var issueId = Number(dataObj[i][0]);
    if (Number.isInteger(issueId) && issueId > maxId) maxId = issueId;
  }
  return maxId + 1;
}

/**
 * apply Issue
 * @param Array vals
 * @return String
 */
function applyIssue(vals) {
  if (!Array.isArray(vals) || vals.length < 10) {
    return getUiLang('invalid-input', 'Invalid input.');
  }

  var ss = getSpreadSheet();
  var sheet = ss.getSheetByName(gIssueSheetName);
  if (!sheet) {
    return getUiLang('issue-sheet-not-exists', 'Issue sheet does not exist.');
  }

  // vals:
  // [0]=ID, [1]=Name, [2]=Solved, [3]=Type, [4]=HTML, [5]=Explanation,
  // [6]=Criteria, [7]=Techniques, [8]=Places(comma), [9]=Memo
  var rawIssueId = String(vals[0] || '').trim();
  var isEdit = rawIssueId != '';
  var requestedIssueId = Number(rawIssueId);
  if (isEdit && (!Number.isInteger(requestedIssueId) || requestedIssueId < 1)) {
    return getUiLang('invalid-issue-id', 'Invalid Issue ID.');
  }

  // Places
  var totalSheets = getAllSheets().length;
  var placesNormalized = vals[8];
  if (placesNormalized && String(placesNormalized).split(",").length === totalSheets) {
    placesNormalized = "all";
  }

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return getUiLang(
      'operation-busy',
      'Another operation is in progress. Please try again.'
    );
  }

  try {
    var dataObj = sheet.getDataRange().getValues();
    var targetRow;
    var issueId;

    if (isEdit) {
      var matchedRows = getIssueRowsById(dataObj, requestedIssueId);
      if (matchedRows.length == 0) {
        return getUiLang('issue-not-found', 'Issue was not found.');
      }
      if (matchedRows.length > 1) {
        return getUiLang('duplicate-issue-id', 'Duplicate Issue IDs were found.');
      }
      targetRow = matchedRows[0];
      issueId = requestedIssueId;
    } else {
      targetRow = sheet.getLastRow() + 1;
      issueId = getNextIssueId(dataObj);
    }

    // write once
    var row = [
      issueId,        // Col 1: ID
      vals[1],        // Col 2: Name
      vals[2],        // Col 3: Solved
      vals[3],        // Col 4: Type
      vals[4],        // Col 5: HTML
      vals[5],        // Col 6: Explanation
      vals[6],        // Col 7: Criteria
      vals[7],        // Col 8: Techniques
      placesNormalized, // Col 9: Places
      vals[9]         // Col10: Memo
    ];
    for (var i = 1; i < row.length; i++) {
      row[i] = toPlainTextCellValue(row[i]);
    }
    sheet.getRange(targetRow, 1, 1, 10).setValues([row]);

    // conditioned rows
    ensureGlobalIssueRules(sheet);

    if (isEdit) {
      return getUiLang('update-value', 'Edited: %s')
        .replace("%s", 'Issue "' + issueId + '"');
    }
    return getUiLang('add-value', 'Added: %s')
      .replace("%s", 'Issue "' + issueId + '"');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ensure the sheet has exactly 3 global conditional-format rules for rows >= 2.
 * - Uses relative row references ($C2, $D2) so one rule applies to all data rows.
 * - Replaces older versions of the same rules if they exist.
 */
/** Add/refresh 3 global conditional-format rules for the issue sheet. */
function ensureGlobalIssueRules(sheet) {
  // Apply to full columns A:J from row 2 (covers direct edits & appended rows)
  // If performance becomes an issue on very large sheets, switch to getLastRow()-based range.
  var startRow = 2, startCol = 1, numCols = 10;
  var numRows  = Math.max(1, sheet.getMaxRows() - (startRow - 1)); // includes blank rows
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);

  var f1 = '=AND($C2="off", $D2="Error")';
  var f2 = '=AND($C2="off", NOT($D2="Error"))';
  var f3 = '=AND(NOT($C2="off"), $D2="Error")';

  var existing = sheet.getConditionalFormatRules();
  var keep = [];
  for (var i = 0; i < existing.length; i++) {
    var r = existing[i], bc = r.getBooleanCondition && r.getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var formula = String(bc.getCriteriaValues()[0] || '');
      if (formula === f1 || formula === f2 || formula === f3) continue; // replace ours
    }
    keep.push(r);
  }

  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(f1).setBold(true).setBackground(gNotYetIssueBgColor).setRanges([dataRange]).build();
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(f2).setBold(false).setBackground(gNotYetIssueBgColor).setRanges([dataRange]).build();
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(f3).setBold(true).setFontColor(null).setBackground(null).setRanges([dataRange]).build();

  sheet.setConditionalFormatRules(keep.concat([rule1, rule2, rule3]));
}

/**
 * Check whether an issue applies to a URL.
 * @param {*} places
 * @param {*} url
 * @return {Bool}
 */
function issueAppliesToUrl(places, url) {
  var normalizedPlaces = String(places || '').trim();
  var normalizedUrl = String(url || '').trim();
  if (normalizedPlaces == '' || normalizedUrl == '') return false;
  if (normalizedPlaces.toLowerCase() == 'all') return true;

  return normalizedPlaces.split(',').some(function(issueUrl) {
    return issueUrl.trim() == normalizedUrl;
  });
}

/**
 * set Issue list
 * @return Object
 */
function setIssueList() {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(gIssueSheetName);
  if (!issueSheet) return {'url': '', 'issues': []};
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName().toString();
  
  // target URL
  if (activeSheetName == gResultSheetName) {
    var targetRow = activeSheet.getActiveCell().getRow();
    var url = activeSheet.getRange(targetRow, 1).getValue();
  } else {
    if (activeSheetName.charAt(0) == '*') {
      return {'url': '', 'issues': []};
    } else {
      var url = getUrlFromSheet(activeSheet);
    }
  }

  var dataObj = issueSheet.getDataRange().getValues();
  var issues = [];
  for (var i = 1; i < dataObj.length; i++) {
    if (issueAppliesToUrl(dataObj[i][8], url)) {
      issues.push(dataObj[i]);
    }
  }
  return {'url': url, 'issues': issues};
}

/**
 * show each issue
 * @param Integer row
 * @return Void
 */
function showEachIssue(row) {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(gIssueSheetName);
  var targetRow = Number(row);
  if (!issueSheet || !Number.isInteger(targetRow) ||
      targetRow < 2 || targetRow > issueSheet.getLastRow()) {
    return getUiLang('issue-not-found', 'Issue was not found.');
  }
  issueSheet.getRange(targetRow, 1).activate();
  showDialog(
    'ui-issue',
    500,
    400,
    getUiLang('edit-issue', 'Edit issue'),
    '',
    {'targetUrl': ''}
  );
  return '';
}
