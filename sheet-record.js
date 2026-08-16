/**
 * Record sheet for COB-CHA
 * functions:
 * - getPulldownMenu
 * - generateSheets
 * - getSheets
 * - addSheet
 */

/**
 * get pulldown menu
 * @return Object
 */
function getPulldownMenu() {
  var pullDown = SpreadsheetApp.newDataValidation();
  pullDown.requireValueInList(getProp('mark'), true);
  return pullDown;
}

/**
 * Build record-sheet rows for a single batch write.
 * @param Array usingCriteria
 * @return Array
 */
function getRecordCriteriaRows(usingCriteria) {
  return usingCriteria.map(function(criterion) {
    return [criterion[1], '', criterion[0]];
  });
}

/**
 * Generate Sheets
 * @param String urlstr
 * @param String lang
 * @param String testType
 * @param String level
 * @param Object targetId
 * @return Object
 */
function generateSheets(urlstr, lang, testType, level, targetId) {
  var ss = getSpreadSheet();
  var returnSheet = ss.getActiveSheet();

  // Generate "SC template sheet" or "Webpage sheet"
  // SC template sheet
  if (urlstr === gScTplSheetName) {
    // Error: already exists
    if (isSheetExist(gScTplSheetName)) return {
      'msg': getUiLang('target-sheet-already-exists', "%s is already exists.").replace('%s', gScTplSheetName),
      'targetId': targetId
    };
    var urls = [[urlstr, urlstr]];
  // Webpage sheet
  } else {
    // Error: URL list was not exist
    if ( ! isSheetExist(gUrlListSheetName)) return {
      'msg': getUiLang('no-target-sheet-exists', "sheet (%s) is not exist.").replace('%s', gUrlListSheetName),
      'targetId': targetId
    };
    var urlListSheet = ss.getSheetByName(gUrlListSheetName);
    var lastRow = urlListSheet.getLastRow();
    if (lastRow < 2) return {
      'msg': getUiLang('no-target-page-exists', "No target Page Exists"),
      'targetId': targetId
    };
    var urls = urlListSheet.getRange(2, 1, lastRow - 1, 4).getValues();

    // Ignore unused rows in the URL list.
    urls = urls.filter(function(row) {
      return row[1].toString().trim() != '';
    });
    
    // sheetname check - check not a number
    var nanFound = false;
    for (var i = 0; i < urls.length; i++) {
      if (urls[i][0].toString().search(/^[0-9]+$/) == -1) {
        nanFound = true;
      }
    }
    // Error: sheetname of not a number was exist
    if (nanFound) return {
      'msg': getUiLang('error-sheetname-must-be-numeric', "Error. All Sheetname must be numeric."),
      'targetId': targetId
    }
  }

  // Error: No URLs exist
  if (urls.length == 0) return {
    'msg': getUiLang('no-target-page-exists', "No target Page Exists"),
    'targetId': targetId
  };

  // start generate
  var alreadyExists = [];
  var added = 0;
  var urlsVals = [];
  var templateSheet = null;

  for (var i = 0; i < urls.length; i++) {
    var targetSheetname = urls[i][0].toString().trim();
    var targetUrl = urls[i][1].toString().trim();

    // check already exists. just add values for update URL Sheet and continue
    if (ss.getSheetByName(targetSheetname)) {
      alreadyExists.push(targetUrl.toString());
      urlsVals = generateUrlsSheetValues(
        urlsVals,
        ss.getSheetByName(targetSheetname),
        targetSheetname,
        false,
        urls[i][3]
      );
      continue;
    }

    // Build the first new sheet normally, then duplicate it for remaining pages.
    addSheet(targetSheetname, templateSheet);
    var targetSheet = ss.getSheetByName(targetSheetname);
    var isEvaluateTarget = targetSheetname.charAt(0) != '*';
    var res = isEvaluateTarget ? getHtmlAndTitle(targetUrl) : {'title':targetUrl, 'html':''};

    if (templateSheet == null) {
      generateASheet(
        testType,
        targetSheet,
        targetUrl,
        res,
        targetSheetname,
        isEvaluateTarget
      );
      templateSheet = targetSheet;
    } else {
      setRecordSheetPageData(targetSheet, targetUrl, res);
    }

    urlsVals = generateUrlsSheetValues(
      urlsVals,
      targetSheet,
      targetSheetname,
      true,
      urls[i][3]
    );
    added++;
  }

  // update url list sheet
  if (urlstr !== gScTplSheetName && urlsVals.length > 0) {
    urlListSheet.getRange(2, 1, urlsVals.length, urlsVals[0].length).setValues(urlsVals);
  }
  
  // clean up
  deleteFallbacksheet();
  
  // return to the sheet that was active before generation
  returnSheet.activate();
  
  var msg = [];
  msg.push(getUiLang('sheet-generated', "%s sheet(s) generated").replace('%s', added));
  if (alreadyExists.length > 0) {
    msg.push(getUiLang('sheet-already-exists', "%s sheet(s) were already exists: \n").replace('%s', alreadyExists.length)+alreadyExists.join("\n"));
  }

  return {'msg': msg.join("\n"), 'targetId': targetId};
}

/**
 * Generate a Sheet
 * @param String testType
 * @param Object sheet
 * @param String targetUrl
 * @param Array res
 * @param String sheetname
 * @param Bool isEvaluateTarget
 * @return Object
 */
function generateASheet(testType, sheet, targetUrl, res, sheetname, isEvaluateTarget) {
  var today = new Date();
  sheet.getRange(1, 1).setValue('URL').setBackground(gLabelColor);
  sheet.getRange(1, 5).setValue('Title').setBackground(gLabelColor);
  sheet.getRange(2, 1).setValue(getUiLang('date', 'Date')).setBackground(gLabelColor);
  sheet.getRange(2, 2).setValue(Utilities.formatDate(today, "JST", "yyyy/MM/dd"));
  sheet.getRange(2, 3).setValue(getUiLang('tester', 'Tester')).setBackground(gLabelColor);
  sheet.getRange(2, 5).setValue(getUiLang('memo', 'Memo')).setBackground(gLabelColor);
  sheet.getRange(3, 1).setValue('HTML').setBackground(gLabelColor);
  setRecordSheetPageData(sheet, targetUrl, res);
  sheet.hideRow(sheet.getRange(3, 1));
  sheet.setFrozenRows(2);
  
  // header
  var firstHeader = testType == 'tt20'
    ? getUiLang('test-id', 'Test ID')
    : getUiLang('criterion', 'Criterion');
  sheet.getRange(4, 1, 1, 4).setValues([[
    firstHeader,
    getUiLang('check', 'Check'),
    getUiLang('level', 'Level'),
    getUiLang('memo', 'Memo')
  ]]);
  sheet.getRange("4:4").setBackground(gLabelColorDark).setFontColor(gLabelColorDarkText).setFontWeight('bold');
  
  // appearance
/*
  // 20231104 try to use default width
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 50);
*/
  sheet.getRange('4:4').setHorizontalAlignment('center');
  
  // test type
  var usingCriteria = getUsingCriteria();
  var criteriaRows = getRecordCriteriaRows(usingCriteria);
  var criteriaRange = sheet.getRange(5, 1, criteriaRows.length, 3);
  criteriaRange.setValues(criteriaRows).setHorizontalAlignment('center');

  var checkRange = sheet.getRange(5, 2, criteriaRows.length, 1);
  checkRange.setDataValidation(getPulldownMenu());
  for (var j = 0; j < usingCriteria.length; j++) {
    checkRange.getCell(j + 1, 1).setComment(usingCriteria[j][2]);
  }
  
  // mark
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];

  // conditioned cell
  var range = sheet.getRange("B:B");
  setCellConditionTF(sheet, range, mT, mF); // see sheet-result.gs
  
  setRowConditionNotYet(sheet, criteriaRange, "=$B5=\"\"");

  setChecklistBorders(
    sheet.getRange(4, 1, criteriaRows.length + 1, 4)
  );
  
  return sheet;
}

/**
 * Set values that differ between duplicated record sheets.
 * @param Object sheet
 * @param String targetUrl
 * @param Object res
 * @return Void
 */
function setRecordSheetPageData(sheet, targetUrl, res) {
  sheet.getRange(1, 2).setValue(toPlainTextCellValue(targetUrl));
  sheet.getRange(1, 6).setValue(toPlainTextCellValue(res['title']));
  // Google Spreadsheet allows 50000 characters per cell.
  sheet.getRange(3, 2).setValue(
    toPlainTextCellValue(res['html'].substring(0, 45000))
  );
}

/**
 * generate Urls sheet Values
 * @param Array urlsVals
 * @param Object sheet
 * @param String sheetname
 * @param Bool isNew
 * @param String memo
 * @return Array
 */
function generateUrlsSheetValues(urlsVals, sheet, sheetname, isNew, memo) {
  var id    = '=HYPERLINK("#gid='+sheet.getSheetId()+'","'+sheetname+'")';
  var url   = getUrlFromSheet(sheet);
  var title = getTitleFromSheet(sheet);
  var apply = isNew ? "o" : "";
  urlsVals.push([id, url, title, memo || '', apply]);
  return urlsVals;
}

/**
 * Get Sheets Names
 * @return String
 */
function getSheets() {
  var allSheets = getAllSheets()
  var str = '';
  for (var i = 0; i < allSheets.length; i++) {
    var sheetname = String(allSheets[i].getName());
    str = str+sheetname+"\n";
  }
  return str;
}

/**
 * add sheet
 * @param String sheetname
 * @param String template
 * @return Bool
 */
function addSheet(sheetname, template) {
  var ss = getSpreadSheet();
 
  sheetname = sheetname.toString()
  var sheet = ss.getSheetByName(sheetname);
  var sheetIndex  = sheetname.charAt(0) == '*' ? 0 : ss.getSheets().length + 1;

  // sheet which name started with * must be refreshed
  if (sheetIndex == 0 && sheet != null) {
    ss.deleteSheet(sheet);
  }
  if (ss.getSheetByName(sheetname)) return false;
  if (template) {
    ss.insertSheet(sheetname, sheetIndex, {template: template});
  } else {
    ss.insertSheet(sheetname, sheetIndex);
  }
  return true;
}
