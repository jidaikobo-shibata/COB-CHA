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
    var urls = urlListSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    
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
  if (urls.length == 1 && urls[0][0] == '') return {
    'msg': getUiLang('no-target-page-exists', "No target Page Exists"),
    'targetId': targetId
  };

  // start generate
  var alreadyExists = [];
  var added = 0;
  var urlsVals = [];
  
  // generate a originalsheet
  for(var i = 0; i < urls.length; i++) {
    var targetSheetname = urls[i][0].toString().trim();
    var targetUrl = urls[i][1].toString().trim();
    if (targetUrl == "") continue;
    
    // check already exists. just add values for update URL Sheet and continue
    if (addSheet(targetSheetname) == false) {
      alreadyExists.push(targetUrl.toString());
      urlsVals = generateUrlsSheetValues(urlsVals, ss.getSheetByName(targetSheetname), targetSheetname, false)
      continue;
    }
    
    // generate a originalsheet
    var isEvaluateTarget = targetSheetname.charAt(0) != '*';
    var originalSheet = ss.getActiveSheet();
    var res = isEvaluateTarget ? getHtmlAndTitle(targetUrl) : {'title':targetUrl, 'html':''};
    
    originalSheet = generateASheet(testType, originalSheet, targetUrl, res, targetSheetname, isEvaluateTarget);
    urlsVals = generateUrlsSheetValues(urlsVals, originalSheet, targetSheetname, true);
    added++;
  }
  var lastSheetKey = i;
  
  // copy sheets from originalsheet - copying seems faster than generate
  if (urlstr !== gScTplSheetName && urls.length > 1) {
    for(var i = lastSheetKey; i < urls.length; i++) {
      var eachSheet = urls[i][1].toString().trim(); 
      if (eachSheet == '') continue;
      if (addSheet(urls[i][0], originalSheet) == false) {
        alreadyExists.push(eachSheet);
        urlsVals = generateUrlsSheetValues(urlsVals, ss.getSheetByName(urls[i][0]), urls[i][0], false);
        continue;
      }
      res = getHtmlAndTitle(eachSheet);
      var activeSheet = ss.getActiveSheet();
      activeSheet.getRange(2, 2).setValue(eachSheet);
      activeSheet.getRange(3, 2).setValue(res['title']);
      activeSheet.getRange(3, 2).setValue(res['html'].substring(0, 45000));
      urlsVals = generateUrlsSheetValues(urlsVals, activeSheet, urls[i][0], true);
      added++;
    }
  }

  // update url list sheet
  if (urlstr !== gScTplSheetName) {
    urlListSheet.getRange(2, 1, urlsVals.length, urlsVals[0].length).setValues(urlsVals);
  }
  
  // clean up
  deleteFallbacksheet();
  
  // return to original sheet
  originalSheet.activate();
  
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
  sheet.getRange(1, 2).setValue(targetUrl);
  sheet.getRange(1, 5).setValue('Title').setBackground(gLabelColor);
  sheet.getRange(1, 6).setValue(res['title']);
  sheet.getRange(2, 1).setValue(getUiLang('date', 'Date')).setBackground(gLabelColor);
  sheet.getRange(2, 2).setValue(Utilities.formatDate(today, "JST", "yyyy/MM/dd"));
  sheet.getRange(2, 3).setValue(getUiLang('tester', 'Tester')).setBackground(gLabelColor);
  sheet.getRange(2, 5).setValue(getUiLang('memo', 'Memo')).setBackground(gLabelColor);
  sheet.getRange(3, 1).setValue('HTML').setBackground(gLabelColor);
  sheet.getRange(3, 2).setValue(res['html'].substring(0, 45000)); //Google Spreadsheet allow 50000 chr/1 cell
  sheet.hideRow(sheet.getRange(3, 1));
  sheet.setFrozenRows(2);
  
  // header
  if (testType == 'tt20') {
    sheet.getRange(4, 1).setValue(getUiLang('test-id', 'Test ID')).setBackground(gLabelColor);
  } else {
    sheet.getRange(4, 1).setValue(getUiLang('criterion', 'Criterion')).setBackground(gLabelColor);
  }
  sheet.getRange(4, 2).setValue(getUiLang('check', 'Check'));
  sheet.getRange(4, 3).setValue(getUiLang('level', 'Level'));
  sheet.getRange(4, 4).setValue(getUiLang('memo', 'Memo'));
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
  
  // each row
  var row = 5;
  for (var j = 0; j < usingCriteria.length; j++) {
    if (testType == 'tt20') {
      sheet.getRange(row, 1).setValue(usingCriteria[j][1]).setHorizontalAlignment('center');
    } else {
//      sheet.getRange(row, 1).setValue('=HYPERLINK("'+usingCriteria[j][5]+'", "'+usingCriteria[j][1]+'")').setHorizontalAlignment('center');
      sheet.getRange(row, 1).setValue(usingCriteria[j][1]).setHorizontalAlignment('center');
    }
    sheet.getRange(row, 2).setDataValidation(getPulldownMenu()).setHorizontalAlignment('center').setComment(usingCriteria[j][2]);
    sheet.getRange(row, 3).setValue(usingCriteria[j][0]).setHorizontalAlignment('center');
    row++;
  }
  
  // mark
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var mD = mark[1];

  // conditioned cell
  var range = sheet.getRange("B:B");
  setCellConditionTF(sheet, range, mT, mF); // see sheet-result.gs
  
  var range = sheet.getRange(5, 1, sheet.getLastRow() - 4, 3);
  setRowConditionNotYet(sheet, range, "=$B5=\"\"");
  
  return sheet;
}

/**
 * generate Urls sheet Values
 * @param Array urlsVals
 * @param Object sheet
 * @param String sheetname
 * @param Bool isNew
 * @return Array
 */
function generateUrlsSheetValues(urlsVals, sheet, sheetname, isNew) {
  var id    = '=HYPERLINK("#gid='+sheet.getSheetId()+'","'+sheetname+'")';
  var url   = getUrlFromSheet(sheet);
  var title = getTitleFromSheet(sheet);
  var apply = isNew ? "o" : "";
  urlsVals.push([id, url, title, '', apply]);
  return urlsVals;
}

/**
 * Get Sheets Names
 * @return String
 */
function getSheets() {
  var allSheets = getAllSheets()
  var str = '';
  for (i = 0; i < allSheets.length; i++) {
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
