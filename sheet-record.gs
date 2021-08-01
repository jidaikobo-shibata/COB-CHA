/**
 * Sheet control for COB-CHA
 * finctions:
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

  if (urlstr === gScTplSheetName) {
    if (isSheetExist(gScTplSheetName)) return {
      'msg': getUiLang('target-sheet-already-exists', "%s is already exists.").replace('%s', gScTplSheetName),
      'targetId': targetId
    };
    var urls = [[urlstr, urlstr]];
  } else {
    if ( ! isSheetExist(gUrlListSheetName)) return {
      'msg': getUiLang('no-target-sheet-exists', "sheet (%s) is not exist.").replace('%s', gUrlListSheetName),
      'targetId': targetId
    };
    var urlListSheet = ss.getSheetByName(gUrlListSheetName);
    var lastRow = urlListSheet.getLastRow();
    var urls = urlListSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  }
  if (urls.length == 1 && urls[0][0] == '') return {'msg': getUiLang('no-target-page-exists', "No target Page Exists"), 'targetId': targetId};
    
  if (urlstr !== gScTplSheetName) {
    var sheetIds = urlListSheet.getRange(2, 1, lastRow - 1, 1).getFormulas();
    var sheetTitles = urlListSheet.getRange(2, 3, lastRow - 1, 1).getValues();
  }
  var alreadyExists = [];
  var added = 0;

  // generate original sheet
  var firstSheetName = urls[0][0].toString();
  var firstUrl = urls[0][1].toString();

  // mark
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var mD = mark[1];

  if (addSheet(firstSheetName) == false) {
    alreadyExists.push(firstUrl);
    var originalSheet = ss.getSheetByName(firstSheetName);
  } else {
    var isEvaluateTarget = firstSheetName.charAt(0) != '*';
    var originalSheet = ss.getActiveSheet();
    var res = isEvaluateTarget ? getHtmlAndTitle(firstUrl) : {'title':firstUrl, 'html':''};
    
    var today = new Date();
    originalSheet.getRange(1, 1).setValue(getUiLang('memo', 'Memo')).setBackground(gLabelColor);
    originalSheet.getRange(2, 1).setValue('URL').setBackground(gLabelColor);
    originalSheet.getRange(2, 2).setValue(firstUrl);
    originalSheet.getRange(2, 5).setValue(getUiLang('screenshot', 'Screenshot')).setBackground(gLabelColor);
    originalSheet.getRange(3, 1).setValue('title').setBackground(gLabelColor);
    originalSheet.getRange(3, 2).setValue(res['title']);
    originalSheet.getRange(3, 5).setValue(getUiLang('tester', 'Tester')).setBackground(gLabelColor);
    originalSheet.getRange(3, 7).setValue(getUiLang('date', 'Date')).setBackground(gLabelColor);
    originalSheet.getRange(3, 8).setValue(Utilities.formatDate(today, "JST", "yyyy/MM/dd"));
    if (isEvaluateTarget) saveHtml(gResourceFolderName, firstSheetName, res['html']);
    originalSheet.setFrozenRows(4);
    
    // header
    if (testType == 'tt20') {
      originalSheet.getRange(4, 1).setValue(getUiLang('test-id', 'Test ID')).setBackground(gLabelColor);
    } else {
      originalSheet.getRange(4, 1).setValue(getUiLang('criterion', 'Criterion')).setBackground(gLabelColor);
    }
    originalSheet.getRange(4, 2).setValue(getUiLang('check', 'Check'));
    originalSheet.getRange(4, 3).setValue(getUiLang('level', 'Level'));
    originalSheet.getRange(4, 4).setValue(getUiLang('memo', 'Memo'));
    originalSheet.getRange("4:4").setBackground(gLabelColorDark).setFontColor(gLabelColorDarkText).setFontWeight('bold');
    
    // appearance
    originalSheet.setColumnWidth(1, 60);
    originalSheet.setColumnWidth(2, 50);
    originalSheet.setColumnWidth(3, 50);
    originalSheet.getRange('4:4').setHorizontalAlignment('center');
    
    // test type
    var usingCriteria = getUsingCriteria();
    
    // complex language selection ...
    var docurl = lang+'-'+testType;
    
    // each row
    var row = 5;
    for (var j = 0; j < usingCriteria.length; j++) {
      var langPointer = testType == 'wcag21' ? usingCriteria[j][4] : usingCriteria[j][3];
      langPointer = testType == 'wcag21' && lang == 'ja' ? usingCriteria[j][3] : langPointer;
      var urlPointer = docurl;
      urlPointer = testType == 'wcag21' && lang == 'ja' ? 'en-wcag21' : urlPointer;
      var url = gUrlbase['understanding'][urlPointer]+langPointer;
      if (lang == 'ja' && testType == 'wcag21' && gCriteria21.indexOf(usingCriteria[j][1]) >= 0) {
        url = gUrlbase['understanding']['en-wcag21']+usingCriteria[j][4];
      }
      
      if (testType == 'tt20') {
        originalSheet.getRange(row, 1).setValue(usingCriteria[j][1]).setHorizontalAlignment('center');
      } else {
        originalSheet.getRange(row, 1).setValue('=HYPERLINK("'+url+'", "'+usingCriteria[j][1]+'")').setHorizontalAlignment('center');
      }
      originalSheet.getRange(row, 2).setDataValidation(getPulldownMenu()).setHorizontalAlignment('center').setComment(usingCriteria[j][2]);
      originalSheet.getRange(row, 3).setValue(usingCriteria[j][0]).setHorizontalAlignment('center');
      row++;
    }
    
    // conditioned cell
    var range = originalSheet.getRange("B:B");
    setCellConditionTF(originalSheet, range, mT, mF); // see sheet-result.gs
    added++;

    if (urlstr !== gScTplSheetName) {
      sheetIds[0]    = ['=HYPERLINK("#gid='+originalSheet.getSheetId()+'","'+firstSheetName+'")'];
      sheetTitles[0] = res['title'] == '' ? [sheetTitles[0][0]] : [res['title']];
    }
  }
  
  // copy sheets
  if (urlstr !== gScTplSheetName && urls.length > 1) {
    for(var i = 1; i < urls.length; i++) {
      var eachSheet = urls[i][1].toString(); 
      if (eachSheet.trim() == '') continue;
      if (addSheet(urls[i][0], originalSheet) == false) {
        alreadyExists.push(eachSheet);
        continue;
      }
      res = getHtmlAndTitle(eachSheet);
      var activeSheet = ss.getActiveSheet();
      activeSheet.getRange(2, 2).setValue(eachSheet);
      activeSheet.getRange(3, 2).setValue(res['title']);
      saveHtml(gResourceFolderName, urls[i][0], res['html']);
      added++;
      
      sheetIds[i]    = ['=HYPERLINK("#gid='+activeSheet.getSheetId()+'","'+urls[i][0]+'")'];
      sheetTitles[i] = [res['title']];
    }
  }

  // update url list sheet
  if (urlstr !== gScTplSheetName) {
    urlListSheet.getRange(2, 1, lastRow - 1, 1).setValues(sheetIds);
    urlListSheet.getRange(2, 3, lastRow - 1, 1).setValues(sheetTitles);
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

  // sheet which name started with * must be refreashed
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
