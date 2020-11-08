/**
 * Sheet control for COB-CHA
 */

/**
 * Get Spreadsheet
 * @return Object
 */
function getSpreadSheet() {
  if (getSpreadSheet.ss) return getSpreadSheet.ss;
  getSpreadSheet.ss = SpreadsheetApp.getActive();
  return getSpreadSheet.ss;
};

/**
 * Get Active Spreadsheet
 * @return Object
 */
function getActiveSheet() {
  if (getActiveSheet.ss) return getActiveSheet.ss;
  var ss = getSpreadSheet();
  getActiveSheet.ss = ss.getActiveSheet();
  return getActiveSheet.ss;
};

/**
 * Get All sheets
 * @return Object
 */
function getAllSheets() {
  if (getAllSheets.ss) return getAllSheets.ss;
  var ss = getSpreadSheet();
  var all = ss.getSheets();
  
  ret = [];
  for (i = 0; i < all.length; i++) {
    if (String(all[i].getName()).charAt(0) == '*') continue;
    ret.push(all[i]);
  }

  getAllSheets.ss = ret;
  return getAllSheets.ss;
};


/**
 * set basic value
 * @param Object sheet
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function setBasicValue(sheet, lang, testType, level) {
  sheet.getRange(1, 1).setValue('Type').setBackground(labelColor);
  sheet.getRange(1, 2).setValue(lang).setHorizontalAlignment('center');
  sheet.getRange(1, 3).setValue(testType).setHorizontalAlignment('center');
  sheet.getRange(1, 4).setValue(level).setHorizontalAlignment('center');
}

/**
 * get pulldown menu
 * @return Object
 */
function getPulldownMenu() {
  var pullDown = SpreadsheetApp.newDataValidation();
  pullDown.requireValueInList(['NT', 'DNA', 'T', 'F'], true);
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

  if (urlstr === templateSheetName) {
    if (ss.getSheetByName(templateSheetName)) return {'msg': getUiLang('template-already-exists', "*Template* is already exists."), 'targetId': targetId};
    var urls = [[urlstr, urlstr]];
  } else {
    if ( ! isUrlListSheetExists()) return {'msg': getUiLang('url-list-sheet-is-not-exists', "URL List sheet is not exist."), 'targetId': targetId};
    var urlListSheet = ss.getSheetByName(urlListSheetName);
    var lastRow = urlListSheet.getLastRow();
    var urls = urlListSheet.getRange(3, 1, lastRow - 2, 3).getValues();
  }
  if (urls.length == 1 && urls[0][0] == '') return {'msg': getUiLang('no-target-page-exists', "No target Page Exists"), 'targetId': targetId};
    
  var alreadyExists = [];
  var added = 0;
  var modUrlSheet = urls;

  // generate original sheet
  var firstSheetName = urls[0][0].toString();
  var firstUrl = urls[0][1].toString();
  
  if (addSheet(firstSheetName) == false) {
    alreadyExists.push(firstUrl);
    var originalSheet = ss.getSheetByName(firstSheetName);
  } else {
    var isEvaluateTarget = firstSheetName.charAt(0) != '*';
    var originalSheet = ss.getActiveSheet();
    var res = isEvaluateTarget ? getHtmlAndTitle(firstUrl) : {'title':firstUrl, 'html':''};
    
    setBasicValue(originalSheet, lang, testType, level);
    var today = new Date();
    originalSheet.getRange(1, 5).setValue(getUiLang('date', 'Date')).setBackground(labelColor);
    originalSheet.getRange(2, 5).setValue(getUiLang('screenshot', 'Screenshot')).setBackground(labelColor);
    originalSheet.getRange(1, 6).setValue(today);
    originalSheet.getRange(1, 7).setValue(getUiLang('memo', 'Memo')).setBackground(labelColor);
    originalSheet.getRange(2, 1).setValue('URL').setBackground(labelColor);
    originalSheet.getRange(2, 2).setValue(firstUrl);
    originalSheet.getRange(3, 1).setValue('title').setBackground(labelColor);
    originalSheet.getRange(3, 5).setValue(getUiLang('tester', 'Tester')).setBackground(labelColor);
    originalSheet.getRange(3, 2).setValue(res['title']);
    if (isEvaluateTarget) saveHtml(resourceFolderName, firstSheetName, res['html']);
    originalSheet.setFrozenRows(4);
    
    // header
    if (testType == 'tt20') {
      originalSheet.getRange(4, 1).setValue(getUiLang('test-id', 'Test ID')).setBackground(labelColor);
    } else {
      originalSheet.getRange(4, 1).setValue(getUiLang('criterion', 'Criterion')).setBackground(labelColor);
    }
    originalSheet.getRange(4, 2).setValue(getUiLang('check', 'Check'));
    originalSheet.getRange(4, 3).setValue(getUiLang('level', 'Level'));
    originalSheet.getRange(4, 4).setValue(getUiLang('tech', 'Techs'));
    originalSheet.getRange(4, 5).setValue(getUiLang('memo', 'Memo'));
    originalSheet.getRange("4:4").setBackground(labelColorDark).setFontColor(labelColorDarkText).setFontWeight('bold');
    
    // appearance
    originalSheet.setColumnWidth(1, 60);
    originalSheet.setColumnWidth(2, 50);
    originalSheet.setColumnWidth(3, 50);
    originalSheet.getRange('4:4').setHorizontalAlignment('center');
    
    // test type
    var usingCriteria = getUsingCriteria(lang, testType, level);
    
    // complex language selection ...
    var docurl = lang+'-'+testType;
    
    // each row
    var row = 5;
    for (var j = 0; j < usingCriteria.length; j++) {
      var langPointer = testType == 'wcag21' ? usingCriteria[j][4] : usingCriteria[j][3];
      langPointer = testType == 'wcag21' && lang == 'ja' ? usingCriteria[j][3] : langPointer;
      var urlPointer = docurl;
      urlPointer = testType == 'wcag21' && lang == 'ja' ? 'en-wcag21' : urlPointer;
      var url = urlbase['understanding'][urlPointer]+langPointer;
      if (lang == 'ja' && testType == 'wcag21' && criteria21.indexOf(usingCriteria[j][1]) >= 0) {
        url = urlbase['understanding']['en-wcag21']+usingCriteria[j][4];
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
    var conditionedRange = originalSheet.getRange("B:B");
    var ruleForF = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("F")
      .setBackground(falseColor)
      .setRanges([conditionedRange])
      .build();
    var ruleForT = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("T")
      .setBackground(trueColor)
      .setRanges([conditionedRange])
      .build();
    var rules = originalSheet.getConditionalFormatRules();
    rules.push(ruleForF);
    rules.push(ruleForT);
    originalSheet.setConditionalFormatRules(rules);
    added++;

    modUrlSheet[0] = [
      '=HYPERLINK("#gid='+originalSheet.getSheetId()+'","'+firstSheetName+'")',
      firstUrl,
      res['title']
    ];
  }
  
  // copy sheets
  if (urlstr !== templateSheetName && urls.length > 1) {
    for(var i = 1; i < urls.length; i++) {
      if (urls[i][1].trim() == '') continue;
      if (addSheet(urls[i][0], originalSheet) == false) {
        alreadyExists.push(urls[i][1]);
        continue;
      }
      res = getHtmlAndTitle(urls[i][1]);
      var activeSheet = ss.getActiveSheet();
      activeSheet.getRange(2, 2).setValue(urls[i][1]);
      activeSheet.getRange(3, 2).setValue(res['title']);
      saveHtml(resourceFolderName, urls[i][0], res['html']);
      added++;
      
      modUrlSheet[i] = [
        '=HYPERLINK("#gid='+activeSheet.getSheetId()+'","'+urls[i][0]+'")',
        urls[i][1],
        res['title']
      ];
    }
    
    // recover fomula
    var numWithFomula = urlListSheet.getRange(3, 1, lastRow - 2, 1).getFormulas();
    Logger.log(numWithFomula);
    for(var i = 0; i < urls.length; i++) {
      modUrlSheet[i][0] = numWithFomula[i][0] ? numWithFomula[i][0] : urls[i][0];
    }
    Logger.log(modUrlSheet);
    
    // update url list sheet
    urlListSheet.getRange(3, 1, lastRow - 2, 3).setValues(modUrlSheet);
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
  // Microsoft Excel compatible
  // Excel's sheetname cannot use : and /
  sheetname = String(sheetname).replace(/https*:\/\//ig, '');
  sheetname = String(sheetname).replace(/\//ig, ' ');
 
  // Excel's sheetname must be under 31 chars
  if (sheetname.length > 28) {
    var tmpbase = sheetname.substr(0, 28);
    var tmp = tmpbase;
    var i = 2;
    while(ss.getSheetByName(tmp)) {
      var tmp = tmpbase+'-'+i;
      i++;
    }
    sheetname = tmp;
  }
  
  var targetSheet = ss.getSheetByName(sheetname);
  var sheetIndex  = sheetname.charAt(0) == '*' ? 0 : ss.getSheets().length+1;

  // sheet which name started with * must be refreashed
  if (sheetIndex == 0 && targetSheet != null) {
    ss.deleteSheet(targetSheet);
  }
  if (ss.getSheetByName(sheetname)) return false;
  if (template) {
    ss.insertSheet(sheetname, sheetIndex, {template: template});
  } else {
    ss.insertSheet(sheetname, sheetIndex);
  }
  return true;
}

/**
 * reset sheets
 * @param Bool isAll
 * @return String
 */
function resetSheets(isAll) {
  if ( ! isDebug()) throw new Error('allowed to developer only');
  
  var ss = getSpreadSheet();
  var all = ss.getSheets();
  
  deleteFallbacksheet();
  ss.insertSheet(fallbacksheetName, 0);
  
  var count = 0;
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName() == fallbacksheetName) continue;
    if (isAll === false && all[i].getName().charAt(0) == '*') continue;
    if (all[i] == null) continue;
    ss.deleteSheet(all[i]);
    count++;
  }
  var all2 = ss.getSheets();
  if (all2.length > 1) {
    deleteFallbacksheet();
  }
  
  return getUiLang('sheet-deleted', "%s sheet(s) deleted.").replace("%s", count);
}

/**
 * delete fallbacksheet
 * @return Void
 */
function deleteFallbacksheet() {
  var ss = getSpreadSheet();
  var fallbacksheet = ss.getSheetByName(fallbacksheetName);
  if (fallbacksheet) {
    ss.deleteSheet(fallbacksheet);
  }
}
