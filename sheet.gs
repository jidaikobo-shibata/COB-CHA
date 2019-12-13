/**
 * Sheet control for COB-CHA
 */

/**
 * set basic value
 * @param Object sheet
 * @param String lang
 * @param String testType
 * @param String level
 */
function setBasicValue(sheet, lang, testType, level) {
  sheet.getRange(1, 1).setValue('Type').setBackground(labelColor);
  sheet.getRange(1, 2).setValue(lang).setHorizontalAlignment('center');
  sheet.getRange(1, 3).setValue(testType).setHorizontalAlignment('center');
  sheet.getRange(1, 4).setValue(level).setHorizontalAlignment('center');
}

/**
 * Generate Sheets
 * @param String urlstr
 * @param String lang
 * @param String testType
 * @param String level
 * @param Object targetId
 * @return Array
 */
function generateSheets(urlstr, lang, testType, level, targetId) {
  var urls = urlstr.replace(/^\s+|\s+$|\n\n/g, '').split(/\n/);
  if (urls.length == 0) return {'msg': getUiLang('no-target-page-exists', "No target Page Exists"), 'targetId': targetId};

  var pullDown = SpreadsheetApp.newDataValidation();
  pullDown.requireValueInList(['Yet', 'DNA', 'T', 'F'], true);
  
  var ss = getSpreadSheet();
  var alreadyExists = [];
  var added = 0;
  
  // generate original sheet
  if (addSheet(urls[0]) == false) {
    alreadyExists.push(urls[0]);
    var originalSheet = ss.getSheetByName(urls[0]);
  } else {
    var isEvaluateTarget = urls[0].charAt(0) != '*';
    
    var originalSheet = ss.getActiveSheet();
    
    var res = isEvaluateTarget ? getHtmlAndTitle(urls[0]) : false;
    
    // meta
    setBasicValue(originalSheet, lang, testType, level);
    var today = new Date();
    originalSheet.getRange(1, 5).setValue(getUiLang('date', 'Date')).setBackground(labelColor);
    originalSheet.getRange(2, 5).setValue(getUiLang('screenshot', 'Screenshot')).setBackground(labelColor);
    originalSheet.getRange(1, 6).setValue(today);
    originalSheet.getRange(1, 7).setValue(getUiLang('memo', 'Memo')).setBackground(labelColor);
    originalSheet.getRange(2, 1).setValue('URL').setBackground(labelColor);
    originalSheet.getRange(2, 2).setValue(urls[0]);
    originalSheet.getRange(3, 1).setValue('title').setBackground(labelColor);
    if (res) {
      originalSheet.getRange(3, 2).setValue(res['title']);
      saveHtml(resourceFolderName, urls[0], res['html']);
    }
    originalSheet.setFrozenRows(4);
    
    // header
    originalSheet.getRange(4, 1).setValue(getUiLang('criterion', 'Criterion')).setBackground(labelColor);
    originalSheet.getRange(4, 2).setValue(getUiLang('check', 'Check')).setBackground(labelColor);
    originalSheet.getRange(4, 3).setValue(getUiLang('level', 'Level')).setBackground(labelColor);
    originalSheet.getRange(4, 4).setValue(getUiLang('tech', 'Techs')).setBackground(labelColor);
    originalSheet.getRange(4, 5).setValue(getUiLang('memo', 'Memo')).setBackground(labelColor);
    
    // appearance
    originalSheet.setColumnWidth(1, 60);
    originalSheet.setColumnWidth(2, 50);
    originalSheet.setColumnWidth(3, 50);
    originalSheet.getRange('4:4').setHorizontalAlignment('center');
    
    // test type
    var set = testType.indexOf('wcag') >= 0 ? 'criteria' : 'ttCriteria' ;
    var usingCriteria = getLangSet(set);
    
    // complex language selection ...
    var docurl = lang+'-'+testType;
    
    // each row
    var row = 5;
    for (var j = 0; j < usingCriteria.length; j++) {
      if (testType == 'wcag20' && criteria21.indexOf(usingCriteria[j][1]) >= 0) continue;
      if (usingCriteria[j][0].length > level.length) continue;

      var langPointer = testType == 'wcag21' ? usingCriteria[j][4] : usingCriteria[j][3];
      langPointer = testType == 'wcag21' && lang == 'ja' ? usingCriteria[j][3] : langPointer;
      var urlPointer = docurl;
      urlPointer = testType == 'wcag21' && lang == 'ja' ? 'en-wcag21' : urlPointer;
      var url = urlbase['understanding'][urlPointer]+langPointer;
      if (lang == 'ja' && testType == 'wcag21' && criteria21.indexOf(usingCriteria[j][1]) >= 0) {
        url = urlbase['understanding']['en-wcag21']+usingCriteria[j][4];
      }
      
      originalSheet.getRange(row, 1).setValue('=HYPERLINK("'+url+'", "'+usingCriteria[j][1]+'")').setHorizontalAlignment('center');
      originalSheet.getRange(row, 2).setDataValidation(pullDown).setHorizontalAlignment('center').setComment(usingCriteria[j][2]);
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
  }
  
  // copy sheets
  if (urls.length > 1) {
    for(var i = 1; i < urls.length; i++) {
      if (addSheet(urls[i], originalSheet) == false) {
        alreadyExists.push(urls[i]);
        continue;
      }
      res = getHtmlAndTitle(urls[i]);
      activeSheet = ss.getActiveSheet();
      activeSheet.getRange(2, 2).setValue(urls[i]);
      activeSheet.getRange(3, 2).setValue(res['title']);
      saveHtml(resourceFolderName, urls[i], res['html']);
      added++;
    }
  }

  // clean up sheet
  cleanUpSheet();
  
  // generate result sheet
  generateResultSheet();
  
  // return to original sheet
  originalSheet.activate();
  
  var msg = [];
  msg.push(getUiLang('sheet-generated', "%s sheet(s) generated").replace('%s', added));
  if (alreadyExists.length > 0) {
    msg.push(getUiLang('sheet-already-exists', "%s sheet(s) were already exists:<br>").replace('%s', alreadyExists.length)+alreadyExists.join('<br>'));
  }

  return({'msg': msg.join('<br>'), 'targetId': targetId});
}

/**
 * Generate Config Sheet
 * @param String lang
 * @param String testType
 * @param String level
 * @return String
 */
function generateConfigSheets(lang, testType, level) {
  var ss = getSpreadSheet();
  var configSheet = ss.getSheetByName(configSheetName);
  if (configSheet) return getUiLang('config-sheet-already-exists', "Config Sheet is already exists.");
  ss.insertSheet(configSheetName, 0);
  var configsheet = ss.getSheetByName(configSheetName);
  configsheet.activate();
  setBasicValue(configsheet, lang, testType, level);
  configsheet.getRange(2, 1).setValue('Name').setBackground(labelColor);
  configsheet.getRange(3, 1).setValue('Report Date').setBackground(labelColor);
  return getUiLang('generate-config-sheet', "Generate Config Sheet.");
}

/**
 * Generate Result Sheet
 */
function generateResultSheet() {
  var ss = getSpreadSheet();
  var resultSheet = ss.getSheetByName(resultSheetName);
  if (resultSheet) {
    ss.deleteSheet(resultSheet);
  }
  ss.insertSheet(resultSheetName, 0);
}

/**
 * Clean up Sheet
 */
function cleanUpSheet() {
  var ss = getSpreadSheet();
  var defaultSheetNames = ['シート1', 'sheet1']; // There must be better code... :(
  for (i = 0; i < defaultSheetNames.length; i++) {
    if (ss.getSheetByName(defaultSheetNames[i])) {
      ss.deleteSheet(ss.getSheetByName(defaultSheetNames[i]));
    }
  }
}

/**
 * Get Sheets Names
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
 * Delete Sheets
 */
function deleteSheets(urlstr) {
  var ss = getSpreadSheet();

  var urls = urlstr.replace(/^\s+|\s+$|\n\n/g,'').split(/\n/);
  var count = 0;
  for (var i = 0; i < urls.length; i++) {
    if (urls[i].charAt(0) == '*') continue;
    var targetSheet = ss.getSheetByName(urls[i]);
    if (targetSheet == null) continue;
    ss.deleteSheet(targetSheet);
    count++;
  }

  generateResultSheet();
  return getUiLang('sheet-deleted', "%s sheet(s) deleted.").replace("%s", count);
}

/**
 * add screenshot
 * @param Object formObj
 */
function screenshotUpload(formObj) {
  var activeSheet = getActiveSheet();
  var activeRow = activeSheet.getActiveCell().getRow();
  Logger.log(activeSheet.getName());
  if (activeSheet.getName().charAt(0) == '*') return getUiLang('current-sheet-is-not-for-webpage', "Current Sheet is not for webpage.");
  
  var file = fileUpload(formObj);
  activeSheet.getRange(2, 6).setValue(file[0])
  activeSheet.getRange(2, 7).setValue('=IMAGE("https://drive.google.com/uc?export=download&id='+file[1]+'",1)')
  
  return getUiLang('screenshot-uploaded', "Screenshot Uploaded.");
}
