/**
 * Sheet control for COB-CHA
 */

/**
 * Generate Sheets
 * @param String urlstr
 * @param String lang
 * @param String testType
 * @param String level
 * @return String
 */
function generateSheets(urlstr, lang, testType, level) {
  var urls = urlstr.replace(/^\s+|\s+$|\n\n/g, '').split(/\n/);
  if (urls.length == 0) return("No target Page Exists");

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
    var isEvalateTarget = urls[0].charAt(0) != '*';
    
    var originalSheet = ss.getActiveSheet();
    
    var res = isEvalateTarget ? getHtmlAndTitle(urls[0]) : false;
    
    // meta
    originalSheet.getRange(1, 1).setValue('Type').setBackground(labelColor);
    originalSheet.getRange(1, 2).setValue(lang).setHorizontalAlignment('center');
    originalSheet.getRange(1, 3).setValue(testType).setHorizontalAlignment('center');
    originalSheet.getRange(1, 4).setValue(level).setHorizontalAlignment('center');
    var today = new Date();
    originalSheet.getRange(1, 5).setValue('Date').setBackground(labelColor);
    originalSheet.getRange(1, 6).setValue(today);
    originalSheet.getRange(1, 7).setValue('Memo').setBackground(labelColor);
    originalSheet.getRange(2, 1).setValue('URL').setBackground(labelColor);
    originalSheet.getRange(2, 2).setValue(urls[0]);
    originalSheet.getRange(3, 1).setValue('title').setBackground(labelColor);
    if (res) {
      originalSheet.getRange(3, 2).setValue(res['title']);
      saveHtml(resourceFolderName, urls[0], res['html']);
    }
    originalSheet.setFrozenRows(4);
    
    // header
    originalSheet.getRange(4, 1).setValue('Criterion').setBackground(labelColor);
    originalSheet.getRange(4, 2).setValue('Check').setBackground(labelColor);
    originalSheet.getRange(4, 3).setValue('Level').setBackground(labelColor);
    originalSheet.getRange(4, 4).setValue('Techs').setBackground(labelColor);
    originalSheet.getRange(4, 5).setValue('Memo').setBackground(labelColor);
    
    // appearance
    originalSheet.setColumnWidth(1, 60);
    originalSheet.setColumnWidth(2, 50);
    originalSheet.setColumnWidth(3, 50);
    originalSheet.getRange('4:4').setHorizontalAlignment('center');
    
    // test type
    var set = testType.indexOf('wcag') >= 0 ? 'criteria' : 'ttCriteria' ;
    var usingCriteria = getLangSet(set);
    
    // each row
    var row = 5;
    for (var j = 0; j < usingCriteria.length; j++) {
      if (testType == 'wcag20' && criteria21.indexOf(usingCriteria[j][1]) >= 0) continue;
      if (usingCriteria[j][0].length > level.length) continue;
      
      originalSheet.getRange(row, 1).setValue(usingCriteria[j][1]).setHorizontalAlignment('center');
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
  msg.push(added+" sheet(s) generated");
  if (alreadyExists.length > 0) {
    msg.push(alreadyExists.length+' sheet(s) were already exists:<br>'+alreadyExists.join('<br>'));
  }

  return(msg.join('<br>'));
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
  return(count+" sheet(s) deleted");
}
