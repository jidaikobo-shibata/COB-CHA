/**
 * Tabulation for COB-CHA
 */

/**
 * Generate Result Sheet
 * @return Void
 */
function generateResultSheet() {
  var ss = getSpreadSheet();
  var all = ss.getSheets();

  var resultSheet = ss.getSheetByName(resultSheetName);
  if (resultSheet && all.length != 1) {
    ss.deleteSheet(resultSheet);
  }
  
  var resultSheet = ss.getSheetByName(resultSheetName);
  if ( ! resultSheet) {
    ss.insertSheet(resultSheetName, 0);
  }
  deleteFallbacksheet();
}

/**
 * evaluate
 * @param String lang
 * @param String testType
 * @param String level
 * @return String
 */
function evaluate(lang, testType, level) {
  var ss = getSpreadSheet();
  var allSheets = getAllSheets();
  if (allSheets.length == 0) {
     throw new Error(getUiLang('no-target-page-exists', "No Target Page Exists."));
  }

  // enable iterative calculation
  if (ss.isIterativeCalculationEnabled() == false) {
    ss.setIterativeCalculationEnabled(true);
    ss.setMaxIterativeCalculationCycles(50); // default value
    ss.setIterativeCalculationConvergenceThreshold(0.05); // default value
  }
  
  // activate and reset sheet
  generateResultSheet();
  var activeSheet = ss.getSheetByName(resultSheetName);
  activeSheet.activate();
  activeSheet.clear();
  activeSheet.setFrozenRows(3);
  activeSheet.setFrozenColumns(2);
  activeSheet.getRange("2:2").setFontSize(8);
  activeSheet.getRange("2:2").setHorizontalAlignment('center');
  activeSheet.getRange("3:2").setFontSize(8);
  activeSheet.getRange("3:2").setHorizontalAlignment('center');
  activeSheet.setColumnWidth(1, 70);

  // headers
  setBasicValue(activeSheet, lang, testType, level);
  var today = new Date();
  activeSheet.getRange(1, 5).setValue(getUiLang('date', 'Date')).setBackground(labelColor);
  activeSheet.getRange(1, 6).setValue(today);

  var col = 3;
  activeSheet.getRange(2, 1).setValue('PAGE');
  activeSheet.getRange(2, 2).setValue(getUiLang('result', 'Result'));
  activeSheet.setColumnWidth(2, 35);
  
  var type4criteria = testType == 'tt20' ? 'wcag20' : testType ;
  var criteria = getUsingCriteria(lang, type4criteria, level);
  
  var headers = [[], []];
  for (var i = 0; i < criteria.length; i++) {
    headers[0].push(criteria[i][1]);
    headers[1].push(criteria[i][0]);
  }
  headers[0].push('NI');
  headers[1].push('');
  headers[0].push('A');
  headers[1].push('');
  var labelcel = 2;
  if (level.length > 1) {
    headers[0].push('AA');
    headers[1].push('');
    labelcel++;
  }
  if (level.length > 2) {
    headers[0].push('AAA');
    headers[1].push('');
    labelcel++;
  }
  var col = headers[0].length;
  activeSheet.getRange(2, 3, 2, col).setValues(headers);
  activeSheet.getRange(2, col, 2, labelcel).setBackground(labelColor);
  activeSheet.setColumnWidths(3, col, 30);
  var maxCol = col + labelcel - 1;
//doubleAColor

  // mark
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var mD = mark[1];
  
  // conditioned cell
  var conditionedRange = activeSheet.getRange(4, 3, allSheets.length, col);
  var ruleForF = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mF)
      .setBackground(falseColor)
      .setRanges([conditionedRange])
      .build();
  var ruleForT = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mT)
      .setBackground(trueColor)
      .setRanges([conditionedRange])
      .build();
  var rules = activeSheet.getConditionalFormatRules();
  rules.push(ruleForF);
  rules.push(ruleForT);
  activeSheet.setConditionalFormatRules(rules);

  // each row
  var vals = [];
  var row = 4;
  var num = allSheets.length + row;
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().charAt(0) == '*') continue;
    var each = [];
    var targetUrl = getUrlFromSheet(allSheets[i]);
    var targetSheet = allSheets[i].getName();
    each.push('=HYPERLINK("#gid='+allSheets[i].getSheetId()+'","'+targetSheet+'")');
    
    // page result
    var resultExpression = '=INDIRECT(ADDRESS(ROW(), '+maxCol+'))';
    each.push(resultExpression);
    
    var chks = allSheets[i].getRange(5, 2, criteria.length, 1).getValues();
    for (var i = 0; i < chks.length; i++) {
      each.push(chks[i][0]);
    }
    
    // Non-Interference
    var aRows = '2:'+num;
    var niExpressions = [];
    for (var j = 0; j < nonInterference.length; j++) {
      niExpressions[j] = 'HLOOKUP("'+nonInterference[j]+'", '+aRows+', ROW() - 1, false) = "'+mF+'"';
    }
    var niExpression = 'OR('+niExpressions.join(', ')+')';
    each.push('=IF('+niExpression+', "NI", "")'); // do not use mT as "NI is OK"

    // single-A
    var singleAExpressions = [];
    for (var j = 0; j < singleACriteria.length; j++) {
      if ((testType == 'wcag20' || testType == 'tt20') && criteria21.indexOf(singleACriteria[j]) >= 0) continue;
      singleAExpressions[j] = 'OR(HLOOKUP("'+singleACriteria[j]+'", '+aRows+', ROW() - 1, false) = "'+mT+'"';
      singleAExpressions[j] = singleAExpressions[j]+', HLOOKUP("'+singleACriteria[j]+'", '+aRows+', ROW() - 1, false) = "'+mD+'")';
    }
    var singleAExpression = 'IF(AND('+singleAExpressions.join(', ')+'), "A", "A-")';
    each.push('=IF('+niExpression+', "NI", '+singleAExpression+')');
    
    // double-A
    var cRow = 'INDIRECT(ADDRESS(ROW(), 3)&":"&ADDRESS(ROW(), COLUMN()))';
    if (level.length > 1){
      var fullAA = (testType == 'wcag20' || testType == 'tt20') ? 38 : 50 ;
      var isAPassed = 'HLOOKUP("A", '+aRows+', ROW() - 1, false) = "A"'; // loop reference...
      var partialAAexpression = 'IF(AND('+isAPassed+', COUNTIF('+cRow+', "'+mT+'") + COUNTIF('+cRow+', "'+mD+'") < '+fullAA+'), "AA-", "A-")';
      var doubleAExpression = 'IF(AND('+isAPassed+', COUNTIF('+cRow+', "'+mT+'") + COUNTIF('+cRow+', "'+mD+'") >= '+fullAA+'), "AA", '+partialAAexpression+')';
      each.push('=IF('+niExpression+', "NI", '+doubleAExpression+')');
    }

    // triple-A
    if (level.length > 2){
      var fullAAA = (testType == 'wcag20' || testType == 'tt20') ? 61 : 78 ;
      var loouUpAA = 'HLOOKUP("AA", '+aRows+', ROW() - 1, false)' ;
      var isAAPassed = 'IF(AND('+loouUpAA+' = "AA", COUNTIF('+cRow+', "'+mT+'") + COUNTIF('+cRow+', "'+mD+'") >= '+fullAA+'), "AAA-", '+loouUpAA+')';
      var tripleAexpression = '=IF(COUNTIF('+cRow+', "T") + COUNTIF('+cRow+', "'+mD+'") = '+fullAAA+', "AAA", '+isAAPassed+')';
      each.push('=IF('+niExpression+', "NI", '+tripleAexpression+')');
    }

    vals.push(each);
    row++;
  }
  activeSheet.getRange(4, 1, vals.length, vals[0].length).setValues(vals);
  activeSheet.getRange(4, 1, vals.length, vals[0].length).setHorizontalAlignment('center');
  
  // conditioned cell
  var targetText = 'A';
  var targetText = level.length > 1 ? 'AA' : targetText;
  var targetText = level.length > 2 ? 'AAA' : targetText;
  var conditionedRange = activeSheet.getRange(3, 2, allSheets.length, 1);
  var ruleForResult = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(targetText)
      .setBackground(trueColor)
      .setRanges([conditionedRange])
      .build();
  var ruleForNI = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('NI')
      .setBackground(falseColor)
      .setRanges([conditionedRange])
      .build();
  var rules = activeSheet.getConditionalFormatRules();
  rules.push(ruleForResult);
  rules.push(ruleForNI);
  activeSheet.setConditionalFormatRules(rules);

  return getUiLang('evaluated', 'Evaluated.');
}

/**
 * evaluate Icl
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function evaluateIcl(lang, testType, level) {
  // template not exists
  var ss = getSpreadSheet();
  var iclTplSheet = ss.getSheetByName(iclTplSheetName);
  if (iclTplSheet == null) {
     throw new Error(getUiLang('error-icl-tpl-not-exists', "ICL sheet is not exists."));
  }

  // generate Sheet
  var iclSheet = ss.getSheetByName(iclSheetName);
  if (iclSheet) {
    ss.deleteSheet(iclSheet);
  }
  iclTplSheet.activate();
  var iclSheet = ss.duplicateActiveSheet().setName(iclSheetName);
    
  iclSheet.setColumnWidth(1, 60);
  iclSheet.deleteColumn(2);
  iclSheet.setColumnWidth(2, 50);
  iclSheet.setColumnWidth(3, 50);
  iclSheet.setFrozenRows(1);
  iclSheet.setFrozenColumns(3);
  
  // detect ICL Rows
  var allSheets = getAllSheets();
  if (allSheets.length == 0) {
     throw new Error(getUiLang('no-target-page-exists', "No Target Page Exists."));
  }
  allSheets[0].activate();
  var found = false;
  var iclFirstRow = 1;
  while ( ! found) {
    if (allSheets[0].getRange(iclFirstRow, 1).getValue() != '') {
      iclFirstRow++;
      continue;
    }
    found = true;
  }
  iclFirstRow = iclFirstRow + 3;
  var iclLastRow = allSheets[0].getLastRow();
  var rows = iclLastRow - iclFirstRow;
  iclSheet.activate();

  // copy value
  var col = 5;
  var numId = 1;

  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().charAt(0) == '*') continue;
    var targetUrl = getUrlFromSheet(allSheets[i]);
    iclSheet.getRange(1, col).setValue('=HYPERLINK("#gid='+allSheets[i].getSheetId()+'","'+numId+'")');
    iclSheet.getRange(1, col).setComment(targetUrl);
    iclSheet.setColumnWidth(col, 40)
    allSheets[i].getRange(iclFirstRow, 2, rows, 1).copyTo(iclSheet.getRange(4, col), {contentsOnly:true});
    iclSheet.getRange(1, col, iclSheet.getLastRow(), 1).setHorizontalAlignment('center');
    numId++;
    col++;
  }

  return getUiLang('generate-icl-sheet', "Generate ICL Sheet.");
}
