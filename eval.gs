/**
 * Evaluate for COB-CHA
 */

/**
 * get Google Spreadsheet Url
 * @param String ssId
 * @param String currentId
 * @return String
 */
function getGoogleSpreadsheetUrl(ssId, currentId) {
  return "https://docs.google.com/spreadsheets/d/"+ssId+"/edit#gid="+currentId;
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
  var activeSheet = ss.getSheetByName(resultSheetName);
  activeSheet.activate();
  activeSheet.clear();
  activeSheet.setFrozenRows(2);
  activeSheet.setFrozenColumns(2);
  activeSheet.getRange("2:2").setFontSize(8);
  activeSheet.getRange("2:2").setHorizontalAlignment('center');
  activeSheet.setColumnWidth(1, 70);

  // headers
  setBasicValue(activeSheet, lang, testType, level);
  var today = new Date();
  activeSheet.getRange(1, 5).setValue(getUiLang('date', 'Date')).setBackground(labelColor);
  activeSheet.getRange(1, 6).setValue(today);

  var col = 3;
  var usingCriterions = {};
  activeSheet.getRange(2, 1).setValue('URL');
  activeSheet.getRange(2, 2).setValue(getUiLang('result', 'Result'));
  activeSheet.setColumnWidth(2, 35);
  var criteria = getUsingCriteria(lang, testType, level);
  for (var i = 0; i < criteria.length; i++) {
    activeSheet.getRange(2, col).setValue(criteria[i][1]);
    activeSheet.setColumnWidth(col, 30);
    if (cCheckVal.indexOf(criteria[i][1]) == -1) {
      activeSheet.getRange(2, col).setBackground(doubleAColor);
    }
    usingCriterions[criteria[i][1]] = col;
    col++;
  }
  activeSheet.getRange(2, col).setValue('NI').setBackground(labelColor);
  activeSheet.getRange(2, col+1).setValue('A').setBackground(labelColor);
  if (level.length > 1) {
    activeSheet.getRange(2, col+2).setValue('AA').setBackground(labelColor);
  }
  if (level.length > 2) {
    activeSheet.getRange(2, col+3).setValue('AAA').setBackground(labelColor);
  }

  // alignment
  activeSheet.getRange(3, 2, allSheets.length, col+1).setHorizontalAlignment('center');

  // conditioned cell
  var conditionedRange = activeSheet.getRange(3, 3, allSheets.length, col);
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
  var rules = activeSheet.getConditionalFormatRules();
  rules.push(ruleForF);
  rules.push(ruleForT);
  activeSheet.setConditionalFormatRules(rules);

  // each row
  var row = 3;
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().charAt(0) == '*') continue;
    var targetUrl = getUrlFromSheet(allSheets[i]);
    var targetSheet = allSheets[i].getName();
    activeSheet.getRange(row, 1).setValue('=HYPERLINK("'+getGoogleSpreadsheetUrl(ss.getId(), allSheets[i].getSheetId())+'","'+targetSheet+'")');
    activeSheet.getRange(row, 1).setComment(allSheets[i].getRange(3, 2).getValue());
   
    // each check value
    for (var key in usingCriterions){
      var col = usingCriterions[key];
      activeSheet.getRange(row, col).setValue(generateExpression(testType, key, 'A'+row));
    }

    // Non-Interference
    var targetRow = row-1;
    var niExpressions = [];
    for (var j = 0; j < nonInterference.length; j++) {
      niExpressions[j] = 'HLOOKUP("'+nonInterference[j]+'", 2:'+row+', '+targetRow+', false) = "F"';
    }
    var niExpression = 'OR('+niExpressions.join(', ')+')';
    activeSheet.getRange(row, col+1).setValue('=IF('+niExpression+', "NI", "-")');

    // single-A
    var singleAExpressions = [];
    for (var j = 0; j < cCheckVal.length; j++) {
      if ((testType == 'wcag20' || testType == 'tt20') && criteria21.indexOf(cCheckVal[j]) >= 0) continue;
      singleAExpressions[j] = 'OR(HLOOKUP("'+cCheckVal[j]+'", 2:'+row+', '+targetRow+', false) = "T"';
      singleAExpressions[j] = singleAExpressions[j]+', HLOOKUP("'+cCheckVal[j]+'", 2:'+row+', '+targetRow+', false) = "DNA")';
    }
    var singleAExpression = 'IF(AND('+singleAExpressions.join(', ')+'), "A", "A-")';
    singleAExpression = '=IF('+niExpression+', "NI", '+singleAExpression+')';
    activeSheet.getRange(row, col+2).setValue(singleAExpression);
    activeSheet.getRange(row, 2).setValue(singleAExpression);
    
    // double-A
    if (level.length > 1){
      var fullAA = (testType == 'wcag20' || testType == 'tt20') ? 38 : 50 ;
      var isAPassed = 'HLOOKUP("A", 2:'+row+', '+targetRow+', false) = "A"'; // loop reference...
      var partialAAexpression = 'IF(AND('+isAPassed+', COUNTIF('+row+':'+row+', "T") + COUNTIF('+row+':'+row+', "DNA") < '+fullAA+'), "AA-", "A-")';
      var doubleAExpression = 'IF(AND('+isAPassed+', COUNTIF('+row+':'+row+', "T") + COUNTIF('+row+':'+row+', "DNA") >= '+fullAA+'), "AA", '+partialAAexpression+')';
      doubleAExpression = '=IF('+niExpression+', "NI", '+doubleAExpression+')';
      activeSheet.getRange(row, col+3).setValue(doubleAExpression);
      activeSheet.getRange(row, 2).setValue(doubleAExpression);
    }

    // triple-A
    if (level.length > 2){
      var fullAAA = (testType == 'wcag20' || testType == 'tt20') ? 61 : 78 ;
      var loouUpAA = 'HLOOKUP("AA", 2:'+row+', '+targetRow+', false)' ;
      var isAAPassed = 'IF(AND('+loouUpAA+' = "AA", COUNTIF('+row+':'+row+', "T") + COUNTIF('+row+':'+row+', "DNA") >= '+fullAA+'), "AAA-", '+loouUpAA+')';
      var tripleAexpression = '=IF(COUNTIF('+row+':'+row+', "T") + COUNTIF('+row+':'+row+', "DNA") = '+fullAAA+', "AAA", '+isAAPassed+')';
      doubleAExpression = '=IF('+niExpression+', "NI", '+tripleAexpression+')';
      activeSheet.getRange(row, col+4).setValue(tripleAexpression);
      activeSheet.getRange(row, 2).setValue(tripleAexpression);
    }

    row++;
  }
  
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
 * generateExpression
 * @param String testType
 * @param String currentCriterion
 * @param String row
 * @return String
 */
function generateExpression(testType, currentCriterion, row) {
  var ret = '';
  if (testType.indexOf('wcag') >= 0) {
    ret = '=VLOOKUP("'+currentCriterion+'", INDIRECT('+row+'&"!A:B") , 2, false)';
  } else {
    var conditions = [];
    var strs = ["DNA", "F", "T"];
    if (typeof ttCheckVal[currentCriterion] === 'undefined') return('');
    for (var j = 0; j < strs.length; j++) {
      for(var k = 0; k < ttCheckVal[currentCriterion].length; k++) {
        conditions[k] = 'VLOOKUP("'+ttCheckVal[currentCriterion][k]+'", INDIRECT('+row+'&"!A:B") , 2, false) = "'+strs[j]+'"';
      }
      if (strs[j] == 'DNA') {
        ret = 'IF(AND('+conditions.join(', ')+'), "'+strs[j]+'", "-")';
      } else {
        if (strs[j] == 'F') {
          ret = 'IF(OR('+conditions.join(', ')+'), "'+strs[j]+'", '+ret+')';
        } else {
          ret = 'IF(AND('+conditions.join(', ')+'), "'+strs[j]+'", '+ret+')';
        }
      }
    }
    ret = '='+ret;
  }
  return ret;
}

/**
 * evaluate Icl
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function evaluateIcl(lang, testType, level) {
  // generate Sheet
  var ss = getSpreadSheet();
  var iclSheet = ss.getSheetByName(iclSheetName);
  if (iclSheet) {
    ss.deleteSheet(iclSheet);
  }
  ss.insertSheet(iclSheetName, 0);
  var iclSheet = ss.getSheetByName(iclSheetName);
  iclSheet.activate();
  generateIcl(iclSheet, 2, getUsingCriteria(lang, testType, level), level);
  
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
    iclSheet.getRange(1, col).setValue('=HYPERLINK("'+getGoogleSpreadsheetUrl(ss.getId(), allSheets[i].getSheetId())+'","'+numId+'")');
    iclSheet.getRange(1, col).setComment(targetUrl);
    iclSheet.setColumnWidth(col, 40)
    allSheets[i].getRange(iclFirstRow, 2, rows, 1).copyTo(iclSheet.getRange(4, col), {contentsOnly:true});
    iclSheet.getRange(1, col, iclSheet.getLastRow(), 1).setHorizontalAlignment('center');
    numId++;
    col++;
  }

  return getUiLang('generate-icl-sheet', "Generate ICL Sheet.");
}
