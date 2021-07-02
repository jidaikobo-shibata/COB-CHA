/**
 * Tabulation for COB-CHA
 * functions:
 * - evaluate
 * - setCellConditionTF
 * - setCellConditionLv
 * - pageResultFormula
 * - criteriaFormula
 * - generateToatalSheet
 */

/**
 * setCellConditionTF
 * @param Object sheet
 * @param Object range
 * @param String mT
 * @param String mF
 * @return Object
 */
function setCellConditionTF(sheet, range, mT, mF) {
  var ruleForF = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mF)
      .setBackground(gFalseColor)
      .setRanges([range])
      .build();
  var ruleForT = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mT)
      .setBackground(gTrueColor)
      .setRanges([range])
      .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(ruleForF);
  rules.push(ruleForT);
  sheet.setConditionalFormatRules(rules);
}

/**
 * setCellConditionLv
 * @param Object sheet
 * @param Object range
 * @param String level
 * @return Object
 */
function setCellConditionLv(sheet, range, level) {
  var targetText = 'A';
  var targetText = level.length > 1 ? 'AA' : targetText;
  var targetText = level.length > 2 ? 'AAA' : targetText;
  var ruleForResult = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(targetText)
      .setBackground(gTrueColor)
      .setRanges([range])
      .build();
  var ruleForNI = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('NI')
      .setBackground(gFalseColor)
      .setRanges([range])
      .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(ruleForResult);
  rules.push(ruleForNI);
  sheet.setConditionalFormatRules(rules);
}

/**
 * evaluate Sc
 * @return String
 */
function evaluateSc() {
  var ss = getSpreadSheet();
  var allSheets = getAllSheets();
  if (allSheets.length == 0) {
     throw new Error(getUiLang('no-target-page-exists', "No Target Page Exists."));
  }
  
  var lang  = getProp('lang');
  var type  = getProp('type');
  var level = getProp('level');

  // enable iterative calculation
  if (ss.isIterativeCalculationEnabled() == false) {
    ss.setIterativeCalculationEnabled(true);
    ss.setMaxIterativeCalculationCycles(50); // default value
    ss.setIterativeCalculationConvergenceThreshold(0.05); // default value
  }
  
  // activate and reset sheet
  var sheet = generateSheetEvenIfAlreadyExists(gResultSheetName);
  sheet.activate();
  sheet.clear();
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);
  sheet.getRange("1:1").setFontSize(8);
  sheet.getRange("1:1").setHorizontalAlignment('center');
  sheet.getRange("2:1").setFontSize(8);
  sheet.getRange("2:1").setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 70);
  sheet.getRange(1, 1).setValue('PAGE');
  sheet.getRange(1, 2).setValue(getUiLang('result', 'Result'));
  sheet.setColumnWidth(2, 35);
  
  // header
  var criteria = type == 'tt20' ? getUsingCriteria('wcag20') : getUsingCriteria();
  var headers = [[], []];
  for (var i = 0; i < criteria.length; i++) {
    headers[0].push(criteria[i][1]);
    headers[1].push(criteria[i][0]);
  }
  var col = headers[0].length;
  var maxCol = col + 2;
  var countLv = function (maxCol, lv) {
    return '=COUNTIF(INDIRECT(ADDRESS(2, 3)&":"&ADDRESS(2, '+maxCol+')), "'+lv+'")'
  }

  headers[0] = headers[0].concat(['NI', 'A', 'AA', 'AAA']);
  headers[1] = headers[1].concat(['', countLv(maxCol, "A"), countLv(maxCol, "AA"), countLv(maxCol, "AAA")]);

  sheet.getRange(1, 3, 2, col + 4).setValues(headers);
  sheet.getRange(1, maxCol + 1, 2, 4).setBackground(gLabelColor);
  sheet.setColumnWidths(3, col + 4, 30);
  maxCol = maxCol + 4;

  // mark
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var mD = mark[1];

  // formulras
  var startRow = 3;
  var num = allSheets.length + startRow;
  var criteriaF = criteriaFormula(maxCol, num, mF);
  var tabulateF = tabulateFormula(maxCol, level);

  // tabulate
  var tabulate = ['Total', tabulateF];

  // each row
  var vals = [];

  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().charAt(0) == '*') continue;
    var each = [];
    var targetUrl = getUrlFromSheet(allSheets[i]);
    var targetSheet = allSheets[i].getName();
    each.push('=HYPERLINK("#gid='+allSheets[i].getSheetId()+'","'+targetSheet+'")');
    each.push(tabulateF);

    var chks = fetchEachResults(type, allSheets[i], criteria, mT, mF, mD);
    for (var j = 0; j < chks.length; j++) {
      each.push(chks[j]);
      var tabcol = j + 2;
      tabulate[tabcol] = tabulate[tabcol] === undefined ? chks[j] : tabulate[tabcol] ;
      tabulate[tabcol] = chks[j][0] == mF ? mF : tabulate[tabcol];
    }
    
    each = each.concat(criteriaF);
    
    vals.push(each);
  }

  // tabulation row
  tabulate = tabulate.concat(criteriaF);
  vals.push(tabulate);
  
  // set vals
  sheet.getRange(3, 1, vals.length, vals[0].length).setValues(vals);
  sheet.getRange(3, 1, vals.length, vals[0].length).setHorizontalAlignment('center');
  
  // conditioned cell for T or F
  var range = sheet.getRange(3, 3, allSheets.length + 1, col);
  setCellConditionTF(sheet, range, mT, mF);

  // conditioned cell for level
  var range = sheet.getRange(3, 2, allSheets.length + 1, 1);
  setCellConditionLv(sheet, range, level);

  // total
  generateToatalSheet();
  
  return getUiLang('evaluated', 'Evaluated.');
}

/**
 * fetchEachResults
 * @param String type
 * @param Object sheet
 * @param Array criteria
 * @return String mT
 * @return String mF
 * @return String mD
 * @return Array
 */
function fetchEachResults(type, sheet, criteria, mT, mF, mD) {
  // WCAG 2.0/2.1
  if (type != 'tt20') {
    return sheet.getRange(5, 2, criteria.length, 1).getValues();
  }  
  
  // Trusted Tester
  var chks = sheet.getRange(5, 1, getUsingCriteria().length, 2).getValues();
  var tmp = {};
  var tmp2 = {};
  var ret = [];
  
  for (var i = 0; i < chks.length; i++) {
    var testId = chks[i][0];
    var chk = chks[i][1];
    for (var key in gRelTtAndCriteria) {
      if (gRelTtAndCriteria[key].indexOf(testId) == -1) continue;
      if (typeof tmp[key] === "undefined") tmp[key] = [];
      tmp[key].push(chk);
    }
  }

  // union  
  for (var key in tmp) {
    // at least one Fail found
    if (tmp[key].indexOf(mF) >= 0) {
      tmp2[key] = mF;
      continue;
    }

    // No comformance and No N/A (find blank and '?')
    if (tmp[key].indexOf(mT) == -1 && tmp[key].indexOf(mD) == -1) {
      tmp2[key] = '?';
      continue;
    }
    
    // No comformance and N/A (N/A Only)
    if (tmp[key].indexOf(mT) == -1 && tmp[key].indexOf(mD) >= 0) {
      tmp2[key] = mD;
      continue;
    }

    tmp2[key] = mT;
  }
  
  // order
  for (var key in gRelTtAndCriteria) {
    ret.push(tmp2[key]);
  }

  return ret;
}

/**
 * tabulateFormula
 * @param Interger maxCol
 * @param String level
 * @return String
 */
function tabulateFormula(maxCol, level) {
  var niCol = maxCol - 3;
  var aCol = maxCol - 2;
  var aaCol = maxCol - 1;

  // non-interference
  var str = '=IF(INDIRECT(ADDRESS(ROW(), '+niCol+')) = "NI", "NI",';
  
  // each condition
  var singleACond = 'INDIRECT(ADDRESS(ROW(), '+aCol+')) = 0';
  var doubleACond = 'INDIRECT(ADDRESS(ROW(), '+aaCol+')) = 0';
  var tripleACond = 'INDIRECT(ADDRESS(ROW(), '+maxCol+')) = 0';
  var singleAF = 'IF('+singleACond+', "A", "A-")';
  var doubleAF = 'IF(AND('+singleACond+', '+doubleACond+'), "AA",';
  doubleAF+= 'IF(AND('+singleACond+', INDIRECT(ADDRESS(ROW(), '+aaCol+')) <= -1), "AA-", '+singleAF+'))';
  
  // A
  if (level.length == 1) {
    str+= singleAF+')';
    return str;
  }

  // AA
  if (level.length == 2) {
    str+= doubleAF+')';
    return str;
  }

  // single AA
  if (level.length == 3) {
    str+= 'IF(AND('+singleACond+', '+doubleACond+', '+tripleACond+'), "AAA",';
    str+= 'IF(AND('+singleACond+', '+doubleACond+', INDIRECT(ADDRESS(ROW(), '+maxCol+')) <= -1), "AAA-",';
    str+= doubleAF+')';
    return str;
  }
}

/**
 * criteriaFormula
 * @param Interger maxCol
 * @param Interger num
 * @param String mF
 * @return Array
 */
function criteriaFormula(maxCol, num, mF) {
  var rets = [];
  var cntcol = maxCol - 4;
  var singleAcol = maxCol - 2;
  
  // Non-Interference
  var aRows = '1:'+num;
  var niExpressions = [];
  var nonInterference = gNonInterference;
  for (var j = 0; j < nonInterference.length; j++) {
    niExpressions[j] = 'HLOOKUP("'+nonInterference[j].toString()+'", '+aRows+', ROW(), false) = "'+mF+'"';
  }
  var niExpression = 'OR('+niExpressions.join(', ')+')';
  rets.push('=IF('+niExpression+', "NI", "")'); // do not use mT as "NI is OK"
  
  var countF = function(cntcol, mF, targetLevelCol) {
    var str = '=COUNTIFS(';
    str+= 'INDIRECT(ADDRESS(ROW(), 3)&":"&ADDRESS(ROW(), '+cntcol+')), "'+mF+'",'; // count ""
    str+= 'INDIRECT("C2:"&ADDRESS(2, '+cntcol+')), INDIRECT(ADDRESS(1, '+targetLevelCol+')))';
    return str+'*-1';
  }
  
  rets.push(countF(cntcol, mF, singleAcol));
  rets.push(countF(cntcol, mF, singleAcol + 1));
  rets.push(countF(cntcol, mF, singleAcol + 2));
  return rets;
}

/**
 * generate toatal sheet
 * @return Void
 */
function generateToatalSheet() {
  var ss = getSpreadSheet();
  var allSheets = getAllSheets();
  var lang = getProp('lang');
  var type = getProp('type');
  var level = getProp('level');

  // activate and reset sheet
  var sheet = generateSheetEvenIfAlreadyExists(gTotalSheetName);
  sheet.activate();
  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setValue(getUiLang('criterion', 'Criterion'));
  sheet.getRange(1, 2).setValue(getUiLang('name', 'Name'));
  sheet.getRange(1, 3).setValue(getUiLang('level', 'Level'));
  sheet.getRange(1, 4).setValue(getUiLang('result', 'Result'));
  sheet.getRange(1, 5).setValue(getUiLang('achievementDna', 'Achievement (DNA)'));
  sheet.getRange("1:1").setBackground(gLabelColorDark).setFontColor(gLabelColorDarkText).setFontWeight('bold');
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 40);
  sheet.setColumnWidth(4, 40);
  
  var resultSheet = ss.getSheetByName(gResultSheetName);
  var criteria = type == 'tt20' ? getUsingCriteria('wcag20') : getUsingCriteria();
  var chks = resultSheet.getRange(1, 3, resultSheet.getLastRow() - 1, criteria.length).getValues();
  var totalResult = resultSheet.getRange(resultSheet.getLastRow(), 2, resultSheet.getLastRow(), 1).getValue();
  
  var transpose = function transpose(a) {
    return Object.keys(a[0]).map(function (c) {
      return a.map(function (r) {
        return r[c];
      });
    });
  }
  var rawrows = transpose(chks);

  var criteriaName = {};
  for(var i = 0; i < criteria.length; i++) {
    var key = criteria[i][1];
    criteriaName[key] = criteria[i][2];
  }
  
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var mD = mark[1];

  var rows = [];
  for(var i = 0; i < rawrows.length; i++) {
    rows[i] = [];
    var eachCriterion = rawrows[i][0];
    rows[i][0] = eachCriterion; // criterion
    rows[i][1] = criteriaName[eachCriterion]; // name
    rows[i][2] = rawrows[i][1]; // level
    rawrows[i].shift();
    rawrows[i].shift();

    var counts = {};
    var numOfcriterion = rawrows[i].length;
    counts[mT] = 0;
    counts[mD] = 0;
    counts[mF] = 0;
    counts['YET'] = 0;
    for(var j = 0; j < numOfcriterion; j++) {
      var key = rawrows[i][j];
      key = key === "" || key === "?" ? 'YET' : key ;
      counts[key] = (counts[key]) ? counts[key] + 1 : 1 ;
    }
    
    // set value by counting
    rows[i][4] = numOfcriterion + '/' + numOfcriterion;
    if (counts[mF] >= 1) {
      rows[i][3] = mF;
      rows[i][4] = counts[mT] + '/' + numOfcriterion;
    } else if (counts[mD] == numOfcriterion) {
      rows[i][3] = mD;
    } else {
      rows[i][3] = mT;
    }

    // Overwrite if BLANK or ? Exists
    if (counts['YET'] >= 1) {
      rows[i][3] = "?";
    }

    var dna = (counts[mD]) ? counts[mD] : 0;
    rows[i][4] = rows[i][4] + ' (' + dna + ')';
  }
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(rows.length + 2, 1).setValue('Total');
  sheet.getRange(rows.length + 2, 4).setValue(totalResult);
  sheet.getRange(2, 1, rows.length + 1, 4).setHorizontalAlignment('center');
  sheet.getRange(2, 2, rows.length + 1, 1).setHorizontalAlignment('left');

  // conditioned cell fot T 0r F
  var range = sheet.getRange(2, 4, rows.length, 1);
  setCellConditionTF(sheet, range, mT, mF)

  // conditioned cell for level
  var range = sheet.getRange(rows.length + 2, 4, 1, 1);
  setCellConditionLv(sheet, range, level);
}
