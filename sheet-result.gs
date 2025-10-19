/**
 * Result sheet for COB-CHA
 * functions:
 * - setCellConditionLv
 * - evaluateSc
 * - pageResultFormula
 * - criteriaFormula
 * - generateTotalSheet
 */

/**
 * setCellConditionLv
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {string} level - "A" | "AA" | "AAA"
 * @return {void}
 */
function setCellConditionLv(sheet, range, level) {
  // Derive the targetText based on level length
  /** @type {string} */
  const targetText = (level && level.length >= 3) ? 'AAA' :
                     (level && level.length === 2) ? 'AA' : 'A';

  const ruleForResult = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(targetText)
    .setBackground(gTrueColor)
    .setRanges([range])
    .build();

  const ruleForNI = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('NI')
    .setBackground(gFalseColor)
    .setRanges([range])
    .build();

  // Replace existing rules that target the same range & same condition (avoid duplication)
  const a1 = range.getA1Notation();
  const rules = sheet.getConditionalFormatRules().filter(function (r) {
    try {
      // keep rules that are not for this A1 or not equal to our 2 rules' conditions
      const rs = r.getRanges().map(function (rg) { return rg.getA1Notation(); });
      const touches = rs.indexOf(a1) !== -1;
      if (!touches) return true; // different range is kept
      // drop any rule that sets background for 'NI' or targetText on this range
      const cond = (r.getBooleanCondition && r.getBooleanCondition()) || null;
      const txt = (cond && cond.getText()) || '';
      return !(txt === 'NI' || txt === targetText);
    } catch (e) {
      return true;
    }
  });

  rules.push(ruleForResult, ruleForNI);
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
      if (tabulate[tabcol] != mF) {
        tabulate[tabcol] = chks[j][0] == mT ? mT : tabulate[tabcol];
      }
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
  generateTotalSheet();
  
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
  // WCAG 2.0/2.1/2.2
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

  // AAA
  if (level.length == 3) {
    str+= 'IF(AND('+singleACond+', '+doubleACond+', '+tripleACond+'), "AAA",';
    str+= 'IF(AND('+singleACond+', '+doubleACond+', INDIRECT(ADDRESS(ROW(), '+maxCol+')) <= -1), "AAA-",';
    str+= doubleAF+')))';
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
 * generate total sheet
 * @return Void
 */
function generateTotalSheet() {
  var ss = getSpreadSheet();
  var allSheets = getAllSheets();
  var lang = getProp('lang');
  var type = getProp('type');
  var level = getProp('level');

  // Activate and reset the total sheet
  var sheet = generateSheetEvenIfAlreadyExists(gTotalSheetName);
  sheet.activate();
  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setValue(getUiLang('criterion', 'Criterion'));
  sheet.getRange(1, 2).setValue(getUiLang('name', 'Name'));
  sheet.getRange(1, 3).setValue(getUiLang('level', 'Level'));
  sheet.getRange(1, 4).setValue(getUiLang('applied', 'Applied'));
  sheet.getRange(1, 5).setValue(getUiLang('result', 'Result'));
  sheet.getRange(1, 6).setValue(getUiLang('note', 'Note'));
  sheet.getRange("1:1").setBackground(gLabelColorDark).setFontColor(gLabelColorDarkText).setFontWeight('bold');
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(4, 60); // Applied
  sheet.setColumnWidth(5, 60); // Result
  sheet.setColumnWidth(6, 120); // Note

  var resultSheet = ss.getSheetByName(gResultSheetName);
  var criteria = (type == 'tt20') ? getUsingCriteria('wcag20') : getUsingCriteria();

  // chks: matrix where each row is [criterion, level, ...per-target marks...]
  var chks = resultSheet.getRange(1, 3, resultSheet.getLastRow() - 1, criteria.length).getValues();
  // totalResult: overall result mark (o/x/-/?) placed at the last row, col B on result sheet
  var totalResult = resultSheet.getRange(resultSheet.getLastRow(), 2, 1, 1).getValue();

  // transpose rows -> per criterion arrays
  var transpose = function transpose(a) {
    return Object.keys(a[0]).map(function (c) {
      return a.map(function (r) { return r[c]; });
    });
  };
  var rawrows = transpose(chks);

  var criteriaName = {};
  for (var i = 0; i < criteria.length; i++) {
    var key = criteria[i][1];
    criteriaName[key] = criteria[i][2];
  }

  var mark = getProp('mark');
  var mT = mark[2]; // "o"
  var mF = mark[3]; // "x"
  var mD = mark[1]; // "-"

  var rows = [];
  for (var r = 0; r < rawrows.length; r++) {
    rows[r] = [];

    // 1) Header cells per row
    var eachCriterion = rawrows[r][0];    // criterion key
    rows[r][0] = eachCriterion;           // Criterion
    rows[r][1] = criteriaName[eachCriterion]; // Name
    rows[r][2] = rawrows[r][1];           // Level

    // Strip header columns to get per-target values only
    rawrows[r].shift(); // criterion
    rawrows[r].shift(); // level

    var values = rawrows[r];              // marks per target: o/x/-/?/''
    var totalTargets = values.length;

    // 2) Count marks
    var counts = {};
    counts[mT] = 0;        // conforming
    counts[mF] = 0;        // non-conforming
    counts[mD] = 0;        // not applicable
    counts['?'] = 0;       // unknown
    counts['BLANK'] = 0;   // empty input

    for (var j = 0; j < totalTargets; j++) {
      var key = values[j];
      if (key === '') key = 'BLANK';
      counts[key] = (counts[key] || 0) + 1;
    }

    // 3) Applied count = total targets excluding N/A ("-")
    var appliedCount = totalTargets - (counts[mD] || 0);

    // 4) Applied cell (col 4): if appliedCount == 0 -> "-", else "o"
    rows[r][3] = (appliedCount === 0) ? mD : mT;

    // 5) Result cell (col 5)
    //    Priority:
    //      a) If appliedCount == 0 -> treated as "conforming without applicability" => "o"
    //      b) If any unknown or blank exists among applicable targets -> "?"
    //      c) If any non-conforming exists -> "x"
    //      d) Else if any conforming exists -> "o"
    //      e) Else "?" (safety net)
    var resultMark;
    if (appliedCount === 0) {
      resultMark = mT; // conforming because no applicable targets
    } else {
      var unknownOrBlank = (counts['?'] || 0) + (counts['BLANK'] || 0);
      if (unknownOrBlank > 0) {
        resultMark = '?';
      } else if ((counts[mF] || 0) > 0) {
        resultMark = mF;
      } else if ((counts[mT] || 0) > 0) {
        resultMark = mT;
      } else {
        resultMark = '?';
      }
    }
    rows[r][4] = resultMark;

    // 6) Note cell (col 6): leave blank as requested
    rows[r][5] = '';
  }

  // Write rows
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // Footer "Total"
  sheet.getRange(rows.length + 2, 1).setValue('Total');
  // Put overall result in the Result column (col 5)
  sheet.getRange(rows.length + 2, 5).setValue(totalResult);

  // Alignments
  sheet.getRange(2, 1, rows.length + 1, 6).setHorizontalAlignment('center');
  sheet.getRange(2, 2, rows.length + 1, 1).setHorizontalAlignment('left');

  // Conditional formatting for Result (o / x)
  var rangeTF = sheet.getRange(2, 5, rows.length, 1);
  setCellConditionTF(sheet, rangeTF, mT, mF);

  // Conditional formatting for total level (if needed)
  var rangeLv = sheet.getRange(rows.length + 2, 5, 1, 1);
  setCellConditionLv(sheet, rangeLv, level);

  // Center-align headers for Level, Applied, Result
  sheet.getRange(1, 3, 1, 3).setHorizontalAlignment('center');

  // generate Report Sheet
  generateReportSheet(level, totalResult);
}

/**
 * generate Report Sheet (non-destructive)
 * @param {String} level
 * @param {String} totalResult
 * @return {String}
 */
function generateReportSheet(level, totalResult) {
  // Pre-build multi-line default text for "Way to choose"
  var wayToChooseText = [
    getUiLang('report-way-to-choose1', 'Per web page'),
    getUiLang('report-way-to-choose2', 'All web pages selected'),
    getUiLang('report-way-to-choose3', 'XX pages selected by random sampling'),
    getUiLang('report-way-to-choose4', 'XX pages selected as representative of the entire website'),
    getUiLang('report-way-to-choose5', 'XX pages selected by random sampling and XX pages as representative')
  ].join('\n');

  // Default rows for a freshly created report sheet
  var defaults = [
    [getUiLang("report-declaration-day", "Declaration day"), ""],
    [getUiLang("report-standard-version", "Standard's version"), ""],
    [getUiLang("report-target-level", "Target level"), level],
    [getUiLang("report-gained-level", "Gained level"), totalResult],
    [getUiLang("report-explanation-pages", "Explanation of pages"), ""],
    [getUiLang("report-depending-tech", "Technology in depend"), ""],
    [getUiLang("report-way-to-choose", "Way to choose"), wayToChooseText],
    [getUiLang("report-urls-pages", "Pages' urls"), getUiLang("report-another-report", "Another sheet")],
    [getUiLang("report-evaluate-sc", "Success Criteria Check Sheet"), getUiLang("report-another-report", "Another sheet")],
    [getUiLang("report-test-days", "Test date"), ""]
  ];

  // Create the sheet only if it does not exist
  var msgOrSheet = generateSheetIfNotExists(gReportSheetName, defaults /*, header*/);

  // If a string is returned, the target sheet already exists; do nothing further
  if (typeof msgOrSheet === 'string') {
    return msgOrSheet; // e.g., "Target sheet (...) is already exists."
  }

  // Newly created: apply presentation only once here
  var title = getUiLang('report-jis-title', 'Accessibility Conformance Report');
  var sheet = msgOrSheet; // prepareTargetSheet(...) returned the created Sheet object
  sheet.insertRows(1); // shift all rows down
  sheet.getRange(1, 1, 1, 2).merge();
  sheet.getRange(1, 1)
       .setValue(title)
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');

  // Apply presentation styles
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 450);
  sheet.getRange(2, 1, defaults.length, 2)
       .setFontWeight('bold')
       .setVerticalAlignment('top')
       .setWrap(true);
  sheet.getRange("A2:A" + (defaults.length + 1)).setHorizontalAlignment('left');
  sheet.getRange("B2:B" + (defaults.length + 1)).setHorizontalAlignment('left');

  return getUiLang(
    'target-sheet-generated',
    "Generated new Report Sheet (%s)."
  ).replace('%s', gReportSheetName);
}
