/**
 * ICL sheet for COB-CHA - Japanese Only
 * functions:
 * - getIclPassPulldownMenu
 * - getIclApplyPulldownMenu
 * - generateIclTplSheet
 * - generateIcl
 * - evaluateIcl
 */

/**
 * get ICL T or F pulldown menu
 * Japanese Only
 * @return Object
 */
function getIclPassPulldownMenu() {
  var pullDown = SpreadsheetApp.newDataValidation();
  pullDown.requireValueInList(["-", "o", "x"], true);
  return pullDown;
}

/**
 * get ICL applicability pulldown menu
 * Japanese Only
 * @return Object
 */
function getIclApplyPulldownMenu() {
  var pullDown = SpreadsheetApp.newDataValidation();
  pullDown.requireValueInList(["x"], true); // "x": eliminated
  return pullDown;
}

/**
 * Find the first ICL result row after the blank separator in column A.
 * @param Array firstColumnValues
 * @return Number
 */
function findIclFirstResultRow(firstColumnValues) {
  for (var i = 0; i < firstColumnValues.length; i++) {
    if (firstColumnValues[i][0] == '') return i + 2;
  }
  return 0;
}

/**
 * Convert per-page result columns into rows for one setValues call.
 * @param Array resultColumns
 * @return Array
 */
function transposeIclResultColumns(resultColumns) {
  if (resultColumns.length == 0) return [];
  var rowCount = resultColumns[0].length;
  var rows = [];
  for (var row = 0; row < rowCount; row++) {
    var values = [];
    for (var col = 0; col < resultColumns.length; col++) {
      values.push(resultColumns[col][row][0]);
    }
    rows.push(values);
  }
  return rows;
}

/**
 * Generate ICL Sheets
 * @param String iclType
 * @param String type
 * @return Void
 */
function generateIclTplSheet(iclType, type) {
  var defaults = [[
    "ID",
    getUiLang("check", "Check"),
    getUiLang("eliminated", "eliminated"),
    getUiLang("memo", "memo"),
    getUiLang("criterion", "criterion"),
    getUiLang("level", "level"),
    getUiLang("related-techniques", "Related Techniques")
//    getUiLang("icl-note", "-:not applied,　o:conformance,　　x:non-conformance. \"eliminated\": test way not applied")
  ]];
  var msgOrSheetObj = generateSheetIfNotExists(gIclTplSheetName, defaults, "row");
  if (typeof msgOrSheetObj == "string") return msgOrSheetObj;
  if (generateIcl(msgOrSheetObj, iclType, type)){
    return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gIclTplSheetName);
  }
  deleteSheetIfExist(gIclTplSheetName);
  return getUiLang('error-no-icl-found', "Error: ICL template was not found.");
}

/**
 * Get the Japanese WAIC technique URL for a technique ID.
 * Use WCAG 2.1 by default and WCAG 2.2 for techniques added in 2.2.
 *
 * @param String techniqueId
 * @return String
 */
function getIclTechniqueUrl(techniqueId) {
  var directories = {
    'ARIA': 'aria',
    'C': 'css',
    'F': 'failures',
    'G': 'general',
    'H': 'html',
    'PDF': 'pdf',
    'SCR': 'client-side-script',
    'SVR': 'server-side-script',
    'T': 'text'
  };
  var wcag22TechniqueIds = {
    'C41': true,
    'C42': true,
    'C43': true,
    'C45': true,
    'F108': true,
    'F109': true,
    'F110': true,
    'F111': true,
    'F112': true,
    'G218': true,
    'G219': true,
    'G220': true,
    'G221': true,
    'G223': true,
    'G224': true,
    'G225': true,
    'H99': true,
    'H100': true,
    'H101': true,
    'H102': true
  };
  var w3cOnlyTechniqueIds = {
    'ARIA25': true,
    'ARIA26': true,
    'C44': true,
    'F113': true,
    'G226': true,
    'SCR40': true
  };
  var match = String(techniqueId).match(/^([A-Z]+)\d+$/);

  if (!match || !directories[match[1]]) return '';

  if (w3cOnlyTechniqueIds[techniqueId]) {
    return 'https://www.w3.org/WAI/WCAG22/Techniques/' +
      directories[match[1]] + '/' + techniqueId;
  }

  var wcagVersion = wcag22TechniqueIds[techniqueId] ? 'WCAG22' : 'WCAG21';

  return 'https://waic.jp/translations/' + wcagVersion + '/Techniques/' +
    directories[match[1]] + '/' + techniqueId;
}

/**
 * Add links to each technique ID in the ICL technique column.
 *
 * @param Object sheet
 * @param Number firstRow
 * @param Number numRows
 * @param String excludedMark
 * @return Void
 */
function setIclTechniqueLinks(sheet, firstRow, numRows, excludedMark) {
  if (numRows < 1) return;

  var range = sheet.getRange(firstRow, 7, numRows, 1);
  var values = range.getDisplayValues();
  var excludedValues = sheet.getRange(firstRow, 3, numRows, 1).getValues();
  var richTextValues = [];
  var linkStyle = SpreadsheetApp.newTextStyle()
      .setForegroundColor('#000000')
      .setUnderline(true)
      .build();
  var excludedLinkStyle = SpreadsheetApp.newTextStyle()
      .setForegroundColor(gLabelColor)
      .setUnderline(false)
      .build();

  for (var i = 0; i < values.length; i++) {
    var text = values[i][0];
    var builder = SpreadsheetApp.newRichTextValue().setText(text);
    var pattern = /[A-Z]+\d+/g;
    var match = pattern.exec(text);
    var textStyle = excludedValues[i][0] === excludedMark ?
      excludedLinkStyle : linkStyle;

    while (match !== null) {
      var url = getIclTechniqueUrl(match[0]);
      if (url) {
        builder
          .setLinkUrl(match.index, match.index + match[0].length, url)
          .setTextStyle(match.index, match.index + match[0].length, textStyle);
      }
      match = pattern.exec(text);
    }

    richTextValues.push([builder.build()]);
  }

  range.setRichTextValues(richTextValues);
}

/**
 * Keep ICL technique link styles in sync with applicability edits.
 *
 * @param Object e
 * @return Void
 */
function onEdit(e) {
  if (!e || !e.range) return;

  var editedRange = e.range;
  var sheet = editedRange.getSheet();
  var firstColumn = editedRange.getColumn();
  var lastColumn = firstColumn + editedRange.getNumColumns() - 1;

  if (sheet.getName() !== gIclTplSheetName ||
      firstColumn > 3 || lastColumn < 3) return;

  var firstRow = Math.max(2, editedRange.getRow());
  var lastRow = Math.min(
    editedRange.getRow() + editedRange.getNumRows() - 1,
    sheet.getLastRow()
  );

  if (lastRow < firstRow) return;

  setIclTechniqueLinks(
    sheet,
    firstRow,
    lastRow - firstRow + 1,
    getProp('mark')[3]
  );
}

/**
 * Generate ICL
 * @param Object sheet
 * @param String iclType
 * @param String type
 * @return Bool
 */
function generateIcl(sheet, iclType, type) {
  // value
  var type          = type || getProp('type');
  var mark          = getProp('mark'); // ex: ['?', '-', 'o', 'x'] or ['NT','DNA','T','F']
  var mF            = mark[3]; // 'x' or 'F'
  var usingCriteria = getUsingCriteria();
  var iclSituation  = getLangSet('iclSituation'+iclType);
  var iclTest       = getLangSet('iclTest'+iclType);
  var techNames     = getLangSet('tech');
  var row           = 2;
  
  if (iclSituation.length == 0) return false;
    
  for (var j = 0; j < usingCriteria.length; j++) {
    // criterion title
    var clevel     = usingCriteria[j][0];
    var cCriterion = usingCriteria[j][1];
    if (typeof iclSituation[cCriterion] === "undefined") continue;

    sheet.getRange(row, 1).setValue(cCriterion+': '+usingCriteria[j][2]); 
    sheet.getRange(row+":"+row).setBackground(gLabelColorDark).setFontColor(gLabelColorDarkText).setFontWeight('bold');
    row++;
    
    // situation
    // Rhino
    // var eachIclSituation = Object.keys(iclSituation[cCriterion]);
    // for (var key in eachIclSituation) {
    //  var testId = eachIclSituation[key]
    // /Rhino
    for (const testId of Object.keys(iclSituation[cCriterion])) {
      if (iclSituation[cCriterion][testId] != '') {
        sheet.getRange(row, 1).setValue(iclSituation[cCriterion][testId]);
        sheet.getRange(row+":"+row).setBackground(gLabelColor);
        row++;
      }
      var eachNum = 1;
      var eachTest = [];
      for (var l = 0; l < iclTest[testId].length; l++) {

        var eachTestId = testId+'-'+eachNum;
        
        var isApply = '';
        if (iclType == 'Waic') {
          // WAIC
          var eachTechId = iclTest[testId][l].join("\n");
          var eachTechNames = [];
          for (var m = 0; m < iclTest[testId][l].length; m++) {
            eachTechNames.push(techNames[iclTest[testId][l][m]]+" ("+iclTest[testId][l][m]+")");
          }
          var eachTechName = eachTechNames.join("\n");
        } else {
          // COB-CHA , Icollabo
          if (iclType.indexOf('Cobcha') != -1) {
            var eachTechId = iclTest[testId][l][0];
          } else {
            var eachTechId = iclTest[testId][l][0].split("/").join("\n");
          }
          var eachTechName = iclTest[testId][l][1];
          isApply = iclTest[testId][l][2] ? mF : isApply;
        }
        
        // WCAG 2.2 4.1.1 sets the default ICL value to "Not Applicable"
        if (type === 'wcag22' && cCriterion === '4.1.1') {
          isApply = mF; // x or F
        }

        eachTest.push([eachTestId, "", isApply, "", cCriterion, clevel, eachTechId, eachTechName]);
        eachNum++;
      }
      sheet.getRange(row, 1, eachTest.length, 8).setValues(eachTest).setVerticalAlignment('top');
      sheet.getRange(row, 2, eachTest.length, 1).setDataValidation(getIclPassPulldownMenu()).setHorizontalAlignment('center');
      sheet.getRange(row, 3, eachTest.length, 1).setDataValidation(getIclApplyPulldownMenu()).setHorizontalAlignment('center');
      sheet.getRange(row, 5, eachTest.length, 2).setHorizontalAlignment('center');
      row = row + eachTest.length;
    }
  }

  // Link each related technique ID in one batch to avoid per-cell writes.
  setIclTechniqueLinks(sheet, 2, row - 2, mF);
  
  // appearance
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
    
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 50);

  var range = sheet.getRange(2, 2, row, 1);
  setCellConditionTF(sheet, range, mT, mF)

  var range = sheet.getRange(2, 1, sheet.getLastRow(), 26);
  setRowConditionApplicability(sheet, range, "=$C2=\""+mF+"\"");
  setRowConditionNotYet(sheet, range, "=AND($B2=\"\", NOT($E2=\"\"))");

  return true;
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
  var iclTplSheet = ss.getSheetByName(gIclTplSheetName);
  if (iclTplSheet == null) {
     throw new Error(getUiLang('icl-tpl-not-exists', "ICL sheet is not exists."));
  }

  // generate Sheet
  var iclSheet = ss.getSheetByName(gIclSheetName);
  if (iclSheet) {
    ss.deleteSheet(iclSheet);
  }
  iclTplSheet.activate();
  var iclSheet = ss.duplicateActiveSheet().setName(gIclSheetName);
    
  iclSheet.setColumnWidth(1, 60);
  iclSheet.deleteColumn(2);
  iclSheet.deleteColumn(3);
  iclSheet.setColumnWidth(2, 50);
  iclSheet.setColumnWidth(3, 50);
  iclSheet.setColumnWidth(4, 50);
  iclSheet.insertRows(1, 1);
  iclSheet.getRange('1:1').setBackground(gLabelColor).setFontColor(gLabelColorText).setFontWeight('bold');
  iclSheet.setFrozenRows(1);
  iclSheet.setFrozenColumns(5);

  // detect ICL Rows
  var allSheets = getAllSheets();
  if (allSheets.length == 0) {
     throw new Error(getUiLang('no-target-page-exists', "No Target Page Exists."));
  }
  var iclLastRow = allSheets[0].getLastRow();
  var firstColumnValues = allSheets[0]
    .getRange(1, 1, iclLastRow, 1)
    .getValues();
  var iclFirstRow = findIclFirstResultRow(firstColumnValues);
  if (iclFirstRow == 0 || iclFirstRow > iclLastRow) {
    throw new Error(getUiLang('error-no-icl-found', "Error: ICL template was not found."));
  }
  var resultRowCount = iclLastRow - iclFirstRow + 1;
  iclSheet.activate();

  // Read one result column from each page, then write the matrix once.
  var headers = [];
  var targetUrls = [];
  var resultColumns = [];
  for (var i = 0; i < allSheets.length; i++) {
    var eachSheetName = allSheets[i].getName();
    headers.push(
      '=HYPERLINK("#gid=' + allSheets[i].getSheetId() + '","' +
      eachSheetName + '")'
    );
    targetUrls.push(getUrlFromSheet(allSheets[i]));
    resultColumns.push(
      allSheets[i]
        .getRange(iclFirstRow, 2, resultRowCount, 1)
        .getValues()
    );
  }

  var firstResultCol = 6;
  var pageCount = allSheets.length;
  var headerRange = iclSheet.getRange(1, firstResultCol, 1, pageCount);
  headerRange.setValues([headers]);
  for (var i = 0; i < pageCount; i++) {
    headerRange.getCell(1, i + 1).setComment(targetUrls[i]);
  }
  iclSheet.setColumnWidths(firstResultCol, pageCount, 40);

  var resultRows = transposeIclResultColumns(resultColumns);
  iclSheet
    .getRange(2, firstResultCol, resultRowCount, pageCount)
    .setValues(resultRows);
  iclSheet
    .getRange(1, firstResultCol, iclSheet.getLastRow(), pageCount)
    .setHorizontalAlignment('center');

  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var range = iclSheet.getRange(5, firstResultCol, resultRowCount, pageCount);
  setCellConditionTF(iclSheet, range, mT, mF)
 
  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gIclSheetName);
}
