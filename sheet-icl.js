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

function parseIclCriteria(value) {
  var matches = String(value || '').match(/\d+\.\d+\.\d+/g);
  return matches ? matches.filter(function (criterion, index) {
    return matches.indexOf(criterion) === index;
  }) : [];
}

function summarizeIclMarks(values, marks) {
  var unknown = marks[0];
  var notApplicable = marks[1];
  var pass = marks[2];
  var fail = marks[3];
  if (values.length === 0) return notApplicable;
  if (values.every(function (value) { return value === notApplicable; })) {
    return notApplicable;
  }
  if (values.indexOf(fail) !== -1) return fail;
  if (values.indexOf('') !== -1 || values.indexOf(unknown) !== -1) return unknown;
  if (values.indexOf(pass) !== -1) return pass;
  return unknown;
}

function calculateIclPageResults(pageRows, marks) {
  var fail = marks[3];
  var output = pageRows.map(function (row) { return [row[1]]; });
  var headingRows = [];
  var criterionValues = {};

  for (var row = 1; row < pageRows.length; row++) {
    var idOrLabel = String(pageRows[row][0] || '');
    var criteria = parseIclCriteria(pageRows[row][4]);
    var level = String(pageRows[row][5] || '');
    if (/^\d+\.\d+\.\d+:/.test(idOrLabel)) {
      var heading = {
        row: row,
        criteria: parseIclCriteria(idOrLabel).concat(criteria)
      };
      heading.criteria = heading.criteria.filter(
        function (criterion, index, values) {
          return values.indexOf(criterion) === index;
        }
      );
      headingRows.push(heading);
    } else if (level !== '' && criteria.length > 0) {
      criteria.forEach(function (criterion) {
        if (!criterionValues[criterion]) criterionValues[criterion] = [];
        if (pageRows[row][2] !== fail) {
          criterionValues[criterion].push(pageRows[row][1]);
        }
      });
    }
  }

  var criterionResults = {};
  Object.keys(criterionValues).forEach(function (criterion) {
    criterionResults[criterion] = summarizeIclMarks(
      criterionValues[criterion],
      marks
    );
  });

  headingRows.forEach(function (heading) {
    var results = {};
    heading.criteria.forEach(function (criterion) {
      results[criterion] = criterionResults[criterion] || marks[1];
    });
    output[heading.row][0] = formatIclAggregateResult(results);
  });

  return {output: output, criteria: criterionResults};
}

function aggregateIclPageResults(pageRows, marks) {
  return calculateIclPageResults(pageRows, marks).output;
}

function formatIclAggregateResult(results) {
  var criteria = Object.keys(results);
  if (criteria.length === 0) return '';
  if (criteria.length === 1) return results[criteria[0]];
  return criteria.map(function (criterion) {
    return criterion + ':' + results[criterion];
  }).join(' / ');
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

function getIclUnderstandingUrl(slug) {
  if (!slug) return '';
  return 'https://waic.jp/translations/WCAG22/Understanding/' + slug;
}

function getIclCriterionHeadingRichText(text, slug) {
  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  var url = getIclUnderstandingUrl(slug);
  if (!url) return builder.build();

  var linkStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor(gLabelColorDarkText)
    .setBold(true)
    .setUnderline(true)
    .build();
  return builder
    .setLinkUrl(0, text.length, url)
    .setTextStyle(0, text.length, linkStyle)
    .build();
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
  var usingCriteria = getUsingCriteria().slice();
  var iclSituation  = getLangSet('iclSituation'+iclType);
  var iclTest       = getLangSet('iclTest'+iclType);
  var techNames     = getLangSet('tech');
  var row           = 2;
  
  if (iclSituation.length == 0) return false;

  if (typeof iclSituation['UI-Common'] !== "undefined") {
    usingCriteria.unshift(['A', 'UI-Common', 'ユーザインタフェースコンポーネント共通', '', '']);
  }
    
  for (var j = 0; j < usingCriteria.length; j++) {
    // criterion title
    var clevel     = usingCriteria[j][0];
    var cCriterion = usingCriteria[j][1];
    if (typeof iclSituation[cCriterion] === "undefined") continue;

    var criterionHeading = cCriterion === 'UI-Common' ?
      usingCriteria[j][2] : cCriterion + ': ' + usingCriteria[j][2];
    sheet.getRange(row, 1).setRichTextValue(
      getIclCriterionHeadingRichText(criterionHeading, usingCriteria[j][4])
    );
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

        var relatedCriteria = iclTest[testId][l][3] || cCriterion;
        eachTest.push([eachTestId, "", isApply, "", relatedCriteria, clevel, eachTechId, eachTechName]);
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
  setRowConditionNotYet(sheet, range, "=AND($B2=\"\", NOT($E2=\"\"), NOT($F2=\"\"))");

  setChecklistBorders(sheet.getRange(1, 1, row - 1, 8));

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
  var allSheets = getAllSheets();
  if (allSheets.length == 0) {
     throw new Error(getUiLang('no-target-page-exists', "No Target Page Exists."));
  }

  var baseSheet = allSheets[0];
  var iclLastRow = baseSheet.getLastRow();
  var firstColumnValues = baseSheet
    .getRange(1, 1, iclLastRow, 1)
    .getValues();
  var iclFirstRow = findIclFirstResultRow(firstColumnValues);
  if (iclFirstRow == 0 || iclFirstRow > iclLastRow) {
    throw new Error(getUiLang('error-no-icl-found', "Error: ICL template was not found."));
  }
  var resultRowCount = iclLastRow - iclFirstRow + 1;

  // Use page 1 as the row structure, retaining its A, E, F, G and H columns.
  var iclSheet = generateSheetEvenIfAlreadyExists(gIclSheetName);
  baseSheet
    .getRange(iclFirstRow, 1, resultRowCount, 8)
    .copyTo(iclSheet.getRange(1, 1));
  for (var i = 0; i < 3; i++) {
    iclSheet.deleteColumn(2);
  }
  iclSheet.insertRows(1, 1);
  iclSheet.setColumnWidth(1, 100);
  iclSheet.setColumnWidth(2, 140);
  iclSheet.setColumnWidth(3, 50);
  iclSheet.setColumnWidth(4, 180);
  iclSheet.setColumnWidth(5, 600);

  var firstResultCol = 6;
  var pageCount = allSheets.length;
  var requiredColumns = firstResultCol + pageCount - 1;
  var currentMaxColumns = iclSheet.getMaxColumns();
  if (requiredColumns > currentMaxColumns) {
    iclSheet.insertColumnsAfter(
      currentMaxColumns,
      requiredColumns - currentMaxColumns
    );
  }

  var headers = [];
  var targetUrls = [];
  var resultColumns = [];
  for (var i = 0; i < pageCount; i++) {
    var pageSheet = allSheets[i];
    var sheetName = pageSheet.getName();
    var escapedSheetName = sheetName.replace(/"/g, '""');
    headers.push(
      '=HYPERLINK("#gid=' + pageSheet.getSheetId() + '","' +
      escapedSheetName + '")'
    );
    targetUrls.push(getUrlFromSheet(pageSheet));

    var pageLastRow = pageSheet.getLastRow();
    var pageFirstColumn = pageSheet
      .getRange(1, 1, pageLastRow, 1)
      .getValues();
    var pageFirstRow = findIclFirstResultRow(pageFirstColumn);
    var pageResults = [];
    if (pageFirstRow > 0 && pageFirstRow <= pageLastRow) {
      var copyRowCount = Math.min(
        resultRowCount,
        pageLastRow - pageFirstRow + 1
      );
      pageResults = pageSheet
        .getRange(pageFirstRow, 2, copyRowCount, 1)
        .getValues();
    }
    while (pageResults.length < resultRowCount) pageResults.push(['']);
    resultColumns.push(pageResults);
  }

  var headerRange = iclSheet.getRange(1, firstResultCol, 1, pageCount);
  headerRange.setValues([headers]);
  for (var i = 0; i < pageCount; i++) {
    headerRange.getCell(1, i + 1).setComment(targetUrls[i]);
  }
  headerRange
    .setBackground(gLabelColorDark)
    .setFontColor(gLabelColorDarkText)
    .setFontWeight('bold');
  iclSheet.setColumnWidths(firstResultCol, pageCount, 40);

  var resultRows = transposeIclResultColumns(resultColumns);
  var pageResultRange = iclSheet
    .getRange(2, firstResultCol, resultRowCount, pageCount);
  pageResultRange
    .setValues(resultRows)
    .setHorizontalAlignment('center');

  // Extend criterion and situation row styles through every page column.
  var baseStructure = baseSheet
    .getRange(iclFirstRow, 1, resultRowCount, 6)
    .getValues();
  var baseStyleRange = baseSheet
    .getRange(iclFirstRow, 1, resultRowCount, 1);
  var baseBackgrounds = baseStyleRange.getBackgrounds();
  var baseFontColors = baseStyleRange.getFontColors();
  var baseFontWeights = baseStyleRange.getFontWeights();
  var resultBackgrounds = pageResultRange.getBackgrounds();
  var resultFontColors = pageResultRange.getFontColors();
  var resultFontWeights = pageResultRange.getFontWeights();
  for (var row = 0; row < resultRowCount; row++) {
    var isHeadingRow = baseStructure[row][0] !== '' &&
      baseStructure[row][5] === '';
    if (!isHeadingRow) continue;
    for (var col = 0; col < pageCount; col++) {
      resultBackgrounds[row][col] = baseBackgrounds[row][0];
      resultFontColors[row][col] = baseFontColors[row][0];
      resultFontWeights[row][col] = baseFontWeights[row][0];
    }
  }
  pageResultRange
    .setBackgrounds(resultBackgrounds)
    .setFontColors(resultFontColors)
    .setFontWeights(resultFontWeights);
  iclSheet.setFrozenRows(1);
  iclSheet.setFrozenColumns(5);
  iclSheet.activate();

  var mark = getProp('mark');
  setCellConditionTF(
    iclSheet,
    pageResultRange,
    mark[2],
    mark[3]
  );
  setChecklistBorders(
    iclSheet.getRange(
      1,
      1,
      resultRowCount + 1,
      firstResultCol + pageCount - 1
    )
  );
 
  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gIclSheetName);
}
