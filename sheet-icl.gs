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
 * Generate ICL Sheets
 * @param String type
 * @return Void
 */
function generateIclTplSheet(type) {
  var defaults = [[
    "ID",
    getUiLang("check", "Check"),
    getUiLang("eliminated", "eliminated"),
    getUiLang("memo", "memo"),
    getUiLang("criterion", "criterion"),
    getUiLang("level", "level"),
    getUiLang("tech", "tech")
//    getUiLang("icl-note", "-:not applied,　o:conformance,　　x:non-conformance. \"eliminated\": test way not applied")
  ]];
  var msgOrSheetObj = generateSheetIfNotExists(gIclTplSheetName, defaults, "row");
  if (typeof msgOrSheetObj == "string") return msgOrSheetObj;
  if (generateIcl(msgOrSheetObj, type)){
    return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gIclTplSheetName);
  }
  deleteSheetIfExist(gIclTplSheetName);
  return getUiLang('error-no-icl-found', "Error: ICL template was not found.");
}

/**
 * Generate ICL
 * @param Object sheet
 * @param String type
 * @return Bool
 */
function generateIcl(sheet, type) {
  // value
  var usingCriteria = getUsingCriteria();
  var iclSituation  = getLangSet('iclSituation'+type);
  var iclTest       = getLangSet('iclTest'+type);
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
        if (type == 'Waic') {
          // WAIC
          var eachTechId = iclTest[testId][l].join("\n");
          var eachTechNames = [];
          for (var m = 0; m < iclTest[testId][l].length; m++) {
            eachTechNames.push(techNames[iclTest[testId][l][m]]+" ("+iclTest[testId][l][m]+")");
          }
          var eachTechName = eachTechNames.join("\n");
        } else {
          // COB-CHA , Icollabo
          if (type.indexOf('Cobcha') != -1) {
            var eachTechId = iclTest[testId][l][0];
          } else {
            var eachTechId = iclTest[testId][l][0].split("/").join("\n");
          }
          var eachTechName = iclTest[testId][l][1];
          isApply = iclTest[testId][l][2] ? "x" : isApply;
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
  allSheets[0].activate();
  var iclFirstRow = 1;
  while (allSheets[0].getRange(iclFirstRow, 1).getValue() != '') {
    iclFirstRow++;
  }
  iclFirstRow++;
  var iclLastRow = allSheets[0].getLastRow();
  var rows = iclLastRow - iclFirstRow;
  iclSheet.activate();

  // copy value
  var col = 6;

  for (var i = 0; i < allSheets.length; i++) {
    var eachSheetName = allSheets[i].getName();
    if (eachSheetName.charAt(0) == '*') continue;
    var targetUrl = getUrlFromSheet(allSheets[i]);
    iclSheet.getRange(1, col).setValue('=HYPERLINK("#gid='+allSheets[i].getSheetId()+'","'+eachSheetName+'")');
    iclSheet.getRange(1, col).setComment(targetUrl);
    iclSheet.setColumnWidth(col, 40)
    allSheets[i].getRange(iclFirstRow, 2, rows + 1, 1).copyTo(iclSheet.getRange(2, col), {contentsOnly:true});
    iclSheet.getRange(1, col, iclSheet.getLastRow(), 1).setHorizontalAlignment('center');
    col++;
  }

  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var range = iclSheet.getRange(5, 6, rows + 1, allSheets.length);
  setCellConditionTF(iclSheet, range, mT, mF)
 
  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gIclSheetName);
}
