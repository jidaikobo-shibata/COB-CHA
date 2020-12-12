/**
 * ICL Sheet control for COB-CHA - Japanese Only
 */

/**
 * Generate ICL Sheets
 * @param String level
 * @return Void
 */
function generateIclTplSheet(level) {
  var ss = getSpreadSheet();
  var iclTplSheet = ss.getSheetByName(gIclTplSheetName);
  if (iclTplSheet) {
     throw new Error(getUiLang('icl-exists', "ICL sheet is already exists. delete it manually."));
  }
  iclTplSheet = ss.insertSheet(gIclTplSheetName, 0);
  iclTplSheet.activate();
  generateIcl(iclTplSheet, level);
  deleteFallbacksheet();
  return getUiLang('icl-tpl-generated', "ICL was generated. It's hard to customize ICL after reflecting the template. Customize first.");
}

/**
 * Generate ICL - Japanese Only
 * @param Object iclTplSheet
 * @param String level
 * @return Void
 */
function generateIcl(iclTplSheet, level) {
  // value
  var usingCriteria = getUsingCriteria('ja', 'wcag20', level);
  var iclSituation  = getLangSet('iclSituation');
  var iclTest       = getLangSet('iclTest');
  var row           = 1;
  
  for (var j = 0; j < usingCriteria.length; j++) {
    var clevel = usingCriteria[j][0];
    var cCriterion = usingCriteria[j][1];
    iclTplSheet.getRange(row, 1).setValue(cCriterion+': '+usingCriteria[j][2]);
    iclTplSheet.getRange(row+":"+row).setBackground(gLabelColorDark).setFontColor(gLabelColorDarkText).setFontWeight('bold');
    row++;
    
    //    for (let testId of Object.keys(iclSituation[cCriterion])) { // Chrome V8 
    var eachIclSituation = Object.keys(iclSituation[cCriterion]);
    for (var key in eachIclSituation) {
      var testId = eachIclSituation[key]
      if (iclSituation[cCriterion][testId] != '') {
        iclTplSheet.getRange(row, 1).setValue(iclSituation[cCriterion][testId]);
        iclTplSheet.getRange(row+":"+row).setBackground(gLabelColor);
        row++;
      }
      var eachNum = 1;
      for (var l = 0; l < iclTest[testId].length; l++) {
        iclTplSheet.getRange(row, 1).setValue(testId+'-'+eachNum);
        iclTplSheet.getRange(row, 2).setDataValidation(getPulldownMenu()).setHorizontalAlignment('center');
        iclTplSheet.getRange(row, 3).setValue(clevel).setHorizontalAlignment('center');
        iclTplSheet.getRange(row, 4).setValue(iclTest[testId][l]['implement'].join("\n"));
        iclTplSheet.getRange(row, 5).setValue(iclTest[testId][l]['test']);
        iclTplSheet.getRange(row+":"+row).setVerticalAlignment('top');
        row++;
        eachNum++;
      }
    }
  }
}

/**
 * Apply ICL Sheet
 * @return String
 */
function applyIclSheet() {
  var ss = getSpreadSheet();
  var iclTpl = ss.getSheetByName(gIclTplSheetName);
  if (iclTpl == null) throw new Error(getUiLang('no-template-found', "No template exists"));
  var allSheets = getAllSheets();
  var tpl = ss.getSheetByName(gTemplateSheetName);
  var inspectSheet = tpl;
  inspectSheet = inspectSheet == null ? allSheets[0] : inspectSheet;
  
  // find row to start
  if (inspectSheet) {
    var row2start = 1;
    var FirstCol = inspectSheet.getRange(row2start, 1).getValue();
    while (FirstCol != '') {
      row2start++;
      FirstCol = inspectSheet.getRange(row2start, 1).getValue();
    }
    row2start++;
  }

  // no sheet exists
  if (tpl) {
    allSheets.push(tpl);
  }
  if (allSheets.length == 0 ) throw new Error(getUiLang('no-target-page-exists', "No target page exists"));
 
  // copy
  var n = 0;
  for (i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getLastRow() > row2start + 5) {
      allSheets[i].deleteRows(row2start, allSheets[i].getLastRow() - row2start);
    } else {
      allSheets[i].getRange(row2start + 1, 1).setValue('');
    }
    iclTpl.getRange(1, 1, iclTpl.getLastRow(), 26).copyTo(allSheets[i].getRange(row2start, 1));
    n++;
  }

  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}
