/**
 * Input Utility control for COB-CHA
 * functions:
 * - applyAllToT
 * - applyScTemplate
 * - applyIclTemplate
 * - doLumpEdit
 */

/**
 * Apply Conformance(T) to All
 * @return String
 */
function applyConformanceToAll() {
  var msg = getUiLang('this-pages-all-check-will-be-overwritten', 'This Page\'s All "Check will be overwritten.');
  if(showConfirm(msg) == "CANCEL") return '';
  
  var activeSheet = getActiveSheet();
  if (
    activeSheet.getName() !== gScTplSheetName &&
    activeSheet.getName() !== gIclTplSheetName &&
    activeSheet.getName().charAt(0) === '*'
  ) {
    return getUiLang('current-sheet-is-not-for-webpage', 'Current sheet is not for webpage');
  }
  
  var vals = [];
  var mark = getProp('mark');

  // SC template
  if (activeSheet.getName() === gScTplSheetName || activeSheet.getName().charAt(0) !== '*') {
    for (var i = 1; i <= getUsingCriteria().length; i++) {
      vals.push([mark[2]]);
    }
    activeSheet.getRange(5, 2, vals.length, 1).setValues(vals);
  }
  
  // ICL template
  if (activeSheet.getName() === gIclTplSheetName) {
    var iclResults = activeSheet.getRange(2, 2, activeSheet.getLastRow(), 4).getValues();
    for (var j = 0; j < iclResults.length; j++) {
      if (iclResults[j][3] == '') continue;
      if (iclResults[j][1] == mark[3]) continue;
      iclResults[j][0] = mark[2];
    }
    activeSheet.getRange(2, 2, activeSheet.getLastRow(), 4).setValues(iclResults);
  }
      
  return getUiLang('edit-done', 'Value Edited');
}

/**
 * applyScTemplate
 * @return String
 */
function applyScTemplate() {
  var msg = getUiLang('caution-using-template', 'CAUTION: All result will be overwritten.');
  if(showConfirm(msg) == "CANCEL") return '';

  var ss = getSpreadSheet();
  var tpl = ss.getSheetByName(gScTplSheetName);
  if (tpl == null) return getUiLang('no-template-found', 'No template exists.');

  var n = 0;
  var allSheets = getAllSheets();
  for (i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).charAt(0) == '*') continue;
    tpl.getRange(5, 1, tpl.getLastRow() - 4, 6).copyTo(allSheets[i].getRange(5, 1));
    n++;
  }

  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}

/**
 * Find first empty row
 * @param Object sheet
 * @return Integer
 */
function findFirstEmptyRow(sheet) {
  var row2start = 1;
  if (sheet) {
    var row2start = 1;
    var FirstCol = sheet.getRange(row2start, 1).getValue();
    while (FirstCol != '') {
      row2start++;
      FirstCol = sheet.getRange(row2start, 1).getValue();
    }
    row2start++;
  }
  
  return row2start;
}

/**
 * Apply ICL template
 * @return String
 */
function applyIclTemplate() {
  var msg = getUiLang('caution-using-template', 'This action cannot revert.');
  if(showConfirm(msg) == "CANCEL") return '';

  var ss = getSpreadSheet();
  var iclTpl = ss.getSheetByName(gIclTplSheetName);
  if (iclTpl == null) throw new Error(getUiLang('no-template-found', "No template exists"));
  var allSheets = getAllSheets();
  if (allSheets.length == 0) throw new Error(getUiLang('no-target-page-exists', "no target page exists"));
  var inspectSheet = allSheets[0];
  
  // find row to start
  var row2start = findFirstEmptyRow(inspectSheet);
 
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

/**
 * Lump Edit
 * @param Integer row
 * @param Integer col
 * @param String val
 * @return String
 */
function doLumpEdit(row, col, val) {
  var msg = getUiLang('template-caution', 'CAUTION: All result will be overwritten.');
  if(showConfirm(msg) == "CANCEL") return '';

  var n = 0;
  var allSheets = getAllSheets();
  for (var i = 0; i < allSheets.length; i++) {
    allSheets[i].getRange(row, col).setValue(val);
    n++;
  }
  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}

/**
 * iclToSc
 * @param Integer col
 * @return String
 */
function doApplyIclToSc(col) {
  var msg = getUiLang('template-caution', 'CAUTION: All result will be overwritten.');
  if(showConfirm(msg) == "CANCEL") return '';
  
  // mark
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var mD = mark[1];

  var n = 0;
  var allSheets = getAllSheets();
  var usingCriteria = getUsingCriteria();

  // find row to start  
  var row2start = findFirstEmptyRow(allSheets[0]);
  row2start++;
  for (var i = 0; i < allSheets.length; i++) {
    var lastrow = allSheets[i].getLastRow();
    if (row2start >= lastrow) continue;
    var iclResults = allSheets[i].getRange(row2start, 2, lastrow, 4).getValues();
    
    var ScResultsTmp = {};
    for (var j = 0; j < iclResults.length; j++) {
      var criterion = iclResults[j][3];
      if (criterion == '') continue;
      if (typeof ScResultsTmp[criterion] === "undefined") ScResultsTmp[criterion] = [];
      ScResultsTmp[criterion].push(iclResults[j][0]);
    }
   
    var ScResults = [];
    for (var j = 0; j < usingCriteria.length; j++) {
      var cCriterion = usingCriteria[j][1];
      if (typeof ScResultsTmp[cCriterion] === "undefined") {
        ScResults.push('');
      } else {
        // at least one Fail found
        if (ScResultsTmp[cCriterion].indexOf(mF) >= 0) {
          ScResults.push([mF]);
          continue;
        }
        
        // No comformance and N/A (N/A Only)
        if (ScResultsTmp[cCriterion].indexOf(mT) == -1 && ScResultsTmp[cCriterion].indexOf(mD) >= 0) {
          ScResults.push([mD]);
          continue;
        }
        
        ScResults.push([mT]);
      }
    }

    // set
    allSheets[i].getRange(5, col, ScResults.length, 1).setValues(ScResults);
    
    n++;
  }

  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}

