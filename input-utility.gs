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
  var activeSheet = getActiveSheet();
  if (activeSheet.getName() !== gTemplateSheetName && activeSheet.getName().charAt(0) === '*') {
    return getUiLang('current-sheet-is-not-for-webpage', 'Current sheet is not for webpage');
  }
  
  var vals = [];
  var mark = getProp('mark');
  for (var i = 1; i <= getUsingCriteria().length; i++) {
    vals.push([mark[2]]);
  }
  activeSheet.getRange(5, 2, vals.length, 1).setValues(vals);
      
  return getUiLang('edit-done', 'Value Edited');
}

/**
 * applyScTemplate
 * @return String
 */
function applyScTemplate() {
  var ss = getSpreadSheet();
  var tpl = ss.getSheetByName(gTemplateSheetName);
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
 * Apply ICL template
 * @return String
 */
function applyIclTemplate() {
  var ss = getSpreadSheet();
  var iclTpl = ss.getSheetByName(gIclTplSheetName);
  if (iclTpl == null) throw new Error(getUiLang('no-template-found', "No template exists"));
  var allSheets = getAllSheets();
  if (allSheets.length == 0) throw new Error(getUiLang('no-target-page-exists', "no target page exists"));
  var inspectSheet = allSheets[0];
  
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
  var n = 0;
  var allSheets = getAllSheets();
  for (i = 0; i < allSheets.length; i++) {
    allSheets[i].getRange(row, col).setValue(val);
    n++;
  }
  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}
