/**
 * Input Utility control for COB-CHA
 * functions:
 * - applyAllToT
 * - applyScTemplate
 * - applyIclTemplate
 * - getCurrentPos
 * - doLumpEdit
 * - iclToSc
 */

/**
 * Apply Conformance(T) to All
 * @return String
 */
function applyConformanceToAll() {
  var msg = getUiLang('this-pages-all-check-will-be-overwritten', 'This Page\'s All "Check will be overwritten.');
  if(showConfirm(msg) != "OK") return getUiLang('canceled', 'canceled');
  
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
  var ss = getSpreadSheet();
  var tpl = ss.getSheetByName(gScTplSheetName);
  if (tpl == null) return getUiLang('no-template-found', 'No template exists.');

  var n = 0;
  var allSheets = getSelectedSheets();
  if (allSheets.length == 0) return getUiLang('no-target-page-exists2', 'No Target Page Exists. set "o" at "*URLs" sheet');

  var msg = getUiLang('caution-using-template', 'CAUTION: All result will be overwritten.');
  if(showConfirm(msg) != "OK") return getUiLang('canceled', 'canceled');

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
  var ss = getSpreadSheet();
  var iclTpl = ss.getSheetByName(gIclTplSheetName);
  if (iclTpl == null) return getUiLang('no-template-found', "No template exists");
  var allSheets = getSelectedSheets();
  if (allSheets.length == 0) return getUiLang('no-target-page-exists2', 'No Target Page Exists. set "o" at "*URLs" sheet');
  
  var msg = getUiLang('caution-using-template', 'This action cannot revert.');
  if(showConfirm(msg) != "OK") return getUiLang('canceled', 'canceled');
  
  var inspectSheet = allSheets[0];
  
  // find row to start
  var row2start = findFirstEmptyRow(inspectSheet);
 
  // copy
  var n = 0;
  for (var i = 0; i < allSheets.length; i++) {
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
 * Get Current Position
 * @return Array
 */
function getCurrentPos() {
  var activeSheet = getActiveSheet();
  var row = activeSheet.getActiveCell().getRow();
  var col = activeSheet.getActiveCell().getColumn();
  var vals = activeSheet.getActiveRange().getValues();
  
  // escape
  vals = vals.map(function(x){return x.map(function(y){return y.toString().replace(/;;/g, "%;%;%")})});
  vals = vals.map(function(x){return x.map(function(y){return y.toString().replace(/::/g, "%:%:%")})});
  vals = vals.map(function(x){return x.map(function(y){return y.toString().replace(/\r?\n/g, "%%")})});
  vals = vals.map(function(x){return x.join(";;")}).join("::");

  return [row, col, vals];
}

/**
 * Open Lump Edit Dialog
 * @param {Number} row
 * @param {Number} col
 * @param {String} vals
 * @param {Boolean} is_append
 * @return {Void}
 */
function openLumpEditDialog(row, col, vals, is_append) {

  // Create hidden inputs (same style as openDialogIssue)
  var html = ''
    + '<input type="hidden" id="lump-pos-row" value="' + row + '">'
    + '<input type="hidden" id="lump-pos-col" value="' + col + '">'
    + '<input type="hidden" id="lump-val" value="' + vals + '">'
    + '<input type="hidden" id="lump-is_append" value="' + (is_append ? '1' : '0') + '">';

  var title = getUiLang('lump-edit', 'Lump Edit');

  // Show dialog using the existing unified dialog function
  showDialog('ui-lump-edit', 500, 200, title, html);
}

/**
 * Lump Edit
 * @param Integer row
 * @param Integer col
 * @param String vals
 * @param bool is_append
 * @return String
 */
/*
function doLumpEdit(row, col, vals, is_append) {
  // is append value
  is_append = is_append === undefined ? false : is_append;

  // cell coordinate must be numeric
  if (row.toString().search(/^[0-9]+$/) == -1 || col.toString().search(/^[0-9]+$/) == -1) {
    return getUiLang('error-lump-edit', 'An invalid value is set in the row/column of the lump edit');
  }
  
  var allSheets = getSelectedSheets();
  if (allSheets.length == 0) {
    var msg = getUiLang('no-target-page-exists2', 'No Target Page Exists. set "o" at "*URLs" sheet');
    showAlert(msg);
    return getUiLang('canceled', 'canceled');
  }
  
  if (is_append) {
    var msg = getUiLang('template-caution-updated', 'CAUTION: All result will be updated.');
  } else {
    var msg = getUiLang('template-caution', 'CAUTION: All result will be overwritten.');
  }
  if(showConfirm(msg) != "OK") return getUiLang('canceled', 'canceled');

  // string to array
  vals = vals.split("::");
  vals = vals.map(function(x){return x.split(";;")});
  vals = vals.map(function(x){return x.map(function(y){return y.toString().replace(/%;%;%/g, ";;")})});
  vals = vals.map(function(x){return x.map(function(y){return y.toString().replace(/%:%:%/g, "::")})});
  vals = vals.map(function(x){return x.map(function(y){return y.toString().replace(/%%/g, "\n")})});
  
  // apply
  var n = 0;
  var targetvals = [];
  for (var i = 0; i < allSheets.length; i++) {
    if (is_append) {
      // append value
      targetvals = allSheets[i].getRange(row, col, vals.length, vals[0].length).getValues();
      for (var tmprow = 0; tmprow < vals.length; tmprow++) {
        for (var tmpcol = 0; tmpcol < vals[0].length; tmpcol++) {
          // ignore single char value. it must be symbol, not words
          if (targetvals[tmprow][tmpcol].toString().length == 1) continue;
          // do not append same value
          if (targetvals[tmprow][tmpcol].toString() == vals[tmprow][tmpcol].toString()) continue;
          // newline or not
          if (targetvals[tmprow][tmpcol].toString().length == 0) {
            targetvals[tmprow][tmpcol] = vals[tmprow][tmpcol].toString();
          } else {
            targetvals[tmprow][tmpcol] = targetvals[tmprow][tmpcol].toString() + "\n" + vals[tmprow][tmpcol].toString();
          }
        }
      }
      // append - update
      allSheets[i].getRange(row, col, vals.length, vals[0].length).setValues(targetvals);
    } else {
      // orverwrite - update
      allSheets[i].getRange(row, col, vals.length, vals[0].length).setValues(vals);
    }
    n++;
  }
  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}
*/

/**
 * Lump Edit with filter key
 * @param {Number} row
 * @param {Number} col
 * @param {String} vals
 * @param {Boolean} is_append
 * @param {String} filterKey
 * @return {String}
 */
function doLumpEditWithFilter(row, col, vals, is_append, filterKey) {

  // Get sheets filtered by the given header
  var sheets = getSelectedSheets(filterKey);

  if (sheets.length === 0) {
    var msg = getUiLang('no-target-page-exists2', 'No Target Page Exists. set "o" at "*URLs" sheet');
    showAlert(msg);
    return getUiLang('canceled', 'canceled');
  }

  // convert string vals into 2D array
  var arr = vals.split("::")
                .map(function(x){ return x.split(";;") })
                .map(function(x){ return x.map(function(y){
                      return y.toString()
                              .replace(/%;%;%/g, ";;")
                              .replace(/%:%:%/g, "::")
                              .replace(/%%/g, "\n");
                })});

  var n = 0;

  sheets.forEach(function(sh){
    if (is_append) {
      var targetvals = sh.getRange(row, col, arr.length, arr[0].length).getValues();

      for (var r = 0; r < arr.length; r++) {
        for (var c = 0; c < arr[0].length; c++) {
          if (targetvals[r][c].toString().length == 1) continue;
          if (targetvals[r][c].toString() == arr[r][c].toString()) continue;
          if (targetvals[r][c].toString().length == 0) {
            targetvals[r][c] = arr[r][c].toString();
          } else {
            targetvals[r][c] = targetvals[r][c].toString() + "\n" + arr[r][c].toString();
          }
        }
      }

      sh.getRange(row, col, arr.length, arr[0].length).setValues(targetvals);

    } else {
      sh.getRange(row, col, arr.length, arr[0].length).setValues(arr);
    }

    n++;
  });

  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}

/**
 * iclToSc
 * @return String
 */
function doApplyTargetIclToSc() {
  var allSheets = getSelectedSheets();
  if (allSheets.length == 0) {
    var msg = getUiLang('no-target-page-exists2', 'No Target Page Exists. set "o" at "*URLs" sheet');
    showAlert(msg);
    return getUiLang('canceled', 'canceled');
  }

  var msg = getUiLang('template-caution', 'CAUTION: All result will be overwritten.');
  if(showConfirm(msg) != "OK") return getUiLang('canceled', 'canceled');
  
  // mark
  var mark = getProp('mark');
  var mT = mark[2];
  var mF = mark[3];
  var mD = mark[1];

  var n = 0;
  var usingCriteria = getUsingCriteria();

  // find row to start  
  var row2start = findFirstEmptyRow(allSheets[0]);
  row2start++;
  for (var i = 0; i < allSheets.length; i++) {
    var lastrow = allSheets[i].getLastRow();
    if (row2start >= lastrow) continue;
    var iclResults = allSheets[i].getRange(row2start, 2, lastrow, 4).getValues();

    var ScResultsTmp = {};
    var ScApplyTmp = {};
    var ScMemoTmp = {};
    for (var j = 0; j < iclResults.length; j++) {
      var criterion = iclResults[j][3];
      if (criterion == '') continue;
      if (typeof ScResultsTmp[criterion] === "undefined") ScResultsTmp[criterion] = [];
      if (typeof ScApplyTmp[criterion] === "undefined") ScApplyTmp[criterion] = [];
      if (typeof ScMemoTmp[criterion] === "undefined") ScMemoTmp[criterion] = [];
      ScResultsTmp[criterion].push(iclResults[j][0]);
      ScApplyTmp[criterion].push(iclResults[j][1]);
      if (iclResults[j][2] != '' && ScMemoTmp[criterion].indexOf(iclResults[j][2])) {
        ScMemoTmp[criterion].push(iclResults[j][2]);
      }
    }
   
    var ScResults = [];
    var ScMemos = [];
    for (var j = 0; j < usingCriteria.length; j++) {
      var cCriterion = usingCriteria[j][1];
      if (typeof ScResultsTmp[cCriterion] === "undefined") {
        ScResults.push(['']);
        ScMemos.push(['']);
      } else {
        ScMemos.push([ScMemoTmp[cCriterion].join("\n").trim()]);

        // N/A Only
        //var arr = ScApplyTmp[cCriterion].join().match(/x/g);
        var xCount = ScApplyTmp[cCriterion].filter(function(v){ return v === 'x'; }).length;
        if (ScApplyTmp[cCriterion].length > 0 && xCount === ScResultsTmp[cCriterion].length) {
          ScResults.push([mD]);
          continue;
        }

        // not yet (no check found)
        if (ScResultsTmp[cCriterion].every(function(v){ return v === ''; })) {
          ScResults.push(['']);
          continue;
        }

        // at least one Fail found
        if (ScResultsTmp[cCriterion].indexOf(mF) >= 0) {
          ScResults.push([mF]);
          continue;
        }
        
        // No comformance and N/A
        if (ScResultsTmp[cCriterion].indexOf(mT) == -1 && ScResultsTmp[cCriterion].indexOf(mD) >= 0) {
          ScResults.push([mD]);
          continue;
        }

        ScResults.push([mT]);
      }
    }
    
    // set
    allSheets[i].getRange(5, 2, ScResults.length, 1).setValues(ScResults);
    allSheets[i].getRange(5, 4, ScMemos.length, 1).setValues(ScMemos);
    
    n++;
  }

  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}

