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

  for (var i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).charAt(0) == '*') continue;
    tpl.getRange(5, 1, tpl.getLastRow() - 4, 6).copyTo(allSheets[i].getRange(5, 1));
    setChecklistBorders(
      allSheets[i].getRange(4, 1, tpl.getLastRow() - 3, 4)
    );
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
  var title = getUiLang('lump-edit', 'Lump Edit');
  var templateVars = {
    'row': row,
    'col': col,
    'vals': vals,
    'isAppend': is_append ? '1' : '0'
  };

  // Show dialog using the existing unified dialog function
  showDialog('ui-lump-edit', 500, 200, title, '', templateVars);
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

  row = Number(row);
  col = Number(col);
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) {
    return getUiLang(
      'error-lump-edit',
      'An invalid value is set in the row/column of the lump edit'
    );
  }
  vals = String(vals == null ? '' : vals);

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

  var cellCount = arr.length * arr[0].length;
  var hasOversizedCell = arr.some(function (items) {
    return items.some(function (item) { return item.length > 50000; });
  });
  if (cellCount > 10000 || hasOversizedCell || sheets.some(function (sheet) {
    return row + arr.length - 1 > sheet.getMaxRows() ||
      col + arr[0].length - 1 > sheet.getMaxColumns();
  })) {
    return getUiLang('error-lump-edit', 'The lump edit range is invalid or too large.');
  }

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
function getIclScTransferRows(iclRows, usingCriteria, marks) {
  var criterionResults = calculateIclPageResults(iclRows, marks).criteria;
  var criterionMemos = {};
  iclRows.forEach(function (row) {
    if (row[2] === marks[3]) return;
    var memo = String(row[3] || '').trim();
    if (memo === '') return;
    parseIclCriteria(row[4]).forEach(function (criterion) {
      if (!criterionMemos[criterion]) criterionMemos[criterion] = [];
      if (criterionMemos[criterion].indexOf(memo) === -1) {
        criterionMemos[criterion].push(memo);
      }
    });
  });

  var results = [];
  var memos = [];
  usingCriteria.forEach(function (criterion) {
    var criterionId = criterion[1];
    results.push([criterionResults[criterionId] || '']);
    memos.push([(criterionMemos[criterionId] || []).join('\n')]);
  });
  return {results: results, memos: memos};
}

function doApplyTargetIclToSc() {
  var allSheets = getSelectedSheets();
  if (allSheets.length == 0) {
    var msg = getUiLang('no-target-page-exists2', 'No Target Page Exists. set "o" at "*URLs" sheet');
    showAlert(msg);
    return getUiLang('canceled', 'canceled');
  }

  var msg = getUiLang('template-caution', 'CAUTION: All result will be overwritten.');
  if(showConfirm(msg) != "OK") return getUiLang('canceled', 'canceled');
  
  var mark = getProp('mark');
  var n = 0;
  var usingCriteria = getUsingCriteria();

  // find row to start  
  var row2start = findFirstEmptyRow(allSheets[0]);
  row2start++;
  for (var i = 0; i < allSheets.length; i++) {
    var lastrow = allSheets[i].getLastRow();
    if (row2start >= lastrow) continue;
    var iclRows = allSheets[i]
      .getRange(row2start, 1, lastrow - row2start + 1, 6)
      .getValues();
    var transfer = getIclScTransferRows(iclRows, usingCriteria, mark);
    
    // set
    allSheets[i].getRange(5, 2, transfer.results.length, 1).setValues(transfer.results);
    allSheets[i].getRange(5, 4, transfer.memos.length, 1).setValues(transfer.memos);
    
    n++;
  }

  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}
