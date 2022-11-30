/**
 * Issue Report for COB-CHA
 * functions:
 * - isEditIssue
 * - dialogValueIssue
 * - generateIssueSheet
 * - openDialogIssue
 * - applyIssue
 * - setIssueList
 * - showEachIssue
 * - uploadIssueImage
 */

/**
 * is Edit or Add
 * @return Bool
 */
function isEditIssue() {
  // if not exist issue sheet then this time is not for edit
  if ( ! isSheetExist(gIssueSheetName)) return false;

  // if current sheet is not issue sheet then this time is not for edit
  if (getActiveSheet().getName() != gIssueSheetName) return false;

  // if current row has id then this time is for edit
  var sheet = getSpreadSheet().getSheetByName(gIssueSheetName);
  var activeRow = sheet.getActiveCell().getRow();
  var issueId = sheet.getRange(activeRow, 1).getValue();
  
  return (String(issueId).length > 0 && activeRow > 1);
}

/**
 * set dialog Value Issue
 * @param Bool isEdit
 * @return Array
 */
function dialogValueIssue(isEdit) {
  var ret = {};
  ret['isEdit'] = isEdit;
  ret['lang'] = getProp('lang');
  ret['type'] = getProp('type');
  ret['level'] = getProp('level');
  ret['usingCriteria'] = getUsingCriteria();
  ret['usingTechs'] = getUsingTechs();

  ret['allPlaces'] = [];
  var all = getAllSheets();
  for (i = 0; i < all.length; i++) {
//    ret['allPlaces'].push(getUrlFromSheet(all[i]));
    ret['allPlaces'].push({
      'url' : getUrlFromSheet(all[i]),
      'title' : getTitleFromSheet(all[i])
    });
  }
  
  ret['vals'] = {};
  var celposes = {
    'issueId': 1,
    'issueName': 2,
    'issueVisibility': 3,
    'errorNotice': 4,
    'html': 5,
    'explanation': 6,
    'checked': 7,
    'techs': 8,
    'places': 9,
    'image': 10,
    'preview': 11,
    'memo': 12
  };

  if (isEdit) {
    // issue sheet must be existed and activated
    var ss = getSpreadSheet();
    var sheet = ss.getSheetByName(gIssueSheetName);
    var activeRow = sheet.getActiveCell().getRow();
    for (var key in celposes) {
      var celpos = celposes[key];
      if (key !== 'preview') {
        ret['vals'][key] = sheet.getRange(activeRow, celpos).getValue().toString();
      } else {
        ret['vals'][key] = sheet.getRange(activeRow, celpos).getFormula();
      }
    }
    ret['vals']['preview'] = removeImageFormula(ret['vals']['preview']);
  }

  return ret;
}

/**
 * generate Issue sheet
 * @return Void
 */
function generateIssueSheet() {
  if (isSheetExist(gIssueSheetName)) return;
  
  // generate Issue sheet
  var defaults = [[
    "ID",
    getUiLang('name', 'Name'),
    getUiLang('issue-solved', 'Solved'),
    'Type',
    'HTML',
    getUiLang('explanation', 'Explanation'),
    getUiLang('criterion', 'Criteria'),
    getUiLang('tech', 'Techniques'),
    getUiLang('places', 'Places'),
    getUiLang('image', 'Image'),
    getUiLang('preview', 'Preview'),
    getUiLang('memo', 'Memo')
  ]];
  var sheet = generateSheetIfNotExists(gIssueSheetName, defaults, "row"); // do not return msg
  sheet.getRange("F:F").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 45);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);
}

/**
 * open dialog Issue
 * @return Void
 */
function openDialogIssue() {
  // to tell current page
  var activeSheet = getActiveSheet();
  var html = '<input type="hidden" id="target-url" value="">';
  if (activeSheet.getName().charAt(0) != '*') {
    html = '<input type="hidden" id="target-url" value="'+getUrlFromSheet(activeSheet)+'">';
  }

  // generate
  generateIssueSheet()  
  
  var title = isEditIssue() ? getUiLang('edit-issue', 'Edit issue') : getUiLang('add-new-issue', 'Add new issue');
  showDialog('ui-issue', 500, 400, title, html);
}

/**
 * apply Issue
 * @param Array vals
 * @return String
 */
function applyIssue(vals) {
  var ss = getSpreadSheet();
  var sheet = ss.getSheetByName(gIssueSheetName);

  // issue id - edit
  if (vals[0] > 0) {
    var targetRow = sheet.getActiveCell().getRow();
    sheet.getRange(targetRow, 1).setValue(vals[0]);
  } else {
    var targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue(targetRow - 1);
  }

  // set values
  for (var i = 1; i < vals.length; i++) {
    if (i == 8) {
      if (vals[i].split(",").length == getAllSheets().length) {
        vals[i] = "all";
      }
    }
    sheet.getRange(targetRow, i + 1).setValue(vals[i]);
  }
  var preview = sheet.getRange(targetRow, 11).getValue();
  if (preview) {
    sheet.getRange(targetRow, 11).setValue('=IMAGE("https://drive.google.com/uc?export=download&id='+preview+'",1)')
  }
  
  var range = sheet.getRange(targetRow+":"+targetRow);
  setRowConditionSolved(sheet, range, targetRow);
  
  if (vals[0] > 0) {
    return getUiLang('edit-done', 'Edited');
  }
  return getUiLang('add-done', 'Added');
}

/**
 * setRowConditionSolved
 * @param Object sheet
 * @param Object range
 * @param String targetRow
 * @return Object
 */
function setRowConditionSolved(sheet, range, targetRow) {
  var NotSolvedAndError = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND($C"+targetRow+"=\"off\", $D"+targetRow+"=\"Error\")")
      .setBold(true)
      .setBackground(gNotYetIssueBgColor)
      .setRanges([range])
      .build();
  var NotSolvedAndNotError = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND($C"+targetRow+"=\"off\", NOT($D"+targetRow+"=\"Error\"))")
      .setBold(false)
      .setBackground(gNotYetIssueBgColor)
      .setRanges([range])
      .build();
  var SolvedAndError = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND(NOT($C"+targetRow+"=\"off\"), $D"+targetRow+"=\"Error\")")
      .setBold(true)
      .setFontColor(null)
      .setBackground(null)
      .setRanges([range])
      .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(NotSolvedAndError);
  rules.push(NotSolvedAndNotError);
  rules.push(SolvedAndError);
  sheet.setConditionalFormatRules(rules);
}

/**
 * set Issue list
 * @return Object
 */
function setIssueList() {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(gIssueSheetName);
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName().toString();
  
  // target URL
  if (activeSheetName == gResultSheetName) {
    var targetRow = activeSheet.getActiveCell().getRow();
    var url = activeSheet.getRange(targetRow, 1).getValue();
  } else {
    if (activeSheetName.charAt(0) == '*') {
      return {'url': '', 'issues': []};
    } else {
      var url = activeSheet.getRange(2, 2).getValue();
    }
  }

  var dataObj = issueSheet.getDataRange().getValues();
  var issues = [];
  for (var i = 1; i < dataObj.length; i++) {
    var urls = dataObj[i][8].toString().split(',');
    for (var j = 0; j < urls.length; j++) {
      var issueurl = urls[j].trim();
      if (issueurl != url) continue;
      issues.push(dataObj[i]);
    }
  }
  return {'url': url, 'issues': issues};
}

/**
 * show each issue
 * @param Integer row
 * @return Void
 */
function showEachIssue(row) {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(gIssueSheetName);
  issueSheet.getRange(row, 1).activate();
  showDialog('ui-issue', 500, 400, getUiLang('edit-issue', 'Edit issue'));
}

/**
 * upload Issue image
 * @param Object formObj
 * @return Object
 */
function uploadIssueImage(formObj) {
  return fileUpload(gImagesFolderName, formObj, "imageFile");
}