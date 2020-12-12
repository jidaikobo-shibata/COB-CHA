/**
 * Config for COB-CHA
 */

/**
 * is URL List Exists
 * @return Bool
 */
function isUrlListSheetExists() {
  var ss = getSpreadSheet();
  var urlListSheet = ss.getSheetByName(gUrlListSheetName);
  return (urlListSheet);
}

/**
 * Generate URL List
 * @param String lang
 * @param String testType
 * @param String level
 * @param String mark
 * @return String
 */
function generateUrlListSheet(lang, testType, level, mark) {
  if (isUrlListSheetExists()) return getUiLang('url-list-sheet-already-exists', "URL List sheet is already exists.");
  prepareUrlListSheet(lang, testType, level, mark);
  return getUiLang('generated-url-list-sheet', "Generate URL List.");
}

/**
 * prepare URL List
 * @param String lang
 * @param String testType
 * @param String level
 * @param String mark
 * @return Object
 */
function prepareUrlListSheet(lang, testType, level, mark) {
  var ss = getSpreadSheet();
  ss.insertSheet(gUrlListSheetName, 0);
  var urlListSheet = ss.getSheetByName(gUrlListSheetName);
  urlListSheet.activate();
  setBasicValue(urlListSheet, lang, testType, level, mark);
  urlListSheet.getRange(1, 5).setValue(getUiLang('additional-criterion', "Additional Criterion")).setBackground(gLabelColor);
  urlListSheet.getRange(2, 1).setValue('No.');
  urlListSheet.getRange(2, 2).setValue('URL');
  urlListSheet.getRange('2:2').setBackground(gLabelColor);
  urlListSheet.setColumnWidth(1, 35);
  urlListSheet.setColumnWidth(2, 200);
  urlListSheet.getRange("A1:A").setHorizontalAlignment('center');
  var nos = [];
  for (var i = 1; i <= 40; i++) {
    nos.push([i]);
  }
  urlListSheet.getRange(3, 1, 40, 1).setValues(nos);
  deleteFallbacksheet();
  return urlListSheet;
}

/**
 * get Additional Criterion
 * @return String
 */
function getAdditionalCriterion() {
  var ss = getSpreadSheet();
  var urlListSheet = ss.getSheetByName(gUrlListSheetName);
  if ( ! urlListSheet) return '';
  return urlListSheet.getRange(1, 6).getValue();
}

/**
 * open dialog Additional Criterion
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function openDialogAdditionalCriterion(lang, testType, level) {
  if (isUrlListSheetExists()) {
    var ss = getSpreadSheet();
    var urlListSheet = ss.getSheetByName(gUrlListSheetName);
    setBasicValue(urlListSheet, lang, testType, level);
  } else {
    var UrlListSheet = prepareUrlListSheet(lang, testType, level);
  }
  showDialog('ui-additional-criterion', 500, 400, getUiLang('set-additional-criterion', 'Set additional criterion'));
}

/**
 * get Dialog Value Additional Criterion Value
 * @return Object
 */
function dialogValueAdditionalCriterionValue() {
  var ret = {};
  ret['lang'] = getProp('lang');
  ret['type'] = getProp('type');
  ret['level'] = getProp('level');

  ret['checked'] = getAdditionalCriterion();
  ret['criteria'] = getAllCriteria(ret['lang'], ret['type']);
  ret['criteria21'] = gCriteria21;
  
  return ret;
}

/**
 * apply Additional Criterion
 * @param Array checked
 * @return String
 */
function applyAdditionalCriterion(checked) {
  var ss = getSpreadSheet();
  var urlListSheet = ss.getSheetByName(gUrlListSheetName);
  urlListSheet.getRange(1, 6).setValue(checked);
  return getUiLang('update-additional-criterion', 'Update additional criterion');
}
