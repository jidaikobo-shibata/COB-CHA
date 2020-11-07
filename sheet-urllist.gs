/**
 * Config for COB-CHA
 */

/**
 * is URL List Exists
 * @return Bool
 */
function isUrlListSheetExists() {
  var ss = getSpreadSheet();
  var urlListSheet = ss.getSheetByName(urlListSheetName);
  return (urlListSheet);
}

/**
 * Generate URL List
 * @param String lang
 * @param String testType
 * @param String level
 * @return String
 */
function generateUrlListSheet(lang, testType, level) {
  if (isUrlListSheetExists()) return getUiLang('url-list-sheet-already-exists', "URL List sheet is already exists.");
  prepareUrlListSheet(lang, testType, level);
  return getUiLang('generated-url-list-sheet', "Generate URL List.");
}

/**
 * prepare URL List
 * @param String lang
 * @param String testType
 * @param String level
 * @return Object
 */
function prepareUrlListSheet(lang, testType, level) {
  var ss = getSpreadSheet();
  ss.insertSheet(urlListSheetName, 0);
  var urlListSheet = ss.getSheetByName(urlListSheetName);
  urlListSheet.activate();
  setBasicValue(urlListSheet, lang, testType, level);
  urlListSheet.getRange(1, 5).setValue(getUiLang('additional-criterion', "Additional Criterion")).setBackground(labelColor);
  urlListSheet.getRange(2, 1).setValue('No.');
  urlListSheet.getRange(2, 2).setValue('URL');
  urlListSheet.getRange('2:2').setBackground(labelColor);
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
  var urlListSheet = ss.getSheetByName(urlListSheetName);
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
    var urlListSheet = ss.getSheetByName(urlListSheetName);
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
  ret['criteria21'] = criteria21;
  
  return ret;
}

/**
 * apply Additional Criterion
 * @param Array checked
 * @return String
 */
function applyAdditionalCriterion(checked) {
  var ss = getSpreadSheet();
  var urlListSheet = ss.getSheetByName(urlListSheetName);
  urlListSheet.getRange(1, 6).setValue(checked);
  return getUiLang('update-additional-criterion', 'Update additional criterion');
}
