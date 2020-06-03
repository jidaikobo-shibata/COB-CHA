/**
 * Config for COB-CHA
 */

/**
 * is Config Sheet Exists
 * @return Bool
 */
function isConfigSheetExists() {
  var ss = getSpreadSheet();
  var configSheet = ss.getSheetByName(configSheetName);
  return (configSheet);
}

/**
 * Generate Config Sheet
 * @param String lang
 * @param String testType
 * @param String level
 * @return String
 */
function generateConfigSheets(lang, testType, level) {
  if (isConfigSheetExists()) return getUiLang('config-sheet-already-exists', "Config Sheet is already exists.");
  doGenerateConfigSheets(lang, testType, level);
  return getUiLang('generate-config-sheet', "Generate Config Sheet.");
}

/**
 * Do Generate Config Sheet
 * @param String lang
 * @param String testType
 * @param String level
 * @return Object
 */
function doGenerateConfigSheets(lang, testType, level) {
  var ss = getSpreadSheet();
  ss.insertSheet(configSheetName, 0);
  var configsheet = ss.getSheetByName(configSheetName);
  configsheet.activate();
  setBasicValue(configsheet, lang, testType, level);
  configsheet.getRange(2, 1).setValue('Name').setBackground(labelColor);
  configsheet.getRange(3, 1).setValue('Report Date').setBackground(labelColor);
  configsheet.getRange(4, 1).setValue(getUiLang('additional-criteria', "Additional Criteria")).setBackground(labelColor);
  return configsheet;
}

/**
 * set Additional Criteria
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function setAdditionalCriteria(lang, testType, level) {
  var ss = getSpreadSheet();
  if (isConfigSheetExists()) {
    var configsheet = ss.getSheetByName(configSheetName);
    setBasicValue(configsheet, lang, testType, level);
  } else {
    var configsheet = doGenerateConfigSheets(lang, testType, level);
  }
  showDialog('additional-criteria', 500, 400, getUiLang('set-additional-criteria', 'Set additional criteria'));
}

/**
 * get Additional Criteria
 * @return String
 */
function getAdditionalCriteria() {
  var ss = getSpreadSheet();
  var configsheet = ss.getSheetByName(configSheetName);
  return configsheet.getRange(4, 2).getValue();
}

/**
 * set Additional Criteria Value
 * @return Object
 */
function setAdditionalCriteriaValue() {
  var ret = {};
  ret['lang'] = getProp('lang');
  ret['type'] = getProp('type');
  ret['level'] = getProp('level');
  
  ret['checked'] = getAdditionalCriteria();
  ret['criteria'] = getAllCriteria(ret['lang'], ret['type']);

  return ret;
}

/**
 * apply Additional Criteria
 * @param Array checked
 * @return String
 */
function applyAdditionalCriteria(checked) {
  var ss = getSpreadSheet();
  var configsheet = ss.getSheetByName(configSheetName);
  configsheet.getRange(4, 2).setValue(checked);
  return getUiLang('update-additional-criteria', 'Update additional criteria');
}
