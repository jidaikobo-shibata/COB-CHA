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
function generateConfigSheet(lang, testType, level) {
  if (isConfigSheetExists()) return getUiLang('config-sheet-already-exists', "Config Sheet is already exists.");
  prepareConfigSheet(lang, testType, level);
  return getUiLang('generate-config-sheet', "Generate Config Sheet.");
}

/**
 * prepare Config Sheet
 * @param String lang
 * @param String testType
 * @param String level
 * @return Object
 */
function prepareConfigSheet(lang, testType, level) {
  var ss = getSpreadSheet();
  ss.insertSheet(configSheetName, 0);
  var configsheet = ss.getSheetByName(configSheetName);
  configsheet.activate();
  setBasicValue(configsheet, lang, testType, level);
  configsheet.getRange(2, 1).setValue('Name').setBackground(labelColor);
  configsheet.getRange(3, 1).setValue('Report Date').setBackground(labelColor);
  configsheet.getRange(4, 1).setValue(getUiLang('additional-criterion', "Additional Criterion")).setBackground(labelColor);
  return configsheet;
}

/**
 * get Additional Criterion
 * @return String
 */
function getAdditionalCriterion() {
  var ss = getSpreadSheet();
  var configsheet = ss.getSheetByName(configSheetName);
  return configsheet.getRange(4, 2).getValue();
}

/**
 * open dialog Additional Criterion
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function openDialogAdditionalCriterion(lang, testType, level) {
  if (isConfigSheetExists()) {
    var ss = getSpreadSheet();
    var configsheet = ss.getSheetByName(configSheetName);
    setBasicValue(configsheet, lang, testType, level);
  } else {
    var configsheet = prepareConfigSheet(lang, testType, level);
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

  return ret;
}

/**
 * apply Additional Criterion
 * @param Array checked
 * @return String
 */
function applyAdditionalCriterion(checked) {
  var ss = getSpreadSheet();
  var configsheet = ss.getSheetByName(configSheetName);
  configsheet.getRange(4, 2).setValue(checked);
  return getUiLang('update-additional-criterion', 'Update additional criterion');
}
