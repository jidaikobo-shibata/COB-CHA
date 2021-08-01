/**
 * Config for COB-CHA
 * functions:
 * - generateConfigSheet
 * - openDialogAdditionalCriteria
 * - dialogValueAdditionalCriteria
 * - setAdditionalCriteria
 */

/**
 * generate Config Sheet
 * @param String lang
 * @param String testType
 * @param String level
 * @param String mark
 * @param Bool force
 * @return String
 */
function generateConfigSheet(lang, testType, level, mark, force) {
  var defaults = [
    ["Lang", lang],
    ["Type", testType],
    ["Level", level],
    ["Mark Type", mark],
    ["Additional Criteria", ""]
  ];
  var msgOrSheetObj = generateSheetIfNotExists(gConfigSheetName, defaults);
  if (force !== true && typeof msgOrSheetObj == "string") return msgOrSheetObj;
  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gConfigSheetName);
}

/**
 * open dialog Additional Criteria
 * @param String lang
 * @param String testType
 * @param String level
 * @param String mark
 * @return Void
 */
function openDialogAdditionalCriteria(lang, testType, level, mark) {
  generateConfigSheet(lang, testType, level, mark, true);
  showDialog('ui-additional-criteria', 500, 400, getUiLang('set-additional-criteria', 'Set additional criteria'));
}

/**
 * get Dialog Value Additional Criteria
 * @return Object
 */
function dialogValueAdditionalCriteria() {
  var ret = {};
  ret['lang'] = getProp('lang');
  ret['type'] = getProp('type');
  ret['level'] = getProp('level');
  ret['checked'] = getProp('additional');
  ret['criteria'] = getAllCriteria();
  ret['criteria21'] = gCriteria21;
  return ret;
}

/**
 * set Additional Criteria
 * @param Array checked
 * @return Void
 */
function setAdditionalCriteria(checked) {
  var sheet = getSheetIfExists(gConfigSheetName);
  if (sheet === false) return getUiLang('no-target-sheet-exists', "Target sheet (%s) is not exists.").replace('%s', gConfigSheetName);
  sheet.getRange(5, 2).setValue(checked);
  showAlert(getUiLang('update-value', 'Update %s').replace('%s', getUiLang('additional-criteria', 'additional criteria')))
}
