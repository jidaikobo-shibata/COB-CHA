/**
 * Config sheet for COB-CHA
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
    [getUiLang('lang-using', "Lang"), lang],
    [getUiLang('standard-using', "Type"), testType],
    [getUiLang('level', "Level"), level],
    [getUiLang('symbol-using', "Mark Type"), mark],
    [getUiLang('additional-criteria', "Additional Criteria"), ""]
  ];
  var msgOrSheetObj = generateSheetIfNotExists(gConfigSheetName, defaults);

  var sheet = getSheetIfExists(gConfigSheetName);
  sheet.setColumnWidth(1, 100);

  //if (force !== true && typeof msgOrSheetObj == "string") return msgOrSheetObj;
  if (force !== true && typeof msgOrSheetObj == "string") {
    var msg = getUiLang('force-update-config', 'Config Sheet is alreadt exists. Update config?');
    if(showConfirm(msg) != "OK") return getUiLang('canceled', 'canceled');
    sheet.getRange(1, 2, 4, 1).setValues([[lang], [testType], [level], [mark]]);
    return getUiLang('target-sheet-updated', "Update Target Sheet (%s).").replace('%s', gConfigSheetName);
  }

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
  ret['criteria22'] = gCriteria22;
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

/**
 * force update Language
 * @param String lang
 * @return Void
 */
function forceUpdateLang(lang) {
  var msg = getUiLang('confirm-control-panel-update', 'Change of language cause reload control panel.');
  if(showConfirm(msg) != "OK") return getProp('lang');
  
  var sheet = getSheetIfExists(gConfigSheetName);
  
  // if sheet exists then update language
  if (sheet !== false) {
    sheet.getRange(1, 2).setValue(lang);
  } else {
    generateConfigSheet(lang, '', '', '', true);
  }
    
  // when language was update, renew control panel
  showControlPanel();
}
