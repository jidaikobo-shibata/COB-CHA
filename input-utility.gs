/**
 * Input Utility control for COB-CHA
 */

/**
 * get contextual techniques
 * @param String criterion
 * @param String checked
 * @return Object
 */
function getContextualTechs(criterion, checked) {
  if (criterion == '') {
    var activeSheet = getActiveSheet();
    var activeRow = activeSheet.getActiveCell().getRow();
    var criterion = activeSheet.getRange(activeRow, 1).getValue();
    var checked = activeSheet.getRange(activeRow, 4).getValue();
  }
  var techLangsSrc = getLangSet('tech');

  var rets = [];
  var type = getProp('type');
  if (type.indexOf('tt') >= 0) {
    for (var key in gRelTtAndCriteria) {
      if (gRelTtAndCriteria[key].indexOf(criterion) < 0) continue;
      var techs = gRelTechsAndCriteria[key] ? gRelTechsAndCriteria[key] : [] ;
      for (var i = 0; i < techs.length; i++) {
        if (techLangsSrc[techs[i]] == null) continue;
        rets.push([techs[i], techLangsSrc[techs[i]]]);
      }
    }
  } else {
    var techs = gRelTechsAndCriteria[criterion] ? gRelTechsAndCriteria[criterion] : [] ;
    for (var i = 0; i < techs.length; i++) {
      if (techLangsSrc[techs[i]] == null) continue;
      rets.push([techs[i], techLangsSrc[techs[i]]]);
    }
  }

  var lang = getProp('lang');
  var type = getProp('type');
  var docurl = lang+'-'+type;
  var docurlEn = 'en'+'-'+type;
  
  return {'criterion': criterion, 'techs': rets, 'checked': checked, 'lang': lang, 'type': type, 'techDirAbbr': gTechDirAbbr, 'urlbase': gUrlbase, 'docurl': docurl, 'docurlEn': docurlEn};
}

/**
 * set contextual techniques
 * @param String techs
 * @return Void
 */
function setContextualTechs(techs) {
  var activeSheet = getActiveSheet();
  var activeRow = activeSheet.getActiveCell().getRow();
  activeSheet.getRange(activeRow, 4).setValue(techs);
}

/**
 * Apply Value to Conformance
 * @param String testType
 * @param String level
 * @return String
 */
function applyAllToT(testType, level) {
  var ttCriteria = getLangSet('ttCriteria');
  var activeSheet = getActiveSheet();
  if (activeSheet.getName() == gResultSheetName) return getUiLang('current-sheet-is-not-for-webpage', 'Current sheet is not for webpage');

  var additionalCriteria = getAdditionalCriterion() ? getAdditionalCriterion().split(/,/) : [];
  var rows = 61; // WCAG 2.0 AAA
  rows = testType == 'wcag20' && level == 'A'   ? 25 + additionalCriteria.length : rows;
  rows = testType == 'wcag20' && level == 'AA'  ? 38 + additionalCriteria.length : rows;
  rows = testType == 'wcag21' && level == 'A'   ? 30 + additionalCriteria.length : rows;
  rows = testType == 'wcag21' && level == 'AA'  ? 50 + additionalCriteria.length : rows;
  rows = testType == 'wcag21' && level == 'AAA' ? 78 : rows;
  rows = testType == 'tt20' ? ttCriteria.length : rows;
  
  var vals = [];
  var mark = getProp('mark');
  for (var i = 1; i <= rows; i++) {
    vals.push([mark[2]]);
  }
  activeSheet.getRange(5, 2, vals.length, 1).setValues(vals);
      
  return getUiLang('edit-done', 'Value Edited');
}

/**
 * Make same as template
 * @return String
 */
function templateApplyAll() {
  var ss = getSpreadSheet();
  var tpl = ss.getSheetByName(gTemplateSheetName);
  if (tpl == null) return getUiLang('no-template-found', 'No template exists.');

  var n = 0;
  var allSheets = getAllSheets();
  for (i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).charAt(0) == '*') continue;
    tpl.getRange(5, 2, tpl.getLastRow(), 4).copyTo(allSheets[i].getRange(5, 2));
    n++;
  }

  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}

/**
 * Make same as template row
 * @return String
 */
function templateApplyRow() {
  var ss = getSpreadSheet();
  var tpl = ss.getSheetByName(gTemplateSheetName);
  if (tpl == null) throw new Error(getUiLang('no-template-found', 'No template exists.'));

  var activeSheet = getActiveSheet();
  if (gTemplateSheetName != activeSheet.getName()) throw new Error(getUiLang('is-not-template', 'Current Sheet is not template.'));

  var activeRow = activeSheet.getActiveCell().getRow();
  if (activeRow < 5) throw new Error(getUiLang('is-not-appropriate-row', 'Current Row is not Result.'));

  var n = 0;
  var allSheets = getAllSheets();
  for (i = 0; i < allSheets.length; i++) {
    tpl.getRange(activeRow, 2, activeRow, tpl.getLastColumn()).copyTo(allSheets[i].getRange(activeRow, 2));
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
  Logger.log(val);
  
  var n = 0;
  var allSheets = getAllSheets();
  for (i = 0; i < allSheets.length; i++) {
    allSheets[i].getRange(row, col).setValue(val);
    n++;
  }
  return getUiLang('sheet-edited', '%s sheet(s) edited.').replace("%s", n);
}
