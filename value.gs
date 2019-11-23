/**
 * Value control for COB-CHA
 */

var relTechsAndCriteria = {
  '1.1.1': ['G68', 'G73', 'G74', 'G82', 'G92', 'G94', 'G95', 'G100', 'G143', 'G144', 'G196', 'H2', 'H24', 'H30', 'H35', 'H36', 'H37', 'H44', 'H45', 'H46', 'H53', 'H65', 'H67', 'H86', 'C9', 'C18', 'ARIA6', 'ARIA9', 'ARIA10', 'ARIA15', 'PDF1', 'PDF4', 'F3', 'F13', 'F20', 'F30', 'F38', 'F39', 'F65', 'F67', 'F71', 'F72'],
  '1.2.1': ['G158', 'G159', 'G166', 'H96', 'F30', 'F67'],
  '1.2.2': ['G87', 'G93', 'H95', 'F8', 'F74', 'F75'],
  '1.2.3': ['G8', 'G58', 'G69', 'G78', 'G173', 'G203', 'H53', 'H96'],
  '1.2.4': ['G9', 'G87', 'G93'],
  '1.2.5': ['G8', 'G78', 'G173', 'G203', 'H96'],
  '1.2.6': ['G54', 'G81'],
  '1.2.7': ['G8', 'H96'],
  '1.2.8': ['G58', 'G69', 'G159', 'H46', 'H53', 'F74'],
  '1.2.9': ['G150', 'G151', 'G157'],
  '1.3.1': ['G115', 'G117', 'G138', 'G140', 'G141', 'G162', 'H39', 'H42', 'H43', 'H44', 'H48', 'H49', 'H51', 'H63', 'H65', 'H71', 'H73', 'H85', 'H97', 'C22', 'SCR21', 'T1', 'T2', 'T3', 'ARIA1', 'ARIA2', 'ARIA11', 'ARIA12', 'ARIA13', 'ARIA16', 'ARIA17', 'ARIA20', 'PDF6', 'PDF9', 'PDF10', 'PDF11', 'PDF12', 'PDF17', 'PDF20', 'PDF21', 'F2', 'F33', 'F34', 'F42', 'F43', 'F46', 'F48', 'F87', 'F90', 'F91', 'F92'],
  '1.3.2': ['G57', 'H34', 'H56', 'C6', 'C8', 'C27', 'PDF3', 'F1', 'F32', 'F33', 'F34', 'F49'],
  '1.3.3': ['G96', 'F14', 'F26'],
  '1.4.1': ['G14', 'G111', 'G182', 'G183', 'G205', 'C15', 'F13', 'F73', 'F81'],
  '1.4.2': ['G60', 'G170', 'G171', 'F23', 'F93'],
  '1.4.3': ['G18', 'G145', 'G148', 'G156', 'G174', 'F24', 'F83'],
  '1.4.4': ['G142', 'G146', 'G178', 'G179', 'C12', 'C13', 'C14', 'C17', 'C20', 'C22', 'C28', 'SCR34', 'F69', 'F80'],
  '1.4.5': ['G140', 'C6', 'C8', 'C12', 'C13', 'C14', 'C22', 'C30', 'PDF7'],
  '1.4.6': ['G17', 'G18', 'G148', 'G156', 'G174', 'F24', 'F83'],
  '1.4.7': ['G56'],
  '1.4.8': ['G146', 'G148', 'G156', 'G169', 'G172', 'G175', 'G188', 'G204', 'G206', 'C12', 'C13', 'C14', 'C19', 'C20', 'C21', 'C23', 'C24', 'C25', 'SCR34', 'F24', 'F88'],
  '1.4.9': ['G140', 'C6', 'C8', 'C12', 'C13', 'C14', 'C22', 'C30', 'PDF7'],
  '2.1.1': ['G90', 'G202', 'H91', 'SCR2', 'SCR20', 'SCR29', 'SCR35', 'PDF3', 'PDF11', 'PDF23', 'F42', 'F54', 'F55'],
  '2.1.2': ['G21', 'F10'],
  '2.1.3': ['G90', 'G202', 'H91', 'SCR2', 'SCR20', 'SCR29', 'SCR35', 'PDF3', 'PDF11', 'PDF23', 'F42', 'F54', 'F55'],
  '2.2.1': ['G4', 'G133', 'G180', 'G198', 'SCR1', 'SCR16', 'SCR33', 'SCR36', 'F40', 'F41', 'F58'],
  '2.2.2': ['G4', 'G11', 'G152', 'G186', 'G187', 'G191', 'SCR22', 'SCR33', 'F4', 'F7', 'F16', 'F47', 'F50'],
  '2.2.3': ['G5'],
  '2.2.4': ['G75', 'G76', 'SCR14', 'F40', 'F41'],
  '2.2.5': ['G105', 'G181', 'F12'],
  '2.3.1': ['G15', 'G19', 'G176'],
  '2.3.2': ['G19'],
  '2.4.1': ['G1', 'G123', 'G124', 'H64', 'H69', 'H70', 'C6', 'SCR28', 'ARIA11', 'PDF9'],
  '2.4.10': ['G141', 'H69'],
  '2.4.2': ['G88', 'G127', 'H25', 'PDF18', 'F25'],
  '2.4.3': ['G59', 'H4', 'C27', 'SCR26', 'SCR27', 'SCR37', 'PDF3', 'F44', 'F85'],
  '2.4.4': ['G53', 'G91', 'G189', 'H2', 'H24', 'H30', 'H33', 'H77', 'H78', 'H79', 'H80', 'H81', 'C7', 'SCR30', 'ARIA7', 'ARIA8', 'PDF11', 'PDF13', 'F63', 'F89'],
  '2.4.5': ['G63', 'G64', 'G125', 'G126', 'G161', 'G185', 'H59', 'PDF2'],
  '2.4.6': ['G130', 'G131'],
  '2.4.7': ['G149', 'G165', 'G195', 'C15', 'SCR31', 'F55', 'F78'],
  '2.4.8': ['G63', 'G65', 'G127', 'G128', 'H59', 'PDF14', 'PDF17'],
  '2.4.9': ['G91', 'G189', 'H2', 'H24', 'H30', 'H33', 'C7', 'SCR30', 'ARIA8', 'PDF11', 'PDF13', 'F84', 'F89'],
  '3.1.1': ['H57', 'SVR5', 'PDF16', 'PDF19'],
  '3.1.2': ['H58', 'PDF19'],
  '3.1.3': ['G55', 'G62', 'G70', 'G101', 'G112', 'H40', 'H54', 'H60'],
  '3.1.4': ['G55', 'G62', 'G70', 'G97', 'G102', 'H28', 'H60', 'PDF8'],
  '3.1.5': ['G79', 'G86', 'G103', 'G153', 'G160'],
  '3.1.6': ['G62', 'G120', 'G121', 'G163', 'H62'],
  '3.2.1': ['G107', 'F52', 'F55'],
  '3.2.2': ['G13', 'G80', 'H32', 'H84', 'SCR19', 'PDF15', 'F36', 'F37'],
  '3.2.3': ['G61', 'PDF14', 'PDF17', 'F66'],
  '3.2.4': ['G197', 'F31'],
  '3.2.5': ['G76', 'G110', 'H76', 'H83', 'SCR19', 'SCR24', 'SVR1', 'F9', 'F22', 'F41', 'F52', 'F60', 'F61'],
  '3.3.1': ['G83', 'G84', 'G85', 'G139', 'G199', 'SCR18', 'SCR32', 'ARIA18', 'ARIA19', 'ARIA21', 'PDF5', 'PDF22'],
  '3.3.2': ['G13', 'G83', 'G89', 'G131', 'G162', 'G167', 'G184', 'H44', 'H65', 'H71', 'H90', 'ARIA1', 'ARIA9', 'ARIA17', 'PDF5', 'PDF10', 'F82'],
  '3.3.3': ['G83', 'G84', 'G85', 'G139', 'G177', 'G199', 'SCR18', 'SCR32', 'ARIA2', 'ARIA18', 'PDF5', 'PDF22'],
  '3.3.4': ['G98', 'G99', 'G155', 'G164', 'G168', 'G199', 'SCR18'],
  '3.3.5': ['G71', 'G89', 'G184', 'G193', 'G194', 'H89'],
  '3.3.6': ['G98', 'G99', 'G155', 'G164', 'G168', 'G199'],
  '4.1.1': ['G134', 'G192', 'H74', 'H75', 'H88', 'H93', 'H94', 'F70', 'F77'],
  '4.1.2': ['G10', 'G108', 'G135', 'H44', 'H64', 'H65', 'H88', 'H91', 'ARIA4', 'ARIA5', 'ARIA14', 'ARIA16', 'PDF10', 'PDF12', 'F15', 'F20', 'F42', 'F59', 'F68', 'F79', 'F86', 'F89']
};

/**
 * get contextual techniques
 * @param String criterion
 * @param String checked
 * @return Array
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
    var ttCheckValSrc = getLangSet('ttCheckVal');
    for (var key in ttCheckValSrc) {
      if (ttCheckValSrc[key].indexOf(criterion) < 0) continue;
      var techs = relTechsAndCriteria[key] ? relTechsAndCriteria[key] : [] ;
      for (var i = 0; i < techs.length; i++) {
        if (techLangsSrc[techs[i]] == null) continue;
        rets.push([techs[i], techLangsSrc[techs[i]]]);
      }
    }
  } else {
    var techs = relTechsAndCriteria[criterion] ? relTechsAndCriteria[criterion] : [] ;
    for (var i = 0; i < techs.length; i++) {
      if (techLangsSrc[techs[i]] == null) continue;
      rets.push([techs[i], techLangsSrc[techs[i]]]);
    }
  }
  
  var lang = getProp('lang');
  var type = getProp('type');
  var techDirAbbr = techDirAbbr;
  var docurl = lang+'-'+type;
  var docurlEn = 'en'+'-'+type;
  
  return {'criterion': criterion, 'techs': rets, 'checked': checked, 'lang': lang, 'type': type, 'techDirAbbr': techDirAbbr, 'urlbase': urlbase, 'docurl': docurl, 'docurlEn': docurlEn};
}

/**
 * set contextual techniques
 * @param String techs
 */
function setContextualTechs(techs) {
  var activeSheet = getActiveSheet();
  var activeRow = activeSheet.getActiveCell().getRow();
  activeSheet.getRange(activeRow, 4).setValue(techs);
}

/**
 * Edit Value to "T"
 * @param String testType
 * @param String level
 * @return String
 */
function editValue(testType, level) {
  var ttCriteria = getLangSet('ttCriteria');
  var activeSheet = getActiveSheet();
  if (activeSheet.getName() == resultSheetName) return getUiLang('current-sheet-is-not-for-webpage', 'Current sheet is not for webpage');

  var rows = 61; // WCAG 2.0 AAA
  rows = testType == 'wcag20' && level == 'A'   ? 25 : rows;
  rows = testType == 'wcag20' && level == 'AA'  ? 38 : rows;
  rows = testType == 'wcag21' && level == 'A'   ? 30 : rows;
  rows = testType == 'wcag21' && level == 'AA'  ? 50 : rows;
  rows = testType == 'wcag21' && level == 'AAA' ? 78 : rows;
  rows = testType == 'tt20' ? ttCriteria.length : rows;

  for (var i = 1; i <= rows; i++) {
    activeSheet.getRange(i+4, 2).setValue('T');
  }
  return getUiLang('edit-done', 'Value Edited');
}

/**
 * Bulk Edit
 * @param String target
 * @param String check
 * @param String tech
 * @param String memo
 * @return String
 */
function bulkEdit(target, check, tech, memo) {
  var ss = getSpreadSheet();

  var n = 0;
  var allSheets = getAllSheets()
  for (i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).charAt(0) == '*') continue;

    // search row
    var dat = allSheets[i].getDataRange().getValues();
    var row = 0;
    for (var j = 1; j < dat.length; j++) {
      if (dat[j][0] === target) {
        row = j + 1;
        break;
      }
    }
    if (row == 0) continue;
    
    // apply
    allSheets[i].getRange(row, 2).setValue(check);
    allSheets[i].getRange(row, 4).setValue(tech);
    allSheets[i].getRange(row, 5).setValue(memo);
    n++;
  }
  return getUiLang('edit-done', '%s sheet(s) edited.').replace("%s", n);
}

/**
 * Make same as template
 * @return String
 */
function makeSameAsTemplate() {
  var ss = getSpreadSheet();
  var tpl = ss.getSheetByName(templateSheetName);
  if (tpl == null) return getUiLang('no-template-found', 'No template exists.');

  // extract template data  
  var dataObj = tpl.getDataRange().getValues();
  var vals = {};
  for (var i = 1; i < dataObj.length; i++) {
    var key = dataObj[i].shift();
    if (['URL', 'title', 'Criterion'].indexOf(key) >= 0) continue;
    if (key.length > 0) {
      vals[i+1] = dataObj[i];
    }
  }

  // edit values
  var n = 0;
  var allSheets = getAllSheets()
  for (i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).charAt(0) == '*') continue;
    for (var row in vals){
      allSheets[i].getRange(row, 2).setValue(vals[row][0]);
      allSheets[i].getRange(row, 4).setValue(vals[row][2]);
      allSheets[i].getRange(row, 5).setValue(vals[row][3]);
    }
    n++;
  }
  
  return getUiLang('edit-done', '%s sheet(s) edited.').replace("%s", n);
}