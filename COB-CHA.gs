/**
 * COB-CHA: CollaBorative CHeck tool for Accessibility
 * Google Spreadsheet Add-on
 * @Author  shibata@jidaikobo.com
 *          arimatsu@jidaikobo.com
 * @Year    2021 - 
 * @Licence MIT
 * 
 * TODO:
 * - 解説書、達成基準へのリンクは削除予定（#del-link）
 * - showConfirm()において、キャンセルボタンでキャンセルしたら問題がないが、エスケープキーではエラーになる。try-catchで捉えようとしたが、捉えられていない（#cannot-esc）
 * 
 * functions:
 * - onInstall
 * - onOpen
 * - askEnabled
 * - addShowControlPanel
 * - showControlPanel
 * - showHelp
 * - showCredit
 * - showDialog
 * - showAlert
 * - getUrlFromSheet
 * - getTitleFromSheet
 * - getSheetByUrl
 * - getProp
 * - getLangSet
 * - getUiLang
 * - getAllCriteria
 * - getUsingCriteria
 * - getUsingTechs
 * - getHtmlAndTitle
 * - getSpreadSheet
 * - getActiveSheet
 * - getAllSheets
 * - resetSheets
 * - deleteFallbacksheet
 * - isSheetExist
 * - getSheetIfExists
 * - generateSheetIfNotExists
 * - generateSheetEvenIfAlreadyExists
 * - prepareSheet
 * - setRowConditionApplicability
 * - setRowConditionNotYet
 * - setCellConditionTF
 * - setCellConditionAttn
 */

/**
 * onInstall
 * @param Object e
 * @return Void
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * onOpen
 * @param Object e
 * @return Void
 */
function onOpen(e) {
  if (e && e.authMode == 'NONE') {
    var menu = SpreadsheetApp.getUi().createAddonMenu();
    menu.addItem('Getting Started', 'askEnabled');
    menu.addToUi();
  } else {
    addShowControlPanel();
  }
}

/**
 * askEnabled
 * @return Void
 */
function askEnabled() {
  var title = 'COB-CHA';
  var msg   = 'Script has been enabled.';
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  addShowControlPanel(menu)
};

/**
 * add "Show Control Panel" to menu
 * @return Void
 */
function addShowControlPanel() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem(getUiLang('show-control-panel', 'Show Control Panel'), 'showControlPanel');
  menu.addItem(getUiLang('help', 'COB-CHA Help'), 'showHelp');
  menu.addItem(getUiLang('credit', 'COB-CHA credit'), 'showCredit');
  menu.addToUi();
};

/**
 * show control panel
 * @return Void
 */
function showControlPanel() {
  var ui = HtmlService.createTemplateFromFile('ui-control-panel')
                      .evaluate()
                      .setTitle('COB-CHA'+getUiLang('control-panel-title', ' Control Panel'));
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * show help
 * @return Void
 */
function showHelp() {
  showDialog('ui-help', 500, 400, getUiLang('help', 'Help'));
}

/**
 * show credit
 * @return Void
 */
function showCredit() {
  showDialog('ui-credit', 500, 400, getUiLang('credit', 'Credit'));
}

/**
 * show dialog
 * @param String sheetname
 * @param Integer width
 * @param Integer height
 * @param String title
 * @param String html
 * @return Void
 */
function showDialog(sheetname, width, height, title, html) {
  var output = HtmlService.createTemplateFromFile(sheetname);
  var ss = getSpreadSheet();
  title = title == null ? '' : title;
  html  = html == null  ? '' : html;
  var html = output.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setWidth(width)
                   .setHeight(height)
                   .setTitle(title)
                   .append(html);
  ss.show(html);
}

/**
 * show alert
 * @param String msg
 * @param String buttonId
 * @return String
 */
function showAlert(msg, buttonId) {
  // #cannot-esc
  try {
    var ui = SpreadsheetApp.getUi();
    var ret = ui.alert(
      'COB-CHA',
      msg,
      ui.ButtonSet.OK
    );
  } catch (e) {
    var ret = '';
  }
  return {'ret': ret, 'buttonId': buttonId};
}

/**
 * show confirm
 * @param String msg
 * @return String
 */
function showConfirm(msg) {
  // #cannot-esc
  try {
    var ui = SpreadsheetApp.getUi();
    var ret = ui.alert(
      'COB-CHA',
      msg,
      ui.ButtonSet.OK_CANCEL
    );
  } catch (e) {
    var ret = '';
  }
  return ret;

/*
  return ui.alert(
    'COB-CHA',
    msg,
    ui.ButtonSet.OK_CANCEL
  );
*/
}

/**
 * Get URL from sheet
 * @param Object sheet
 * @return String
 */
function getUrlFromSheet(sheet) {
  return sheet.getRange(1, 2).getValue().toString();
}

/**
 * Get title from sheet
 * @param Object sheet
 * @return String
 */
function getTitleFromSheet(sheet) {
  return sheet.getRange(1, 6).getValue().toString();
}

/**
 * Get sheet by URL
 * @param String url
 * @return Object
 */
function getSheetByUrl(url) {
  if (getSheetByUrl.vals && getSheetByUrl.vals[url]) return getSheetByUrl.vals[url];

  var vals = {};
  var allSheets = getAllSheets();
  for (var i = 0; i < allSheets.length; i++) {
    var url = getUrlFromSheet(allSheets[i]);
    vals[url] = allSheets[i];
  }
  getSheetByUrl.vals = vals;
  return getSheetByUrl.vals[url];
}

/**
 * Get Property
 * @param String prop [lang, type, level]
 * @return String
 */
function getProp(prop) {
  if (getProp.vals && getProp.vals[prop]) return getProp.vals[prop];
  var userLocale = Session.getActiveUserLocale();
  
  var vals = {
    "lang"      : userLocale,
    "type"      : "wcag20",
    "level"     : "AA",
    "mark"      : ['?', '-', 'o', 'x'],
    "additional": ""
  };
  
  var sheet = getSheetIfExists(gConfigSheetName);
  if (sheet === false) return vals[prop];
  var rets = sheet.getRange(1, 2, 5, 2).getValues();
  
  vals['lang']  = ['en', 'ja'].indexOf(rets[0][0]) > -1 ? rets[0][0] : vals['lang'];
  vals['type']  = ['wcag20', 'wcag21', 'wcag22', 'tt20'].indexOf(rets[1][0]) > -1 ? rets[1][0] : vals['type'];
  vals['level'] = ['A', 'AA', 'AAA'].indexOf(rets[2][0]) > -1 ? rets[2][0] : vals['level'];
  vals['mark']  = rets[3][0].toString().charAt(0) == 'o' ? vals['mark'] : ['NT', 'DNA', 'T', 'F'];
  vals['additional']  = rets[4][0].toString();
  //  vals['lang']  = 'en';

  getProp.vals = vals;
  return getProp.vals[prop];
}

/**
 * Get Language Set
 * this function is language hard coding
 * @param String setName
 * @return Array
 */
function getLangSet(setName) {
  // ja
  if (getProp('lang') == 'ja') {
    switch (setName) {
      case 'version':    return getVersion();
      case 'criteria':   return getCriteriaJa();
      case 'ttCriteria': return getTtCriteriaJa();
      case 'tech':       return getTechValJa();
      case 'ui':         return getUiJa();
      // ICL: Japanese Only
      case 'iclList':                return getIclListJa();
      case 'iclSituationWaic':       return getIclSituationWaic();
      case 'iclTestWaic':            return getIclTestWaic();
      case 'iclSituationCobcha':     return getIclSituationCobcha();
      case 'iclTestCobcha':          return getIclTestCobcha();
      case 'iclSituationCobchaEasy': return getIclSituationCobchaEasy();
      case 'iclTestCobchaEasy':      return getIclTestCobchaEasy();
      case 'iclSituationIcollabo':   return getIclSituationIcollabo();
      case 'iclTestIcollabo':        return getIclTestIcollabo();
    }
  }
  
  // fallback - en
  switch (setName) {
    case 'version':    return getVersion();
    case 'criteria':   return getCriteriaEn();
    case 'ttCriteria': return getTtCriteriaEn();
    case 'tech':       return getTechValEn();
    case 'ui':         return {};
  }
  
  return {};
}

/**
 * Get Language UI Set
 * @param String uiname
 * @param String defaultStr
 * @return String
 */
function getUiLang(uiname, defaultStr) {
  var ui = getLangSet('ui');
  if (ui.length == 0 || ui[uiname] == null) {
    return defaultStr;
  }
  return ui[uiname];
}

/**
 * Get All Criteria Set
 * @param String type
 * @return Array
 */
function getAllCriteria(type) {
  var lang = getProp('lang');
  var type = type === undefined ? getProp('type') : type;
  var set = type.indexOf('wcag') >= 0 ? 'criteria' : 'ttCriteria' ;
  var allCriteria = getLangSet(set);

  // Trusted Tester does not apply additional criteria
  if (set == 'ttCriteria') return allCriteria;
  if (getAllCriteria.vals) return getAllCriteria.vals;
  
/*
  // #del-link
  // add URL
  var urlPointer = lang+'-'+type;
  for (var i = 0; i < allCriteria.length; i++) {
    var langPointer = type == 'wcag21' ? allCriteria[i][4] : allCriteria[i][3];
    allCriteria[i].push(gUrlbase['understanding'][urlPointer]+langPointer);
  }
  getAllCriteria.vals = allCriteria;
*/

  return allCriteria;
}

/**
 * Get Using Criteria Set
 * @param String type
 * @return Array
 */
function getUsingCriteria(type) {
  var type = type === undefined ? getProp('type') : type;
  var level = getProp('level');
  var usingCriteria = getAllCriteria(type);
  
  // Trusted Tester does not apply additional criteria
  if (type.indexOf('tt') >= 0) return usingCriteria;
//  if (getUsingCriteria.vals) return getUsingCriteria.vals;

  // additional criteria
  var additionalCriteriaArr = getProp('additional').split(/,/);
  var additionalCriteria = [];
  for (var i = 0; i < additionalCriteriaArr.length; i++) {
    additionalCriteria.push(additionalCriteriaArr[i].trim());
  }
  
  // eliminate unuse criteria
  for (var i = 0; i < usingCriteria.length; i++) {
    if (typeof usingCriteria[i] === 'undefined') continue;
    if (
      (type == 'wcag20' && (gCriteria21.indexOf(usingCriteria[i][1]) >= 0 || gCriteria22.indexOf(usingCriteria[i][1]) >= 0)) ||
      (type == 'wcag21' && gCriteria22.indexOf(usingCriteria[i][1]) >= 0) ||
      (type == 'wcag22' && usingCriteria[i][1] == '4.1.1') ||
      usingCriteria[i][0].length > level.length
    ) {
      if (additionalCriteria.indexOf(usingCriteria[i][1]) >= 0) continue;
      delete usingCriteria[i];
    }
  }
  
  usingCriteria = usingCriteria.filter(function(x){
	  return !(x === null || x === undefined || x === ""); 
  });
  
//  getUsingCriteria.vals = usingCriteria;
  
  return usingCriteria;
}

/**
 * Get Using Tech Set
 * @return Array
 */
function getUsingTechs() {
  if (getUsingTechs.vals) return getUsingTechs.vals;

  var lang = getProp('lang');
  var type = getProp('type');

  var techNames = getLangSet('tech');
  var urlPointer = lang+'-'+type;
  var usingCriteria = getUsingCriteria();
  var usingTechs = [];
  
  for (i = 0; i < usingCriteria.length; i++) {
    var criteria = usingCriteria[i][1];
    if (gRelTechsAndCriteria[criteria] == null) continue;
    for (j = 0; j < gRelTechsAndCriteria[criteria].length; j++) {
      var each = gRelTechsAndCriteria[criteria][j];
      
/*
      // #del-link
      // Techniques for WCAG 2.1 has directory
      var url = gUrlbase['tech'][urlPointer];
      if (type == 'wcag21' && lang == 'en') {
        var dir = each.charAt(0)+each.charAt(1);
        if (['M', 'L', 'V', 'C'].indefOf(each.charAt(1)) < 0) {
          dir = dir.charAt(0);
        }
        url += gTechDirAbbr[dir]+'/'+each;
      } else {
        url += each+'.html';
      }
      usingTechs.push([criteria, gRelTechsAndCriteria[criteria][j], techNames[each], url]);
*/
      usingTechs.push([criteria, gRelTechsAndCriteria[criteria][j], techNames[each]]);
    }
  }
  
  getUsingTechs.vals = usingTechs;
    
  return usingTechs;
}

/**
 * Get HTML and its title
 * @param String url
 * @return Object
 */
function getHtmlAndTitle(url) {
  var ret = {'title': '', 'html': ''};
  if (url.indexOf('http') < 0) {
    return ret;
  }
  
  var options = {
    "muteHttpExceptions" : true,
    "validateHttpsCertificates" : false,
    "followRedirects" : false,
  }

  try {
    var res = UrlFetchApp.fetch(url, options).getContentText();
    res = res == '' ? UrlFetchApp.fetch(url+'/', options).getContentText() : res;
    var title = res.match(/<title>.+?<\/title>/ig);
    title = String(title).replace(/<\/*title>/ig, '');
    title = title == null ? '' : title;
    return {'title': title, 'html': res};
  } catch(e) {
    return ret;
  }
}

/**
 * Get Spreadsheet
 * @return Object
 */
function getSpreadSheet() {
  if (getSpreadSheet.ss) return getSpreadSheet.ss;
  getSpreadSheet.ss = SpreadsheetApp.getActive();
  return getSpreadSheet.ss;
};

/**
 * Get Active Spreadsheet
 * @return Object
 */
function getActiveSheet() {
  if (getActiveSheet.ss) return getActiveSheet.ss;
  var ss = getSpreadSheet();
  getActiveSheet.ss = ss.getActiveSheet();
  return getActiveSheet.ss;
};

/**
 * Get All sheets
 * @return Object
 */
function getAllSheets() {
  if (getAllSheets.ss) return getAllSheets.ss;
  var ss = getSpreadSheet();
  var all = ss.getSheets();
  
  var ret = [];
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].getName()).charAt(0) == '*') continue;
    ret.push(all[i]);
  }

  getAllSheets.ss = ret;
  return getAllSheets.ss;
};

/**
 * Get Selected sheets
 * @return Array
 */
function getSelectedSheets() {
  if (getSelectedSheets.ss) return getSelectedSheets.ss;

  var ss = getSpreadSheet();
  var urlsheet = ss.getSheetByName(gUrlListSheetName);
  if (urlsheet == null) return [];
  var rownum = urlsheet.getLastRow() - 1;
  if (rownum <= 1) return [];
  
  var urls = urlsheet.getRange(2, 1, rownum, 4).getValues();
  var targets = []
  for (var i = 0; i < urls.length; i++) {
    if (urls[i][3] != "o") continue;
//    if (urls[i][3] == "x") continue;
    targets.push(urls[i][0]);
  }
  
  var all = getAllSheets();
  var ret = [];
  for (var i = 0; i < all.length; i++) {
    if (targets.indexOf(String(all[i].getName())) == -1) continue;
    ret.push(all[i]);
  }

  getSelectedSheets.ss = ret;
  return getSelectedSheets.ss;
};

/**
 * reset sheets
 * @param Bool isAll
 * @return String
 */
function resetSheets(isAll) {
  var id = isAll ? 'reset-all-sheets' : 'reset-record-sheets';
  var msg = getUiLang('reset-caution', 'CAUTION: Reset Sheets?');
  if(showConfirm(msg) != "OK") return {'msg':getUiLang('canceled', 'canceled'), 'id': id};

  var ss = getSpreadSheet();
  var all = ss.getSheets();
  
  deleteFallbacksheet();
  ss.insertSheet(gFallbackSheetName, 0);
  
  var count = 0;
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName() == gFallbackSheetName) continue;
    if (isAll === false && all[i].getName().charAt(0) == '*') continue;
    if (all[i] == null) continue;
    ss.deleteSheet(all[i]);
    count++;
  }
  var all2 = ss.getSheets();
  if (all2.length > 1) {
    deleteFallbacksheet();
  }
    
  return {
    'msg': getUiLang('sheet-deleted', "%s sheet(s) deleted.").replace("%s", count),
    'id': id
  };
}

/**
 * delete fallbacksheet
 * @return Void
 */
function deleteFallbacksheet() {
  deleteSheetIfExist(gFallbackSheetName);
}

/**
 * delete sheet if exists
 * @return Void
 */
function deleteSheetIfExist(sheetName) {
  var sheet = getSheetIfExists(sheetName);
  if ( ! sheet) return;
  getSpreadSheet().deleteSheet(sheet);
}

/**
 * is sheet Exist
 * @param String sheetName
 * @return Bool
 */
function isSheetExist(sheetName) {
  return (getSheetIfExists(sheetName));
}

/**
 * get target sheet if Exists
 * @param String sheetName
 * @return Bool | Object
 */
function getSheetIfExists(sheetName) {
  var ss = getSpreadSheet();
  var targetSheet = ss.getSheetByName(sheetName);
  return (targetSheet) ? targetSheet : false;
}

/**
 * generate sheet if not exists
 * @param String sheetName
 * @param Array defaults
 * @param String [header = "row"]
 * @return Object | String
 */
function generateSheetIfNotExists(sheetName, defaults, header) {
  if (isSheetExist(sheetName)) return getUiLang('target-sheet-already-exists', "Target sheet (%s) is already exists.").replace('%s', sheetName);
  return prepareTargetSheet(sheetName, defaults, header);
}

/**
 * generate sheet even if already exists
 * @param String sheetName
 * @return Object
 */
function generateSheetEvenIfAlreadyExists(sheetName) {
  var ss = getSpreadSheet();
  if (targetSheet = getSheetIfExists(sheetName)) {
    ss.deleteSheet(targetSheet);
  }
  var sheet = ss.insertSheet(sheetName, 0);
  deleteFallbacksheet();
  return sheet;
}

/**
 * prepare target sheet
 * @param String sheetName
 * @param Array defaults
 * @param String [header = "row"]
 * @return Object
 */
function prepareTargetSheet(sheetName, defaults, header) {
  var ss = getSpreadSheet();
  ss.insertSheet(sheetName, 0);
  var sheet = ss.getSheetByName(sheetName);
  sheet.activate();
  if (defaults)
  {
    sheet.getRange(1, 1, defaults.length, defaults[0].length).setValues(defaults);
    if (header === 'row') {
      sheet.getRange("1:1").setBackground(gLabelColor);
      sheet.setFrozenRows(1);
    } else {
      sheet.getRange(1, 1, defaults.length, 1).setBackground(gLabelColor);
    }
    sheet.autoResizeColumn(1);
  }
  deleteFallbacksheet();
  return sheet;
}

/**
 * setRowConditionApplicability
 * @param Object sheet
 * @param Object range
 * @param String cond
 * @return Object
 */
function setRowConditionApplicability(sheet, range, cond) {
  var ruleNonApp = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(cond)
      .setFontColor(gLabelColor)
      .setRanges([range])
      .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(ruleNonApp);
  sheet.setConditionalFormatRules(rules);
}

/**
 * setRowConditionNotYet
 * @param Object sheet
 * @param Object range
 * @param String cond
 * @param bool isBold
 * @param String textColor
 * @param String bgColor
 * @return Object
 */
function setRowConditionNotYet(sheet, range, cond, isBold, textColor, bgColor) {
  var isBold    = isBold    === undefined ? true : isBold;
  var textColor = textColor === undefined ? "#000000" : textColor; 
  var bgColor   = bgColor   === undefined ? gNotYetBgColor : bgColor;
  
  var ruleNonApp = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(cond)
      .setBold(isBold)
      .setFontColor(textColor)
      .setBackground(bgColor)
      .setRanges([range])
      .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(ruleNonApp);
  sheet.setConditionalFormatRules(rules);
}

/**
 * setCellConditionTF
 * @param Object sheet
 * @param Object range
 * @param String mT
 * @param String mF
 * @return Object
 */
function setCellConditionTF(sheet, range, mT, mF) {
  var ruleForF = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mF)
      .setBackground(gFalseColor)
      .setRanges([range])
      .build();
  var ruleForT = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(mT)
      .setBackground(gTrueColor)
      .setRanges([range])
      .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(ruleForF);
  rules.push(ruleForT);
  sheet.setConditionalFormatRules(rules);
}
