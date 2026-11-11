/**
 * COB-CHA: CollaBorative CHeck tool for Accessibility
 * Google Spreadsheet Add-on.
 * @Author  shibata@jidaikobo.com
 *          arimatsu@jidaikobo.com
 * @Year    2021 -
 * @Licence MIT
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

    // ensure issue sheet rules
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(gIssueSheetName);
    if (sheet) ensureGlobalIssueRules(sheet);
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
  menu.addItem(getUiLang('cob-cha-help', 'COB-CHA Help'), 'showHelp');
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
  showDialog('ui-help', 500, 400, getUiLang('cob-cha-help', 'COB-CHA Help'));
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
 * @param Object templateVars
 * @return Void
 */
function showDialog(sheetname, width, height, title, html, templateVars) {
  var t = HtmlService.createTemplateFromFile(sheetname);

  t.I18N = getLangSet('ui') || {};        // dictionary
  t.LANG = getProp('lang') || 'en';       // lang
  t.TEMPLATE_VARS = templateVars || {};

  var out = t.evaluate()
    .setWidth(width || 800)
    .setHeight(height || 600)
    .setTitle(title || '');

  if (html) out.append(html);

  var ss = getSpreadSheet();
  ss.show(out);
}

/**
 * show alert
 * @param String msg
 * @param String buttonId
 * @return String
 */
function showAlert(msg, buttonId) {
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
  var targetUrl = String(url || '');
  if (getSheetByUrl.vals && getSheetByUrl.vals[targetUrl]) {
    return getSheetByUrl.vals[targetUrl];
  }

  var vals = {};
  var allSheets = getAllSheets();
  for (var i = 0; i < allSheets.length; i++) {
    var sheetUrl = getUrlFromSheet(allSheets[i]);
    vals[sheetUrl] = allSheets[i];
  }
  getSheetByUrl.vals = vals;
  return getSheetByUrl.vals[targetUrl];
}

/**
 * Get Property
 * @param String prop [lang, type, level, mark, additional, criteriaNameSet]
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
    "additional": "",
    "criteriaNameSet": "wcag-ja"
  };

  var sheet = getSheetIfExists(gConfigSheetName);
  if (sheet === false) return vals[prop];
  var rets = sheet.getRange(1, 2, 6, 2).getValues();

  vals['lang']  = ['en', 'ja'].indexOf(rets[0][0]) > -1 ? rets[0][0] : vals['lang'];
  vals['type']  = ['wcag20', 'wcag21', 'wcag22', 'tt20'].indexOf(rets[1][0]) > -1 ? rets[1][0] : vals['type'];
  vals['level'] = ['A', 'AA', 'AAA'].indexOf(rets[2][0]) > -1 ? rets[2][0] : vals['level'];
  vals['mark']  = rets[3][0].toString().charAt(0) == 'o' ? vals['mark'] : ['NT', 'DNA', 'T', 'F'];
  vals['additional']  = rets[4][0].toString();
  vals['criteriaNameSet'] = ['wcag-ja', 'jis2016-ja'].indexOf(rets[5][0]) > -1 ?
    rets[5][0] : vals['criteriaNameSet'];
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
      case 'iclSituationCobchaPdf':  return getIclSituationCobchaPdf();
      case 'iclTestCobchaPdf':       return getIclTestCobchaPdf();
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

  var translatedNames = lang == 'ja' ? getCriteriaNamesJaWcagVersion(type) : {};
  if (lang == 'ja' && Object.keys(translatedNames).length > 0) {
    allCriteria = allCriteria.map(function (criterion) {
      var copied = criterion.slice();
      if (translatedNames[copied[1]]) copied[2] = translatedNames[copied[1]];
      return copied;
    });
  }

  if (lang == 'ja' && getProp('criteriaNameSet') == 'jis2016-ja') {
    var jisNames = getCriteriaNamesJaJis2016();
    allCriteria = allCriteria.map(function (criterion) {
      var copied = criterion.slice();
      if (jisNames[copied[1]]) copied[2] = jisNames[copied[1]];
      return copied;
    });
  }

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
      // If 4.1.1 is excluded from WCAG 2.2, JIS X 8341-3:2016 cannot add the success criteria of WCAG 2.2
      // (type == 'wcag22' && usingCriteria[i][1] == '4.1.1') ||
      usingCriteria[i][0].length > level.length
    ) {
      if (additionalCriteria.indexOf(usingCriteria[i][1]) >= 0) continue;
      delete usingCriteria[i];
    }
  }

  usingCriteria = usingCriteria.filter(function(x){
	  return !(x === null || x === undefined || x === "");
  });

  // --- annotate labels: WCAG 2.2 section 4.1.1 is marked as "(deprecated/deleted)" ---
  if (type === 'wcag22' && getProp('criteriaNameSet') !== 'jis2016-ja') {
    var suffix = ' ' + getUiLang('obsolete-removed', '(Obsolete and removed)');
    for (var k = 0; k < usingCriteria.length; k++) {
      if (!usingCriteria[k]) continue;
      if (usingCriteria[k][1] === '4.1.1' && usingCriteria[k][2].indexOf(suffix) === -1) {
        usingCriteria[k][2] += suffix;
      }
    }
  }

//  getUsingCriteria.vals = usingCriteria;

  return usingCriteria;
}

/**
 * Get Using Tech Set
 * @return Array
 */
function getUsingTechs() {
  if (getUsingTechs.vals) return getUsingTechs.vals;

  var type = getProp('type');

  var techNames = getLangSet('tech');
  var usingCriteria = getUsingCriteria();
  var usingTechs = [];

  for (var i = 0; i < usingCriteria.length; i++) {
    var criteria = usingCriteria[i][1];
    if (gRelTechsAndCriteria[criteria] == null) continue;
    for (var j = 0; j < gRelTechsAndCriteria[criteria].length; j++) {
      var each = gRelTechsAndCriteria[criteria][j];

      usingTechs.push([criteria, gRelTechsAndCriteria[criteria][j], techNames[each]]);
    }
  }

  getUsingTechs.vals = usingTechs;

  return usingTechs;
}

/**
 * Normalize a URL for UrlFetchApp.
 * @param {*} url
 * @return {String}
 */
function normalizeFetchUrl(url) {
  var normalized = String(url || '').trim();
  if (normalized.length > 2082) return '';
  if ( ! /^https?:\/\/[^\s]+$/i.test(normalized)) return '';
  var parts = getHttpUrlParts(normalized);
  if (!parts || !isAllowedFetchHostname(parts.hostname)) return '';
  return normalized;
}

/**
 * Parse the parts needed to validate and resolve an HTTP(S) URL.
 * @param String url
 * @return Object|null
 */
function getHttpUrlParts(url) {
  var match = String(url || '').match(
    /^(https?):\/\/([^\/?#]+)([^?#]*)(\?[^#]*)?(?:#.*)?$/i
  );
  if (!match || match[2].indexOf('@') >= 0 || match[2].indexOf('%') >= 0) {
    return null;
  }

  var authority = match[2];
  var hostname = authority;
  if (authority.charAt(0) == '[') {
    var bracketEnd = authority.indexOf(']');
    if (bracketEnd < 0) return null;
    hostname = authority.substring(1, bracketEnd);
    var ipv6Port = authority.substring(bracketEnd + 1);
    if (ipv6Port != '' && !/^:\d+$/.test(ipv6Port)) return null;
  } else {
    var portMatch = authority.match(/^(.*?):(\d+)$/);
    if (portMatch) {
      hostname = portMatch[1];
    } else if (authority.indexOf(':') >= 0) {
      return null;
    }
  }

  if (hostname == '') return null;
  return {
    'scheme': match[1].toLowerCase(),
    'authority': authority,
    'hostname': hostname.toLowerCase().replace(/\.$/, ''),
    'path': match[3] || '/',
    'query': match[4] || ''
  };
}

/**
 * Reject local, private, reserved, and ambiguous hostnames.
 * @param String hostname
 * @return Bool
 */
function isAllowedFetchHostname(hostname) {
  var host = String(hostname || '').toLowerCase();
  if (host == '' || host == 'localhost' || /\.(?:localhost|local|internal)$/.test(host)) {
    return false;
  }

  // Domain names may resolve to IPv6, but direct IPv6 literals are ambiguous
  // to validate safely in Apps Script and are therefore rejected.
  if (host.indexOf(':') >= 0) return false;

  // Reject single-label and ambiguous numeric hosts such as 2130706433,
  // 0177.0.0.1, and 0x7f.0.0.1.
  if (host.indexOf('.') < 0) return false;
  var numericHost = /^(?:0x[0-9a-f]+|\d+)(?:\.(?:0x[0-9a-f]+|\d+)){0,3}$/i;
  var canonicalIpv4 = /^(?:0|[1-9]\d{0,2})(?:\.(?:0|[1-9]\d{0,2})){3}$/;
  if (numericHost.test(host) && !canonicalIpv4.test(host)) return false;

  if (!canonicalIpv4.test(host)) return true;
  var octets = host.split('.').map(Number);
  if (octets.some(function(octet) { return octet > 255; })) return false;

  var first = octets[0];
  var second = octets[1];
  if (first == 0 || first == 10 || first == 127 || first >= 224) return false;
  if (first == 100 && second >= 64 && second <= 127) return false;
  if (first == 169 && second == 254) return false;
  if (first == 172 && second >= 16 && second <= 31) return false;
  if (first == 192 && (second == 0 || second == 168)) return false;
  if (first == 198 && (second == 18 || second == 19)) return false;
  return true;
}

/**
 * Resolve an HTTP redirect and validate the resulting URL.
 * @param String baseUrl
 * @param String location
 * @return String
 */
function resolveFetchRedirectUrl(baseUrl, location) {
  var next = String(location || '').trim();
  if (next == '') return '';
  if (/^https?:\/\//i.test(next)) return normalizeFetchUrl(next);

  var base = getHttpUrlParts(baseUrl);
  if (!base) return '';
  var origin = base.scheme + '://' + base.authority;
  if (next.indexOf('//') == 0) {
    return normalizeFetchUrl(base.scheme + ':' + next);
  }
  if (next.charAt(0) == '/') return normalizeFetchUrl(origin + next);
  if (next.charAt(0) == '?') return normalizeFetchUrl(origin + base.path + next);

  var directory = base.path.replace(/[^\/]*$/, '');
  var segments = (directory + next).split('/');
  var resolved = [];
  for (var i = 0; i < segments.length; i++) {
    if (segments[i] == '' || segments[i] == '.') continue;
    if (segments[i] == '..') {
      resolved.pop();
    } else {
      resolved.push(segments[i]);
    }
  }
  return normalizeFetchUrl(origin + '/' + resolved.join('/'));
}

/**
 * Get an HTTP response header without depending on header-name casing.
 * @param Object response
 * @param String targetName
 * @return String
 */
function getResponseHeader(response, targetName) {
  var headers = response.getHeaders();
  var normalizedTarget = targetName.toLowerCase();
  for (var name in headers) {
    if (name.toLowerCase() == normalizedTarget) return String(headers[name]);
  }
  return '';
}

/**
 * Whether a response represents HTML that COB-CHA can inspect.
 * @param Object response
 * @return Bool
 */
function isHtmlFetchResponse(response) {
  var contentType = getResponseHeader(response, 'Content-Type')
    .toLowerCase()
    .split(';')[0]
    .trim();
  return contentType == '' || contentType == 'text/html' ||
    contentType == 'application/xhtml+xml';
}

/**
 * Extract a normalized title from HTML.
 * @param {*} html
 * @return {String}
 */
function extractHtmlTitle(html) {
  var match = String(html || '').match(/<title\b[^>]*>([\s\S]*?)<\/title>/i);
  if ( ! match) return '';
  return match[1].replace(/\s+/g, ' ').trim();
}

/**
 * Preserve an untrusted string as text when writing it to a spreadsheet cell.
 * @param {*} value
 * @return {*}
 */
function toPlainTextCellValue(value) {
  if (typeof value != 'string') return value;
  var formulaPrefixes = ['=', '+', '-', '@'];
  if (formulaPrefixes.indexOf(value.charAt(0)) >= 0) return "'" + value;
  return value;
}

/**
 * Summarize check marks with one shared precedence.
 *
 * Not applicable is returned only when every value is not applicable. A
 * confirmed failure takes precedence over incomplete values; incomplete
 * values take precedence over conformance.
 *
 * @param Array values
 * @param Array marks [unknown, not applicable, pass, fail]
 * @return String
 */
function summarizeEvaluationMarks(values, marks) {
  var unknown = marks[0];
  var notApplicable = marks[1];
  var pass = marks[2];
  var fail = marks[3];
  if (values.length === 0) return notApplicable;
  if (values.every(function (value) { return value === notApplicable; })) {
    return notApplicable;
  }
  if (values.indexOf(fail) !== -1) return fail;
  if (values.indexOf('') !== -1 || values.indexOf(unknown) !== -1) {
    return unknown;
  }
  if (values.indexOf(pass) !== -1) return pass;
  return unknown;
}

/**
 * Get HTML and its title
 * @param String url
 * @return Object
 */
function getHtmlAndTitle(url, skipSlashFallback) {
  var ret = {'title': '', 'html': '', 'status': null, 'error': ''};
  var targetUrl = normalizeFetchUrl(url);
  if (targetUrl == '') {
    ret.error = 'invalid URL';
    return ret;
  }

  var options = {
    "muteHttpExceptions" : true,
    "validateHttpsCertificates" : true,
    "followRedirects" : false,
    "timeoutSeconds" : 30,
  }

  try {
    var response;
    var status;
    var redirectCodes = [301, 302, 303, 307, 308];
    for (var i = 0; i <= 5; i++) {
      response = UrlFetchApp.fetch(targetUrl, options);
      status = response.getResponseCode();
      ret.status = status;
      if (redirectCodes.indexOf(status) < 0) break;
      if (i == 5) {
        ret.error = 'too many redirects';
        return ret;
      }
      targetUrl = resolveFetchRedirectUrl(
        targetUrl,
        getResponseHeader(response, 'Location')
      );
      if (targetUrl == '') {
        ret.error = 'invalid redirect URL';
        return ret;
      }
    }

    if (status < 200 || status >= 300) return ret;
    if (!isHtmlFetchResponse(response)) {
      ret.error = 'non-HTML response';
      return ret;
    }
    var contentLength = Number(getResponseHeader(response, 'Content-Length'));
    if (contentLength > 5000000) {
      ret.error = 'response too large';
      return ret;
    }
    var html = response.getContentText();
    if (html.length > 5000000) {
      ret.error = 'response too large';
      return ret;
    }
    if (
      html == '' && !skipSlashFallback &&
      targetUrl.charAt(targetUrl.length - 1) != '/'
    ) {
      return getHtmlAndTitle(targetUrl + '/', true);
    }

    if (html == '') {
      ret.error = 'empty response';
      return ret;
    }
    return {'title': extractHtmlTitle(html), 'html': html, 'status': status, 'error': ''};
  } catch(e) {
    // A previous redirect status does not describe a failed subsequent request.
    ret.status = null;
    ret.error = 'fetch error';
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
 * Get selected sheets.
 *
 * - If filterKey is omitted, use the header of column E.
 * - Always filter based on the column's "o".
 *
 * @param {String=} filterKey
 * @return {Sheet[]}
 */
function getSelectedSheets(filterKey) {

  var ss = getSpreadSheet();
  var sheet = ss.getSheetByName(gUrlListSheetName);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  // Read header row
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // ------------------------------------------------------------------
  // If filterKey is not given → use column E's header
  // (E column = 5th column = index 4)
  // ------------------------------------------------------------------
  if (!filterKey) {
    filterKey = headers[4];   // E列のヘッダ名
  }

  // Find the index of the header
  var colIndex = headers.indexOf(filterKey);
  if (colIndex === -1) return [];  // no such header

  // Read all data rows
  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Find sheet names where the filter column has "o"
  var targets = [];
  rows.forEach(function(row){
    if (row[colIndex] === "o") {
      targets.push(String(row[0])); // A列 = sheet name
    }
  });

  // Return only matched sheets
  var allSheets = getAllSheets();
  return allSheets.filter(function(sh){
    return targets.indexOf(sh.getName()) !== -1;
  });
}

/**
 * Get headers and their count of "o" in the column
 * @return {Array<{header: string, count: number}>}
 */
function getUrlSheetHeadersWithCount() {
  var ss = getSpreadSheet();
  var sheet = ss.getSheetByName(gUrlListSheetName);
  if (!sheet) return [];

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastCol < 5 || lastRow < 2) return [];

  var headerRow = sheet.getRange(1, 5, 1, lastCol - 4).getValues()[0];
  var data = sheet.getRange(2, 5, lastRow - 1, lastCol - 4).getValues();

  var result = [];

  for (var col = 0; col < headerRow.length; col++) {
    var header = headerRow[col].toString().trim();
    var count = 0;

    for (var r = 0; r < data.length; r++) {
      if (data[r][col] === "o") count++;
    }

    result.push({ header: header, count: count });
  }

  return result;
}

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
  var targetSheet = getSheetIfExists(sheetName);
  if (targetSheet) {
    ss.deleteSheet(targetSheet);
  }
  var sheet = ss.insertSheet(sheetName, 0);
  deleteFallbacksheet();
  return sheet;
}

/**
 * Apply a lightweight grid to a checklist range in one batch operation.
 * @param Object range
 * @return Object
 */
function setChecklistBorders(range) {
  return range.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    '#d9d9d9',
    SpreadsheetApp.BorderStyle.SOLID
  );
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
