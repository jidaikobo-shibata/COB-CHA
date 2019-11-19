/**
 * COB-CHA: CollaBorative CHeck tool for Accessibility
 * powered by Google Spreadsheet
 * @Author shibata@jidaikobo.com
 * @Licence CC0
 */

/**
 * WCAG 2.1
 */
var criteria21 = [
  '1.3.4', '1.3.5', '1.3.6', '1.4.10', '1.4.11', '1.4.12', '1.4.13',
  '2.1.4', '2.2.6', '2.3.3', '2.5.1', '2.5.2', '2.5.3', '2.5.4', '2.5.5', '2.5.6',
  '4.1.3'
]

/**
 * WCAG 2.0/2.1 Single-A criteria
 */
var cCheckVal = [
  '1.1.1', '1.2.1', '1.2.2', '1.2.3', '1.3.1', '1.3.2', '1.3.3', '1.4.1',
  '1.4.2', '2.1.1', '2.1.2', '2.1.4', '2.2.1', '2.2.2', '2.3.1', '2.4.1',
  '2.4.2', '2.4.3', '2.4.4', '2.5.1', '2.5.2', '2.5.3', '2.5.4', '3.1.1',
  '3.2.1', '3.2.2', '3.3.1', '3.3.2', '4.1.1', '4.1.2'
];

/**
 * Non-Interference
 */
var nonInterference = [
  '1.4.2', '2.1.2', '2.2.2', '2.3.1'
];

/**
 * URL
 */
var urlbase = {
  'understanding': {
    'en-wcag20': 'https://www.w3.org/TR/UNDERSTANDING-WCAG20/',
    'en-wcag21': 'https://www.w3.org/WAI/WCAG21/Understanding/', // and directory
    'ja-wcag20': 'https://waic.jp/docs/UNDERSTANDING-WCAG20/',
    'ja-wcag21': 'https://waic.jp/docs/UNDERSTANDING-WCAG20/'
  },
  'tech': {
    'en-wcag20': 'https://www.w3.org/TR/WCAG20-TECHS/',
    'en-wcag21': 'https://www.w3.org/WAI/WCAG21/Techniques/',
    'ja-wcag20': 'https://waic.jp/docs/WCAG-TECHS/',
    'ja-wcag21': 'https://waic.jp/docs/WCAG-TECHS/'
  }
};

var techDirAbbr = {
  'G': 'general',
  'H': 'html',
  'C': 'css',
  'A': 'aria',
  'T': 'text',
  'P': 'pdf',
  'F': 'failures',
  'FL': 'flash',
  'SM': 'smil',
  'SL': 'silverlight',
  'SV': 'server-side-script',
  'SC': 'client-side-script'
}

/**
 * global variables
 */
var resultSheetName = '*Result*';
var issueSheetName = '*Issue*';
var templateSheetName = '*Template*';
var resourceFolderName = 'resource';
var exportFolderName = 'export';
var imagesFolderName = 'images';
var issueFileName = 'issue-report';
var trueColor    = '#f5fff3';
var falseColor   = '#f7f3ff';
var labelColor   = '#eeeeee';
var doubleAColor = '#eeeefe';

/**
 * onInstall
 * @param Object e
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * onOpen
 * @param Object e
 */
function onOpen (e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Show Control Panel', 'showSidebar');
  menu.addToUi();
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
  
  ret = [];
  for (i = 0; i < all.length; i++) {
    if (String(all[i].getName()).charAt(0) == '*') continue;
    ret.push(all[i]);
  }

  getAllSheets.ss = ret;
  return getAllSheets.ss;
};

/**
 * Get Spreadsheet's folder
 * @return Object
 */
function getCurrentFolder() {
  if (getCurrentFolder.folder) return getCurrentFolder.folder;
  var ss = getSpreadSheet();
  var ssId = ss.getId();
  var parentFolder = DriveApp.getFileById(ssId).getParents();
  getCurrentFolder.folder = parentFolder.next();
  return getCurrentFolder.folder; 
};

/**
 * Get target folder object
 * @param String target [resourceFolderName, exportFolderName, imagesFolderName]
 * @return Object
 */
function getTargetFolder(target) {
  if (getTargetFolder.folder && getTargetFolder.folder[target]) return getTargetFolder.folder[target];
  var currentFolder = getCurrentFolder();
  var children = currentFolder.getFolders();
  getTargetFolder.folder = {};
  
  // is already exists?
  var folders = [resourceFolderName, exportFolderName, imagesFolderName]
  while (children.hasNext()){
    var folder = children.next();
    for (var i = 0; i < folders.length; i++) {
      if (folder.getName().indexOf(folders[i]) != -1) {
        getTargetFolder.folder[folders[i]] = folder;
      }
    }
    if (getTargetFolder.folder[target]) return getTargetFolder.folder[target];
  }
  
  // create folder
  for (var i = 0; i < folders.length; i++) {
    getTargetFolder.folder[folders[i]] = currentFolder.createFolder(folders[i]);
  }
  return getTargetFolder.folder[target];
};

/**
 * delete file if exists
 * @param String target
 * @param String name
 */
function deleteFileIfExists(targetFolder, name) {
  var targetFolder = getTargetFolder(targetFolder);
  var children = targetFolder.getFiles();
  while (children.hasNext()) {
    var current = children.next();
    if (current.getName() == name) {
      targetFolder.removeFile(current);
    }
  };
};

/**
 * save HTML
 * @param String target
 * @param String name
 * @param String html
 * @param Bool overwrite
 */
function saveHtml(targetFolder, name, html, overwrite) {
  if (overwrite) {
    deleteFileIfExists(targetFolder, name);
  }
  var targetFolder = getTargetFolder(targetFolder);
  targetFolder.createFile(name, html, 'text/html');
};

/**
 * show control pannel
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle("COB-CHA Control Panel");
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Get Property
 * @param String prop [lang, type, level]
 * @return String
 */
function getProp(prop) {
  if (getProp.vals && getProp.vals[prop]) return getProp.vals[prop];
  
  var activeSheet = getActiveSheet();
  var rets = activeSheet.getRange(1, 2, 1, 3).getValues();
  var vals = {};
  vals['lang']  = ['en', 'ja'].indexOf(rets[0][0]) > -1 ? rets[0][0] : 'en';
  vals['type']  = ['wcag20', 'wcag21', 'tt20'].indexOf(rets[0][1]) > -1 ? rets[0][1] : 'wcag21';
  vals['level'] = ['A', 'AA', 'AAA'].indexOf(rets[0][2]) > -1 ? rets[0][2] : 'AA';
  getProp.vals = vals;
  
  return getProp.vals[prop] ? getProp.vals[prop] : 'en' ;
}

/**
 * Get Language Set
 * @param String setName
 * @return Array
 */
function getLangSet(setName) {
  // ja
  if (getProp('lang') == 'ja') {
    switch (setName) {
      case 'criteria':   return getCriteriaJa();
      case 'ttCriteria': return getTtCriteriaJa();
      case 'ttCheckVal': return getTtCheckValJa();
      case 'tech':       return getTechValJa();
    }
  }

  // fallback - en
  switch (setName) {
    case 'criteria':   return getCriteriaEn();
    case 'ttCriteria': return getTtCriteriaEn();
    case 'ttCheckVal': return getTtCheckValEn();
    case 'tech':       return getTechValEn();
  }
}

/**
 * add sheet
 * @param String sheetname
 * @param String template
 * @return Bool
 */
function addSheet(sheetname, template) {
  var ss = getSpreadSheet();
  var targetSheet = ss.getSheetByName(sheetname);
  var sheetIndex  = sheetname.charAt(0) == '*' ? 0 : ss.getSheets().length+1;

  // sheet which name started with * must be refreashed
  if (sheetIndex == 0 && targetSheet != null) {
    ss.deleteSheet(targetSheet);
  }
  if (ss.getSheetByName(sheetname)) return false;
  if (template) {
    ss.insertSheet(sheetname, sheetIndex, {template: template});
  } else {
    ss.insertSheet(sheetname, sheetIndex);
  }
  return true;
}

/**
 * Get HTML and it's title
 * @param String url
 * @return Array
 */
function getHtmlAndTitle(url) {
  try {
    var res = UrlFetchApp.fetch(url).getContentText();
    var title = res.match(/<title>.+?<\/title>/ig);
    title = String(title).replace(/<\/*title>/ig, '');
    return {'title': title, 'html': res};
  } catch(e) {
    return {'title': '', 'html': ''};
  }
}
