/**
 * COB-CHA: CollaBorative CHeck tool for Accessibility
 * powered by Google Spreadsheet
 * @Author shibata@jidaikobo.com
 *         arimatsu@jidaikobo.com
 * @Licence MIT
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
function onOpen (e) {
  if(e && e.authMode == 'NONE'){
    var menu = SpreadsheetApp.getUi().createAddonMenu();
    menu.addItem('Getting Started', 'askEnabled');
    menu.addToUi();
  } else {
    addShowControlPannel();
  }
}

/**
 * askEnabled
 * @param Object e
 * @return Void
 */
function askEnabled() {
  var title = 'COB-CHA';
  var msg = 'Script has been enabled.';
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  addShowControlPannel(menu)
};

/**
 * add "Show Control Pannel" to menu
 * @return Void
 */
function addShowControlPannel() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem(getUiLang('show-control-panel', 'Show Control Panel'), 'showControlPannel');
  menu.addItem(getUiLang('help', 'COB-CHA Help'), 'showHelp');
  menu.addToUi();
};

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
 * @return Void
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
 * @return Void
 */
function saveHtml(targetFolder, name, html, overwrite) {
  if (overwrite) {
    deleteFileIfExists(targetFolder, name);
  }
  var targetFolder = getTargetFolder(targetFolder);
  targetFolder.createFile(name, html, 'text/html');
};

/**
 * file upload
 * @param String targetFolder
 * @param Object formObj
 * @param String nameAttr
 * @return Array [fileName, fileId]
 */
function fileUpload(targetFolder, formObj, nameAttr) {
  if (formObj[nameAttr].length == 0) throw new Error('Empty File Uploaded');
  var formBlob = formObj[nameAttr];
  var driveFile = DriveApp.createFile(formBlob);
  var targetFolder = getTargetFolder(targetFolder);
  deleteFileIfExists(targetFolder, driveFile.getName());
  targetFolder.addFile(driveFile);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getRootFolder().removeFile(driveFile);
  return [driveFile.getName() ,driveFile.getId()];
}

/**
 * add image formula
 * @param String id
 * @return String
 */
function addImageFormula(id) {
  return '=IMAGE("https://drive.google.com/uc?export=download&id='+id+'",1)';
};

/**
 * remove image formula
 * @param String id
 * @return String
 */
function removeImageFormula(id) {
  id = id.replace('=IMAGE("https://drive.google.com/uc?export=download&id=' ,'');
  id = id.replace('",1)', '');
  return id;
};

/**
 * show control pannel
 * @return Void
 */
function showControlPannel() {
  var ui = HtmlService.createTemplateFromFile('ui-control-pannel')
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
  html = html == null ? '' : html;
  var html = output.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setWidth(width)
                   .setHeight(height)
                   .setTitle(title)
                   .append(html);
  ss.show(html);
}

/**
 * Get First Column
 * @return String
 */
function getFirstColumn() {
  var activeSheet = getActiveSheet();
  var activeRow = activeSheet.getActiveCell().getRow();
  var criterion = activeSheet.getRange(activeRow, 1).getValue();
  criterion = criterion.match(/^\d\.\d\.\d+/) || criterion.match(/^\d+\.\w/) ? criterion : '';
  return criterion;
}

/**
 * Get URL from sheet
 * @param Object sheet
 * @return String
 */
function getUrlFromSheet(sheet) {
  return sheet.getRange(2, 2).getValue();
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

  var activeSheet = getActiveSheet();
  var rets = activeSheet.getRange(1, 2, 1, 3).getValues();
  var vals = {};
  var userLocale = Session.getActiveUserLocale();
  //  activeSheet.getRange(1,1).setValue(userLocale);
  
  userLocale    = ['en', 'ja'].indexOf(userLocale) > -1 ? userLocale : 'en';
  vals['lang']  = ['en', 'ja'].indexOf(rets[0][0]) > -1 ? rets[0][0] : userLocale;
  vals['type']  = ['wcag20', 'wcag21', 'tt20'].indexOf(rets[0][1]) > -1 ? rets[0][1] : 'wcag21';
  vals['level'] = ['A', 'AA', 'AAA'].indexOf(rets[0][2]) > -1 ? rets[0][2] : 'AA';
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
      case 'criteria':   return getCriteriaJa();
      case 'ttCriteria': return getTtCriteriaJa();
      case 'tech':       return getTechValJa();
      case 'ui':         return getUiJa();
      // ICL: Japanese Only
      case 'iclSituation': return getIclSituation();
      case 'iclTest':      return getIclTest();
    }
  }
  
  // fallback - en
  switch (setName) {
    case 'criteria':   return getCriteriaEn();
    case 'ttCriteria': return getTtCriteriaEn();
    case 'tech':       return getTechValEn();
    case 'ui':         return {};
  }
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
 * @param String lang
 * @param String type
 * @return Array
 */
function getAllCriteria(lang, type) {
  var set = type.indexOf('wcag') >= 0 ? 'criteria' : 'ttCriteria' ;
  var allCriteria = getLangSet(set);
  
  // Trusted Tester does not apply additional criteria
  if (set == 'ttCriteria') return allCriteria;
  if (getAllCriteria.vals) return getAllCriteria.vals;
  
  // add URL
  var urlPointer = lang+'-'+type;
  for (var i = 0; i < allCriteria.length; i++) {
    var langPointer = type == 'wcag21' ? allCriteria[i][4] : allCriteria[i][3];
    allCriteria[i].push(urlbase['understanding'][urlPointer]+langPointer);
  }
  getAllCriteria.vals = allCriteria;
    
  return allCriteria;
}

/**
 * Get Using Criteria Set
 * @param String lang
 * @param String type
 * @param String level
 * @return Array
 */
function getUsingCriteria(lang, type, level) {
  var usingCriteria = getAllCriteria(lang, type);
  
  // Trusted Tester does not apply additional criteria
  if (type.indexOf('tt') >= 0) return usingCriteria;
  if (getUsingCriteria.vals) return getUsingCriteria.vals;
 
  // additional criteria
  var additionalCriteriaArr = getAdditionalCriterion().split(/,/);
  var additionalCriteria = [];
  for (var i = 0; i < additionalCriteriaArr.length; i++) {
    additionalCriteria.push(additionalCriteriaArr[i].trim());
  }
  
  // eliminate unuse criteria
  for (var i = 0; i < usingCriteria.length; i++) {
    if (
      (type == 'wcag20' && criteria21.indexOf(usingCriteria[i][1]) >= 0) ||
      usingCriteria[i][0].length > level.length
    ) {
      if (additionalCriteria.indexOf(usingCriteria[i][1]) >= 0) continue;
      delete usingCriteria[i];
    }
  }
  
  usingCriteria = usingCriteria.filter(function(x){
	return !(x === null || x === undefined || x === ""); 
  });
  
  getUsingCriteria.vals = usingCriteria;
  
  return usingCriteria;
}

/**
 * Get Using Tech Set
 * @param String lang
 * @param String type
 * @param String level
 * @return Array
 */
function getUsingTechs(lang, type, level) {
  if (getUsingTechs.vals) return getUsingTechs.vals;

  var techNames = getLangSet('tech');
  var urlPointer = lang+'-'+type;
  var usingCriteria = getUsingCriteria(lang, type, level);
  var usingTechs = [];
  
  for (i = 0; i < usingCriteria.length; i++) {
    var criteria = usingCriteria[i][1];
    if (relTechsAndCriteria[criteria] == null) continue;
    for (j = 0; j < relTechsAndCriteria[criteria].length; j++) {
      var url = urlbase['tech'][urlPointer];
      var each = relTechsAndCriteria[criteria][j];
      
      // Techniques for WCAG 2.1 has directory
      if (type == 'wcag21' && lang == 'en') {
        var dir = each.charAt(0)+each.charAt(1);
        if (['M', 'L', 'V', 'C'].indefOf(each.charAt(1)) < 0) {
          dir = dir.charAt(0);
        }
        url += techDirAbbr[dir]+'/'+each;
      } else {
        url += each+'.html';
      }

      usingTechs.push([criteria, relTechsAndCriteria[criteria][j], techNames[each], url]);
    }
  }
  
  getUsingTechs.vals = usingTechs;
    
  return usingTechs;
}

/**
 * add sheet
 * @param String sheetname
 * @param String template
 * @return Bool
 */
function addSheet(sheetname, template) {
  var ss = getSpreadSheet();
  // Microsoft Excel compatible
  // Excel's sheetname cannot use : and /
  sheetname = String(sheetname).replace(/https*:\/\//ig, '');
  sheetname = String(sheetname).replace(/\//ig, ' ');
 
  // Excel's sheetname must be under 31 chars
  if (sheetname.length > 28) {
    var tmpbase = sheetname.substr(0, 28);
    var tmp = tmpbase;
    var i = 2;
    while(ss.getSheetByName(tmp)) {
      var tmp = tmpbase+'-'+i;
      i++;
    }
    sheetname = tmp;
  }
  
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
 * Get HTML and its title
 * @param String url
 * @return Object
 */
function getHtmlAndTitle(url) {
  var options = {
    "muteHttpExceptions" : true,
    "validateHttpsCertificates" : false,
    "followRedirects" : false,
  }
  
  try {
    var res = UrlFetchApp.fetch(url, options).getContentText();
    var title = res.match(/<title>.+?<\/title>/ig);
    title = String(title).replace(/<\/*title>/ig, '');
    return {'title': title, 'html': res};
  } catch(e) {
    return {'title': '', 'html': ''};
  }
}

/**
 * wrapHtmlHeaderAndFooter
 * @param String title
 * @param String body
 * @return String
 */
function wrapHtmlHeaderAndFooter(title, body) {
  return '<!DOCTYPE html><html lang="'
  +getProp('lang')
  +'"><head><meta charset="utf-8"><title>'
  +title
  +'</title></head><body>'
  +body
  +'</body></html>';
}

/**
 * escape html
 * @thx https://qiita.com/saekis/items/c2b41cd8940923863791
 * @return Void
 */
function escapeHtml (string) {
  if (typeof string !== 'string') {
    return string;
  }
  return string.replace(/[&'`"<>]/g, function(match) {
    return {
      '&': '&amp;',
      "'": '&#x27;',
      '`': '&#x60;',
      '"': '&quot;',
      '<': '&lt;',
      '>': '&gt;',
    }[match]
  });
}
