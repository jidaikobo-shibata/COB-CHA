/**
 * Issue Report for COB-CHA
 */

/**
 * is Edit or Add
 * @return Bool
 */
function isEditIssue() {
  // if not exist issue sheet then this time is not for edit
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);
  if (issueSheet == null) return false;

  // if current sheet is not issue sheet then this time is not for edit
  var activeSheet = getActiveSheet();
  if (activeSheet.getName() != issueSheetName) return false;

  // if current row has id then this time is for edit
  var activeRow = issueSheet.getActiveCell().getRow();
  var issueId = issueSheet.getRange(activeRow, 1).getValue();
  
  return (String(issueId).length > 0 && activeRow > 2);
}

/**
 * set dialog Value Issue
 * @param Bool isEdit
 * @return Array
 */
function dialogValueIssue(isEdit) {
  var ret = {};
  ret['vals'] = {
    'issueId': 0,
    'issueName': '',
    'issueVisibility': '',
    'errorNotice': '',
    'html': '',
    'explanation': '',
    'testId': '',
    'checked': '',
    'techs': '',
    'places': '',
    'image': '',
    'preview': '',
    'memo': ''
  };

  ret['lang'] = getProp('lang');
  ret['type'] = getProp('type');
  ret['level'] = getProp('level');

  if (ret['type'] != 'tt20') {
    delete ret['vals'].testId;
  }
  
  if (isEdit) {
    var cellPlace = {
      'issueId': 1,
      'issueName': 2,
      'issueVisibility': 3,
      'errorNotice': 4,
      'html': 5,
      'explanation': 6,
      'testId': 7,
      'checked': 8,
      'techs': 9,
      'places': 10,
      'image': 11,
      'preview': 12,
      'memo': 13
    };
    
    // issue sheet must be existed and activated
    var ss = getSpreadSheet();
    var issueSheet = ss.getSheetByName(issueSheetName);
    var activeRow = issueSheet.getActiveCell().getRow();
    for (var key in ret['vals']) {
      var i = cellPlace[key];
      if (ret['type'] != 'tt20' && key == 'testId') continue;
      i = ret['type'] != 'tt20' && i >= 7 ? i - 1 : i;
      var val = issueSheet.getRange(activeRow, i).getValue();
      if (val) {
        ret['vals'][key] = issueSheet.getRange(activeRow, i).getValue();
      } else {
        ret['vals'][key] = issueSheet.getRange(activeRow, i).getFormula();
      }
      i++;
    }
    ret['vals']['preview'] = removeImageFormula(ret['vals']['preview']);
  } else {
    ret['vals']['places'] = getUrlFromSheet(getActiveSheet());
  }
 
  ret['usingCriteria'] = getUsingCriteria(ret['lang'], ret['type'], ret['level']);
  ret['usingTechs'] = getUsingTechs(ret['lang'], ret['type'], ret['level']);
  
  ret['places'] = [];
  var all = getAllSheets();
  for (i = 0; i < all.length; i++) {
    ret['places'].push(getUrlFromSheet(all[i]));
  }
 
  return ret;
}

/**
 * open dialog Issue
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function openDialogIssue(lang, testType, level) {
  // to tell current page
  var activeSheet = getActiveSheet();
  var html = '';
  if (activeSheet.getName().charAt(0) != '*') {
    html = '<input type="hidden" id="target-url" value="'+activeSheet.getName()+'">';
  }
  
  // generate Issue sheet
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);
  if (issueSheet == null) {
    addSheet(issueSheetName);
    var issueSheet = ss.getActiveSheet();
    issueSheet.getRange("2:2").setBackground(labelColor).setHorizontalAlignment('center');
    issueSheet.setFrozenRows(2);
    
    setBasicValue(issueSheet, lang, testType, level);
    issueSheet.getRange(2,  1).setValue('ID');
    issueSheet.getRange(2,  2).setValue(getUiLang('name', 'Name'));
    issueSheet.getRange(2,  3).setValue(getUiLang('issue-visibility', 'Issue Visibility'));
    issueSheet.getRange(2,  4).setValue('Error/Notice');
    issueSheet.getRange(2,  5).setValue('HTML');
    issueSheet.getRange(2,  6).setValue(getUiLang('explanation', 'Explanation'));
    var col = 7;
    if (testType == 'tt20') {
      issueSheet.getRange(2, col).setValue(getUiLang('test-id', 'Test ID')); col++;
    }
    issueSheet.getRange(2, col).setValue(getUiLang('criterion', 'Criteria')); col++;
    issueSheet.getRange(2, col).setValue(getUiLang('tech', 'Techniques')); col++;
    issueSheet.getRange(2, col).setValue(getUiLang('places', 'Places')); col++;
    issueSheet.getRange(2, col).setValue(getUiLang('image', 'Image')); col++;
    issueSheet.getRange(2, col).setValue(getUiLang('preview', 'Preview')); col++;
    issueSheet.getRange(2, col).setValue(getUiLang('memo', 'Memo'));
  };

  var title = isEditIssue() ? getUiLang('edit-issue', 'Edit issue') : getUiLang('add-new-issue', 'Add new issue');
  showDialog('ui-issue', 500, 400, title, html);
}

/**
 * apply Issue
 * @param Array vals
 * @return String
 */
function applyIssue(vals) {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);
  var testType = getProp('type');
  if (testType != 'tt20') {
    vals.splice(6, 1);
  }
  
  // issue id - edit
  if (vals[0] > 0) {
    var targetRow = issueSheet.getActiveCell().getRow();
    issueSheet.getRange(targetRow, 1).setValue(vals[0]);
  } else {
    var targetRow = issueSheet.getLastRow() + 1;
    issueSheet.getRange(targetRow, 1).setValue(targetRow - 2);
  }

  for (i = 1; i < vals.length; i++) {
    issueSheet.getRange(targetRow, i + 1).setValue(vals[i]);
  }
  var preview = issueSheet.getRange(targetRow, 11).getValue();
  if (preview) {
    issueSheet.getRange(targetRow, 11).setValue('=IMAGE("https://drive.google.com/uc?export=download&id='+preview+'",1)')
  }
    
  if (vals[0] > 0) {
    return getUiLang('edit-done', 'Issue Edited');
  }
  return getUiLang('add-done', 'Issue Added');
}

/**
 * set Issue list
 * @return Object
 */
function setIssueList() {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName();

  // target URL
  if (activeSheetName == resultSheetName) {
    var targetRow = activeSheet.getActiveCell().getRow();
    var url = activeSheet.getRange(targetRow, 1).getValue();
  } else {
    if (activeSheetName.charAt(0) == '*') {
      return {'url': '', 'issues': []};
    } else {
      var url = activeSheet.getRange(2, 2).getValue();
    }
  }

  var dataObj = issueSheet.getDataRange().getValues();
  var issues = [];
  for (var i = 2; i < dataObj.length; i++) {
    var urls = dataObj[i][8].split(',');
    for (var j = 0; j < urls.length; j++) {
      var issueurl = urls[j].trim();
      if (issueurl != url) continue;
      issues.push(dataObj[i]);
    }
  }
  return {'url': url, 'issues': issues};
}

/**
 * show each issue
 * @param Integer row
 * @return Void
 */
function showEachIssue(row) {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);
  issueSheet.getRange(row, 1).activate();
  showDialog('ui-issue', 500, 400);
}

/**
 * export Issue
 * @return String
 */
function exportIssue() {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);
  var dataObj = issueSheet.getDataRange().getValues();
  var date = new Date();
  var filename = 'issue-report-'+Utilities.formatDate( date, 'Asia/Tokyo', 'yyyy-MM-dd-hh-mm')+'.html';

  var lang = getProp('lang');
  var common = [];
  var each = {};
  var commonEach = {};
  var testType = getProp('type');
  var place_col = testType == 'tt20' ? 9 : 8;
  var targetPlaces = [];

  // prepare vals
  for (var i = 2; i < dataObj.length; i++) {
    // issueVisibility
    if (dataObj[i][2] == 'off') continue;

    // if target URL is not exists continue
    var places = dataObj[i][place_col];
    targetPlaces = places.split(',');
    if (targetPlaces.length == 0) continue;

    if (testType == 'tt20') {
      var vals = {
        'issueId':     dataObj[i][0],
        'name':        dataObj[i][1],
        'errorNotice': dataObj[i][3],
        'html':        escapeHtml(dataObj[i][4]),
        'explanation': dataObj[i][5],
        'testId':      dataObj[i][6],
        'criteria':    dataObj[i][7],
        'techs':       dataObj[i][8],
        'places':      targetPlaces,
        'image':       dataObj[i][10]
      }
    } else {
      var vals = {
        'issueId':     dataObj[i][0],
        'name':        dataObj[i][1],
        'errorNotice': dataObj[i][3],
        'html':        escapeHtml(dataObj[i][4]),
        'explanation': dataObj[i][5],
        'criteria':    dataObj[i][6],
        'techs':       dataObj[i][7],
        'places':      targetPlaces,
        'image':       dataObj[i][9]
      }
    }

    if (targetPlaces.length > 1) {
      common.push(vals);
      for (var j = 0; j < targetPlaces.length; j++) {
        var url = targetPlaces[j].trim();
        if (commonEach[url] == null) {
          commonEach[url] = [];
        }
        commonEach[url].push(vals);
      }
    } else {
      var url = targetPlaces[0].trim();
      if (each[url] == null) {
        each[url] = [];
      }
      each[url].push(vals);
    }
  }
  
  // generate html
  var str = '';
  str += '<h1>'+getUiLang('issue-report', 'Issue Report')+'</h1>';
  str += '<h2>'+getUiLang('common-issue', 'Common Issue')+'</h2>';
  if (common.length == 0) {
    str += '<p>'+getUiLang('no-common-issue-was-reported', 'No Common Issue was reported.')+'</p>';
  } else {
    str = generateIssueReportHtml(str, common, lang);
  }

  var allSheets = getAllSheets();
  for (var i = 0; i < allSheets.length; i++) {
    var activeSheet = allSheets[i];
    var url = activeSheet.getRange(2, 2).getValue();
    var title = activeSheet.getRange(3, 2).getValue();
    var screenshot = activeSheet.getRange(2, 6).getValue();

    str += '<h2>'+title+'<br>'+url+'</h2>';
    if (screenshot != '') {
      str += '<div class="screenshot page"><img src="'+screenshot+'" alt="'+getUiLang('screenshot', 'screenshot')+'"></div>';
    }
    
    if (each[url] == null) {
      str += '<p>'+getUiLang('no-particular-issue-was-reported-on-this-page', 'No particular issue was reported on this page.')+'</p>';
    } else {
      str = generateIssueReportHtml(str, each[url], lang);
    }

    if (commonEach[url] == null) {
      str += '<p>'+getUiLang('no-common-issue-was-reported-on-this-page', 'No common issue was reported on this page.')+'</p>';
    } else {
      var lis = [];
      for (var j = 0; j < commonEach[url].length; j++) {
        lis.push('<li>'+commonEach[url][j]['issueId']+': '+commonEach[url][j]['name']+'</li>');
      }
      str += '<h3>'+getUiLang('common-issue-on-this-page', 'common issue on this page.')+'</h3>';
      str += '<ul>'+lis.join('')+'</ul>';
    }
  }

  saveHtml(exportFolderName, filename, wrapHtmlHeaderAndFooter('Issue Report', str));
  
  return getUiLang('issue-exported', "Issue Exported");
}

/**
 * generate issue report html
 * @param String str
 * @param Array vals
 * @param String lang
 * @return String
 */
function generateIssueReportHtml(str, vals, lang) {
  var ss = getSpreadSheet();
  var allSheets = getAllSheets();
  var criteriaVals = getLangSet('criteria');
  var techVals = getLangSet('tech');
  var testType = getProp('type');
  
  for (var i = 0; i < vals.length; i++) {
    str += '<h3>'+vals[i]['issueId']+': '+vals[i]['name']+'</h3>';
    str += '<table>';

    str += '<tr><th>';
    str += getUiLang('priority', 'Priority');
    str += '</th><td>'+vals[i]['errorNotice']+'</td></tr>';

    str += '<tr><th>';
    str += lang == 'ja' ? 'HTML' : 'Explanation';
    str += '</th><td>'+vals[i]['html']+'</td></tr>';

    str += '<tr><th>';
    str += getUiLang('explanation', 'Explanation');
    str += '</th><td>'+vals[i]['explanation']+'</td></tr>';

    if (testType == 'tt20') {
      str += '<tr><th>';
      str += getUiLang('test-id', 'Test ID');
      var tmp = [];
      var n = 0;
      var targetTestIds = vals[i]['testId'].split(',');
      for (var j = 0; j < relTtAndCriteria.length; j++) {
        for (var k = 0; k < targetTestIds.length; k++) {
          var cTestId = targetTestIds[k].trim();
          if (targetTestIds[j][1] != cTestId) continue;
          tmp[n] = '<li>'+targetTestIds[j][1]+'</li>';
        }
        n++;
      }
      str += '</th><td><ul>'+tmp.join()+'</ul></td></tr>';
    }

    str += '<tr><th>';
    str += getUiLang('criterion', 'Criteria');
    var tmp = [];
    var n = 0;
    var targetCriteria = vals[i]['criteria'].split(',');
    for (var j = 0; j < criteriaVals.length; j++) {
      for (var k = 0; k < targetCriteria.length; k++) {
        var cCriteria = targetCriteria[k].trim();
        if (criteriaVals[j][1] != cCriteria) continue;
        tmp[n] = '<li>'+criteriaVals[j][0]+': '+criteriaVals[j][2]+' ('+criteriaVals[j][0]+')</li>';
      }
      n++;
    }
    str += '</th><td><ul>'+tmp.join()+'</ul></td></tr>';
    
    str += '<tr><th>';
    str += getUiLang('tech', 'Techniques');
    var tmp = [];
    var targetTechs = vals[i]['techs'].split(',');
    for (var j = 0; j < targetTechs.length; j++) {
      var cTech = targetTechs[j].trim();
      if (techVals[cTech] == null) continue;
      tmp.push('<li>'+cTech+': '+techVals[cTech]+'</li>');
    }
    str += '</th><td><ul>'+tmp.join('')+'</ul></td></tr>';
    
    str += '<tr><th>';
    str += 'URL';
    var tmp = [];
    if (vals[i]['places'].length == allSheets.length) {
      str += '</th><td>'+getUiLang('all-page', 'All Page')+'</td></tr>';
    } else {
      for (var j = 0; j < vals[i]['places'].length; j++) {
        var url = vals[i]['places'][j].trim();
        var eachTargetSheet = getSheetByUrl(url);
        if (eachTargetSheet == null) continue;
        var pageTitle = eachTargetSheet.getRange(3, 2).getValue();
        tmp.push('<li><a href="'+url+'">'+pageTitle+'</a></li>');
      }
      str += '</th><td><ul>'+tmp.join('')+'</ul></td></tr>';
    }

    str += '</table>';
  }
  return str;
}

/**
 * upload Issue image
 * @param Object formObj
 * @return Void
 */
function uploadIssueImage(formObj) {
  return fileUpload(imagesFolderName, formObj, "imageFile");
}