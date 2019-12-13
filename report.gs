/**
 * Report for COB-CHA
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
  
  return (String(issueId).length > 0);
}

/**
 * set issue value
 * @param Bool isEdit
 * @return Array
 */
function setIssueValue(isEdit) {
  var ret = {};
  ret['vals'] = {
    'issueId': 0,
    'issueName': '',
    'issueVisibility': '',
    'errorNotice': '',
    'html': '',
    'explanation': '',
    'checked': '',
    'techs': '',
    'places': '',
    'image': '',
    'preview': '',
    'memo': ''
  };

  if (isEdit) {
    // issue sheet must be existed and activated
    var ss = getSpreadSheet();
    var issueSheet = ss.getSheetByName(issueSheetName);
    var activeRow = issueSheet.getActiveCell().getRow();
    var i = 1;
    for (var key in ret['vals']) {
      var val = issueSheet.getRange(activeRow, i).getValue();
      if (val) {
        ret['vals'][key] = issueSheet.getRange(activeRow, i).getValue();
      } else {
        ret['vals'][key] = issueSheet.getRange(activeRow, i).getFormula();
      }
      i++;
    }
    ret['vals']['preview'] = removeImageFormula(ret['vals']['preview']);
  }

  // to keep array order
  var criteria = getLangSet('criteria');
  var tmp = relTechsAndCriteria;
  ret['techs'] = [];
  for (i = 0; i < criteria.length; i++) {
    if (relTechsAndCriteria[criteria[i][1]] == null) continue;
    ret['techs'].push([criteria[i][1], relTechsAndCriteria[criteria[i][1]]]);
  }
  
  ret['lang'] = getProp('lang');
  ret['type'] = getProp('type');
  ret['level'] = getProp('level');
  
  ret['techDirAbbr'] = techDirAbbr;
  ret['criteria21'] = criteria21;
  ret['criteria'] = getLangSet('criteria');
  ret['techNames'] = getLangSet('tech');

  ret['places'] = [];
  var all = getAllSheets();
  for (i = 0; i < all.length; i++) {
    ret['places'].push(all[i].getName());
  }
 
  ret['urlbase'] = urlbase;
  ret['docurl'] = ret['lang']+'-'+ret['type'];
  ret['docurlEn'] = 'en'+'-'+ret['type'];
 
  return ret;
}

/**
 * add Issue
 * @param String lang
 * @param String testType
 * @param String level
 * @return Void
 */
function addIssue(lang, testType, level) {
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
    issueSheet.getRange(2,  5).setValue('Error/Notice');
    issueSheet.getRange(2,  6).setValue('HTML');
    issueSheet.getRange(2,  7).setValue(getUiLang('explanation', 'Explanation'));
    issueSheet.getRange(2,  8).setValue(getUiLang('criterion', 'Criteria'));
    issueSheet.getRange(2,  9).setValue(getUiLang('tech', 'Techniques'));
    issueSheet.getRange(2, 10).setValue(getUiLang('places', 'Places'));
    issueSheet.getRange(2, 11).setValue(getUiLang('image', 'Image'));
    issueSheet.getRange(2, 12).setValue(getUiLang('preview', 'Preview'));
    issueSheet.getRange(2, 13).setValue(getUiLang('memo', 'Memo'));
  };

  var title = isEditIssue() ? getUiLang('edit-issue', 'Edit issue') : getUiLang('add-new-issue', 'Add new issue');
  showDialog('issue', 500, 400, title);
}

/**
 * apply Issue
 * @param Array vals
 * @return String
 */
function applyIssue(vals) {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);

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
  showDialog('issue', 500, 400);
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

  /*
  // Google Document - too heavey to use...
  var document = DocumentApp.create(filename);
  document.getBody().setText('issue-report-text');
  var targetFolder = getTargetFolder(exportFolderName);
  var docFile = DriveApp.getFileById(document.getId());
  targetFolder.addFile(docFile);
  var parentFolder = DriveApp.getFileById(docFile.getId()).getParents();
  parentFolder.next().removeFile(docFile);
  */

  var lang = getProp('lang');
  var common = [];
  var each = {};
  var commonEach = {};
  
  // prepare vals
  for (var i = 2; i < dataObj.length; i++) {
    // issueVisibility
    if (dataObj[i][2] == 'off') continue;

    var places = dataObj[i][8];
    var targetPlaces = places.split(',');
    if (targetPlaces.length == 0) continue;
    var vals = {
      'issueId':         dataObj[i][0],
      'name':            dataObj[i][1],
      'errorNotice':     dataObj[i][3],
      'html':            escapeHtml(dataObj[i][4]),
      'explanation':     dataObj[i][5],
      'criteria':        dataObj[i][6],
      'techs':           dataObj[i][7],
      'places':          targetPlaces,
      'image':           dataObj[i][9]
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
    str += '<div class="screenshot"><img src="'+screenshot+'" alt="'+getUiLang('screenshot', 'screenshot')+'"></div>';
    
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

  saveHtml(exportFolderName, filename, str);
  
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
    
    str += '<tr><th>';
    str += getUiLang('Criterion', 'Criteria');
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
        var eachTargetSheet = ss.getSheetByName(url);
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
 * export Html
 * @param String testType
 * @param String level
 * @return String
 */
function exportHtml(testType, level) {
  var ss = getSpreadSheet();
  var resultSheet = ss.getSheetByName(resultSheetName);
  if (resultSheet == null) return 'result page not found. Evalute fisrt';
  var dataObj = resultSheet.getDataRange().getValues();
    
  // evaluate total
  var levels = {'-': -2, 'NI': -1, 'A-': 1, 'A': 2, 'AA-': 3, 'AA': 4, 'AAA-': 5, 'AAA': 6}
  var levelsR = {'-2':'-' , '-1': 'NI', '1': 'A-', '2': 'A', '3': 'AA-', '4': 'AA', '5': 'AAA-', '6': 'AAA'}
  var vals = {'Yet': -2 , 'F': -1, 'DNA': 1, 'T': 2}

  var currentLevel = 6;
  var currentVals = dataObj[2]; // first row
  currentVals.shift();
  currentVals.shift();
  for (var i = 2; i < dataObj.length; i++) {
    var eachlevel = dataObj[i][1];
    currentLevel = currentLevel >= levels[eachlevel] ? levels[eachlevel] : currentLevel;
    
    var n = 0;
    for (var j = 2; j < dataObj[i].length; j++) {
      var targetVal = dataObj[i][j] == '' ? 'Yet' : dataObj[i][j];

      // update
      if (vals[currentVals[n]] > vals[targetVal]) {
        currentVals[n] = vals[targetVal];
      }
      n++;
    }
  }
  var totalLevel = levelsR[currentLevel];
  
  // criterion
  var criteriaVals = getLangSet('criteria');

  var str = '';
  str += '<table>';
  str += '<tr><th>サイトで達成している達成レベル</th>';
  str += '<td>'+totalLevel+'</td>';
  str += '</tr>';
  str += '</table>';
  
  str += '<table>';
  var n = 0;
  for (var i = 0; i < criteriaVals.length; i++) {
    if (criteriaVals[i][0].length > level.length) continue;
    if ((testType == 'wcag20' || testType == 'tt20') && criteria21.indexOf(criteriaVals[i][1]) >= 0) continue;
    str += '<tr>';
    str += '<th>'+criteriaVals[i][1]+': '+criteriaVals[i][2]+'</th>';
    str += '<td>'+currentVals[n]+'</td>';
    str += '</tr>';
    n++;
  }
  str += '<table>';
 
  saveHtml(exportFolderName, 'index.html', str, true);
  
  // each page
  var dataObj = resultSheet.getDataRange().getValues();
  
  var n = 1;
  for (var i = 2; i < dataObj.length; i++) {
    var eachlevel = dataObj[i][1];
    var filename = 'report_'+n+'.html';
    var str = '';

    str += '<table>';
    str += '<tr><th>ページで達成している達成レベル</th>';
    str += '<td>'+eachlevel+'</td>';
    str += '</tr>';
    str += '</table>';
  
    str += '<table>';
    var nn = 0;
    for (var j = 0; j < criteriaVals.length; j++) {
      if (criteriaVals[j][0].length > level.length) continue;
      if ((testType == 'wcag20' || testType == 'tt20') && criteria21.indexOf(criteriaVals[j][1]) >= 0) continue;
      str += '<tr>';
      str += '<th>'+criteriaVals[j][1]+': '+criteriaVals[j][2]+'</th>';
      str += '<td>'+dataObj[i][nn]+'</td>';
      str += '</tr>';
      nn++;
    }
    str += '<table>';
 
    saveHtml(exportFolderName, filename, str, true);
    n++;
 }
  
  return getUiLang('file-exported', "%s file(s) exported").replace('%s', n);
}