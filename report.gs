/**
 * Report for COB-CHA
 */

/**
 * show Issue dialog
 */
function showIssueDialog() {
  var output = HtmlService.createTemplateFromFile('Issue');
  var ss = getSpreadSheet();
  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(500).setHeight(500);
  ss.show(html);
}

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
  
  return(String(issueId).length > 0);
}

/**
 * set issue value
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
    'memo': ''
  };

  if (isEdit) {
    // issue sheet must be existed and activated
    var ss = getSpreadSheet();
    var issueSheet = ss.getSheetByName(issueSheetName);
    var activeRow = issueSheet.getActiveCell().getRow();
    var i = 1;
    for (var key in ret['vals']) {
      ret['vals'][key] = issueSheet.getRange(activeRow, i).getValue();
      i++;
    }
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
    
    issueSheet.getRange(1, 1).setValue('Type').setBackground(labelColor);
    issueSheet.getRange(1, 2).setValue(lang).setHorizontalAlignment('center');
    issueSheet.getRange(1, 3).setValue(testType).setHorizontalAlignment('center');
    issueSheet.getRange(1, 4).setValue(level).setHorizontalAlignment('center');
    issueSheet.getRange(2,  1).setValue('ID');
    issueSheet.getRange(2,  2).setValue('Name');
    issueSheet.getRange(2,  3).setValue('Issue Visibility');
    issueSheet.getRange(2,  5).setValue('Error/Notice');
    issueSheet.getRange(2,  6).setValue('HTML');
    issueSheet.getRange(2,  7).setValue('Explanation');
    issueSheet.getRange(2,  8).setValue('Criteria');
    issueSheet.getRange(2,  9).setValue('Techniques');
    issueSheet.getRange(2, 10).setValue('Places');
    issueSheet.getRange(2, 11).setValue('Memo');
    issueSheet.getRange(2, 12).setValue('Create Date');
    issueSheet.getRange(2, 13).setValue('Update Date');
  };

  var today = new Date();

  showIssueDialog();
}

/**
 * apply Issue
 * @param Array vals
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
    
  if (vals[0] > 0) {
    return("Issue Edited");
  } else {
    return("Issue Added");
  }
}

/**
 * show Issue
 */
function showIssue() {
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
      return ("could not specify URL");
    } else {
      var url = activeSheet.getRange(2, 2).getValue();
    }
  }

  var dataObj = issueSheet.getDataRange().getValues();
  for (var i = 2; i < dataObj.length; i++) {
    Logger.log(dataObj[i]);
    
    /*
    ここで、URLを含んだものをさがして、エラー名称を返す
    */
    
  }
}

/**
 * escape html
 * thx https://qiita.com/saekis/items/c2b41cd8940923863791
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
 */
function exportIssue() {
  var ss = getSpreadSheet();
  var issueSheet = ss.getSheetByName(issueSheetName);
  var dataObj = issueSheet.getDataRange().getValues();
//  var vals = {};
  
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
  var criteriaVals = getLangSet('criteria');
  var techVals = getLangSet('tech');

  var str = '';
  for (var i = 2; i < dataObj.length; i++) {
    var issueId         = dataObj[i][0];
    var name            = dataObj[i][1];
    var issueVisibility = dataObj[i][2];
    var errorNotice     = dataObj[i][3];
    var html            = escapeHtml(dataObj[i][4]);
    var explanation     = dataObj[i][5];
    var criteria        = dataObj[i][6];
    var techs           = dataObj[i][7];
    var places          = dataObj[i][8];
//    var memo            = dataObj[i][9];
    
    if (issueVisibility == 'off') continue;

    str += '<h2>'+issueId+': '+name+'</h2>';
    str += '<table>';

    str += '<tr><th>';
    str += lang == 'ja' ? '重要度' : 'Priority';
    str += '</th><td>'+errorNotice+'</td></tr>';

    str += '<tr><th>';
    str += lang == 'ja' ? 'HTML' : 'Explanation';
    str += '</th><td>'+html+'</td></tr>';

    str += '<tr><th>';
    str += lang == 'ja' ? '解説' : 'Explanation';
    str += '</th><td>'+explanation+'</td></tr>';
    
    str += '<tr><th>';
    str += lang == 'ja' ? '関連する達成基準' : 'Criteria';
    var tmp = [];
    var n = 0;
    var targetCriteria = criteria.split(',');
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
    str += lang == 'ja' ? '関連する達成方法' : 'Techniques';
    var tmp = [];
    var targetTechs = techs.split(',');
    for (var j = 0; j < targetTechs.length; j++) {
      var cTech = targetTechs[j].trim();
      if (techVals[cTech] == null) continue;
      tmp.push('<li>'+cTech+': '+techVals[cTech]+'</li>');
    }
    str += '</th><td><ul>'+tmp.join('')+'</ul></td></tr>';
    
    str += '<tr><th>';
    str += lang == 'ja' ? '問題が存在するページ' : 'URL';
    var tmp = [];
    var targetPlaces = places.split(',');
    for (var j = 0; j < targetPlaces.length; j++) {
      var url = targetPlaces[j].trim();
      var eachTargetSheet = ss.getSheetByName(url);
      if (eachTargetSheet == null) continue;
      var pageTitle = eachTargetSheet.getRange(3, 2).getValue();
      tmp.push('<li><a href="'+url+'">'+pageTitle+'</a></li>');
    }
    str += '</th><td><ul>'+tmp.join('')+'</ul></td></tr>';

    str += '</table>';
  }
  
  saveHtml(exportFolderName, filename, str);
  
  return("Issue Exported");
}

/**
 * export Html
 * @param String testType
 * @param String level
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
  
  return(n+" file(s) exported");
}