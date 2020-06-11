/**
 * Report for COB-CHA
 */

/**
 * export Result
 * @param String lang
 * @param String testType
 * @param String level
 * @return String
 */
function exportResult(lang, testType, level) {
  var ss = getSpreadSheet();
  var resultSheet = ss.getSheetByName(resultSheetName);
  if (resultSheet == null) throw new Error(getUiLang('no-target-page-exists', "No Target Page Exists."));
  var dataObj = resultSheet.getDataRange().getValues();
    
  // evaluate total
  var levels = {'-': -2, 'NI': -1, 'A-': 1, 'A': 2, 'AA-': 3, 'AA': 4, 'AAA-': 5, 'AAA': 6}
  var levelsR = {'-2':'-' , '-1': 'NI', '1': 'A-', '2': 'A', '3': 'AA-', '4': 'AA', '5': 'AAA-', '6': 'AAA'}
  var vals = {'NT': -2 , 'F': -1, 'DNA': 1, 'T': 2}

  var currentLevel = 6;
  var currentVals = dataObj[2]; // first row
  currentVals.shift();
  currentVals.shift();
  for (var i = 2; i < dataObj.length; i++) {
    var eachlevel = dataObj[i][1];
    currentLevel = currentLevel >= levels[eachlevel] ? levels[eachlevel] : currentLevel;
    
    var n = 0;
    for (var j = 2; j < dataObj[i].length; j++) {
      var targetVal = dataObj[i][j] == '' ? 'NT' : dataObj[i][j];

      // update apply lower result
      if (vals[currentVals[n]] > vals[targetVal] && targetVal != 'DNA') {
        currentVals[n] = targetVal;
      }
      n++;
    }
  }
  var totalLevel = levelsR[currentLevel];
  
  // criterion
  var criteriaVals = getUsingCriteria(lang, testType, level);

  var str = '';
  str += '<table>';
  str += '<tr><th>'+getUiLang('report-criterion-site', 'Achievement level achieved on the site')+'</th>';
  str += '<td>'+totalLevel+'</td>';
  str += '</tr>';
  str += '</table>';
  
  str += '<table>';
  var n = 0;
  for (var i = 0; i < criteriaVals.length; i++) {
    str += '<tr>';
    str += '<th>'+criteriaVals[i][1]+': '+criteriaVals[i][2]+'</th>';
    str += '<td>'+currentVals[n]+'</td>';
    str += '</tr>';
    n++;
  }
  str += '<table>';
 
  saveHtml(exportFolderName, 'index.html', wrapHtmlHeaderAndFooter('Report Index', str), true);
  
  // each page
  var dataObj = resultSheet.getDataRange().getValues();
  
  var n = 1;
  for (var i = 2; i < dataObj.length; i++) {
    var eachlevel = dataObj[i][1];
    var filename = 'report_'+n+'.html';
    var str = '';

    str += '<table>';
    str += '<tr><th>'+getUiLang('report-criterion-page', 'Achievement level achieved on the page')+'</th>';
    str += '<td>'+eachlevel+'</td>';
    str += '</tr>';
    str += '</table>';
  
    str += '<table>';
    var nn = 2;
    for (var j = 0; j < criteriaVals.length; j++) {
      str += '<tr>';
      str += '<th>'+criteriaVals[j][1]+': '+criteriaVals[j][2]+'</th>';
      str += '<td>'+dataObj[i][nn]+'</td>';
      str += '</tr>';
      nn++;
    }
    str += '<table>';
 
    saveHtml(exportFolderName, filename, wrapHtmlHeaderAndFooter('Report', str), true);
    n++;
 }
  
  return getUiLang('file-exported', "%s file(s) exported").replace('%s', n);
}