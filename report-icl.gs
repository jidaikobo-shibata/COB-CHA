/**
 * Report of Implementation Check List(ICL) for COB-CHA
 */

/**
 * export ICL Html
 * @return Void
 */
function exportIclHtml() {
  var allSheets = getAllSheets();
  var techs = getLangSet('tech');

  // ここから手付かず
  
  var str = '';
  str += '<table>';
  str += '<tr><th></th>';
  str += '<td>'+totalLevel+'</td>';
  str += '</tr>';
  str += '</table>';

  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().charAt(0) == '*') continue;
    var dataObj = allSheets[i].getDataRange().getValues();

  }

  return 'done2';
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
 
  saveHtml(exportFolderName, 'index.html', wrapHtmlHeaderAndFooter('Report Index', str), true);
  
  return getUiLang('file-exported', "%s file(s) exported").replace('%s', 1);
}