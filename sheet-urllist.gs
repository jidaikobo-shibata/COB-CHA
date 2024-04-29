/**
 * URL List sheet for COB-CHA
 * functions:
 * - generateUrlListSheet
 */

/**
 * Generate URL List
 * @return String
 */
function generateUrlListSheet() {
  var defaults = [
    [
      "No.",
      "URL",
      "title",
      getUiLang('target-lump-edit', "Target for lump edit (o)"),
      getUiLang('memo', "Memo"),
      getUiLang('video', "Video"),
      getUiLang('time-based', "Time-based"),
      getUiLang('form', "Form"),
      getUiLang('table', "Table"),
      getUiLang('help', "Help"),
      getUiLang('ads', "Ads")
    ]
  ];
  for (var i = 1; i <= 40; i++) {
    defaults.push([i, "",  "",  "",  "",  "",  "",  "",  "",  "",  ""]);
  }
  var msgOrSheetObj = generateSheetIfNotExists(gUrlListSheetName, defaults, "row");
  if (typeof msgOrSheetObj == "string") return msgOrSheetObj;

  msgOrSheetObj.getRange('D:K').setHorizontalAlignment('center');

  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gUrlListSheetName);
}
