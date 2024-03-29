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
    ["No.", "URL", "title", getUiLang('target-lump-edit', "Target for lump edit (o)")],
  ];
  for (var i = 1; i <= 40; i++) {
    defaults.push([i, "", "", ""]);
  }
  var msgOrSheetObj = generateSheetIfNotExists(gUrlListSheetName, defaults, "row");
  if (typeof msgOrSheetObj == "string") return msgOrSheetObj;
  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gUrlListSheetName);
}
