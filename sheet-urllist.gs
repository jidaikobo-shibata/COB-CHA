/**
 * Config for COB-CHA
 * functions:
 * - generateUrlListSheet
 */

/**
 * Generate URL List
 * @return String
 */
function generateUrlListSheet() {
  var defaults = [
    ["No.", "URL"],
  ];
  for (var i = 1; i <= 40; i++) {
    defaults.push([i, ""]);
  }
  generateSheetIfNotExists(gUrlListSheetName, defaults, "row");
  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gUrlListSheetName);
}
