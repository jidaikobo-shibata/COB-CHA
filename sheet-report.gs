/**
 * Report Index for COB-CHA
 * finctions:
 * - generateReportSheet
 */

/**
 * generate Report Sheet
 * @param String level
 * @return String
 */
function generateReportSheet(level) {
  var defaults = [
    [getUiLang("report-declaration-day", "Declaration day"), ""],
    [getUiLang("report-standard-version", "Standard's version"), ""],
    [getUiLang("report-target-level", "Target level"), level],
    [getUiLang("report-gained-level", "Gained level"), ""],
    [getUiLang("report-explanation-pages", "Explanation of pages"), ""],
    [getUiLang("report-way-to-choose", "Way to choose"), ""],
    [getUiLang("report-depending-tech", "Technology in depend"), ""],
//    [getUiLang("report-urls-pages", "Pages' urls"), ""],
    [getUiLang("report-test-days", "Test date"), ""]
  ];
  generateSheetIfNotExists(gReportSheetName, defaults);
  return getUiLang('target-sheet-generated', "Generate Target Sheet (%s).").replace('%s', gReportSheetName);
}
