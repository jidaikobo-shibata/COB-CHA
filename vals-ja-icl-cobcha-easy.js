/**
 * Easy adapter for the COB-CHA Implementation Check List.
 *
 * The complete COB-CHA ICL is the source of truth. Checks marked with
 * {easy: true} are selected for this short level-A screening checklist.
 * LICENSE: [CC0](https://creativecommons.org/share-your-work/public-domain/cc0)
 */

/**
 * Return checks selected for the easy ICL.
 *
 * @return Object
 */
function getIclTestCobchaEasy() {
  var sourceTests = getIclTestCobcha();
  var tests = {};

  for (var situationId in sourceTests) {
    if (/-Other$/.test(situationId)) continue;

    var selected = [];
    for (var i = 0; i < sourceTests[situationId].length; i++) {
      var sourceItem = sourceTests[situationId][i];
      var options = sourceItem[2];
      if (!options || typeof options !== 'object' || options.easy !== true) {
        continue;
      }

      var item = sourceItem.slice();
      var easyOptions = {};
      for (var optionName in options) {
        if (optionName !== 'easy' && optionName !== 'pdf' &&
            optionName !== 'pdfOnly') {
          easyOptions[optionName] = options[optionName];
        }
      }

      if (easyOptions.criteria) {
        var criteria = easyOptions.criteria.split(',').filter(function (criterion) {
          return gSingleACriteria.indexOf(criterion) !== -1;
        });
        if (criteria.length > 0) {
          easyOptions.criteria = criteria.join(',');
        } else {
          delete easyOptions.criteria;
        }
      }

      if (Object.keys(easyOptions).length > 0) {
        item[2] = easyOptions;
      } else {
        item.splice(2, 1);
      }
      selected.push(item);
    }

    if (selected.length > 0) tests[situationId] = selected;
  }

  return tests;
}

/**
 * Return situations containing at least one easy ICL check.
 *
 * @return Object
 */
function getIclSituationCobchaEasy() {
  var sourceSituations = getIclSituationCobcha();
  var tests = getIclTestCobchaEasy();
  var situations = {};

  for (var criterion in sourceSituations) {
    if (criterion !== 'UI-Common' &&
        gSingleACriteria.indexOf(criterion) === -1) continue;

    var selected = {};
    for (var situationId in sourceSituations[criterion]) {
      if (tests[situationId]) {
        selected[situationId] = sourceSituations[criterion][situationId];
      }
    }
    if (Object.keys(selected).length > 0) situations[criterion] = selected;
  }

  return situations;
}
