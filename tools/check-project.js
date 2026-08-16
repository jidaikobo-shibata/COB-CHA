#!/usr/bin/env node

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const projectRoot = path.resolve(__dirname, '..');
const rootFiles = fs.readdirSync(projectRoot, {withFileTypes: true});
const jsFiles = rootFiles
  .filter((entry) => entry.isFile() && entry.name.endsWith('.js'))
  .map((entry) => entry.name)
  .sort();
const htmlFiles = rootFiles
  .filter((entry) => entry.isFile() && entry.name.endsWith('.html'))
  .map((entry) => entry.name)
  .sort();

const errors = [];
const checks = [];

function read(relativePath) {
  return fs.readFileSync(path.join(projectRoot, relativePath), 'utf8');
}

function reportError(check, message) {
  errors.push(`${check}: ${message}`);
}

function runCheck(name, callback) {
  const before = errors.length;

  try {
    callback();
  } catch (error) {
    reportError(name, error.stack || error.message);
  }

  checks.push({name, passed: errors.length === before});
}

function stripHtmlScripts(html) {
  return html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '');
}

function extractInlineScripts(html) {
  const scripts = [];
  const pattern = /<script\b[^>]*>([\s\S]*?)<\/script>/gi;
  let match;

  while ((match = pattern.exec(html)) !== null) {
    scripts.push(match[1].replace(/<\?!=[\s\S]*?\?>/g, 'null'));
  }

  return scripts;
}

runCheck('GAS JavaScript syntax', () => {
  for (const file of jsFiles) {
    try {
      new vm.Script(read(file), {filename: file});
    } catch (error) {
      reportError('GAS JavaScript syntax', `${file}: ${error.message}`);
    }
  }
});

runCheck('No implicit loop globals or condition assignments', () => {
  const implicitLoopPattern = /\bfor\s*\(\s*[A-Za-z_$][\w$]*\s*=/;
  const conditionAssignmentPattern = /\b(?:if|while)\s*\([^)]*[^=!<>]=[^=]/;

  for (const file of jsFiles) {
    const lines = read(file).split(/\r?\n/);

    lines.forEach((line, index) => {
      if (implicitLoopPattern.test(line)) {
        reportError(
          'No implicit loop globals or condition assignments',
          `${file}:${index + 1}: loop counter must be declared`
        );
      }

      if (conditionAssignmentPattern.test(line)) {
        reportError(
          'No implicit loop globals or condition assignments',
          `${file}:${index + 1}: assignment inside condition`
        );
      }
    });
  }
});

runCheck('URL fetch helper behavior', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('COB-CHA.js'), context, {filename: 'COB-CHA.js'});

  const validUrls = [
    'https://example.com/',
    'https://example.com:443/',
    'https://8.8.8.8/',
    'http://example.com/path?key=value',
    '  https://example.com/trimmed  '
  ];
  const invalidUrls = [
    '',
    'javascript:http://example.com',
    'data:text/html,http://example.com',
    'ftp://example.com/',
    'https://example.com/has space',
    'http://localhost/',
    'http://127.0.0.1/',
    'http://10.0.0.1/',
    'http://169.254.169.254/',
    'http://192.168.1.1/',
    'http://2130706433/',
    'http://0177.0.0.1/',
    'http://0x7f.0.0.1/',
    'http://999.1.1.1/',
    'http://[::1]/',
    'http://intranet/',
    'http://user:password@example.com/'
  ];

  validUrls.forEach((url) => {
    if (context.normalizeFetchUrl(url) === '') {
      reportError('URL fetch helper behavior', `valid URL rejected: ${url}`);
    }
  });

  invalidUrls.forEach((url) => {
    if (context.normalizeFetchUrl(url) !== '') {
      reportError('URL fetch helper behavior', `invalid URL accepted: ${url}`);
    }
  });

  const titleCases = [
    ['<title>Example</title>', 'Example'],
    ['<TITLE class="test"> Multi\n line </TITLE>', 'Multi line'],
    ['<html><body>No title</body></html>', '']
  ];

  titleCases.forEach(([html, expected]) => {
    const actual = context.extractHtmlTitle(html);
    if (actual !== expected) {
      reportError(
        'URL fetch helper behavior',
        `expected title "${expected}", received "${actual}"`
      );
    }
  });

  const redirectCases = [
    ['https://example.com/a/page', '/next', 'https://example.com/next'],
    ['https://example.com/a/page', '../next', 'https://example.com/next'],
    ['https://example.com/a/page', '?next=1', 'https://example.com/a/page?next=1'],
    ['https://example.com/a/page', '//other.example/path', 'https://other.example/path'],
    ['https://example.com/a/page', 'http://127.0.0.1/', '']
  ];
  redirectCases.forEach(([base, location, expected]) => {
    const actual = context.resolveFetchRedirectUrl(base, location);
    if (actual !== expected) {
      reportError(
        'URL fetch helper behavior',
        `redirect "${location}": expected "${expected}", received "${actual}"`
      );
    }
  });

  function fakeResponse(status, headers, content) {
    return {
      getResponseCode: () => status,
      getHeaders: () => headers,
      getContentText: () => content
    };
  }

  const responses = [
    fakeResponse(302, {Location: '/final'}, ''),
    fakeResponse(200, {'Content-Type': 'text/html; charset=UTF-8'}, '<title>Final</title>')
  ];
  const requests = [];
  context.UrlFetchApp = {
    fetch(url, options) {
      requests.push({url, options});
      return responses.shift();
    }
  };
  const fetched = context.getHtmlAndTitle('https://example.com/start');
  if (fetched.title !== 'Final' || requests.length !== 2) {
    reportError(
      'URL fetch helper behavior',
      'validated redirect was not fetched successfully'
    );
  }
  if (
    requests[0].options.followRedirects !== false ||
    requests[0].options.timeoutSeconds !== 30
  ) {
    reportError(
      'URL fetch helper behavior',
      'fetch must disable automatic redirects and use a bounded timeout'
    );
  }

  const nonHtml = fakeResponse(200, {'Content-Type': 'application/pdf'}, '%PDF');
  if (context.isHtmlFetchResponse(nonHtml)) {
    reportError('URL fetch helper behavior', 'non-HTML response was accepted');
  }
});

runCheck('Spreadsheet formula injection protection', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('COB-CHA.js'), context, {filename: 'COB-CHA.js'});

  const unsafeValues = ['=1+1', '+1+1', '-1+1', '@SUM(A1:A2)'];
  unsafeValues.forEach((value) => {
    if (context.toPlainTextCellValue(value) !== `'${value}`) {
      reportError(
        'Spreadsheet formula injection protection',
        `unsafe prefix was not escaped: ${value}`
      );
    }
  });

  const safeValues = ['', 'Example', ' https://example.com', 123, false];
  safeValues.forEach((value) => {
    if (context.toPlainTextCellValue(value) !== value) {
      reportError(
        'Spreadsheet formula injection protection',
        `safe value was changed: ${value}`
      );
    }
  });
});

runCheck('Record sheet duplication behavior', () => {
  const context = {};
  const written = {};
  const fakeSheet = {
    getRange(row, col) {
      return {
        setValue(value) {
          written[`${row},${col}`] = value;
        }
      };
    }
  };

  vm.createContext(context);
  vm.runInContext(read('COB-CHA.js'), context, {filename: 'COB-CHA.js'});
  vm.runInContext(read('sheet-record.js'), context, {filename: 'sheet-record.js'});
  context.setRecordSheetPageData(fakeSheet, 'https://example.com/', {
    title: '=unsafe title',
    html: '<html>Example</html>'
  });

  const expected = {
    '1,2': 'https://example.com/',
    '1,6': "'=unsafe title",
    '3,2': '<html>Example</html>'
  };
  Object.entries(expected).forEach(([cell, value]) => {
    if (written[cell] !== value) {
      reportError(
        'Record sheet duplication behavior',
        `cell ${cell}: expected "${value}", received "${written[cell]}"`
      );
    }
  });

  const source = read('sheet-record.js');
  if (!source.includes('addSheet(targetSheetname, templateSheet);')) {
    reportError(
      'Record sheet duplication behavior',
      'new record sheets must be created from the reusable template sheet'
    );
  }
  if (source.includes('lastSheetKey')) {
    reportError(
      'Record sheet duplication behavior',
      'unreachable lastSheetKey loop must not be restored'
    );
  }
});

runCheck('Record criteria batch write behavior', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('sheet-record.js'), context, {filename: 'sheet-record.js'});

  const rows = Array.from(
    context.getRecordCriteriaRows([
      ['A', '1.1.1', 'Criterion one'],
      ['AA', '1.2.1', 'Criterion two']
    ]),
    (row) => Array.from(row)
  );
  const expected = [
    ['1.1.1', '', 'A'],
    ['1.2.1', '', 'AA']
  ];
  if (JSON.stringify(rows) !== JSON.stringify(expected)) {
    reportError(
      'Record criteria batch write behavior',
      `expected ${JSON.stringify(expected)}, received ${JSON.stringify(rows)}`
    );
  }

  const source = read('sheet-record.js');
  if (!source.includes('criteriaRange.setValues(criteriaRows)')) {
    reportError(
      'Record criteria batch write behavior',
      'record criteria must be written with one setValues call'
    );
  }
  if (source.includes('sheet.getRange(row, 1).setValue')) {
    reportError(
      'Record criteria batch write behavior',
      'per-row criterion value writes must not be restored'
    );
  }
});

runCheck('ICL evaluation batch behavior', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('sheet-icl.js'), context, {filename: 'sheet-icl.js'});

  const firstRow = context.findIclFirstResultRow([
    ['URL'],
    ['Date'],
    ['HTML'],
    ['Criterion'],
    [''],
    ['ICL-1']
  ]);
  if (firstRow !== 6) {
    reportError(
      'ICL evaluation batch behavior',
      `expected first ICL result row 6, received ${firstRow}`
    );
  }
  if (context.findIclFirstResultRow([['A'], ['B']]) !== 0) {
    reportError(
      'ICL evaluation batch behavior',
      'missing ICL separator must return 0'
    );
  }

  const rows = Array.from(
    context.transposeIclResultColumns([
      [['o'], ['x'], ['-']],
      [['x'], ['o'], ['']]
    ]),
    (row) => Array.from(row)
  );
  const expected = [['o', 'x'], ['x', 'o'], ['-', '']];
  if (JSON.stringify(rows) !== JSON.stringify(expected)) {
    reportError(
      'ICL evaluation batch behavior',
      `expected ${JSON.stringify(expected)}, received ${JSON.stringify(rows)}`
    );
  }

  const markCases = [
    [[], '-'],
    [['-'], '-'],
    [['-', 'o'], 'o'],
    [['', 'o'], '?'],
    [['?', 'x'], 'x']
  ];
  for (const [values, expectedMark] of markCases) {
    const actualMark = context.summarizeIclMarks(
      values,
      ['?', '-', 'o', 'x']
    );
    if (actualMark !== expectedMark) {
      reportError(
        'ICL evaluation batch behavior',
        `expected ${JSON.stringify(values)} to summarize as ` +
          `${expectedMark}, received ${actualMark}`
      );
    }
  }

  const pageRows = [
    ['ID', 'Check', 'Excluded', 'Memo', 'Criterion', 'Level'],
    ['9.9.8: Shared check', '', '', '', '', ''],
    ['Shared situation', '', '', '', '', ''],
    ['9.9.8-1', 'o', '', 'memo', '9.9.8,9.9.9', 'A']
  ];
  const aggregated = Array.from(
    context.aggregateIclPageResults(
      pageRows,
      ['?', '-', 'o', 'x']
    ),
    (row) => Array.from(row)
  );
  if (aggregated[1][0] !== 'o' || aggregated[2][0] !== '') {
    reportError(
      'ICL evaluation batch behavior',
      'criterion headings must be aggregated without assigning results to situations'
    );
  }

  vm.runInContext(read('input-utility.js'), context, {
    filename: 'input-utility.js'
  });
  const transfer = context.getIclScTransferRows(
    pageRows,
    [['A', '9.9.8'], ['A', '9.9.9']],
    ['?', '-', 'o', 'x']
  );
  const transferredResults = Array.from(
    transfer.results,
    (row) => Array.from(row)
  );
  if (JSON.stringify(transferredResults) !== JSON.stringify([['o'], ['o']])) {
    reportError(
      'ICL evaluation batch behavior',
      `expected ICL-to-SC transfer [["o"],["o"]], received ` +
        JSON.stringify(transferredResults)
    );
  }

  const nestedCriteriaRows = [
    ['ID', 'Check', 'Excluded', 'Memo', 'Criterion', 'Level'],
    ['3.3.4: Error prevention', '', '', '', '', ''],
    ['3.3.4-A-1', 'o', '', '', '3.3.4,3.3.6', 'AA'],
    ['3.3.6: Error prevention all', '', '', '', '', ''],
    ['3.3.6-1', 'x', '', '', '3.3.6', 'AAA']
  ];
  const nestedResults = context.calculateIclPageResults(
    nestedCriteriaRows,
    ['?', '-', 'o', 'x']
  ).criteria;
  if (nestedResults['3.3.4'] !== 'o' || nestedResults['3.3.6'] !== 'x') {
    reportError(
      'ICL evaluation batch behavior',
      'shared lower-scope checks and higher-scope-only checks must aggregate independently'
    );
  }

  const source = read('sheet-icl.js');
  const evaluateSource = source.slice(
    source.indexOf('function evaluateIcl('),
    source.indexOf('\n}', source.indexOf('function evaluateIcl(')) + 2
  );
  if (evaluateSource.includes('gIclTplSheetName') ||
      evaluateSource.includes('iclTplSheet')) {
    reportError(
      'ICL evaluation batch behavior',
      'ICL evaluation must not depend on the disposable ICL template sheet'
    );
  }
  if (!evaluateSource.includes(
        '.getRange(iclFirstRow, 1, resultRowCount, 8)') ||
      !evaluateSource.includes('.copyTo(iclSheet.getRange(1, 1))') ||
      !evaluateSource.includes('iclSheet.deleteColumn(2)')) {
    reportError(
      'ICL evaluation batch behavior',
      'the result sheet must retain page 1 columns A, E, F, G and H'
    );
  }
  if (evaluateSource.includes('aggregateIclPageResults') ||
      !evaluateSource.includes(
        '.getRange(pageFirstRow, 2, copyRowCount, 1)') ||
      !evaluateSource.includes('.setValues(resultRows)')) {
    reportError(
      'ICL evaluation batch behavior',
      'the ICL list must place each page raw column B from column F onward'
    );
  }
  if (!evaluateSource.includes('iclSheet.setFrozenColumns(5)')) {
    reportError(
      'ICL evaluation batch behavior',
      'the five ICL heading columns must remain frozen'
    );
  }
  if (!evaluateSource.includes('requiredColumns - currentMaxColumns') ||
      !evaluateSource.includes('iclSheet.insertColumnsAfter(')) {
    reportError(
      'ICL evaluation batch behavior',
      'the result sheet must allocate enough columns for every page result'
    );
  }
});

runCheck('ICL technique link behavior', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('sheet-icl.js'), context, {filename: 'sheet-icl.js'});

  const cases = [
    ['ARIA16', 'WCAG21', 'aria'],
    ['G82', 'WCAG21', 'general'],
    ['H30', 'WCAG21', 'html'],
    ['C9', 'WCAG21', 'css'],
    ['F3', 'WCAG21', 'failures'],
    ['PDF1', 'WCAG21', 'pdf'],
    ['SCR20', 'WCAG21', 'client-side-script'],
    ['SVR1', 'WCAG21', 'server-side-script'],
    ['T1', 'WCAG21', 'text'],
    ['C41', 'WCAG22', 'css'],
    ['F108', 'WCAG22', 'failures'],
    ['G218', 'WCAG22', 'general'],
    ['H100', 'WCAG22', 'html']
  ];

  for (const [techniqueId, wcagVersion, directory] of cases) {
    const expected =
      `https://waic.jp/translations/${wcagVersion}/Techniques/` +
      `${directory}/${techniqueId}`;
    const actual = context.getIclTechniqueUrl(techniqueId);
    if (actual !== expected) {
      reportError(
        'ICL technique link behavior',
        `expected ${expected}, received ${actual}`
      );
    }
  }

  const w3cOnlyCases = [
    ['ARIA25', 'aria'],
    ['ARIA26', 'aria'],
    ['C44', 'css'],
    ['F113', 'failures'],
    ['G226', 'general'],
    ['SCR40', 'client-side-script']
  ];
  for (const [techniqueId, directory] of w3cOnlyCases) {
    const expected =
      `https://www.w3.org/WAI/WCAG22/Techniques/${directory}/${techniqueId}`;
    const actual = context.getIclTechniqueUrl(techniqueId);
    if (actual !== expected) {
      reportError(
        'ICL technique link behavior',
        `expected ${expected}, received ${actual}`
      );
    }
  }

  if (context.getIclTechniqueUrl('-') !== '') {
    reportError(
      'ICL technique link behavior',
      'non-technique markers must not receive links'
    );
  }

  const source = read('sheet-icl.js');
  const understandingUrl = context.getIclUnderstandingUrl(
    'error-prevention-all'
  );
  if (understandingUrl !==
      'https://waic.jp/translations/WCAG22/Understanding/' +
      'error-prevention-all') {
    reportError(
      'ICL technique link behavior',
      `unexpected Understanding URL: ${understandingUrl}`
    );
  }
  if (!source.includes('getIclCriterionHeadingRichText(') ||
      !source.includes('.setRichTextValue(') ||
      !source.includes('.setForegroundColor(gLabelColorDarkText)')) {
    reportError(
      'ICL technique link behavior',
      'criterion headings must link to WCAG 2.2 Understanding pages'
    );
  }
  const wcag22UnderstandingSlugs = [
    'focus-not-obscured-minimum',
    'focus-not-obscured-enhanced',
    'focus-appearance',
    'dragging-movements',
    'target-size-minimum',
    'consistent-help',
    'redundant-entry',
    'accessible-authentication-minimum',
    'accessible-authentication-enhanced'
  ];
  for (const dictionaryFile of ['vals-ja.js', 'vals-en.js']) {
    const dictionarySource = read(dictionaryFile);
    for (const slug of wcag22UnderstandingSlugs) {
      if (!dictionarySource.includes(`'${slug}'`)) {
        reportError(
          'ICL technique link behavior',
          `${dictionaryFile} must define the WCAG 2.2 slug ${slug}`
        );
      }
    }
  }
  if (!source.includes('range.setRichTextValues(richTextValues)')) {
    reportError(
      'ICL technique link behavior',
      'technique links must be written in one batch'
    );
  }
  if (!source.includes(".setForegroundColor('#000000')") ||
      !source.includes('.setUnderline(true)')) {
    reportError(
      'ICL technique link behavior',
      'applicable technique links must have a consistent visible style'
    );
  }
  if (!source.includes('.setForegroundColor(gLabelColor)') ||
      !source.includes('.setUnderline(false)')) {
    reportError(
      'ICL technique link behavior',
      'excluded technique links must follow the excluded row style'
    );
  }
  if (!source.includes("sheet.getName() !== gIclTplSheetName") ||
      !source.includes('firstColumn > 3 || lastColumn < 3')) {
    reportError(
      'ICL technique link behavior',
      'applicability edits must refresh only relevant ICL technique links'
    );
  }
});

runCheck('COB-CHA ICL other option behavior', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(
    read('vals-ja-icl-cobcha.js'),
    context,
    {filename: 'vals-ja-icl-cobcha.js'}
  );

  const situations = context.getIclSituationCobcha();
  const tests = context.getIclTestCobcha();

  if (!tests['2.4.9'] || tests['2.4.9'].length !== 3 ||
      tests['2.4.4'].some((item) => item.length > 3)) {
    reportError(
      'COB-CHA ICL other option behavior',
      '2.4.4 and 2.4.9 must retain independent implementation rows'
    );
  }
  const sharedErrorPreventionRows = []
    .concat(tests['3.3.4-A'], tests['3.3.4-B'], tests['3.3.4-C']);
  if (sharedErrorPreventionRows.some((item) =>
      item[3] !== '3.3.4,3.3.6') ||
      !tests['3.3.6'] || tests['3.3.6'].length !== 1 ||
      tests['3.3.6'][0][0] !== '-' ||
      tests['3.3.6-A'] || tests['3.3.6-B']) {
    reportError(
      'COB-CHA ICL other option behavior',
      '3.3.4 checks must be shared with 3.3.6 while broader forms remain independent'
    );
  }
  if (tests['2.1.1'].slice(0, 3).some((item) =>
      item[3] !== '2.1.1,2.1.3') ||
      tests['2.1.1'][3][3] !== undefined ||
      !tests['2.1.3'] || tests['2.1.3'].length !== 1 ||
      tests['2.1.3'][0][0] !== '-') {
    reportError(
      'COB-CHA ICL other option behavior',
      '2.1.1 keyboard checks must be shared while its exception remains specific to 2.1.1'
    );
  }
  if (tests['3.3.8'].slice(0, 3).some((item) =>
      item[3] !== '3.3.8,3.3.9') ||
      tests['3.3.8'][4][3] !== '3.3.8,3.3.9' ||
      tests['3.3.9'].length !== 1 || tests['3.3.9'][0][0] !== '-') {
    reportError(
      'COB-CHA ICL other option behavior',
      '3.3.8 checks must be shared while 3.3.9 retains its narrower exceptions'
    );
  }
  if (tests['1.4.5-B'][1][3] !== '1.4.5,1.4.9' ||
      tests['1.4.5-B'][2][3] !== '1.4.5,1.4.9' ||
      tests['1.4.5-A'][0][3] !== undefined ||
      tests['1.4.9'].length !== 2) {
    reportError(
      'COB-CHA ICL other option behavior',
      '1.4.5 and 1.4.9 must share only checks with the same image-of-text scope'
    );
  }
  const commonSituations = situations['UI-Common'];
  const commonItems = []
    .concat(tests['UI-Common-A'], tests['UI-Common-B']);
  const expectedCommonItems = [
    ['G108/H44/ARIA6/ARIA10/ARIA14/ARIA16/PDF10', '1.1.1,1.3.1,4.1.2'],
    ['G108/H65', '1.1.1,1.3.1,4.1.2'],
    ['F111', '1.3.1,2.5.3,4.1.2'],
    ['G83/PDF5', '3.3.1,3.3.3']
  ];
  if (!commonSituations || commonSituations['UI-Common-Other'] ||
      tests['UI-Common-Other'] || commonItems.length !== 4 ||
      expectedCommonItems.some(([technique, criteria]) =>
        !commonItems.some((item) =>
          item[0] === technique && item[2] === false && item[3] === criteria))) {
    reportError(
      'COB-CHA ICL other option behavior',
      'shared user-interface checks must be visible and map to their criteria'
    );
  }
  for (const criterion of Object.keys(situations)) {
    if (!/^\d+\.\d+\.\d+$/.test(criterion)) continue;
    const otherId = `${criterion}-Other`;
    const situationIds = Object.keys(situations[criterion]);

    if (situationIds[situationIds.length - 1] !== otherId ||
        situations[criterion][otherId] !== 'その他') {
      reportError(
        'COB-CHA ICL other option behavior',
        `${criterion} must end with an Other situation`
      );
    }

    const otherIsApplicable = criterion === '1.3.1' || criterion === '4.1.2';
    const expectedOther = otherIsApplicable ?
      [['-', 'その他']] : [['-', 'その他', true]];
    if (JSON.stringify(tests[otherId]) !== JSON.stringify(expectedOther)) {
      reportError(
        'COB-CHA ICL other option behavior',
        otherIsApplicable ?
          `${criterion} Other must be applicable by default` :
          `${criterion} must have an initially excluded Other check item`
      );
    }
  }

  const obsolete412CatchAll = tests['4.1.2-A'].some(
    (item) => item[1].includes('4.1.2の趣旨に沿って')
  );
  if (obsolete412CatchAll) {
    reportError(
      'COB-CHA ICL other option behavior',
      'the former 4.1.2-A catch-all item must be removed'
    );
  }

  const imageAlternative = tests['1.1.1-A'];
  if (imageAlternative.some((item) => item[1].includes('iframeの内容が地図')) ||
      imageAlternative.length !== 2) {
    reportError(
      'COB-CHA ICL other option behavior',
      'the former 1.1.1-A iframe item must be removed'
    );
  }

  const interactiveAlternative = tests['1.1.1-C'];
  if (interactiveAlternative.length !== 3 ||
      interactiveAlternative[2][0] !== 'H24' ||
      interactiveAlternative[2][2] !== true) {
    reportError(
      'COB-CHA ICL other option behavior',
      'shared form-control checks must leave the 1.1.1-C area check intact'
    );
  }

  const customControl = tests['4.1.2-B'][0];
  if (customControl[0] !== 'ARIA4/ARIA5/ARIA16/G202' ||
      !customControl[1].includes('役割に応じたキーボード操作およびフォーカス管理')) {
    reportError(
      'COB-CHA ICL other option behavior',
      '4.1.2-B-1 must cover custom control semantics and operation'
    );
  }

  const ariaHeading = tests['1.3.1-B'][1];
  if (ariaHeading[0] !== 'ARIA12' ||
      !ariaHeading[1].includes('見出し階層に対応した適切なaria-level')) {
    reportError(
      'COB-CHA ICL other option behavior',
      '1.3.1-B-2 must prefer native headings and require a proper aria-level'
    );
  }

  const ariaRegion = tests['1.3.1-B'][5];
  if (ariaRegion[0] !== 'ARIA20' ||
      !ariaRegion[1].includes('アクセシブルな名前を持つsection要素') ||
      !ariaRegion[1].includes('重要な領域に限って')) {
    reportError(
      'COB-CHA ICL other option behavior',
      '1.3.1-B-6 must limit named regions to important unmatched areas'
    );
  }

  const errorDialog = tests['3.3.1'][3];
  const correctionDialog = tests['3.3.3-B'][2];
  if (!errorDialog[1].includes('適切に実装されたalertdialog') ||
      !errorDialog[1].includes('エラー箇所を特定できる') ||
      !correctionDialog[1].includes('適切に実装されたalertdialog') ||
      !correctionDialog[1].includes('エラーの修正方法が提供されている')) {
    reportError(
      'COB-CHA ICL other option behavior',
      'ARIA18 checks must identify a proper alertdialog and its purpose'
    );
  }

  const allItems = Object.values(tests).flat();
  if (tests['3.3.2'].some((item) => item[0] === 'H65') ||
      allItems.some((item) => item[0] === 'F87')) {
    reportError(
      'COB-CHA ICL other option behavior',
      'obsolete H65 for 3.3.2 and F87 checks must be removed'
    );
  }

  if (tests['2.4.3'][0][0] !== 'G59/C27/PDF3' ||
      tests['2.4.8'][3][0] !== 'G127' ||
      tests['2.4.8'][3][1].includes('メタデータ') ||
      tests['3.1.3'][0][0] !== 'G101/G55/H40/G62') {
    reportError(
      'COB-CHA ICL other option behavior',
      'obsolete H4, H59, and H60 references must be removed'
    );
  }

  const expectedNewTechniques = [
    ['1.2.3', 'G226'],
    ['1.2.5', 'F113'],
    ['1.3.1-A', 'H101'],
    ['UI-Common-A', 'F111'],
    ['1.4.10', 'G224'],
    ['1.4.10', 'G225'],
    ['2.2.2', 'F112'],
    ['2.3.3', 'SCR40'],
    ['2.4.3', 'H102'],
    ['2.4.7', 'C45'],
    ['2.4.8', 'ARIA26'],
    ['2.5.5', 'C44'],
    ['4.1.3-A', 'ARIA25']
  ];
  for (const [situation, technique] of expectedNewTechniques) {
    if (!tests[situation].some((item) => item[0].split('/').includes(technique))) {
      reportError(
        'COB-CHA ICL other option behavior',
        `${situation} must include ${technique}`
      );
    }
  }

  if (tests['1.3.1-A'][5][0] !== 'H101' ||
      tests['1.3.1-B'][0][0] !== 'ARIA11') {
    reportError(
      'COB-CHA ICL other option behavior',
      'H101 must be 1.3.1-A-6 and ARIA11 must remain an independent B item'
    );
  }

  if (commonItems.filter((item) => item[0] === 'F111').length !== 1 ||
      allItems.filter((item) => item[0].split('/').includes('F111')).length !== 1) {
    reportError(
      'COB-CHA ICL other option behavior',
      'F111 must be one shared item for 1.3.1, 2.5.3, and 4.1.2'
    );
  }

  if (!tests['2.5.8'][0][1].includes('直径24 CSSピクセル') ||
      tests['2.5.8'][0][1].includes('44ピクセル') ||
      allItems.some((item) => item[0].split('/').includes('ARIA27'))) {
    reportError(
      'COB-CHA ICL other option behavior',
      'C42 must describe 24px spacing and experimental ARIA27 must stay excluded'
    );
  }
});

runCheck('Issue page matching behavior', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('sheet-issue.js'), context, {filename: 'sheet-issue.js'});

  const targetUrl = 'https://example.com/page';
  const cases = [
    ['all', targetUrl, true],
    [' ALL ', targetUrl, true],
    [targetUrl, targetUrl, true],
    [`https://example.com/other, ${targetUrl}`, targetUrl, true],
    [`${targetUrl}, ${targetUrl}`, targetUrl, true],
    ['https://example.com/other', targetUrl, false],
    ['', targetUrl, false],
    ['all', '', false]
  ];

  cases.forEach(([places, url, expected]) => {
    const actual = context.issueAppliesToUrl(places, url);
    if (actual !== expected) {
      reportError(
        'Issue page matching behavior',
        `places="${places}", url="${url}": expected ${expected}, received ${actual}`
      );
    }
  });
});

runCheck('Issue ID allocation and edit targeting', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('sheet-issue.js'), context, {filename: 'sheet-issue.js'});

  const data = [
    ['ID', 'Name'],
    [1, 'First'],
    [5, 'Fifth'],
    [5, 'Duplicate fifth'],
    ['', 'No ID'],
    ['invalid', 'Invalid ID']
  ];
  const matchedRows = Array.from(context.getIssueRowsById(data, 5));
  if (JSON.stringify(matchedRows) !== JSON.stringify([3, 4])) {
    reportError(
      'Issue ID allocation and edit targeting',
      `expected duplicate ID rows [3,4], received ${JSON.stringify(matchedRows)}`
    );
  }
  if (context.getNextIssueId(data) !== 6) {
    reportError(
      'Issue ID allocation and edit targeting',
      `expected next Issue ID 6, received ${context.getNextIssueId(data)}`
    );
  }

  const source = read('sheet-issue.js');
  if (!source.includes('LockService.getDocumentLock()')) {
    reportError(
      'Issue ID allocation and edit targeting',
      'Issue writes must use a document lock'
    );
  }
  if (source.includes('targetRow = sheet.getActiveCell().getRow();')) {
    reportError(
      'Issue ID allocation and edit targeting',
      'Issue edits must not target the active row'
    );
  }
});

runCheck('Issue dialog checked ordering and page numbers', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('sheet-issue.js'), context, {filename: 'sheet-issue.js'});

  const pageNumbers = context.buildPageNumberMap([
    ['1', 'https://example.com/one'],
    [2, 'https://example.com/two'],
    [0, 'https://example.com/zero'],
    ['', '']
  ]);
  const pageNumberCases = [
    ['1', 'https://example.com/one', '1'],
    ['different-sheet', 'https://example.com/two', '2'],
    ['0', 'https://example.com/zero', '0'],
    ['1', 'https://example.com/changed', '1'],
    ['missing', 'https://example.com/missing', '']
  ];
  pageNumberCases.forEach(([sheetName, url, expected]) => {
    const actual = context.findPageNumber(pageNumbers, sheetName, url);
    if (actual === expected) return;
    reportError(
      'Issue dialog checked ordering and page numbers',
      `sheet="${sheetName}", url="${url}": expected "${expected}", received "${actual}"`
    );
  });

  const issueUi = read('ui-issue.html');
  const requiredUiFragments = [
    'function sortCheckedCriteria()',
    'function sortCheckedTechniques()',
    'dataset.originalIndex',
    'data-i18n="checked-items-move-note"',
    "pageNumberLabel.className = 'page-number'",
    "pageNumberLabel.setAttribute('aria-hidden', 'true')",
    "t('page-number-accessible-prefix', 'Page number')"
  ];
  requiredUiFragments.forEach((fragment) => {
    if (!issueUi.includes(fragment)) {
      reportError(
        'Issue dialog checked ordering and page numbers',
        `missing UI behavior: ${fragment}`
      );
    }
  });

  const dictionaryContext = {};
  vm.createContext(dictionaryContext);
  vm.runInContext(read('vals-basic.js'), dictionaryContext, {filename: 'vals-basic.js'});
  vm.runInContext(read('vals-ja.js'), dictionaryContext, {filename: 'vals-ja.js'});
  if (
    dictionaryContext.getUiJa()['page-number-accessible-prefix'] !==
    'ページ番号'
  ) {
    reportError(
      'Issue dialog checked ordering and page numbers',
      'Japanese accessible page number prefix is missing'
    );
  }
});

runCheck('UI inline JavaScript syntax', () => {
  for (const file of htmlFiles) {
    const scripts = extractInlineScripts(read(file));

    scripts.forEach((script, index) => {
      try {
        new vm.Script(script, {filename: `${file}#script-${index + 1}`});
      } catch (error) {
        reportError('UI inline JavaScript syntax', `${file}: ${error.message}`);
      }
    });
  }
});

runCheck('Duplicate static HTML ids', () => {
  for (const file of htmlFiles) {
    const markup = stripHtmlScripts(read(file));
    const ids = new Map();
    const pattern = /\bid\s*=\s*(["'])([^"']+)\1/gi;
    let match;

    while ((match = pattern.exec(markup)) !== null) {
      ids.set(match[2], (ids.get(match[2]) || 0) + 1);
    }

    for (const [id, count] of ids) {
      if (count > 1) {
        reportError('Duplicate static HTML ids', `${file}: id="${id}" appears ${count} times`);
      }
    }
  }
});

runCheck('Japanese UI translation keys', () => {
  const context = {};
  vm.createContext(context);
  vm.runInContext(read('vals-basic.js'), context, {filename: 'vals-basic.js'});
  vm.runInContext(read('vals-ja.js'), context, {filename: 'vals-ja.js'});

  const dictionary = context.getUiJa();
  const usedKeys = new Map();
  const pattern = /\bdata-i18n(?:-html|-href|-title|-aria-label|-value|-placeholder)?\s*=\s*(["'])([^"']+)\1/gi;

  for (const file of htmlFiles) {
    const markup = stripHtmlScripts(read(file));
    let match;

    while ((match = pattern.exec(markup)) !== null) {
      if (!usedKeys.has(match[2])) usedKeys.set(match[2], []);
      usedKeys.get(match[2]).push(file);
    }
  }

  for (const [key, files] of usedKeys) {
    if (!Object.prototype.hasOwnProperty.call(dictionary, key)) {
      reportError(
        'Japanese UI translation keys',
        `missing "${key}" used by ${[...new Set(files)].join(', ')}`
      );
    }
  }
});

runCheck('HTML template references', () => {
  const referencedTemplates = new Set();
  const pattern = /(?:createTemplateFromFile|createHtmlOutputFromFile|showDialog)\(\s*(["'])(ui-[^"']+)\1/g;

  for (const file of [...jsFiles, ...htmlFiles]) {
    const source = read(file);
    let match;

    while ((match = pattern.exec(source)) !== null) {
      referencedTemplates.add(match[2]);
    }
  }

  for (const template of referencedTemplates) {
    if (!htmlFiles.includes(`${template}.html`)) {
      reportError('HTML template references', `missing ${template}.html`);
    }
  }
});

runCheck('Spreadsheet menu handlers', () => {
  const functions = new Set();
  const functionPattern = /^function\s+([A-Za-z_$][\w$]*)\s*\(/gm;

  for (const file of jsFiles) {
    const source = read(file);
    let match;

    while ((match = functionPattern.exec(source)) !== null) {
      functions.add(match[1]);
    }
  }

  const menuPattern = /\.addItem\([\s\S]*?,\s*(["'])([A-Za-z_$][\w$]*)\1\s*\)/g;

  for (const file of jsFiles) {
    const source = read(file);
    let match;

    while ((match = menuPattern.exec(source)) !== null) {
      if (!functions.has(match[2])) {
        reportError('Spreadsheet menu handlers', `${file}: missing function ${match[2]}()`);
      }
    }
  }
});

runCheck('Deprecated client dependencies', () => {
  const forbidden = [
    {name: 'jQuery', pattern: /jquery|ajax\.googleapis\.com\/ajax\/libs\/jquery|\$\s*\(/i},
    {name: 'Font Awesome', pattern: /fontawesome|use\.fontawesome\.com|\bfa-(?:spin|spinner|external-link)\b/i}
  ];

  for (const file of htmlFiles) {
    const source = read(file);

    for (const dependency of forbidden) {
      if (dependency.pattern.test(source)) {
        reportError('Deprecated client dependencies', `${file}: ${dependency.name} reference found`);
      }
    }
  }
});

runCheck('Unsafe dynamic HTML sinks', () => {
  const sinkPattern = /(?:\.innerHTML\s*=|\.outerHTML\s*=|insertAdjacentHTML\s*\(|document\.write\s*\()/g;

  for (const file of htmlFiles) {
    let source = read(file);

    // This one sink is intentional: only trusted, repository-owned translations
    // referenced through data-i18n-html may contain markup.
    if (file === 'ui-common.html') {
      source = source.replace('el.innerHTML = dict[key];', '');
    }

    if (sinkPattern.test(source)) {
      reportError('Unsafe dynamic HTML sinks', `${file}: dynamic HTML sink found`);
    }
  }
});

runCheck('Issue dialog template escaping', () => {
  const issueServer = read('sheet-issue.js');
  const issueTemplate = read('ui-issue.html');

  if (/id=["']target-url["'][^\n]*\+/.test(issueServer)) {
    reportError(
      'Issue dialog template escaping',
      'target URL must not be concatenated into HTML'
    );
  }

  if (!/value=["']<\?=\s*TEMPLATE_VARS\.targetUrl\s*\|\|\s*["']["']\s*\?>["']/.test(issueTemplate)) {
    reportError(
      'Issue dialog template escaping',
      'target URL must use an escaped template expression'
    );
  }
});

runCheck('Lump edit dialog template escaping', () => {
  const lumpServer = read('input-utility.js');
  const lumpTemplate = read('ui-lump-edit.html');
  const fields = ['row', 'col', 'vals', 'isAppend'];

  if (/id=["']lump-(?:pos-row|pos-col|val|is_append)["'][^\n]*\+/.test(lumpServer)) {
    reportError(
      'Lump edit dialog template escaping',
      'lump edit values must not be concatenated into HTML'
    );
  }

  fields.forEach((field) => {
    const expression = `<?= TEMPLATE_VARS.${field} ?>`;
    if (!lumpTemplate.includes(expression)) {
      reportError(
        'Lump edit dialog template escaping',
        `${field} must use an escaped template expression`
      );
    }
  });
});

runCheck('clasp excludes maintenance tools', () => {
  const claspIgnore = read('.claspignore');

  if (!/^\*\*\/tools\/\*\*$/m.test(claspIgnore)) {
    reportError('clasp excludes maintenance tools', 'add **/tools/** to .claspignore');
  }
});

for (const check of checks) {
  const mark = check.passed ? 'PASS' : 'FAIL';
  process.stdout.write(`${mark} ${check.name}\n`);
}

if (errors.length > 0) {
  process.stderr.write('\n');
  errors.forEach((error) => process.stderr.write(`- ${error}\n`));
  process.exitCode = 1;
} else {
  process.stdout.write(`\nAll ${checks.length} checks passed.\n`);
}
