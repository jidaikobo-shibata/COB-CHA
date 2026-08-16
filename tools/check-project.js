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

  const source = read('sheet-icl.js');
  if (!source.includes('.setValues(resultRows);')) {
    reportError(
      'ICL evaluation batch behavior',
      'ICL result matrix must be written with one setValues call'
    );
  }
  if (source.includes('.copyTo(iclSheet.getRange(2, col)')) {
    reportError(
      'ICL evaluation batch behavior',
      'per-page destination copy must not be restored'
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

  if (context.getIclTechniqueUrl('-') !== '') {
    reportError(
      'ICL technique link behavior',
      'non-technique markers must not receive links'
    );
  }

  const source = read('sheet-icl.js');
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
