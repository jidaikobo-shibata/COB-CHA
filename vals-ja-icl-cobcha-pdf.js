/**
 * PDF adapter for the COB-CHA Implementation Check List.
 *
 * The complete COB-CHA ICL is the source of truth. Checks marked with
 * {pdf: true} are selected, and this file overrides only descriptions that
 * require PDF-specific terminology.
 * LICENSE: [CC0](https://creativecommons.org/share-your-work/public-domain/cc0)
 */

/**
 * Remove references to implementation techniques that do not apply to PDF.
 * This does not decide whether the check is included; inclusion is controlled
 * exclusively by the {pdf: true} option in the complete ICL.
 */
function getCobchaPdfRelatedTechniques(techniques) {
  return String(techniques || '').split('/').filter(function (id) {
    return id === '-' || /^PDF\d+$/.test(id) || /^G\d+$/.test(id) ||
      /^F\d+$/.test(id);
  }).join('/') || '-';
}

function getCobchaPdfOverrides() {
  return {
    'UI-Common-A' : [
      ['PDF10/PDF12', 'すべてのPDFフォームコントロールに、内容を表すアクセシブルな名前が提供されている', {criteria: '1.1.1,1.3.1,4.1.2'}],
      ['F111', '可視ラベルを持つすべてのPDFフォームコントロールに、アクセシブルな名前が提供されている', {criteria: '1.3.1,2.5.3,4.1.2'}]
    ],
    '1.1.1-A' : [
      ['PDF1', 'PDFのAltエントリは、画像と同等の記述を短いテキストで提供している。短いテキストだけでは同等の記述ができない場合は、画像を識別できる短いテキストが提供されている。画像がリンクの場合は1.1.1-Cを優先する']
    ],
    '1.1.1-C' : [
      ['G82/PDF13', 'リンクとなる画像のAltエントリは、リンクの目的を短いテキストで提供している（2.4.4も確認）']
    ],
    '1.1.1-F' : [
      ['PDF4', '装飾的な非テキストコンテンツは、Artifactタグ等によって支援技術が無視できるように実装されている']
    ],
    '1.3.1-A' : [
      ['G115/PDF9', '見出しは見出しタグによってセマンティックにマークアップされている'],
      ['PDF21', 'リストはリストタグによってマークアップされている'],
      ['G117', 'フォント、太字、斜体などの見た目によって伝えている情報を、テキストでも理解できる'],
      ['PDF6/PDF20', '表形式の情報はテーブルタグによってマークアップされ、見出しセルとデータセルの関係を理解できる']
    ],
    '2.4.2' : [
      ['G88/PDF18/F25', 'PDF文書には内容を説明するタイトルがある']
    ],
    '3.1.1' : [
      ['PDF16', 'PDF文書のプロパティで、文書全体の主要な言語が正しく設定されている']
    ],
    '3.1.2' : [
      ['PDF19', '文書全体の主要な言語と異なる言語の範囲に、適切なLangエントリが設定されている']
    ],
    '4.1.1' : [
      ['-', 'PDFの構造とタグ付けに、支援技術による解釈を妨げるエラーがない']
    ],
    '4.1.2-C' : [
      ['PDF10/PDF12', 'PDFフォームコントロールがPDFの標準的な機能で作成され、名前、役割、状態および値を支援技術が理解できる']
    ]
  };
}

function getIclTestCobchaPdf() {
  var sourceTests = getIclTestCobcha(true);
  var overrides = getCobchaPdfOverrides();
  var tests = {};

  for (var situationId in sourceTests) {
    var selected = [];
    for (var i = 0; i < sourceTests[situationId].length; i++) {
      var sourceItem = sourceTests[situationId][i];
      var options = sourceItem[2];
      if (!options || typeof options !== 'object' || options.pdf !== true) {
        continue;
      }

      var item = sourceItem.slice();
      item[0] = getCobchaPdfRelatedTechniques(item[0]);
      var pdfOptions = {};
      for (var optionName in options) {
        if (optionName !== 'defaultExcluded' && optionName !== 'pdf' &&
            optionName !== 'pdfOnly' && optionName !== 'easy') {
          pdfOptions[optionName] = options[optionName];
        }
      }
      if (Object.keys(pdfOptions).length > 0) {
        item[2] = pdfOptions;
      } else {
        item.splice(2, 1);
      }
      selected.push(item);
    }
    if (selected.length === 0) continue;
    tests[situationId] = overrides[situationId] || selected;
  }

  return tests;
}

function getIclSituationCobchaPdf() {
  var sourceSituations = getIclSituationCobcha();
  var tests = getIclTestCobchaPdf();
  var situations = {};

  for (var criterion in sourceSituations) {
    var selected = {};
    var source = sourceSituations[criterion];
    for (var situationId in source) {
      if (tests[situationId]) selected[situationId] = source[situationId];
    }
    if (Object.keys(selected).length > 0) situations[criterion] = selected;
  }

  return situations;
}
