/**
 * PDF adapter for the COB-CHA Implementation Check List.
 *
 * The complete COB-CHA ICL is the source of truth. This file selects checks
 * applicable to PDF and overrides only descriptions that require PDF-specific
 * terminology.
 * LICENSE: [CC0](https://creativecommons.org/share-your-work/public-domain/cc0)
 */

function getCobchaPdfTechniqueIds(techniques) {
  var ids = String(techniques || '').split('/');
  var hasPdfTechnique = ids.some(function (id) {
    return /^PDF\d+$/.test(id);
  });
  var hasGeneralTechnique = ids.some(function (id) {
    return /^G\d+$/.test(id);
  });
  var isGenericSet = ids.every(function (id) {
    return id === '-' || /^PDF\d+$/.test(id) || /^G\d+$/.test(id) ||
      /^F\d+$/.test(id);
  });
  var isF111 = ids.length === 1 && ids[0] === 'F111';
  var isImplementationIndependent = ids.every(function (id) {
    return id === '-';
  });
  if (!hasPdfTechnique && !(hasGeneralTechnique && isGenericSet) &&
      !isF111 && !isImplementationIndependent) return '';

  var selected = ids.filter(function (id) {
    return id === '-' || /^PDF\d+$/.test(id) || /^G\d+$/.test(id) ||
      /^F\d+$/.test(id);
  });
  if (selected.length === 0) return '';

  // G108 describes use of standard markup in general, but the complete ICL
  // wording for it is HTML-specific. PDF form semantics are covered by PDF12.
  selected = selected.filter(function (id) { return id !== 'G108'; });
  return selected.join('/');
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
    '3.1.1' : [
      ['PDF16', 'PDF文書のプロパティで、文書全体の主要な言語が正しく設定されている']
    ],
    '3.1.2' : [
      ['PDF19', '文書全体の主要な言語と異なる言語の範囲に、適切なLangエントリが設定されている']
    ],
    '4.1.1' : [
      ['-', 'PDFの構造とタグ付けに、支援技術による解釈を妨げるエラーがない', {defaultExcluded: true}]
    ],
    '4.1.2-C' : [
      ['PDF10/PDF12', 'PDFフォームコントロールがPDFの標準的な機能で作成され、名前、役割、状態および値を支援技術が理解できる']
    ]
  };
}

function getIclTestCobchaPdf() {
  var sourceTests = getIclTestCobcha();
  var overrides = getCobchaPdfOverrides();
  var webOnlySituations = {
    '2.5.6' : true
  };
  var tests = {};

  for (var situationId in sourceTests) {
    if (webOnlySituations[situationId]) continue;
    if (overrides[situationId]) {
      tests[situationId] = overrides[situationId];
      continue;
    }

    var selected = [];
    for (var i = 0; i < sourceTests[situationId].length; i++) {
      var sourceItem = sourceTests[situationId][i];
      var techniques = getCobchaPdfTechniqueIds(sourceItem[0]);
      if (techniques === '') continue;
      var item = sourceItem.slice();
      item[0] = techniques;
      selected.push(item);
    }
    if (selected.length > 0) tests[situationId] = selected;
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
