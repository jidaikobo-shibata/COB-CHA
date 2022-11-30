/**
 * variables for COB-CHA
 */

/**
 * COB-CHA Version
 * @return Integer
 */
function getVersion ()
{
  return 40;
}

/**
 * WCAG 2.1
 */
var gCriteria21 = [
  '1.3.4', '1.3.5', '1.3.6', '1.4.10', '1.4.11', '1.4.12', '1.4.13',
  '2.1.4', '2.2.6', '2.3.3', '2.5.1', '2.5.2', '2.5.3', '2.5.4', '2.5.5', '2.5.6',
  '4.1.3'
];

/**
 * WCAG 2.0/2.1 Single-A criteria
 */
var gSingleACriteria = [
  '1.1.1', '1.2.1', '1.2.2', '1.2.3', '1.3.1', '1.3.2', '1.3.3', '1.4.1',
  '1.4.2', '2.1.1', '2.1.2', '2.1.4', '2.2.1', '2.2.2', '2.3.1', '2.4.1',
  '2.4.2', '2.4.3', '2.4.4', '2.5.1', '2.5.2', '2.5.3', '2.5.4', '3.1.1',
  '3.2.1', '3.2.2', '3.3.1', '3.3.2', '4.1.1', '4.1.2'
];

/**
 * WCAG 2.0/2.1 double-A criteria
 */
var gDoubleACriteria = [
  '1.1.1', '1.2.1', '1.2.2', '1.2.3', '1.2.4', '1.2.5', '1.3.1', '1.3.2', 
  '1.3.3', '1.3.4', '1.3.5', '1.4.1', '1.4.2', '1.4.3', '1.4.4', '1.4.5', 
  '1.4.10', '1.4.11', '1.4.12', '1.4.13', '2.1.1', '2.1.2', '2.1.4', '2.2.1', 
  '2.2.2', '2.3.1', '2.4.1', '2.4.2', '2.4.3', '2.4.4', '2.4.5', '2.4.6', 
  '2.4.7', '2.5.1', '2.5.2', '2.5.3', '2.5.4', '3.1.1', '3.1.2', '3.2.1',
  '3.2.2', '3.2.3', '3.2.4', '3.3.1', '3.3.2', '3.3.3', '3.3.4', '4.1.1',
  '4.1.2', '4.1.3'
];

/**
 * trusted tester's check value
 */
var gSingleATestTt = [
  '2.A', '2.B', '2.C', '2.D', '3.A', '4.A', '4.B', '4.C',
  '4.E', '4.F', '4.G', '4.H', '5.C', '5.D', '5.E', '6.A',
  '6.B', '7.A', '7.B', '7.C', '7.D', '7.E', '8.A', '9.A',
  '10.B', '10.C', '10.D', '11.A', '12.A', '12.B', '12.C',
  '12.D', '13.A', '13.B', '14.A', '14.B', '14.C', '15.A',
  '15.B', '16.A', '16.B', '17.A', '20.A'
]

var gDoubleATestTt = [
  '4.D', '5.A', '5.B', '5.F', '5.G', '5.H', '7.E', '9.B',
  '9.C', '10.A', '11.B', '13.C', '17.B', '17.C', '18.A', '19.A'
];

var gRelTtAndCriteria = {
  '1.1.1': ['7.A', '7.B', '7.C', '7.D', '7.E'],
  '1.2.1': ['16.A', '16.B'],
  '1.2.2': ['17.A'],
  '1.2.3': ['16.A', '16.B'], // same as 1.2.1 temporary
  '1.2.4': ['17.C'],
  '1.2.5': ['17.B'],
  '1.3.1': ['5.C', '10.B', '10.C', '10.D', '14.A', '14.B', '14.C', '15.A'],
  '1.3.2': ['15.B'],
  '1.3.3': ['13.B'],
  '1.4.1': ['13.A'],
  '1.4.2': ['2.A'],
  '1.4.3': ['13.C'],
  '1.4.4': ['18.A'],
  '1.4.5': ['7.E'],
  '2.1.1': ['4.A', '4.B'],
  '2.1.2': ['4.C'],
  '2.2.1': ['8.A'],
  '2.2.2': ['2.B', '2.C'],
  '2.3.1': ['3.A'],
  '2.4.1': ['9.A'],
  '2.4.2': ['12.A', '12.B'],
  '2.4.3': ['4.F', '4.G', '4.H'],
  '2.4.4': ['6.A'],
  '2.4.5': ['19.A'],
  '2.4.6': ['5.B', '10.A'],
  '2.4.7': ['4.D'],
  '3.1.1': ['11.A'],
  '3.1.2': ['11.B'],
  '3.2.1': ['4.E'],
  '3.2.2': ['5.D'],
  '3.2.3': ['9.B'],
  '3.2.4': ['9.C'],
  '3.3.1': ['5.F'],
  '3.3.2': ['5.A'],
  '3.3.3': ['5.G'],
  '3.3.4': ['5.H'],
  '4.1.1': ['20.A'],
  '4.1.2': ['2.D', '5.E', '6.B', '12.C', '12.D']
};

/**
 * Non-Interference
 */
var gNonInterference = [
  '1.4.2', '2.1.2', '2.2.2', '2.3.1'
];

/**
 * Non-Interference Trusted Tester
 */
var gNonInterferenceTt = [
  '2.A', '2.B', '2.C', '3.A', '4.C'
];

/**
 * URL
 */
var gUrlbase = {
  'understanding': {
    'en-wcag20': 'https://www.w3.org/TR/UNDERSTANDING-WCAG20/',
    'en-wcag21': 'https://www.w3.org/WAI/WCAG21/Understanding/', // and directory
    'ja-wcag20': 'https://waic.jp/docs/UNDERSTANDING-WCAG20/',
    'ja-wcag21': 'https://waic.jp/docs/WCAG21/Understanding/'
  },
  'tech': {
    'en-wcag20': 'https://www.w3.org/TR/WCAG20-TECHS/',
    'en-wcag21': 'https://www.w3.org/WAI/WCAG21/Techniques/',
    'ja-wcag20': 'https://waic.jp/docs/WCAG-TECHS/',
    'ja-wcag21': 'https://waic.jp/docs/WCAG-TECHS/'
  }
};

var gTechDirAbbr = {
  'G': 'general',
  'H': 'html',
  'C': 'css',
  'A': 'aria',
  'T': 'text',
  'P': 'pdf',
  'F': 'failures',
  'FL': 'flash',
  'SM': 'smil',
  'SL': 'silverlight',
  'SV': 'server-side-script',
  'SC': 'client-side-script'
};

/*
 * Techs
 */
var gRelTechsAndCriteria = {
  '1.1.1': ['G68', 'G73', 'G74', 'G82', 'G92', 'G94', 'G95', 'G100', 'G143', 'G144', 'G196', 'H2', 'H24', 'H30', 'H35', 'H36', 'H37', 'H44', 'H45', 'H46', 'H53', 'H65', 'H67', 'H86', 'C9', 'C18', 'ARIA6', 'ARIA9', 'ARIA10', 'ARIA15', 'PDF1', 'PDF4', 'F3', 'F13', 'F20', 'F30', 'F38', 'F39', 'F65', 'F67', 'F71', 'F72'],
  '1.2.1': ['G158', 'G159', 'G166', 'H96', 'F30', 'F67'],
  '1.2.2': ['G87', 'G93', 'H95', 'F8', 'F74', 'F75'],
  '1.2.3': ['G8', 'G58', 'G69', 'G78', 'G173', 'G203', 'H53', 'H96'],
  '1.2.4': ['G9', 'G87', 'G93'],
  '1.2.5': ['G8', 'G78', 'G173', 'G203', 'H96'],
  '1.2.6': ['G54', 'G81'],
  '1.2.7': ['G8', 'H96'],
  '1.2.8': ['G58', 'G69', 'G159', 'H46', 'H53', 'F74'],
  '1.2.9': ['G150', 'G151', 'G157'],
  '1.3.1': ['G115', 'G117', 'G138', 'G140', 'G141', 'G162', 'H39', 'H42', 'H43', 'H44', 'H48', 'H49', 'H51', 'H63', 'H65', 'H71', 'H73', 'H85', 'H97', 'C22', 'SCR21', 'T1', 'T2', 'T3', 'ARIA1', 'ARIA2', 'ARIA11', 'ARIA12', 'ARIA13', 'ARIA16', 'ARIA17', 'ARIA20', 'PDF6', 'PDF9', 'PDF10', 'PDF11', 'PDF12', 'PDF17', 'PDF20', 'PDF21', 'F2', 'F33', 'F34', 'F42', 'F43', 'F46', 'F48', 'F87', 'F90', 'F91', 'F92'],
  '1.3.2': ['G57', 'H34', 'H56', 'C6', 'C8', 'C27', 'PDF3', 'F1', 'F32', 'F33', 'F34', 'F49'],
  '1.3.3': ['G96', 'F14', 'F26'],
  '1.4.1': ['G14', 'G111', 'G182', 'G183', 'G205', 'C15', 'F13', 'F73', 'F81'],
  '1.4.2': ['G60', 'G170', 'G171', 'F23', 'F93'],
  '1.4.3': ['G18', 'G145', 'G148', 'G156', 'G174', 'F24', 'F83'],
  '1.4.4': ['G142', 'G146', 'G178', 'G179', 'C12', 'C13', 'C14', 'C17', 'C20', 'C22', 'C28', 'SCR34', 'F69', 'F80'],
  '1.4.5': ['G140', 'C6', 'C8', 'C12', 'C13', 'C14', 'C22', 'C30', 'PDF7'],
  '1.4.6': ['G17', 'G18', 'G148', 'G156', 'G174', 'F24', 'F83'],
  '1.4.7': ['G56'],
  '1.4.8': ['G146', 'G148', 'G156', 'G169', 'G172', 'G175', 'G188', 'G204', 'G206', 'C12', 'C13', 'C14', 'C19', 'C20', 'C21', 'C23', 'C24', 'C25', 'SCR34', 'F24', 'F88'],
  '1.4.9': ['G140', 'C6', 'C8', 'C12', 'C13', 'C14', 'C22', 'C30', 'PDF7'],
  '2.1.1': ['G90', 'G202', 'H91', 'SCR2', 'SCR20', 'SCR29', 'SCR35', 'PDF3', 'PDF11', 'PDF23', 'F42', 'F54', 'F55'],
  '2.1.2': ['G21', 'F10'],
  '2.1.3': ['G90', 'G202', 'H91', 'SCR2', 'SCR20', 'SCR29', 'SCR35', 'PDF3', 'PDF11', 'PDF23', 'F42', 'F54', 'F55'],
  '2.2.1': ['G4', 'G133', 'G180', 'G198', 'SCR1', 'SCR16', 'SCR33', 'SCR36', 'F40', 'F41', 'F58'],
  '2.2.2': ['G4', 'G11', 'G152', 'G186', 'G187', 'G191', 'SCR22', 'SCR33', 'F4', 'F7', 'F16', 'F47', 'F50'],
  '2.2.3': ['G5'],
  '2.2.4': ['G75', 'G76', 'SCR14', 'F40', 'F41'],
  '2.2.5': ['G105', 'G181', 'F12'],
  '2.3.1': ['G15', 'G19', 'G176'],
  '2.3.2': ['G19'],
  '2.4.1': ['G1', 'G123', 'G124', 'H64', 'H69', 'H70', 'C6', 'SCR28', 'ARIA11', 'PDF9'],
  '2.4.10': ['G141', 'H69'],
  '2.4.2': ['G88', 'G127', 'H25', 'PDF18', 'F25'],
  '2.4.3': ['G59', 'H4', 'C27', 'SCR26', 'SCR27', 'SCR37', 'PDF3', 'F44', 'F85'],
  '2.4.4': ['G53', 'G91', 'G189', 'H2', 'H24', 'H30', 'H33', 'H77', 'H78', 'H79', 'H80', 'H81', 'C7', 'SCR30', 'ARIA7', 'ARIA8', 'PDF11', 'PDF13', 'F63', 'F89'],
  '2.4.5': ['G63', 'G64', 'G125', 'G126', 'G161', 'G185', 'H59', 'PDF2'],
  '2.4.6': ['G130', 'G131'],
  '2.4.7': ['G149', 'G165', 'G195', 'C15', 'SCR31', 'F55', 'F78'],
  '2.4.8': ['G63', 'G65', 'G127', 'G128', 'H59', 'PDF14', 'PDF17'],
  '2.4.9': ['G91', 'G189', 'H2', 'H24', 'H30', 'H33', 'C7', 'SCR30', 'ARIA8', 'PDF11', 'PDF13', 'F84', 'F89'],
  '3.1.1': ['H57', 'SVR5', 'PDF16', 'PDF19'],
  '3.1.2': ['H58', 'PDF19'],
  '3.1.3': ['G55', 'G62', 'G70', 'G101', 'G112', 'H40', 'H54', 'H60'],
  '3.1.4': ['G55', 'G62', 'G70', 'G97', 'G102', 'H28', 'H60', 'PDF8'],
  '3.1.5': ['G79', 'G86', 'G103', 'G153', 'G160'],
  '3.1.6': ['G62', 'G120', 'G121', 'G163', 'H62'],
  '3.2.1': ['G107', 'F52', 'F55'],
  '3.2.2': ['G13', 'G80', 'H32', 'H84', 'SCR19', 'PDF15', 'F36', 'F37'],
  '3.2.3': ['G61', 'PDF14', 'PDF17', 'F66'],
  '3.2.4': ['G197', 'F31'],
  '3.2.5': ['G76', 'G110', 'H76', 'H83', 'SCR19', 'SCR24', 'SVR1', 'F9', 'F22', 'F41', 'F52', 'F60', 'F61'],
  '3.3.1': ['G83', 'G84', 'G85', 'G139', 'G199', 'SCR18', 'SCR32', 'ARIA18', 'ARIA19', 'ARIA21', 'PDF5', 'PDF22'],
  '3.3.2': ['G13', 'G83', 'G89', 'G131', 'G162', 'G167', 'G184', 'H44', 'H65', 'H71', 'H90', 'ARIA1', 'ARIA9', 'ARIA17', 'PDF5', 'PDF10', 'F82'],
  '3.3.3': ['G83', 'G84', 'G85', 'G139', 'G177', 'G199', 'SCR18', 'SCR32', 'ARIA2', 'ARIA18', 'PDF5', 'PDF22'],
  '3.3.4': ['G98', 'G99', 'G155', 'G164', 'G168', 'G199', 'SCR18'],
  '3.3.5': ['G71', 'G89', 'G184', 'G193', 'G194', 'H89'],
  '3.3.6': ['G98', 'G99', 'G155', 'G164', 'G168', 'G199'],
  '4.1.1': ['G134', 'G192', 'H74', 'H75', 'H88', 'H93', 'H94', 'F70', 'F77'],
  '4.1.2': ['G10', 'G108', 'G135', 'H44', 'H64', 'H65', 'H88', 'H91', 'ARIA4', 'ARIA5', 'ARIA14', 'ARIA16', 'PDF10', 'PDF12', 'F15', 'F20', 'F42', 'F59', 'F68', 'F79', 'F86', 'F89']
};

/**
 * global variables
 */
var gFallbackSheetName  = '*Fallback';
var gConfigSheetName    = '*Config';
var gReportSheetName    = '*Report';
var gResultSheetName    = '*Result';
var gTotalSheetName     = '*Total';
var gIssueSheetName     = '*Issue';
var gUrlListSheetName   = '*URLs';
var gScTplSheetName     = '*SC Template';
var gIclSheetName       = '*ICL Result'; // Japanese Only
var gIclTplSheetName    = '*ICL Template'; // Japanese Only
var gResourceFolderName = 'cob-cha-resource';
var gImagesFolderName   = 'cob-cha-images';
var gTrueColor          = '#f5fff3';
var gFalseColor         = '#f7f3ff';
var gLabelColor         = '#eeeeee';
var gLabelColorText     = '#000000';
var gLabelColorDark     = '#87823e';
var gLabelColorDarkText = '#ffffff';
var gNotYetBgColor      = '#fff1ac';
var gNotYetIssueBgColor = '#f0f5f7';
