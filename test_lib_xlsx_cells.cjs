// Tests cellRefsPresentInSheetXml / xlsxTableCellRefs /
// xlsxRequestedCellRefs, the pure functions the post-write cell-presence check is built on. Pure
// -function coverage (no browser, no Document Server, no network). The sheet1.xml fixture in [1]
// is NOT hand-written: it is the VERBATIM output of a live, disposable Document Server probe run
// today against buildCreateScript's `table`/`formula` xlsx operations (op batch: table F2, a
// syntactically-"=" but semantically-broken formula at E3, table F3) -- E3 is measured absent,
// not assumed absent. See the full probe and the emit()-sharing argument for why this also bears on the
// coedit (already-open-document) route, which this fixture did not itself exercise.

const { cellRefsPresentInSheetXml, xlsxTableCellRefs, xlsxRequestedCellRefs, xlsxCellVerificationReport, applyCellVerificationToReport, OPERATIONS } = require('./lib.cjs')

let osszes = 0
const hibak = []
function check(cimke, felteteles, reszlet = '') {
  osszes += 1
  if (felteteles) { console.log(`  ok    ${cimke}`) } else { hibak.push(cimke); console.log(`  BUKAS ${cimke}${reszlet ? ' -- ' + reszlet : ''}`) }
}

const ELO_SHEET1_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView zoomScale="100" workbookViewId="0"><selection activeCell="A1" activeCellId="0" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="2"><c r="F2" s="1" t="s"><v>0</v></c></row><row r="3"><c r="F3" s="1" t="s"><v>1</v></c></row></sheetData></worksheet>'

console.log('\n[1] cellRefsPresentInSheetXml -- ELO probaból mentett sheet1.xml (table F2, formula E3=torott, table F3)')
{
  const refs = cellRefsPresentInSheetXml(ELO_SHEET1_XML)
  check('  F2 jelen van', refs.has('F2'))
  check('  F3 jelen van', refs.has('F3'))
  check('  *** E3 HIANYZIK -- a torott keplet nem hozott letre cellat ***', !refs.has('E3'))
  check('  pontosan ket cella all a lapon (nincs kitalalt harmadik)', refs.size === 2)
}

console.log('\n[2] cellRefsPresentInSheetXml -- ONMAGAT-ZARO <c> elem is JELEN LEVO cella (stilus/tipus ertek nelkul, nem hianyzo)')
{
  const xml = '<sheetData><row r="1"><c r="A1" s="2" t="n"/></row></sheetData>'
  const refs = cellRefsPresentInSheetXml(xml)
  check('  onmagat-zaro cella is szamit', refs.has('A1'))
}

console.log('\n[3] cellRefsPresentInSheetXml -- ures/null bemenet: ures halmaz, nem dobas')
{
  check('  ures string -> ures halmaz', cellRefsPresentInSheetXml('').size === 0)
  check('  null -> ures halmaz, nem dob', cellRefsPresentInSheetXml(null).size === 0)
}

console.log('\n[4] xlsxTableCellRefs -- egysoros tabla, alapertelmezett A1-tol')
{
  const refs = xlsxTableCellRefs({ rows: [['a', 'b', 'c']] })
  check('  harom cella, A1/B1/C1', JSON.stringify(refs) === JSON.stringify(['A1', 'B1', 'C1']))
}

console.log('\n[5] xlsxTableCellRefs -- tobbsoros tabla, elt tolt `at`-tal')
{
  const refs = xlsxTableCellRefs({ at: 'F2', rows: [['fejlec'], ['sor2']] })
  check('  F2 es F3, sorrendben', JSON.stringify(refs) === JSON.stringify(['F2', 'F3']))
}

console.log('\n[6] xlsxTableCellRefs -- rongyos sorok (elteru oszlopszam), es ures `rows` -> ures lista')
{
  const refs = xlsxTableCellRefs({ at: 'B1', rows: [['x', 'y'], ['z']] })
  check('  minden sor a SAJAT hosszaval szamol', JSON.stringify(refs) === JSON.stringify(['B1', 'C1', 'B2']))
  check('  ures rows -> ures lista, nem dob', xlsxTableCellRefs({ rows: [] }).length === 0)
  check('  hianyzo rows -> ures lista', xlsxTableCellRefs({}).length === 0)
}

console.log('\n[7] xlsxTableCellRefs -- UGYANAZ, amit OPERATIONS.table.emit() ténylegesen ir (nincs drift a ket oldal kozott)')
{
  const op = { type: 'table', at: 'D3', rows: [['h1', 'h2'], ['v1', 'v2'], ['v3', 'v4']] }
  const refs = xlsxTableCellRefs(op)
  const script = OPERATIONS.table.emit(op, 'xlsx').join('\n')
  const emittedRefs = refs.map((r) => new RegExp(`GetRange\\("${r}"\\)\\.SetValue`).test(script))
  check('  a xlsxTableCellRefs altal adott MINDEN ref valoban SetValue-t kap a generalt scriptben', emittedRefs.every(Boolean))
  check('  hat cella (2x3)', refs.length === 6)
}

console.log('\n[8] xlsxRequestedCellRefs -- diszpecser: table/formula elfogva, minden mas ures (NEM verifikalt, nem "ures a valasz")')
{
  check('  table op -> xlsxTableCellRefs-szel egyezo', JSON.stringify(xlsxRequestedCellRefs({ type: 'table', at: 'A1', rows: [['x']] })) === JSON.stringify(['A1']))
  check('  formula op -> egyetlen cella, az `at`', JSON.stringify(xlsxRequestedCellRefs({ type: 'formula', at: 'E3', formula: '=1' })) === JSON.stringify(['E3']))
  check('  formula op `at` nelkul -> ures (a validator ugyis megtagadja, de itt nem dob)', xlsxRequestedCellRefs({ type: 'formula', formula: '=1' }).length === 0)
  check('  ismeretlen tipus -> ures lista', xlsxRequestedCellRefs({ type: 'fillColor', at: 'A1' }).length === 0)
  check('  null/undefined -> ures lista, nem dob', xlsxRequestedCellRefs(null).length === 0 && xlsxRequestedCellRefs(undefined).length === 0)
}

console.log('\n[9] xlsxCellVerificationReport -- a KARTYA SAJAT VEGYES KOTEGE: 3 ervenyes + 1 ervenytelen, KULON jelezve (3. pont)')
{
  // A kartya pontosan ezt a kotegformat nevezi meg (94415/12553): table F2 = ELOTTE, formula E3
  // torott, table F3 = UTANA. Ugyanaz az op-sorrend, mint az elo probaban.
  const operations = [
    { type: 'table', at: 'F2', rows: [['ELOTTE']] },
    { type: 'formula', at: 'E3', formula: '=SUM(NEMLETEZO!A1:A9' },
    { type: 'table', at: 'F3', rows: [['UTANA']] },
  ]
  const out = xlsxCellVerificationReport(operations, ELO_SHEET1_XML)
  check('  harom bejegyzes, EGY-EGY az erintett muveletenkent (nem egyetlen osszesitett bool)', out.length === 3)
  check('  0. index (table F2): mindLetrejott=true, hianyzoCellak ures', out[0].mindLetrejott === true && out[0].hianyzoCellak.length === 0)
  check('  *** 1. index (formula E3): mindLetrejott=FALSE, hianyzoCellak MEGNEVEZI E3-at ***', out[1].mindLetrejott === false && JSON.stringify(out[1].hianyzoCellak) === JSON.stringify(['E3']))
  check('  2. index (table F3): mindLetrejott=true -- a 2. hibaja NEM torte meg a 3. iras jelzeset', out[2].mindLetrejott === true && out[2].hianyzoCellak.length === 0)
  check('  a "NEM alkalmazva" mezo (nemMertIndok) minden sikeres/sikertelen sornal null (ez nem NEM-MERT eset)', out.every((r) => r.nemMertIndok === null))
}

console.log('\n[10] xlsxCellVerificationReport -- sheetXml=null (a visszaolvasas HIBAZOTT): NEM-MERT, NEM "sikertelen" (5. pont)')
{
  const operations = [{ type: 'formula', at: 'A1', formula: '=1' }]
  const out = xlsxCellVerificationReport(operations, null)
  check('  egy bejegyzes', out.length === 1)
  check('  mindLetrejott NULL (nem false -- a keplet lehet, hogy landolt, csak nem tudjuk)', out[0].mindLetrejott === null)
  check('  hianyzoCellak NULL (nem ures lista -- azt jelentene, hogy MINDEN cella igazolva van)', out[0].hianyzoCellak === null)
  check('  nemMertIndok nevesitve, nem ures', typeof out[0].nemMertIndok === 'string' && out[0].nemMertIndok.length > 0)
}

console.log('\n[11] xlsxCellVerificationReport -- ha EGY operacio sem celoz cellat, a valasz NULL (nem ures lista -- "nem alkalmazhato", nem "0 hiany")')
{
  check('  csak fillColor -> null', xlsxCellVerificationReport([{ type: 'fillColor', at: 'A1', color: '#fff' }], ELO_SHEET1_XML) === null)
  check('  ures operacio-lista -> null', xlsxCellVerificationReport([], ELO_SHEET1_XML) === null)
}

console.log('\n[12] xlsxCellVerificationReport -- EGY ervenyes cella-iras: a valasz POZITIVAN allitja, hogy LETREJOTT (1. pont)')
{
  const out = xlsxCellVerificationReport([{ type: 'table', at: 'F2', rows: [['ELOTTE']] }], ELO_SHEET1_XML)
  check('  mindLetrejott=true, POZITIV allitas, nem csak "nincs hiba"', out[0].mindLetrejott === true)
}

console.log('\n[13] applyCellVerificationToReport -- *** a cella-hiany a REPORTBA is bekerul, nem csak a cellaEllenorzes oldalcsatornaba ***')
{
  const report = [{ index: 0, type: 'formula', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  const verification = [{ index: 0, type: 'formula', kertCellak: ['E3'], hianyzoCellak: ['E3'], mindLetrejott: false, nemMertIndok: null }]
  const out = applyCellVerificationToReport(report, verification)
  check('  ugyanaz a tomb-referencia (in-place mutacio)', out === report)
  check('  *** outcome MAR NEM "alkalmazva" -- "megtagadva" ***', report[0].outcome === 'megtagadva')
  check('  reason nevesiti a hianyzo cellat', report[0].reason.includes('E3'))
}

console.log('\n[14] applyCellVerificationToReport -- NEG. KONTROLL: mindLetrejott=true -> "vegrehajtva", a sikeres eset tovabbra is sikert jelent')
{
  const report = [{ index: 0, type: 'table', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  const verification = [{ index: 0, type: 'table', kertCellak: ['A1'], hianyzoCellak: [], mindLetrejott: true, nemMertIndok: null }]
  applyCellVerificationToReport(report, verification)
  check('  mindLetrejott=true -> "vegrehajtva"', report[0].outcome === 'vegrehajtva')
  check('  reason valtozatlan null (nincs mit nevesiteni egy sikeres esetnel)', report[0].reason === null)
}

console.log('\n[15] applyCellVerificationToReport -- mindLetrejott=null (a visszaolvasas hibazott) NEM valtoztatja a reportot')
{
  const report = [{ index: 0, type: 'formula', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  const verification = [{ index: 0, type: 'formula', kertCellak: ['A1'], hianyzoCellak: null, mindLetrejott: null, nemMertIndok: 'nem kicsomagolhato' }]
  applyCellVerificationToReport(report, verification)
  check('  outcome VALTOZATLAN "elkuldve-nem-verifikalt"', report[0].outcome === 'elkuldve-nem-verifikalt')
}

console.log('\n[16] applyCellVerificationToReport -- csak a NEVESITETT indexen valtoztat, tobb-muveletes kotegben')
{
  const report = [
    { index: 0, type: 'fillColor', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null },
    { index: 1, type: 'formula', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null },
  ]
  const verification = [{ index: 1, type: 'formula', kertCellak: ['E3'], hianyzoCellak: ['E3'], mindLetrejott: false, nemMertIndok: null }]
  applyCellVerificationToReport(report, verification)
  check('  0. bejegyzes (fillColor) erintetlen', report[0].outcome === 'elkuldve-nem-verifikalt' && report[0].reason === null)
  check('  1. bejegyzes (formula) javitva', report[1].outcome === 'megtagadva')
}

console.log('\n[17] applyCellVerificationToReport -- ha az outcome MAR NEM "alkalmazva", NEM irja felul')
{
  const report = [{ index: 0, type: 'formula', outcome: 'nem-alkalmazva', sourceRoute: 'coedit', reason: 'mar megnevezve egy masik okbol' }]
  const verification = [{ index: 0, type: 'formula', kertCellak: ['E3'], hianyzoCellak: ['E3'], mindLetrejott: false, nemMertIndok: null }]
  applyCellVerificationToReport(report, verification)
  check('  outcome es reason VALTOZATLAN', report[0].outcome === 'nem-alkalmazva' && report[0].reason === 'mar megnevezve egy masik okbol')
}

console.log('\n[18] applyCellVerificationToReport -- ures/null bemenetek: nem dob')
{
  check('  report null -> null visszaadva, nem dob', applyCellVerificationToReport(null, []) === null)
  check('  verification null -> a report valtozatlanul visszaadva', applyCellVerificationToReport([], null) !== undefined)
}

console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
if (hibak.length) { console.log('BUKOTT:', hibak.join(' | ')); process.exit(1) }
