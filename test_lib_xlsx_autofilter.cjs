// Tests xlsxAutoFilterRangeInSheetXml /
// xlsxRequestsAutoFilterVerification / xlsxAutoFilterVerificationReport /
// applyAutoFilterVerificationToReport, the pure functions the post-write autoFilter "third case"
// check is built on. Pure-function coverage (no browser, no Document Server, no network). Mirrors
// test_lib_xlsx_cells.cjs's shape for the sibling (cell) check.
//
// The "third case" this covers (19c1c83's own commit
// message): a DIFFERENT range already carries a filter -> the runtime guard throws INSIDE the
// generated Document Server script, the original element survives untouched, BUT the throw happens
// AFTER buildCoeditScript's own client-side report already marked the operation 'alkalmazva' --
// so a caller sees a success report for something that did not happen.

const {
  xlsxAutoFilterRangeInSheetXml,
  xlsxRequestsAutoFilterVerification,
  xlsxAutoFilterVerificationReport,
  applyAutoFilterVerificationToReport,
} = require('./lib.cjs')

let osszes = 0
const hibak = []
function check(cimke, felteteles, reszlet = '') {
  osszes += 1
  if (felteteles) { console.log(`  ok    ${cimke}`) } else { hibak.push(cimke); console.log(`  BUKAS ${cimke}${reszlet ? ' -- ' + reszlet : ''}`) }
}

const SHEET_WITH_A1C10 = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><autoFilter ref="A1:C10"/></worksheet>'
const SHEET_NO_FILTER = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'

console.log('\n[1] xlsxAutoFilterRangeInSheetXml -- jelen levo <autoFilter ref="..."> visszaadja a tartomanyt')
{
  check('  A1:C10 kiolvasva', xlsxAutoFilterRangeInSheetXml(SHEET_WITH_A1C10) === 'A1:C10')
}

console.log('\n[2] xlsxAutoFilterRangeInSheetXml -- nincs <autoFilter> elem -> null (nem ures string)')
{
  check('  null, nem ""', xlsxAutoFilterRangeInSheetXml(SHEET_NO_FILTER) === null)
}

console.log('\n[3] xlsxAutoFilterRangeInSheetXml -- ures/null/nem-string bemenet: null, nem dobas')
{
  check('  ures string -> null', xlsxAutoFilterRangeInSheetXml('') === null)
  check('  null -> null, nem dob', xlsxAutoFilterRangeInSheetXml(null) === null)
  check('  undefined -> null, nem dob', xlsxAutoFilterRangeInSheetXml(undefined) === null)
}

console.log('\n[4] xlsxRequestsAutoFilterVerification -- diszpecser: csak autoFilter+range szamit')
{
  check('  autoFilter + range -> true', xlsxRequestsAutoFilterVerification({ type: 'autoFilter', range: 'A1:C1' }) === true)
  check('  autoFilter range nelkul -> false (a validator ugyis megtagadja, de itt nem dob)', xlsxRequestsAutoFilterVerification({ type: 'autoFilter' }) === false)
  check('  mas tipus -> false', xlsxRequestsAutoFilterVerification({ type: 'fillColor', range: 'A1:C1' }) === false)
  check('  null/undefined -> false, nem dob', xlsxRequestsAutoFilterVerification(null) === false && xlsxRequestsAutoFilterVerification(undefined) === false)
}

console.log('\n[5] xlsxAutoFilterVerificationReport -- SIKERES eset (ugyanaz a tartomany a mentett lapon): egyezik=true, NEG. KONTROLL a kartya sajat kikotesehez ("a sikeres eset TOVABBRA IS sikert jelent")')
{
  const out = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'A1:C10' }], SHEET_WITH_A1C10)
  check('  egy bejegyzes', out.length === 1)
  check('  egyezik=true, POZITIV allitas', out[0].egyezik === true)
  check('  tenylegesRange a mentett ertek', out[0].tenylegesRange === 'A1:C10')
  check('  nemMertIndok null (ez nem NEM-MERT eset)', out[0].nemMertIndok === null)
}

console.log('\n[6] xlsxAutoFilterVerificationReport -- *** A HARMADIK ESET ***: mas tartomany van szurve a mentett lapon -> egyezik=FALSE, tenylegesRange megnevezi a valodit')
{
  const out = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'D1:F5' }], SHEET_WITH_A1C10)
  check('  egyezik=FALSE', out[0].egyezik === false)
  check('  tenylegesRange a VALODI (A1:C10), nem a kert', out[0].tenylegesRange === 'A1:C10')
  check('  kertRange a KERT (D1:F5)', out[0].kertRange === 'D1:F5')
}

console.log('\n[7] xlsxAutoFilterVerificationReport -- a mentett lapon EGYALTALAN nincs autoFilter (pl. egy MASIK ok miatt sem jott letre) -> egyezik=FALSE, tenylegesRange=null')
{
  const out = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'A1:C10' }], SHEET_NO_FILTER)
  check('  egyezik=FALSE', out[0].egyezik === false)
  check('  tenylegesRange=null', out[0].tenylegesRange === null)
}

console.log('\n[7b] xlsxAutoFilterVerificationReport -- *** korabbanMarOtt: a HARMADIK, opcionalis parameter ***')
{
  const noBefore = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'A1:C10' }], SHEET_WITH_A1C10)
  check('  a parameter nelkul (2 argumentumos legacy hivas) korabbanMarOtt=null -- nem talal ki adatot', noBefore[0].korabbanMarOtt === null)

  const wasThere = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'A1:C10' }], SHEET_WITH_A1C10, SHEET_WITH_A1C10)
  check('  ELOTTE IS A1:C10 volt szurve, UTANA IS -> korabbanMarOtt=true (idempotens no-op eset)', wasThere[0].korabbanMarOtt === true)

  const wasNotThere = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'A1:C10' }], SHEET_WITH_A1C10, SHEET_NO_FILTER)
  check('  ELOTTE nem volt szuro, UTANA mar A1:C10 -> korabbanMarOtt=false (ez a hivas hozta letre)', wasNotThere[0].korabbanMarOtt === false)

  const beforeReadFailed = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'A1:C10' }], SHEET_WITH_A1C10, null)
  check('  az ELOTTE-lap kicsomagolasa hibazott (null) -> korabbanMarOtt=null, NEM false (a hianyzo bizonyitek nem allitas)', beforeReadFailed[0].korabbanMarOtt === null)
}

console.log('\n[8] xlsxAutoFilterVerificationReport -- sheetXml=null (a visszaolvasas HIBAZOTT): NEM-MERT (egyezik=null), NEM "sikertelen"')
{
  const out = xlsxAutoFilterVerificationReport([{ type: 'autoFilter', range: 'A1:C10' }], null)
  check('  egyezik NULL (nem false)', out[0].egyezik === null)
  check('  tenylegesRange NULL', out[0].tenylegesRange === null)
  check('  nemMertIndok nevesitve', typeof out[0].nemMertIndok === 'string' && out[0].nemMertIndok.length > 0)
}

console.log('\n[9] xlsxAutoFilterVerificationReport -- perOp tomb (kulon lap muveletenkent), tobb autoFilter op keverve mas tipusokkal (3. pont: soronkent, nem osszesitve)')
{
  const operations = [
    { type: 'fillColor', at: 'A1', color: '#fff' },
    { type: 'autoFilter', range: 'A1:C10' }, // ez SIKERES lesz
    { type: 'autoFilter', range: 'X1:X5' },  // ez a HARMADIK ESET
  ]
  const perOp = [null, SHEET_WITH_A1C10, SHEET_WITH_A1C10]
  const out = xlsxAutoFilterVerificationReport(operations, perOp)
  check('  ket bejegyzes (a fillColor nem autoFilter, kimarad)', out.length === 2)
  check('  0. bejegyzes az 1. indexu op (autoFilter A1:C10), egyezik=true', out[0].index === 1 && out[0].egyezik === true)
  check('  1. bejegyzes a 2. indexu op (autoFilter X1:X5), egyezik=FALSE -- a masodik hibaja nem torolte az elso jelzeset', out[1].index === 2 && out[1].egyezik === false)
}

console.log('\n[10] xlsxAutoFilterVerificationReport -- ha EGY operacio sem autoFilter+range, a valasz NULL (nem ures lista -- "nem alkalmazhato", nem "0 hiany")')
{
  check('  csak fillColor -> null', xlsxAutoFilterVerificationReport([{ type: 'fillColor', at: 'A1', color: '#fff' }], SHEET_WITH_A1C10) === null)
  check('  autoFilter range nelkul -> null', xlsxAutoFilterVerificationReport([{ type: 'autoFilter' }], SHEET_WITH_A1C10) === null)
  check('  ures operacio-lista -> null', xlsxAutoFilterVerificationReport([], SHEET_WITH_A1C10) === null)
}

console.log('\n[11] applyAutoFilterVerificationToReport -- *** A KARTYA SAJAT KOVETELMENYE ***: a HARMADIK ESET a reportban is megjelenik, nevesitett hibakent, nem "alkalmazva"-kent')
{
  const report = [{ index: 0, type: 'autoFilter', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  const verification = [{ index: 0, type: 'autoFilter', kertRange: 'D1:F5', tenylegesRange: 'A1:C10', egyezik: false, korabbanMarOtt: null, nemMertIndok: null }]
  const out = applyAutoFilterVerificationToReport(report, verification)
  check('  ugyanaz a tomb-referencia (in-place mutacio, mint rejectCoeditBatch)', out === report)
  check('  *** outcome MAR NEM "alkalmazva" *** -- es MOSTANTOL "megtagadva" (nem tovabb "hiba")', report[0].outcome === 'megtagadva')
  check('  reason nevesiti MINDKET tartomanyt (kert ES tenyleges)', report[0].reason.includes('D1:F5') && report[0].reason.includes('A1:C10'))
}

console.log('\n[11b] applyAutoFilterVerificationToReport -- *** A HAROM-ALLAPOTU MODELL: vegrehajtva vs kihagyva-idempotens ***')
{
  const vegrehajtva = [{ index: 0, type: 'autoFilter', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  applyAutoFilterVerificationToReport(vegrehajtva, [{ index: 0, type: 'autoFilter', kertRange: 'A1:C10', tenylegesRange: 'A1:C10', egyezik: true, korabbanMarOtt: false, nemMertIndok: null }])
  check('  korabbanMarOtt=false (a hivas ELOTT meg nem volt ott) -> "vegrehajtva"', vegrehajtva[0].outcome === 'vegrehajtva')

  const kihagyva = [{ index: 0, type: 'autoFilter', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  applyAutoFilterVerificationToReport(kihagyva, [{ index: 0, type: 'autoFilter', kertRange: 'A1:C10', tenylegesRange: 'A1:C10', egyezik: true, korabbanMarOtt: true, nemMertIndok: null }])
  check('  korabbanMarOtt=true (mar ott volt a hivas ELOTT is) -> "kihagyva-idempotens"', kihagyva[0].outcome === 'kihagyva-idempotens')
  check('  a kihagyva-idempotens reason nevesiti, hogy biztonsagos no-op volt', /biztonsagos no-op/.test(kihagyva[0].reason))

  const nincsElotteAdat = [{ index: 0, type: 'autoFilter', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  applyAutoFilterVerificationToReport(nincsElotteAdat, [{ index: 0, type: 'autoFilter', kertRange: 'A1:C10', tenylegesRange: 'A1:C10', egyezik: true, korabbanMarOtt: null, nemMertIndok: null }])
  check('  korabbanMarOtt=null (nincs elotte-pillanatkep) -> a regi ket-allapotu viselkedes, "elkuldve-nem-verifikalt" marad, NEM talal ki harmadik allapotot', nincsElotteAdat[0].outcome === 'elkuldve-nem-verifikalt')
}

console.log('\n[12] applyAutoFilterVerificationToReport -- NEG. KONTROLL: SIKERES eset (egyezik=true) NEM valtoztatja a reportot -- "a sikeres eset TOVABBRA IS sikert jelent" (kartya MI A KESZ)')
{
  const report = [{ index: 0, type: 'autoFilter', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  const verification = [{ index: 0, type: 'autoFilter', kertRange: 'A1:C10', tenylegesRange: 'A1:C10', egyezik: true, nemMertIndok: null }]
  applyAutoFilterVerificationToReport(report, verification)
  check('  outcome VALTOZATLAN "elkuldve-nem-verifikalt"', report[0].outcome === 'elkuldve-nem-verifikalt')
  check('  reason VALTOZATLAN null', report[0].reason === null)
}

console.log('\n[13] applyAutoFilterVerificationToReport -- egyezik=null (NEM-MERT, a visszaolvasas hibazott) SEM valtoztatja a reportot -- a hianyzo bizonyitek nem allitas a hibarol')
{
  const report = [{ index: 0, type: 'autoFilter', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null }]
  const verification = [{ index: 0, type: 'autoFilter', kertRange: 'A1:C10', tenylegesRange: null, egyezik: null, nemMertIndok: 'a lap nem volt kicsomagolhato' }]
  applyAutoFilterVerificationToReport(report, verification)
  check('  outcome VALTOZATLAN "elkuldve-nem-verifikalt"', report[0].outcome === 'elkuldve-nem-verifikalt')
}

console.log('\n[14] applyAutoFilterVerificationToReport -- csak a NEVESITETT indexen valtoztat, a tobbi report-bejegyzest erintetlenul hagyja (tobb-muveletes koteg)')
{
  const report = [
    { index: 0, type: 'fillColor', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null },
    { index: 1, type: 'autoFilter', outcome: 'elkuldve-nem-verifikalt', sourceRoute: 'coedit', reason: null },
  ]
  const verification = [{ index: 1, type: 'autoFilter', kertRange: 'X1:X5', tenylegesRange: 'A1:C10', egyezik: false, korabbanMarOtt: null, nemMertIndok: null }]
  applyAutoFilterVerificationToReport(report, verification)
  check('  0. bejegyzes (fillColor) erintetlen', report[0].outcome === 'elkuldve-nem-verifikalt' && report[0].reason === null)
  check('  1. bejegyzes (autoFilter) javitva', report[1].outcome === 'megtagadva')
}

console.log('\n[15] applyAutoFilterVerificationToReport -- ha az outcome MAR NEM "alkalmazva" (pl. a koteg mar elutasitva rejectCoeditBatch altal), NEM irja felul -- nem masodik hibauzenetet ir egy mar megnevezett hiba fole')
{
  const report = [{ index: 0, type: 'autoFilter', outcome: 'nem-alkalmazva', sourceRoute: 'coedit', reason: 'mar megnevezve egy masik okbol' }]
  const verification = [{ index: 0, type: 'autoFilter', kertRange: 'A1:C10', tenylegesRange: 'D1:D1', egyezik: false, nemMertIndok: null }]
  applyAutoFilterVerificationToReport(report, verification)
  check('  outcome es reason VALTOZATLAN (mar volt egy nevesitett ok)', report[0].outcome === 'nem-alkalmazva' && report[0].reason === 'mar megnevezve egy masik okbol')
}

console.log('\n[16] applyAutoFilterVerificationToReport -- ures/null bemenetek: nem dob, visszaadja a reportot valtozatlanul')
{
  check('  report null -> null visszaadva, nem dob', applyAutoFilterVerificationToReport(null, []) === null)
  check('  verification null -> a report valtozatlanul visszaadva', applyAutoFilterVerificationToReport([], null) !== undefined)
}

console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
if (hibak.length) { console.log('BUKOTT:', hibak.join(' | ')); process.exit(1) }
