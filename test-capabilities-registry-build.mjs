// Pure-function tests for capabilities-registry-build.mjs.
// No live Document Server involved; this only tests the FORMATTING/GATE half.

import { isGenuineFunctionSource, buildMethodsRegistry, buildBuildFingerprint, buildRegistry } from './capabilities-registry-build.mjs'

let ok = 0
let bukas = 0
function check(label, cond, detail) {
  if (cond) { ok += 1; console.log('  ok   ', label) }
  else { bukas += 1; console.log('  BUKAS', label, detail !== undefined ? '-- ' + JSON.stringify(detail) : '') }
}

const REAL_MOVE = 'function(before,after){var bb=before instanceof ApiWorksheet;var ba=after instanceof ApiWorksheet;if(bb&&ba||!bb&&!ba)throwException(new Error("Incorrect parametrs."));else{var curIndex=this.GetIndex();var newIndex=bb?before.GetIndex():after.GetIndex()+1;this.worksheet.workbook.oApi.asc_moveWorksheet(newIndex,[curIndex])}}'

console.log('\n[1] isGenuineFunctionSource -- *** A KAPU MASIK IRANYA: EGY VALODI reflektalt szignatura ATMEGY ***')
{
  check('  a MA elo reflexioval mert Move-szignatura verified-nek szamit', isGenuineFunctionSource(REAL_MOVE) === true)
}

console.log('\n[2] isGenuineFunctionSource -- *** A KAPU: EGY NEM LETEZO METODUS reflexioja ("undefined") NEM verified ***')
{
  check('  "undefined" (a String(fn) kimenete nem letezo tulajdonsagra) -> FALSE', isGenuineFunctionSource('undefined') === false)
  check('  "null" -> FALSE', isGenuineFunctionSource('null') === false)
  check('  ures string -> FALSE', isGenuineFunctionSource('') === false)
  check('  csak whitespace -> FALSE', isGenuineFunctionSource('   ') === false)
  check('  nem-string (pl. undefined ertek maga) -> FALSE, nem dob', isGenuineFunctionSource(undefined) === false)
  check('  szam -> FALSE', isGenuineFunctionSource(42) === false)
}

console.log('\n[3] isGenuineFunctionSource -- egy NEM function-alaku, de nem-ures string (pl. "[native code]") sem szamit verifikaltnak')
{
  check('  natv-kod placeholder -> FALSE (ha az engine ilyet adna, ne higgyuk el vakon)', isGenuineFunctionSource('[native code]') === false)
}

console.log('\n[4] buildMethodsRegistry -- *** A KARTYA SAJAT MOTIVALO ESETE: egy nem-letezo/eltavolitott metodus NEM kap verified:true ***')
{
  const reg = buildMethodsRegistry({
    'ApiWorksheet.Move': REAL_MOVE,
    'ApiWorksheet.NemLetezoMetodusXyz': 'undefined',
  }, '2026-08-17T21:00:00Z')
  check('  Move verified=true, verifiedAt kitoltve', reg['ApiWorksheet.Move'].verified === true && reg['ApiWorksheet.Move'].verifiedAt === '2026-08-17T21:00:00Z')
  check('  a nem-letezo metodus verified=FALSE', reg['ApiWorksheet.NemLetezoMetodusXyz'].verified === false)
  check('  a nem-letezo metodus verifiedAt=null (nem a hivas idopontja -- az FELTETELEZNE, hogy mertunk valamit)', reg['ApiWorksheet.NemLetezoMetodusXyz'].verifiedAt === null)
  check('  a nem-verifikalt bejegyzes signature mezeje MEGORZI a nyers stringet (nem tunik el, auditalhato)', reg['ApiWorksheet.NemLetezoMetodusXyz'].signature === 'undefined')
}

console.log('\n[5] buildMethodsRegistry -- ures/hianyzo bemenetek')
{
  let threw = false
  try { buildMethodsRegistry(null, '2026-08-17T21:00:00Z') } catch { threw = true }
  check('  null rawSignatures -> dob (nem csendben ures regisztert ad)', threw)

  let threw2 = false
  try { buildMethodsRegistry({ x: REAL_MOVE }, null) } catch { threw2 = true }
  check('  hianyzo nowIso -> dob (a hivo felelossege az idobelyeg, nem talalhato ki itt)', threw2)

  check('  ures objektum -> ures regiszter, nem dob', Object.keys(buildMethodsRegistry({}, '2026-08-17T21:00:00Z')).length === 0)
}

console.log('\n[6] buildBuildFingerprint -- MIND AZ OT KOTELEZO MEZO, a mai NEM-MERT allapot')
{
  const fields = { packageVersion: 'NEM-MERT', imageId: 'NEM-MERT', wordSdkHash: 'NEM-MERT', cellSdkHash: 'NEM-MERT', slideSdkHash: 'NEM-MERT' }
  const reasons = {
    packageVersion: 'nincs hoszt-ut a build-hoszthoz (E1)', imageId: 'nincs hoszt-ut a build-hoszthoz (E1)',
    wordSdkHash: 'egress-kapu, owner-dontesre var (E1)', cellSdkHash: 'egress-kapu, owner-dontesre var (E1)',
    slideSdkHash: 'egress-kapu, owner-dontesre var (E1)',
  }
  const out = buildBuildFingerprint(fields, reasons)
  check('  mind az ot mezo jelen van', Object.keys(out).length === 5)
  check('  ertekuk a NEM-MERT literal, nem null', out.packageVersion === 'NEM-MERT' && out.slideSdkHash === 'NEM-MERT')
}

console.log('\n[7] buildBuildFingerprint -- *** A KAPU: NEM-MERT mezo OK NELKUL dob (nem csendes null) ***')
{
  const fields = { packageVersion: 'NEM-MERT', imageId: 'x', wordSdkHash: 'x', cellSdkHash: 'x', slideSdkHash: 'x' }
  let threw = false
  try { buildBuildFingerprint(fields, {}) } catch { threw = true }
  check('  NEM-MERT ok nelkul -> dob', threw)
}

console.log('\n[8] buildBuildFingerprint -- hianyzo mezo (SEMA-HIBA) dob, nem csendben kimarad')
{
  const fields = { packageVersion: 'x', imageId: 'x', wordSdkHash: 'x', cellSdkHash: 'x' } // slideSdkHash hianyzik
  let threw = false
  try { buildBuildFingerprint(fields, {}) } catch { threw = true }
  check('  hianyzo mezo -> dob', threw)
}

console.log('\n[9] buildRegistry -- teljes osszeallitas, a VEGSO alak')
{
  const reg = buildRegistry({
    rawSignatures: { 'ApiWorksheet.Move': REAL_MOVE },
    buildFields: { packageVersion: 'NEM-MERT', imageId: 'NEM-MERT', wordSdkHash: 'NEM-MERT', cellSdkHash: 'NEM-MERT', slideSdkHash: 'NEM-MERT' },
    nemMertOkok: { packageVersion: 'ok1', imageId: 'ok2', wordSdkHash: 'ok3', cellSdkHash: 'ok4', slideSdkHash: 'ok5' },
    nowIso: '2026-08-17T21:00:00Z',
    generatedBy: 'fixture-generator',
    source: 'live reflection, run_builder_script',
  })
  check('  generatedAt/generatedBy/source jelen', reg.generatedAt === '2026-08-17T21:00:00Z' && reg.generatedBy === 'fixture-generator' && reg.source.includes('reflection'))
  check('  build ES methods egyutt all', reg.build.packageVersion === 'NEM-MERT' && reg.methods['ApiWorksheet.Move'].verified === true)
  check('  buildNemMertOkok athozva', reg.buildNemMertOkok.packageVersion === 'ok1')
}

console.log(`\nellenorzesek: ${ok} ok, ${bukas} bukas`)
if (bukas) { process.exit(1) }
