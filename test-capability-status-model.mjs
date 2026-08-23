// Pure-function tests for capability-status-model.mjs.

import { STATUS, sixConditionsMet, ceilingFor, classify, comparable, flipDetectable } from './capability-status-model.mjs'

let ok = 0
let bukas = 0
function check(label, cond, detail) {
  if (cond) { ok += 1; console.log('  ok   ', label) }
  else { bukas += 1; console.log('  BUKAS', label, detail !== undefined ? '-- ' + JSON.stringify(detail) : '') }
}

const FULL_EVIDENCE_WITH_FINGERPRINT = {
  minimalReproducerRan: true, runtimeVerifiedSignature: true, repeatCount: 3,
  sameSdkFingerprint: true, controlOperationsRun: true, separateFixture: true,
  sdkFingerprintMeasured: true,
}

console.log('\n[1] sixConditionsMet -- *** A KAPU MASIK IRANYA: mind a hat feltetel teljesul -> met=true ***')
{
  const r = sixConditionsMet(FULL_EVIDENCE_WITH_FINGERPRINT)
  check('  met=true, nincs hianyzo', r.met === true && r.missing.length === 0)
}

console.log('\n[2] sixConditionsMet -- EGYETLEN hianyzo feltetel is met=false (the owner: "egyetlen furcsa eredmeny utan a maximum SUSPECT")')
{
  const only2repeats = { ...FULL_EVIDENCE_WITH_FINGERPRINT, repeatCount: 2 }
  const r = sixConditionsMet(only2repeats)
  check('  2 ismetles < 3 -> met=false, nevesitve', r.met === false && r.missing.includes('repeatCount>=3'))

  const noControl = { ...FULL_EVIDENCE_WITH_FINGERPRINT, controlOperationsRun: false }
  check('  kontrollmuvelet nelkul -> met=false', sixConditionsMet(noControl).met === false)

  const empty = sixConditionsMet({})
  check('  ures evidence -> MIND az ot kotelezo feltetel hianyzik (a 6. csak ajanlott)', empty.missing.length === 5)
}

console.log('\n[3] sixConditionsMet -- a "kulon fixture" AJANLOTT, nem kotelezo (the owner: "lehetoleg")')
{
  const noSeparateFixture = { ...FULL_EVIDENCE_WITH_FINGERPRINT, separateFixture: false }
  const r = sixConditionsMet(noSeparateFixture)
  check('  a tobbi ot feltetel teljesul -> met=true MEG separateFixture NELKUL is', r.met === true)
  check('  DE a hianya jelolve van (nem elnyelve csendben)', r.separateFixtureRecommendedButMissing === true)
}

console.log('\n[4] ceilingFor -- *** A KARTYA SAJAT JAVITASA: sdkFingerprintMeasured=false -> plafon SUSPECT ***')
{
  const r = ceilingFor({ sdkFingerprintMeasured: false })
  check('  maxStatus=SUSPECT', r.maxStatus === STATUS.SUSPECT)
  check('  az ok nevezi az E1-et es a meres commitjat', /E1/.test(r.reason) && /e41978ab/.test(r.reason))
}

console.log('\n[5] ceilingFor -- NEG. KONTROLL: sdkFingerprintMeasured=true -> NINCS plafon EBBOL a szabalybol')
{
  const r = ceilingFor({ sdkFingerprintMeasured: true })
  check('  maxStatus=null (ez a szabaly nem tiltja)', r.maxStatus === null)
}

console.log('\n[6] classify -- *** A MAI VALODI ESET: ApplyLayout, harom ismetlessel megerositve, DE fingerprint NEM-MERT -> SUSPECT, nem ENGINE_BUG ***')
{
  const evidence = { ...FULL_EVIDENCE_WITH_FINGERPRINT, sdkFingerprintMeasured: false, sameSdkFingerprint: false }
  const r = classify(STATUS.ENGINE_BUG, evidence)
  check('  a kert ENGINE_BUG helyett SUSPECT-re esik vissza', r.status === STATUS.SUSPECT && r.downgraded === true)
  check('  az ok nevezi a fingerprint-hianyt', /fingerprint/.test(r.reason))
}

console.log('\n[7] classify -- GREEN: mind a hat feltetel teljesul, fingerprint MERVE -> a kert stabil statusz all')
{
  const r = classify(STATUS.ENGINE_BUG, FULL_EVIDENCE_WITH_FINGERPRINT)
  check('  ENGINE_BUG marad, nincs leminositve', r.status === STATUS.ENGINE_BUG && r.downgraded === false)
}

console.log('\n[8] classify -- JOB_ABORT UGYANIGY kapuzott, mint ENGINE_BUG (mindketto STABLE_NEGATIVE)')
{
  const insufficientEvidence = { sdkFingerprintMeasured: true, sameSdkFingerprint: true, repeatCount: 1 }
  const r = classify(STATUS.JOB_ABORT, insufficientEvidence)
  check('  1 ismetles -> SUSPECT-re esik, JOB_ABORT-ra is', r.status === STATUS.SUSPECT && r.downgraded === true)
}

console.log('\n[9] classify -- *** NEG. KONTROLL: a POZITIV statuszokat ez a szabaly NEM kapuzza *** (the owner\'s 7th point csak a stabil blocked/ENGINE_BUG-rol szol)')
{
  const noEvidence = {}
  check('  SUPPORTED_VERIFIED evidence NELKUL is atmegy -- ez NEM ENGINE_BUG-tipusu allitas', classify(STATUS.SUPPORTED_VERIFIED, noEvidence).status === STATUS.SUPPORTED_VERIFIED)
  check('  NO_OP szinten atmegy kapuzas nelkul', classify(STATUS.NO_OP, noEvidence).status === STATUS.NO_OP)
  check('  REQUIRES_EDITOR_CONTEXT szinten', classify(STATUS.REQUIRES_EDITOR_CONTEXT, noEvidence).status === STATUS.REQUIRES_EDITOR_CONTEXT)
}

console.log('\n[10] comparable -- *** A 3. KIKOTES: csak AZONOS belepesi ponton vethetok ossze ***')
{
  const a = { entryPoint: 'coedit', sdkSha256: 'abc' }
  const b = { entryPoint: 'create_document', sdkSha256: 'abc' }
  const r = comparable(a, b)
  check('  kulonbozo belepesi pont -> NEM osszevethetok', r.comparable === false)
  check('  az ok nevezi mindket belepesi pontot', r.reason.includes('coedit') && r.reason.includes('create_document'))

  const c = { entryPoint: 'coedit', sdkSha256: 'xyz' }
  check('  AZONOS belepesi pont -> osszevethetok (a fingerprint elteresenek itt nincs jelentosege)', comparable(a, c).comparable === true)
}

console.log('\n[11] flipDetectable -- *** (A) A VALTOZAS TENYE megallapithato hianyzo fingerprinttel is (a korabbi felvetes HELYES volt, DE tul szeles alakban) ***')
{
  const before = { entryPoint: 'create_document', sdkSha256: 'NEM-MERT', status: STATUS.NO_OP }
  const after = { entryPoint: 'create_document', sdkSha256: 'NEM-MERT', status: STATUS.SUPPORTED_VERIFIED }
  const r = flipDetectable(before, after)
  check('  flipKnown=true -- a valtozas TENYE all', r.flipKnown === true)
  check('  attributableToBuildChange=false -- az OKA nem attribualhato (E8 FLIP_FINGERPRINT_UNKNOWN mintaja)', r.attributableToBuildChange === false)
  check('  a reason megnevezi a hianyzo attribuciot', /nem attribualhato/.test(r.reason))
}

console.log('\n[12] flipDetectable -- *** (B) EGY STATUSZ NEM ALLITHATJA, HOGY MAS FUTASRA IS AZ ALL, ha a fingerprint hianyzik -- se elore, se visszafele ***')
{
  const same = { entryPoint: 'create_document', sdkSha256: 'NEM-MERT', status: STATUS.SUPPORTED_VERIFIED }
  const r = flipDetectable(same, same)
  check('  ugyanaz a statusz, DE a fingerprint hianya miatt MEGIS nevesitve marad a nem-oroklés indoka', r.reason !== null)
  check('  flipKnown=false, mert nincs valtozas', r.flipKnown === false)
}

console.log('\n[13] flipDetectable -- NEG. KONTROLL: MINDKET oldal fingerprintje ISMERT, KULONBOZIK -> attributableToBuildChange=true')
{
  const before = { entryPoint: 'create_document', sdkSha256: 'sha-old', status: STATUS.NO_OP }
  const after = { entryPoint: 'create_document', sdkSha256: 'sha-new', status: STATUS.SUPPORTED_VERIFIED }
  const r = flipDetectable(before, after)
  check('  mindket fingerprint ismert es kulonbozik -> a build-valtozas ATTRIBUALHATO', r.attributableToBuildChange === true)
  check('  reason=null (nincs mit nevesiteni, a teljes lanc ismert)', r.reason === null)
}

console.log('\n[14] flipDetectable -- kulonbozo belepesi pont eseten meg sem probal itelni')
{
  const a = { entryPoint: 'coedit', sdkSha256: 'x', status: STATUS.NO_OP }
  const b = { entryPoint: 'create_document', sdkSha256: 'x', status: STATUS.SUPPORTED_VERIFIED }
  const r = flipDetectable(a, b)
  check('  flipKnown=false, mert nem osszevethetok', r.flipKnown === false)
}

console.log(`\nellenorzesek: ${ok} ok, ${bukas} bukas`)
if (bukas) { process.exit(1) }
