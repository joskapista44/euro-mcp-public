const crypto = require('crypto')
const lib = require('./lib.cjs')

// Gate for the EURO-MCP core. Needs no Document Server, no secret and no
// box access - everything here is decided locally, so it can run before the service half is
// unblocked.
//
// It also has to be able to FAIL. Run with `--self-test` and one expectation is deliberately
// broken, so the red branch is something we have seen rather than something we assume:
//   node test-lib.js              -> rc=0, all checks pass
//   node test-lib.js --self-test  -> rc=1, and it names which check went red

const SELF_TEST = process.argv.includes('--self-test')

let passed = 0
const failures = []

function check(name, condition, detail = '') {
  if (condition) {
    passed++
    return
  }
  failures.push(detail ? `${name}: ${detail}` : name)
}

// ---------------------------------------------------------------------------
// 1. JWT. Four separate failure classes, each one an assertion of its own. A helper that
//    merely "runs" is the one that later breaks silently, and a wrong signature is invisible
//    from the outside until the service rejects everything.
// ---------------------------------------------------------------------------
const SECRET = 'teszt-titok-nem-eles-' + 'a'.repeat(20)
const payload = { async: false, url: 'http://example.invalid/script.docbuilder' }
const token = lib.signJwt(SECRET, payload)
const [rawHeader, rawBody, rawSig] = token.split('.')

// (a) Recompute the HMAC independently instead of trusting the helper's own arithmetic.
const expectedSig = crypto
  .createHmac('sha256', SECRET)
  .update(`${rawHeader}.${rawBody}`)
  .digest('base64')
  .replace(/=+$/, '')
  .replace(/\+/g, '-')
  .replace(/\//g, '_')
check('jwt/signature matches an independent HMAC', rawSig === expectedSig, `got ${rawSig}, expected ${expectedSig}`)

// (b) Header and payload survive the round trip.
const decodedHeader = JSON.parse(Buffer.from(rawHeader, 'base64url').toString())
const decodedBody = JSON.parse(Buffer.from(rawBody, 'base64url').toString())
check('jwt/header declares HS256', decodedHeader.alg === 'HS256' && decodedHeader.typ === 'JWT')
check('jwt/payload round-trips', JSON.stringify(decodedBody) === JSON.stringify(payload))

// (c) base64url-clean. A stray '+', '/' or '=' travels fine over HTTP and then fails
//     verification on the other side, which reads as an auth problem rather than an encoding one.
check('jwt/is base64url-clean', !/[+/=]/.test(token), `token contained a non-base64url character`)

// (d) The secret itself must not be recoverable from the token.
check('jwt/does not leak the secret', !token.includes(SECRET))

// A different secret must produce a different signature - otherwise the "signature" is not
// keyed at all and every one of the checks above would still pass.
const otherToken = lib.signJwt(SECRET + 'x', payload)
check('jwt/signature is keyed to the secret', otherToken.split('.')[2] !== rawSig)

check(
  'jwt/refuses to sign without a secret',
  (() => {
    try {
      lib.signJwt('', payload)
      return false
    } catch {
      return true
    }
  })(),
)

// ---------------------------------------------------------------------------
// 2. Script generation.
// ---------------------------------------------------------------------------
const marker = lib.makeMarker()
check('marker/is unique per call', lib.makeMarker() !== lib.makeMarker())

const script = lib.buildEditScript({
  docUrl: 'http://192.0.2.1:9999/bemenet.docx',
  operations: [{ type: 'append_paragraph', text: 'szia' }],
  marker,
})
check('script/opens rather than creates', script.includes('builder.OpenFile(') && !script.includes('builder.CreateFile('))
check('script/saves and closes', script.includes('builder.SaveFile(') && script.includes('builder.CloseFile()'))
check('script/carries the marker', script.includes(marker))

// A quote or a newline in caller text must not be able to end the statement early.
const nasty = lib.buildEditScript({
  docUrl: 'http://192.0.2.1:9999/b.docx',
  operations: [{ type: 'append_paragraph', text: 'a"; oDocument.Delete(); //' }],
  marker,
})
check(
  'script/escapes caller text instead of splicing it',
  nasty.includes('oDocument.Delete()') === false || nasty.includes('\\"'),
  'caller-supplied quotes reached the script unescaped',
)
check(
  'script/keeps the injected text inside one string literal',
  (nasty.match(/oParagraph\.AddText\(/g) || []).length === 1,
)

check(
  'script/rejects an unknown operation',
  (() => {
    try {
      lib.buildEditScript({ docUrl: 'http://x/y.docx', operations: [{ type: 'nincs_ilyen' }], marker })
      return false
    } catch {
      return true
    }
  })(),
)

check(
  'script/rejects a no-op request',
  (() => {
    try {
      lib.buildEditScript({ docUrl: 'http://x/y.docx', operations: [] })
      return false
    } catch {
      return true
    }
  })(),
)

// ---------------------------------------------------------------------------
// 3. Response classification. Every outcome the live service can produce, including the two
//    that cost the first live run a morning (-4) and the one that is a licence question (-3).
// ---------------------------------------------------------------------------
const cases = [
  [{ key: 'bld_1', urls: { 'eredmeny.docx': 'http://x/y' }, end: true }, lib.OUTCOME.OK],
  [{ error: -3 }, lib.OUTCOME.BLOCKED],
  [{ error: -8 }, lib.OUTCOME.AUTH],
  [{ error: 6 }, lib.OUTCOME.AUTH],
  [{ error: -4 }, lib.OUTCOME.FETCH],
  [{ error: 0 }, lib.OUTCOME.UNKNOWN],
  [{ error: -999 }, lib.OUTCOME.UNKNOWN],
  ['nem objektum', lib.OUTCOME.UNKNOWN],
  [null, lib.OUTCOME.UNKNOWN],
]
for (const [body, expected] of cases) {
  const got = lib.classifyDsResponse(body).outcome
  check(`classify/${JSON.stringify(body)} -> ${expected}`, got === expected, `got ${got}`)
}

// MEASURED live on 2026-08-14: a single non-existent method
// (SetTextSpacing) got -3 while the service was working perfectly two minutes either side. A
// discriminating probe: nine API calls one by one, EIGHT went through, one gave -3. If the
// advanced_api gate were shut, all nine would have failed -- so -3 is NOT only a licence answer.
//
// The classification stays BLOCKED (the call really did not run); only the EXPLANATION changes.
// Asserted as two separate conditions rather than one regex, so a message that names just one
// cause cannot pass by accident.
{
  const detail = lib.classifyDsResponse({ error: -3 }).detail || ''
  check('a -3 megnevezi a PELDANY-kapu okot', /advanced_api/.test(detail), detail)
  check('a -3 megnevezi a SZKRIPT-hiba okot is', /script|szkript/i.test(detail), detail)
  check(
    'a -3 megmondja, MELYIK PROBA donti el a kettot',
    /minimal|minimum/i.test(detail),
    detail,
  )
  check('a -3 besorolasa VALTOZATLANUL blocked', lib.classifyDsResponse({ error: -3 }).outcome === lib.OUTCOME.BLOCKED)
}

// A success must not be reported when there is nothing to fetch: that is the "HTTP 200 means
// it worked" trap, one layer down.
check('classify/success requires output urls', lib.classifyDsResponse({ key: 'k', end: true }).outcome !== lib.OUTCOME.OK)

// ---------------------------------------------------------------------------
// 4. Content verification.
// ---------------------------------------------------------------------------
check(
  'content/finds a marker split across runs',
  lib.markerInDocumentXml(`<w:t>EURO-MCP-1</w:t><w:t>23-abcd</w:t>`, 'EURO-MCP-123-abcd'),
)
check('content/says no when the marker is absent', !lib.markerInDocumentXml('<w:t>semmi</w:t>', 'EURO-MCP-1'))
check('content/says no on empty input', !lib.markerInDocumentXml('', 'EURO-MCP-1'))

// ---------------------------------------------------------------------------
// 5. resolveXlsxCellText -- a specific cell's shared-string
// VALUE, not just presence. Fixtures below are the SAME shape the Document Server actually saved
// (verified against two live-downloaded probe packages, not invented XML).
// ---------------------------------------------------------------------------
const SHEET_SHARED = '<worksheet><sheetData><row r="1"><c r="Z1" t="s"><v>0</v></c></row></sheetData></worksheet>'
const SST_ONE = '<sst><si><t>marker-1|OK|B2:C3</t></si></sst>'
check('resolveXlsxCellText/shared-string cell resolves through sharedStrings.xml', lib.resolveXlsxCellText(SHEET_SHARED, SST_ONE, 'Z1') === 'marker-1|OK|B2:C3')
check('resolveXlsxCellText/missing cell -> null', lib.resolveXlsxCellText(SHEET_SHARED, SST_ONE, 'Z2') === null)
check('resolveXlsxCellText/missing sheetXml -> null, not a throw', lib.resolveXlsxCellText(null, SST_ONE, 'Z1') === null)
check('resolveXlsxCellText/missing sharedStrings but cell IS a shared-string ref -> null (cannot resolve, not a crash)', lib.resolveXlsxCellText(SHEET_SHARED, null, 'Z1') === null)

const SHEET_NUMERIC = '<worksheet><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>'
check('resolveXlsxCellText/numeric (non-string) cell -> null, not the raw number', lib.resolveXlsxCellText(SHEET_NUMERIC, null, 'A1') === null)

const SHEET_SELFCLOSING = '<worksheet><sheetData><row r="1"><c r="B1" s="2"/></row></sheetData></worksheet>'
check('resolveXlsxCellText/self-closing (no value at all) cell -> null', lib.resolveXlsxCellText(SHEET_SELFCLOSING, null, 'B1') === null)

// Real fixtures, downloaded from two independent live probes this morning (POZ + NEG) -- not
// fabricated, so a future refactor of this function gets checked against what the Document Server
// ACTUALLY produced, not an idealized shape.
const REAL_POZ_SHEET = '<c r="Z1" t="s"><v>0</v></c>'
const REAL_POZ_SST = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>B2:C3</t></si></sst>'
check('resolveXlsxCellText/real POZ-probe fixture (Api.Intersect result address)', lib.resolveXlsxCellText(REAL_POZ_SHEET, REAL_POZ_SST, 'Z1') === 'B2:C3')

const REAL_NEG_SST = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t xml:space="preserve">HIBA: Ranges do not intersect.</t></si><si><t xml:space="preserve">HIBA-NEV: Error</t></si></sst>'
check('resolveXlsxCellText/real NEG-probe fixture (thrown-error message, xml:space="preserve" attribute present)', lib.resolveXlsxCellText(REAL_POZ_SHEET, REAL_NEG_SST, 'Z1') === 'HIBA: Ranges do not intersect.')

// The deliberately broken expectation for --self-test. It asserts something that is false, so
// the harness must go red; if it does not, the harness is not judging anything.
if (SELF_TEST) {
  check('SELF-TEST/deliberately false expectation', lib.signJwt(SECRET, payload) === 'ez-nem-a-token')
}

// ---------------------------------------------------------------------------
console.log(`ellenorzesek: ${passed} ok, ${failures.length} bukas`)
for (const f of failures) console.log(`  BUKAS: ${f}`)
if (SELF_TEST && failures.length === 1 && failures[0].startsWith('SELF-TEST/')) {
  console.log('SELF-TEST: a kapu PIROSAT adott, pontosan egy szandekosan hamis allitasra -- a harness itel.')
  process.exit(1)
}
process.exit(failures.length ? 1 : 0)
