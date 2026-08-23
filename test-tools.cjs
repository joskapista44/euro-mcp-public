const fs = require('fs')
const os = require('os')
const path = require('path')
const { spawn } = require('child_process')

// End-to-end gate for the EURO-MCP tools, driven through the tool interface (tools/call) - the
// zod schema, the argument passing and the response shape only show up on this path.
//
// It runs the PRODUCTION code path: the same box-helper.py, the same runner, the same tools. Only
// the transport is swapped (EURO_EXEC=local) and the Document Server is a throwaway fake. That
// matters more than it sounds: a gate that exercised a separate "test path" would be measuring
// the half nobody deploys.
//
// What a fake can honestly prove is the FAILURE branches. The branch that matters most - a real
// document really changing on the real service - was certified separately against the live
// instance and is not claimed here.

const DIR = __dirname
const SECRET = 'fake-titok-csak-teszthez'
let passed = 0
const failures = []

function check(name, condition, detail = '') {
  if (condition) {
    passed++
    console.log(`  ok    ${name}`)
  } else {
    failures.push(detail ? `${name} (${detail})` : name)
    console.log(`  BUKAS ${name}${detail ? ' -- ' + detail : ''}`)
  }
}

function startFake(mode) {
  return new Promise((resolve, reject) => {
    const proc = spawn('node', [path.join(DIR, 'fake-ds.cjs')], {
      env: { ...process.env, MODE: mode, FAKE_SECRET: SECRET, PORT: '0' },
    })
    let buf = ''
    const onData = (chunk) => {
      buf += chunk
      if (buf.includes('"port"')) {
        proc.stdout.off('data', onData)
        resolve({ proc, port: JSON.parse(buf.split('\n')[0]).port })
      }
    }
    proc.stdout.on('data', onData)
    proc.on('error', reject)
    setTimeout(() => reject(new Error('a hamis DS nem indult el 5 mp alatt')), 5000)
  })
}

// A temp env file so the helper reads a known secret instead of the deployment's own.
function envFileWith(secret) {
  const file = path.join(os.tmpdir(), `euro-mcp-test-env-${process.pid}-${Math.abs(secret.length)}`)
  fs.writeFileSync(file, `# teszt\nJWT_SECRET=${secret}\n`, { mode: 0o600 })
  return file
}

function callTool(env, name, args) {
  return new Promise((resolve, reject) => {
    const proc = spawn('node', [path.join(DIR, 'euro-mcp.cjs')], { env: { ...process.env, ...env } })
    let out = ''
    proc.stdout.on('data', (c) => (out += c))
    proc.on('error', reject)
    proc.on('close', () => {
      for (const line of out.split('\n')) {
        if (!line.trim().startsWith('{')) continue
        const msg = JSON.parse(line)
        if (msg.id === 2) {
          try {
            return resolve(JSON.parse(msg.result.content[0].text))
          } catch {
            return resolve({ _raw: msg.result })
          }
        }
      }
      reject(new Error(`nem jott tool-valasz. Nyers: ${out.slice(0, 400)}`))
    })
    proc.stdin.write(
      JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        method: 'initialize',
        params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'gate', version: '0' } },
      }) + '\n',
    )
    proc.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n')
    proc.stdin.write(
      JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name, arguments: args } }) + '\n',
    )
    proc.stdin.end()
    setTimeout(() => proc.kill(), 40000)
  })
}

const OPS = [{ type: 'append_paragraph', text: 'szia' }]

function envFor(port, secretFile) {
  return {
    EURO_EXEC: 'local',
    EURO_BOX_IP: '127.0.0.1',
    EO_DS_URL: `http://127.0.0.1:${port}`,
    EO_ENV_FILE: secretFile,
  }
}

async function main() {
  const goodEnvFile = envFileWith(SECRET)
  const badEnvFile = envFileWith('ROSSZ-TITOK-nem-ez')

  // --- 1. Rejected token.
  console.log('\n[1] elutasitott token (rossz titok)')
  const f1 = await startFake('ok')
  const badEdit = await callTool(envFor(f1.port, badEnvFile), 'edit_document', { operations: OPS })
  check('rossz titok -> edit_document nem ok', badEdit.ok === false, JSON.stringify(badEdit).slice(0, 160))
  check('rossz titok -> az AUTH agat nevezi meg', badEdit.outcome === 'auth', JSON.stringify(badEdit).slice(0, 160))
  f1.proc.kill()

  // --- 2. The server cannot download what it needs.
  console.log('\n[2] a szolgaltatas nem tudja letolteni a scriptet')
  const f2 = await startFake('fetch')
  const fetchEdit = await callTool(envFor(f2.port, goodEnvFile), 'edit_document', { operations: OPS })
  check('letoltesi hiba -> nem ok', fetchEdit.ok === false)
  check('letoltesi hiba -> a FETCH agat nevezi meg', fetchEdit.outcome === 'fetch', JSON.stringify(fetchEdit).slice(0, 160))
  f2.proc.kill()

  // --- 3. The service is not there at all. A transport failure must not be dressed up as a
  //        service verdict - "unreachable" and "refused" are different answers.
  console.log('\n[3] a szolgaltatas egyaltalan nem valaszol')
  const deadEdit = await callTool(envFor(1, goodEnvFile), 'edit_document', { operations: OPS })
  check('nem valaszol -> nem ok', deadEdit.ok === false)
  check('nem valaszol -> NEM-MERT, nem szolgaltatasi verdikt', deadEdit.outcome === 'nem-mert', JSON.stringify(deadEdit).slice(0, 160))

  // --- 4. The one that matters: a successful answer over a document that did not change.
  console.log('\n[4] sikeres valasz, de a dokumentum NEM valtozott')
  const f4 = await startFake('silent')
  const silentEdit = await callTool(envFor(f4.port, goodEnvFile), 'edit_document', { operations: OPS })
  check('nema kudarc -> NEM jelent sikert', silentEdit.ok === false, JSON.stringify(silentEdit).slice(0, 200))
  check('nema kudarc -> a tartalmi ellenorzes mond nemet', silentEdit.contentVerified === false)
  f4.proc.kill()

  // --- 5. A missing input file must be caught here, not turned into a service error.
  //
  // *** EZ AZ ESET KETTEVALT (2026-08-15): *** a protokoll szerint
  // a DocBuilder CSAK letrehozasra hasznalhato, ezert az `edit_document` mostantol MINDEN
  // `document_path`-ot elutasit -- letezot es nem letezot egyarant. Az eredeti aggaly (egy rossz
  // utat NEVEZZUK MEG, ne szolgaltatas-hibava valjon) attol meg ervenyes, csak mar a
  // `run_builder_script`-re all, ahol a mag-fajl tovabbra is jogos bemenet.
  // *A lefedettseget nem toroltuk, hanem athelyeztuk oda, ahol a szerzodes meg ervenyes.*
  console.log('\n[5] a DocBuilder-ut hatoköre: letrehozas igen, szerkesztes nem')
  const noFile = await callTool(envFor(1, goodEnvFile), 'edit_document', {
    operations: OPS,
    document_path: '/tmp/nincs-ilyen-fajl-euro-mcp.docx',
  })
  check('edit_document + document_path -> protokoll-tilt', noFile.ok === false && noFile.outcome === 'protokoll-tilt',
    JSON.stringify(noFile).slice(0, 200))
  check('a tiltas MEGMONDJA a helyes utat (co-editing)', /co-editing/.test(noFile.error || ''),
    (noFile.error || '').slice(0, 120))
  // Kikötés 1 (2026-08-16): a bool mellett az INDOK is jelen kell legyen -- ez a regresszio, amit
  // a mai javitas megszuntetett, es amit egy kesobbi atiras csendben visszahozhatna, ha itt nincs
  // rajta assert.
  check('  ES az INDOK is jelen van, nem csak a bool (coeditUtAllapot)',
    typeof noFile.coeditUtAllapot === 'string' && noFile.coeditUtAllapot.length > 0,
    JSON.stringify(noFile.coeditUtAllapot))

  const noSeed = await callTool(envFor(1, goodEnvFile), 'run_builder_script', {
    script: 'builder.OpenFile("__DOC_URL__");builder.CloseFile();',
    document_path: '/tmp/nincs-ilyen-fajl-euro-mcp.docx',
  })
  check('run_builder_script + nem letezo mag -> megnevezi a fajlt',
    noSeed.ok === false && /nincs-ilyen-fajl/.test(noSeed.error || ''), JSON.stringify(noSeed).slice(0, 200))

  // --- 6. Positive control. Without it every check above would also pass against a tool that can
  //        only ever fail, and the gate would be measuring nothing.
  console.log('\n[6] POZ. KONTROLL: a jo agnak at KELL mennie')
  const f6 = await startFake('ok')
  const okEdit = await callTool(envFor(f6.port, goodEnvFile), 'edit_document', { operations: OPS })
  check('jo ut -> ok', okEdit.ok === true, JSON.stringify(okEdit).slice(0, 220))
  check('jo ut -> a tartalom igazolva', okEdit.contentVerified === true)
  // Over the local transport the fake also calls from 127.0.0.1, so its fetches cannot be told
  // apart from our own precondition check. The runner reports null rather than 0 for exactly that
  // reason, and this asserts the honesty of that report - not a fetch count the gate cannot see.
  check('lokalis szallitason a fetch-szamlalo NEM-MERT-et mond, nem nullat', okEdit.serverFetches === null, `serverFetches=${okEdit.serverFetches}`)

  const okStatus = await callTool(envFor(f6.port, goodEnvFile), 'service_status', {})
  check('jo ut -> service_status ok', okStatus.ok === true, JSON.stringify(okStatus).slice(0, 200))

  const rawOk = await callTool(envFor(f6.port, goodEnvFile), 'run_builder_script', {
    script: 'builder.OpenFile("__DOC_URL__", "docx");\nvar d = Api.GetDocument();\nbuilder.SaveFile("docx", "eredmeny.docx");\nbuilder.CloseFile();\n',
  })
  check('run_builder_script lefut', rawOk.outcome !== undefined, JSON.stringify(rawOk).slice(0, 200))

  const rawBad = await callTool(envFor(f6.port, goodEnvFile), 'run_builder_script', { script: 'builder.CreateFile("docx");' })
  check('run_builder_script elutasitja a CreateFile-t, indoklassal', rawBad.ok === false && /OpenFile/.test(rawBad.error || ''))

  // MERT lelet (2026-08-14, eles): a legerosebb primitiv eredmenye elerhetetlen volt -- a
  // runJob-ot returnDoc NELKUL hivtuk, es output_path parameter sem volt. Emiatt a mai
  // Rimini-dokumentum NEM ezzel az eszkozzel keszult, hanem a runner.cjs kozvetlen hivasaval:
  // ugyanaz a kodut, csak az eszkoz kerult megkerulesre.
  const outPath = path.join(os.tmpdir(), `zero-rbs-out-${process.pid}-${Date.now()}.docx`)
  fs.rmSync(outPath, { force: true })
  const rawOut = await callTool(envFor(f6.port, goodEnvFile), 'run_builder_script', {
    script: 'builder.OpenFile("__DOC_URL__", "docx");\nvar d = Api.GetDocument();\nbuilder.SaveFile("docx", "eredmeny.docx");\nbuilder.CloseFile();\n',
    output_path: outPath,
  })
  // NOT asserting `ok === true` here, and the reason is measured, not stylistic: box-helper.py
  // computes `"ok": bool(markers)`, i.e. ok means "the per-call marker was found in the saved
  // document". A raw builder script plants no marker, so ok is false for EVERY successful raw
  // run -- which is this tool's documented contract ("No content verification is performed:
  // the script decides what it produces"). The existing 'run_builder_script lefut' check above
  // already avoids ok for the same reason. What CAN be asserted is what the card is about: the
  // produced document actually came back.
  check(
    'run_builder_script output_path-szal VISSZAADJA a dokumentumot',
    rawOut.outcome !== undefined && fs.existsSync(outPath) && fs.statSync(outPath).size > 0,
    `outcome=${rawOut.outcome} letezik=${fs.existsSync(outPath)} meret=${fs.existsSync(outPath) ? fs.statSync(outPath).size : 'n/a'}`,
  )
  check(
    'run_builder_script megnevezi, HOVA irt (written)',
    rawOut.written === outPath,
    `written=${JSON.stringify(rawOut.written)}`,
  )
  fs.rmSync(outPath, { force: true })

  // NEG. KONTROLL ugyanabban a korben: output_path NELKUL semmi nem irodik ki, es a valasz
  // alakja a mai marad -- kulonben a fenti ket allitas azt is elfogadna, ha az eszkoz MINDIG ir.
  const rawNoOut = await callTool(envFor(f6.port, goodEnvFile), 'run_builder_script', {
    script: 'builder.OpenFile("__DOC_URL__", "docx");\nvar d = Api.GetDocument();\nbuilder.SaveFile("docx", "eredmeny.docx");\nbuilder.CloseFile();\n',
  })
  check(
    'output_path NELKUL nem ir fajlt, es a valasz alakja valtozatlan',
    rawNoOut.outcome !== undefined && (rawNoOut.written === null || rawNoOut.written === undefined),
    JSON.stringify(rawNoOut).slice(0, 200),
  )

  // --- 7. create_document: the first tool wiring of
  //        buildCreateScript/OPERATIONS -- POZ. KONTROLL and the itemized report in one shot.
  console.log('\n[7] create_document -- vegyes koteg, csomag-konzisztencia, DS-hiba')

  const vegyesOut = path.join(os.tmpdir(), `euro-e8-vegyes-${process.pid}-${Date.now()}.docx`)
  fs.rmSync(vegyesOut, { force: true })
  const vegyes = await callTool(envFor(f6.port, goodEnvFile), 'create_document', {
    core: 'docx',
    operations: [
      { type: 'text', text: 'a' },
      // listType 2026-08-17 ota MAR TAMOGATOTT -- ez a batch a "nem-tamogatott kozepen" alakot
      // meri, nem magat a listType-ot, ezert egy MASIK, tovabbra is nevesitetten refuzalt mezore
      // (highlight ismeretlen erteke) valtott
      { type: 'text', text: 'x', highlight: 'nincsilyenszin' },
      { type: 'text', text: 'b' },
    ],
    output_path: vegyesOut,
  })
  check('POZ. KONTROLL: 2 tamogatott + 1 nem -> ok=true (a ket sikeres muvelet a DS-hiba/csomag-hiba hianyaban eleg)',
    vegyes.ok === true, JSON.stringify(vegyes).slice(0, 300))
  check('  muveletek 3 tetelt sorol fel, tetelesen', Array.isArray(vegyes.muveletek) && vegyes.muveletek.length === 3, JSON.stringify(vegyes.muveletek))
  check('  a 0. es 2. alkalmazva, az 1. nem-tamogatott', vegyes.muveletek?.[0]?.outcome === 'alkalmazva' && vegyes.muveletek?.[1]?.outcome === 'nem-tamogatott' && vegyes.muveletek?.[2]?.outcome === 'alkalmazva', JSON.stringify(vegyes.muveletek))
  check('  a nem-tamogatott bejegyzes megnevezi a highlightot', /highlight/.test(vegyes.muveletek?.[1]?.reason || ''), vegyes.muveletek?.[1]?.reason)
  check('  csomagKonzisztens.ok true (a fake DS minimal docx-e nem hordoz torott hivatkozast)', vegyes.csomagKonzisztens?.ok === true, JSON.stringify(vegyes.csomagKonzisztens))
  // PIROS AG: ok=true, DE a koteg NEM teljes -- a ket uj mezo
  // ezt mondja ki, kulon az `ok`-tol (ami itt szandekosan igaz marad, lasd a fenti POZ. KONTROLL-t).
  check('  *** UJ: mindenMuveletAlkalmazva=false, pedig ok=true (a koteg nem teljes) ***',
    vegyes.mindenMuveletAlkalmazva === false, JSON.stringify(vegyes.mindenMuveletAlkalmazva))
  check('  *** UJ: nemAlkalmazottMuveletSzam=1 ***',
    vegyes.nemAlkalmazottMuveletSzam === 1, JSON.stringify(vegyes.nemAlkalmazottMuveletSzam))
  check('  a fajl TENYLEG kiirodott a kert utvonalra', fs.existsSync(vegyesOut) && fs.statSync(vegyesOut).size > 0, `letezik=${fs.existsSync(vegyesOut)}`)
  check('  outputPath megnevezi, hova irt', vegyes.outputPath === vegyesOut, JSON.stringify(vegyes.outputPath))
  fs.rmSync(vegyesOut, { force: true })

  console.log('\n[8] create_document -- csupa tamogatott: a hianylista URES, nem hianyzik')
  const csupaJo = await callTool(envFor(f6.port, goodEnvFile), 'create_document', {
    core: 'docx', operations: [{ type: 'text', text: 'a' }, { type: 'text', text: 'b' }],
  })
  check('ok=true', csupaJo.ok === true, JSON.stringify(csupaJo).slice(0, 200))
  check('a muveletek MIND alkalmazva -- egyetlen nem-tamogatott/hiba SEM szerepel (ures, nem hianyzo)',
    Array.isArray(csupaJo.muveletek) && csupaJo.muveletek.length === 2 && csupaJo.muveletek.every((m) => m.outcome === 'alkalmazva'),
    JSON.stringify(csupaJo.muveletek))
  // ZOLD AG, POZ. KONTROLL a fenti PIROS mellett: csupa ervenyes koteg -> a ket uj mezo is teljes sikert jelent.
  check('  *** UJ: mindenMuveletAlkalmazva=true (ZOLD, POZ. KONTROLL a [7] PIROS agahoz) ***',
    csupaJo.mindenMuveletAlkalmazva === true, JSON.stringify(csupaJo.mindenMuveletAlkalmazva))
  check('  *** UJ: nemAlkalmazottMuveletSzam=0 ***',
    csupaJo.nemAlkalmazottMuveletSzam === 0, JSON.stringify(csupaJo.nemAlkalmazottMuveletSzam))
  check('output_path nelkul NEM irodik tartos fajl, DE a konzisztencia meresehez a byte-ok akkor is elleneorzottek', csupaJo.csomagKonzisztens?.ok === true, JSON.stringify(csupaJo.csomagKonzisztens))
  check('output_path nelkul outputPath null', csupaJo.outputPath === null, JSON.stringify(csupaJo.outputPath))

  console.log('\n[9] create_document -- csupa NEM-tamogatott: a hivas nem "sikeres jarat", es a hiba MAGAVAL VISZI a reportot')
  const csupaRossz = await callTool(envFor(f6.port, goodEnvFile), 'create_document', {
    core: 'docx', operations: [{ type: 'text', text: 'x', highlight: 'nincsilyenszin' }],
  })
  check('ok=false, nevesitett outcome', csupaRossz.ok === false && csupaRossz.outcome === 'nincs-alkalmazhato-muvelet', JSON.stringify(csupaRossz).slice(0, 250))
  check('a muveletek meg IGY is tetelesen szerepelnek (nem csak egy osszevont uzenet)', Array.isArray(csupaRossz.muveletek) && csupaRossz.muveletek.length === 1 && csupaRossz.muveletek[0].outcome === 'nem-tamogatott', JSON.stringify(csupaRossz.muveletek))
  check('  *** UJ: mindenMuveletAlkalmazva=false a "nincs-alkalmazhato-muvelet" agon is (nem csak a sikeres uton) ***',
    csupaRossz.mindenMuveletAlkalmazva === false, JSON.stringify(csupaRossz.mindenMuveletAlkalmazva))
  check('  *** UJ: nemAlkalmazottMuveletSzam=1 ***',
    csupaRossz.nemAlkalmazottMuveletSzam === 1, JSON.stringify(csupaRossz.nemAlkalmazottMuveletSzam))

  console.log('\n[10] create_document -- DS-szintu hiba (rossz token) NEM keveredik ossze egy csomag-hibaval')
  const dsHiba = await callTool(envFor(f6.port, badEnvFile), 'create_document', {
    core: 'docx', operations: [{ type: 'text', text: 'a' }],
  })
  check('ok=false, a DS-hiba nevesitve (auth)', dsHiba.ok === false && dsHiba.dsOutcome === 'auth', JSON.stringify(dsHiba).slice(0, 200))
  check('a muveletek listaja MEGVAN (a caller latja, hogy a szandeka helyes lett volna)', Array.isArray(dsHiba.muveletek) && dsHiba.muveletek.length === 1 && dsHiba.muveletek[0].outcome === 'alkalmazva', JSON.stringify(dsHiba.muveletek))
  // *** A MEZO SZANDEKOS HATARA, NEVESITVE: a forditas (muveletek) FUGGETLEN a DS-hivas sikeretol --
  // itt a forditas teljes volt (mindenMuveletAlkalmazva=true), MIKOZBEN ok=false, mert a DS-hivas
  // maga bukott. A mezo tehat "a KERT koteg forditasa teljes volt-e", NEM "tenylegesen letrejott-e
  // a dokumentum" -- azt az `ok` es a `dsOutcome` mondja meg, ez a ket mezo nem helyettesiti oket.
  check('  *** UJ: mindenMuveletAlkalmazva=true a forditasi szinten, DE ok=false a DS-hiba miatt (a mezo hatara) ***',
    dsHiba.mindenMuveletAlkalmazva === true && dsHiba.ok === false, JSON.stringify({ m: dsHiba.mindenMuveletAlkalmazva, ok: dsHiba.ok }))
  check('csomagKonzisztens NULL marad -- nincs kiirt fajl, amit meresre erdemes lenne allitani', dsHiba.csomagKonzisztens === null, JSON.stringify(dsHiba.csomagKonzisztens))

  console.log('\n[11] create_document -- a protokoll-kapu (letezo output_path) UGYANUGY all, mint run_builder_script-nel')
  const gateOut = path.join(os.tmpdir(), `euro-e8-gate-${process.pid}-${Date.now()}.docx`)
  fs.writeFileSync(gateOut, 'mar letezik')
  const gated = await callTool(envFor(f6.port, goodEnvFile), 'create_document', {
    core: 'docx', operations: [{ type: 'text', text: 'a' }], output_path: gateOut,
  })
  check('mar letezo output_path -> protokoll-tilt, nem futtatja le a jobot', gated.ok === false && gated.outcome === 'protokoll-tilt', JSON.stringify(gated).slice(0, 200))
  // Kikötés 1, ugyanaz mint [5]-nel: lasd ott a komment.
  check('  ES az INDOK is jelen van, nem csak a bool (coeditUtAllapot)',
    typeof gated.coeditUtAllapot === 'string' && gated.coeditUtAllapot.length > 0,
    JSON.stringify(gated.coeditUtAllapot))
  fs.rmSync(gateOut, { force: true })

  // [11b] run_builder_script -- UGYANEZ a kapu, a masik eszkoznel. Korabban ez a par NEM allt
  // kulon tesztkent (a [11] cime hivatkozott ra, de a tenyleges hivas hianyzott) -- most potolva,
  // mert a kikötés 1 assert-je harom hely NELKUL nem teljes.
  console.log('\n[11b] run_builder_script -- a protokoll-kapu (letezo output_path) UGYANUGY all, mint create_document-nel')
  const gateOutRbs = path.join(os.tmpdir(), `euro-rbs-gate-${process.pid}-${Date.now()}.docx`)
  fs.writeFileSync(gateOutRbs, 'mar letezik')
  const gatedRbs = await callTool(envFor(f6.port, goodEnvFile), 'run_builder_script', {
    script: 'builder.OpenFile("__DOC_URL__");builder.SaveFile("docx","eredmeny.docx");builder.CloseFile();',
    output_path: gateOutRbs,
  })
  check('mar letezo output_path -> protokoll-tilt, nem futtatja le a jobot', gatedRbs.ok === false && gatedRbs.outcome === 'protokoll-tilt', JSON.stringify(gatedRbs).slice(0, 200))
  check('  ES az INDOK is jelen van, nem csak a bool (coeditUtAllapot)',
    typeof gatedRbs.coeditUtAllapot === 'string' && gatedRbs.coeditUtAllapot.length > 0,
    JSON.stringify(gatedRbs.coeditUtAllapot))
  fs.rmSync(gateOutRbs, { force: true })

  f6.proc.kill()

  for (const f of [goodEnvFile, badEnvFile]) fs.rmSync(f, { force: true })

  console.log(`\nellenorzesek: ${passed} ok, ${failures.length} bukas`)
  for (const f of failures) console.log(`  BUKAS: ${f}`)
  process.exit(failures.length ? 1 : 0)
}

main().catch((err) => {
  console.error('A KAPU NEM FUTOTT LE:', err.message)
  process.exit(2)
})
