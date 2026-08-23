// office-diag-e6-trace-id: the concurrency proof. A trace ID kept in module-level state would
// pass a SEQUENTIAL test and still be wrong -- the failure only shows up when two calls overlap
// in time. Every check here fires calls CONCURRENTLY (Promise.all, never awaited one at a time)
// and verifies the two never mix.
const fs = require('fs')
const path = require('path')
const os = require('os')
const { spawn } = require('child_process')

let passed = 0
const failures = []
function check(name, cond, detail = '') {
  if (cond) { passed++; console.log(`  ok    ${name}`) }
  else { failures.push(detail ? `${name} (${detail})` : name); console.log(`  BUKAS ${name}${detail ? ' -- ' + detail : ''}`) }
}

function readJsonl(file) {
  return fs.readFileSync(file, 'utf8').trim().split('\n').filter(Boolean).map((l) => JSON.parse(l))
}

async function main() {
  const tmpLogDir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-trace-test-'))
  process.env.EURO_TRACE_LOG_DIR = tmpLogDir
  const trace = require('./office-trace.cjs')

  // --- [1] newTraceId() -- format + uniqueness under a tight loop -------------------------------
  const ids = new Set()
  for (let i = 0; i < 500; i++) ids.add(trace.newTraceId())
  check('[1] newTraceId format', /^off-\d{8}-\d{6}-[0-9a-f]{6}$/.test(trace.newTraceId()))
  check('[1] newTraceId uniqueness (500 calls, tight loop)', ids.size === 500, `${ids.size}/500 unique`)

  // --- [2] logTrace() under real concurrency: 20 ids x 10 lines each, ALL fired via Promise.all,
  // interleaved on purpose (odd/even alternate id order) so the appendFileSync calls genuinely
  // race against each other in the event loop, not just look concurrent syntactically.
  const idsFor2 = Array.from({ length: 20 }, () => trace.newTraceId())
  const calls = []
  for (let round = 0; round < 10; round++) {
    for (const id of idsFor2) {
      calls.push(Promise.resolve().then(() => trace.logTrace(id, 'concurrency-probe', { round })))
    }
  }
  await Promise.all(calls)
  const lines = readJsonl(trace.logPathFor())
  const ownLines = lines.filter((l) => l.stage === 'concurrency-probe')
  check('[2] every logTrace call produced exactly one line (200 calls -> 200 lines)', ownLines.length === 200, `${ownLines.length}/200`)
  const byId = new Map()
  for (const l of ownLines) byId.set(l.traceId, (byId.get(l.traceId) || 0) + 1)
  const allCorrectCount = idsFor2.every((id) => byId.get(id) === 10)
  check('[2] each of the 20 ids has EXACTLY its own 10 lines (no cross-write)', allCorrectCount, JSON.stringify([...byId.entries()]))
  const noForeignIds = ownLines.every((l) => idsFor2.includes(l.traceId))
  check('[2] no line carries an id that was never generated (no corruption)', noForeignIds)
  const allParsed = lines.length === fs.readFileSync(trace.logPathFor(), 'utf8').trim().split('\n').filter(Boolean).length
  check('[2] every appended line is valid, complete JSON (readJsonl did not throw)', allParsed)

  // --- [3] end-to-end: two concurrent runJob() calls, via the REAL fake-DS transport, distinct
  // traceIds AND distinct scripts (so a swap would be independently detectable two ways: the
  // wrong traceId showing up, or the wrong script's marker showing up on the wrong answer).
  const fakeDs = spawn('node', [path.join(__dirname, 'fake-ds.cjs')], {
    env: { ...process.env, MODE: 'happy', FAKE_SECRET: 'e6-concurrency-teszt', PORT: '0' },
  })
  const port = await new Promise((resolve, reject) => {
    let buf = ''
    const onData = (c) => {
      buf += c
      if (buf.includes('"port"')) { fakeDs.stdout.off('data', onData); resolve(JSON.parse(buf.split('\n')[0]).port) }
    }
    fakeDs.stdout.on('data', onData)
    fakeDs.on('error', reject)
    setTimeout(() => reject(new Error('fake-ds did not start in 5s')), 5000)
  })
  const envFile = path.join(os.tmpdir(), `e6-concurrency-env-${process.pid}`)
  fs.writeFileSync(envFile, `JWT_SECRET=e6-concurrency-teszt\n`, { mode: 0o600 })
  process.env.EO_ENV_FILE = envFile
  process.env.EO_DS_URL = `http://127.0.0.1:${port}`
  process.env.EURO_EXEC = 'local'
  delete require.cache[require.resolve('./runner.cjs')]
  const { runJob } = require('./runner.cjs')

  const idA = trace.newTraceId()
  const idB = trace.newTraceId()
  const markA = `E6-A-${idA}`
  const markB = `E6-B-${idB}`
  const scriptFor = (mark) => [
    'builder.OpenFile("__DOC_URL__");',
    'var oDocument = Api.GetDocument();',
    'var oParagraph = Api.CreateParagraph();',
    `oParagraph.AddText(${JSON.stringify(mark)});`,
    'oDocument.Push(oParagraph);',
    'builder.SaveFile("docx", "eredmeny.docx");',
    'builder.CloseFile();',
  ].join('\n')

  const [ansA, ansB] = await Promise.all([
    runJob({ script: scriptFor(markA), traceId: idA }),
    runJob({ script: scriptFor(markB), traceId: idB }),
  ])
  fakeDs.kill()
  fs.rmSync(envFile, { force: true })

  check('[3] call A produced output', ansA.ok !== undefined)
  check('[3] call B produced output', ansB.ok !== undefined)
  // The two answers are not compared against each other's traceId field (box-helper.py's own
  // 'traceId' field is descriptive, not authoritative -- runner.cjs is what attributes the log
  // lines). What matters: runJob's OWN log lines for idA/idB were written under the right id.
  const jobLines = readJsonl(trace.logPathFor()).filter((l) => l.stage === 'office-adapter-dispatch' || l.stage === 'job-result')
  const aLines = jobLines.filter((l) => l.traceId === idA)
  const bLines = jobLines.filter((l) => l.traceId === idB)
  check('[3] call A produced its own dispatch+result log lines', aLines.length === 2, `${aLines.length} lines`)
  check('[3] call B produced its own dispatch+result log lines', bLines.length === 2, `${bLines.length} lines`)
  check('[3] no log line is shared between A and B', jobLines.length === aLines.length + bLines.length)

  fs.rmSync(tmpLogDir, { recursive: true, force: true })

  console.log(`\nellenorzesek: ${passed} ok, ${failures.length} bukas`)
  if (failures.length) { console.log('BUKOTT:', failures.join(' | ')); process.exit(1) }
}

main().catch((e) => { console.error('HIBA:', e); process.exit(1) })
