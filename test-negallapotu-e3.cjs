const fs = require('fs')
const os = require('os')
const path = require('path')
const { spawn } = require('child_process')

// RED/GREEN gate for the NEGYALLAPOTU KIMENET: does
// create_document's `contentVerified` (and the `ok` it now derives from) correctly distinguish
// "the DocBuilder call reported success" from "the requested image is actually in the package"?
//
// Uses the SAME production code path as test-tools.cjs (real euro-mcp.cjs over stdio, EURO_EXEC=
// local, throwaway fake-ds.cjs) -- only the fake's response shape is new (MODE=image-empty /
// MODE=image-real, added alongside the existing marker-based modes, see fake-ds.cjs's own header
// comment). The RED case reproduces the same named fixture exactly: an AddImage-shaped
// call whose saved package has a `word/media/` entry that exists but is empty.

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

function envFileWith(secret) {
  const file = path.join(os.tmpdir(), `euro-mcp-e3-test-env-${process.pid}-${Date.now()}`)
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
        jsonrpc: '2.0', id: 1, method: 'initialize',
        params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'e3-gate', version: '0' } },
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

function envFor(port, secretFile) {
  return { EURO_EXEC: 'local', EURO_BOX_IP: '127.0.0.1', EO_DS_URL: `http://127.0.0.1:${port}`, EO_ENV_FILE: secretFile }
}

// A vilag legkisebb PNG-je nem kell -- a fake nem nezi a bajtokat, csak azt, hogy `src` jol
// formalt data: URI legyen (lib.cjs resolveImageSrc), tehat barmilyen ervenyes base64 eleg.
const TINY_PNG_DATA_URI = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='

const IMAGE_OPS = [{ type: 'image', src: TINY_PNG_DATA_URI }]

async function main() {
  const goodEnvFile = envFileWith(SECRET)

  console.log('\n[R] PIROS: az AddImage-hivas ma ok:true-t adna ures word/media/-val -- most NEM')
  const fRed = await startFake('image-empty')
  const outRed = path.join(os.tmpdir(), `euro-e3-red-${process.pid}-${Date.now()}.docx`)
  fs.rmSync(outRed, { force: true })
  const red = await callTool(envFor(fRed.port, goodEnvFile), 'create_document', {
    core: 'docx', operations: IMAGE_OPS, output_path: outRed,
  })
  fRed.proc.kill()
  check('a script-szintu hivas nem dobott, es a muvelet magat ALKALMAZVA-nak jelenti (a regi csapda pontosan ez)',
    red.muveletek?.[0]?.outcome === 'alkalmazva', JSON.stringify(red.muveletek))
  check('*** contentVerified:false (a media-resz LETEZIK, de URES) ***',
    red.contentVerified === false, JSON.stringify(red).slice(0, 300))
  check('*** osszesitett ok:false, PEDIG a muvelet "alkalmazva"-t jelentett ***',
    red.ok === false, JSON.stringify(red).slice(0, 300))
  check('transportOk/executionOk/outputProduced mind true -- csak a TARTALOM bukik, a tobbi harom nem',
    red.transportOk === true && red.executionOk === true && red.outputProduced === true,
    JSON.stringify(red).slice(0, 300))
  fs.rmSync(outRed, { force: true })

  console.log('\n[G] ZOLD, POZ. KONTROLL: valodi media-tartalom -- a mero tud igent is mondani')
  const fGreen = await startFake('image-real')
  const outGreen = path.join(os.tmpdir(), `euro-e3-green-${process.pid}-${Date.now()}.docx`)
  fs.rmSync(outGreen, { force: true })
  const green = await callTool(envFor(fGreen.port, goodEnvFile), 'create_document', {
    core: 'docx', operations: IMAGE_OPS, output_path: outGreen,
  })
  fGreen.proc.kill()
  check('*** contentVerified:true (a media-resz LETEZIK ES nemnulla meretu) ***',
    green.contentVerified === true, JSON.stringify(green).slice(0, 300))
  check('*** osszesitett ok:true mind a negy mezon ***',
    green.ok === true && green.transportOk === true && green.executionOk === true && green.outputProduced === true,
    JSON.stringify(green).slice(0, 300))
  fs.rmSync(outGreen, { force: true })

  console.log('\n[N] NEGATIV KONTROLL: sima szoveges create_document, media-op nelkul -- a regi ut nem ternek el')
  const fPlain = await startFake('ok')
  const plain = await callTool(envFor(fPlain.port, goodEnvFile), 'create_document', {
    core: 'docx', operations: [{ type: 'text', text: 'sima szoveg, kep nelkul' }],
  })
  fPlain.proc.kill()
  check('kep-muvelet nelkul a contentVerified a MEGLEVO csomag-konzisztencia jelre esik vissza, es igaz',
    plain.contentVerified === true && plain.ok === true, JSON.stringify(plain).slice(0, 300))

  console.log(`\nellenorzesek: ${passed} ok, ${failures.length} bukas`)
  if (failures.length) {
    console.log('BUKOTT:')
    failures.forEach((f) => console.log('  - ' + f))
    process.exit(1)
  }
}

main().catch((err) => {
  console.error('VEGZETES HIBA:', err)
  process.exit(2)
})
