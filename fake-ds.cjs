const http = require('http')
const crypto = require('crypto')

// A throwaway stand-in for the Document Server, so the tool's failure branches can be proven
// without the real service, without a secret and without box access.
//
// It is deliberately NOT a simulator: it answers only the shapes we have actually observed on
// the live instance, and it is here to make the RED branches reachable. The green branch that
// matters - a real document really changing - can only be certified against the real server,
// and this file must never be mistaken for that.
//
//   MODE=auth        every call -> {"error":-8}      (token rejected)
//   MODE=fetch       every call -> {"error":-4}      (server could not download the script/doc)
//   MODE=silent      answers with urls, but the output does NOT contain the marker
//                    -> the "it ran, so it worked" trap, one layer down
//   MODE=ok          answers with urls, and the output DOES contain the marker
//   MODE=image-empty answers with urls; the output is a docx with a word/media/ ENTRY that is
//                    EMPTY (0 bytes) -- the red
//                    fixture for the AddImage call that reports ok:true with no actual image data
//   MODE=image-real  same shape, but the media entry carries real bytes -- the POZ. KONTROLL
//                    proving the check can also say yes

const MODE = process.env.MODE || 'ok'
const PORT = Number(process.env.PORT || 0)
const SECRET = process.env.FAKE_SECRET || 'fake-titok'

function b64url(buf) {
  return Buffer.from(buf).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_')
}

// The fake verifies the signature for real. A stand-in that accepted anything would let a
// broken signer pass here and only fail against the live service - the exact failure this
// harness exists to catch early.
function tokenValid(token) {
  if (!token || token.split('.').length !== 3) return false
  const [h, b, s] = token.split('.')
  return s === b64url(crypto.createHmac('sha256', SECRET).update(`${h}.${b}`).digest())
}

let lastMarker = null

const { execFileSync } = require('child_process')

// python3's zipfile writes the archive; see the comment at the /out/ handler for why this is not
// assembled here in node.
function buildDocx(text) {
  const py = `
import base64, io, json, sys, zipfile
text = json.loads(sys.stdin.read())["text"]
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    z.writestr("word/document.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' + text + '</w:t></w:r></w:p></w:body></w:document>')
sys.stdout.write(base64.b64encode(buf.getvalue()).decode())
`
  const out = execFileSync('python3', ['-c', py], { input: JSON.stringify({ text }) })
  return Buffer.from(out.toString(), 'base64')
}

// Az AddImage-piros fixture -- egy docx, ahol a
// `word/media/` BEJEGYZES LETEZIK, de EMPTY-vel (0 bajt, mert=false) vagy VALODI tartalommal
// (mert=true) sul el. Ugyanaz a python3-stdlib-mintaju epites, mint buildDocx() -- kezzel irt
// zip itt sem jon szoba.
function buildDocxWithMedia(mert) {
  const py = `
import base64, io, json, sys, zipfile
mert = json.loads(sys.stdin.read())["mert"]
media = b"\\x89PNG-fake-bytes-not-a-real-image" if mert else b""
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    z.writestr("word/document.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>')
    z.writestr("word/media/image1.png", media)
sys.stdout.write(base64.b64encode(buf.getvalue()).decode())
`
  const out = execFileSync('python3', ['-c', py], { input: JSON.stringify({ mert }) })
  return Buffer.from(out.toString(), 'base64')
}

const server = http.createServer(async (req, res) => {
  const send = (code, obj) => {
    res.writeHead(code, { 'Content-Type': 'application/json' })
    res.end(JSON.stringify(obj))
  }

  if (req.url.startsWith('/out/')) {
    // The "saved document" has to be a real docx, because the production helper unzips it and
    // reads word/document.xml - a stand-in that served bare XML would make the tool report
    // "not a docx" and the gate would be measuring the fake's shortcut, not the tool.
    // The zip is built by python3's stdlib rather than by hand: hand-rolled central directories
    // produce quietly corrupt archives, which is a worse failure than not having a fake at all.
    const text = MODE === 'silent' ? 'valtozatlan' : lastMarker || ''
    try {
      const buf =
        MODE === 'image-empty' ? buildDocxWithMedia(false) :
        MODE === 'image-real' ? buildDocxWithMedia(true) :
        buildDocx(text)
      res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
      res.end(buf)
    } catch (err) {
      res.writeHead(500)
      res.end(String(err.message))
    }
    return
  }

  if (!req.url.startsWith('/docbuilder')) {
    res.writeHead(404)
    res.end('nincs ilyen vegpont')
    return
  }

  let raw = ''
  for await (const chunk of req) raw += chunk
  let body
  try {
    body = JSON.parse(raw)
  } catch {
    return send(200, { error: -999 })
  }

  const header = (req.headers.authorization || '').replace(/^Bearer\s+/i, '')
  if (MODE === 'auth' || !tokenValid(body.token) || !tokenValid(header)) return send(200, { error: -8 })
  if (MODE === 'fetch') return send(200, { error: -4 })

  // Fetch the script the way the real server does, so an unreachable script url produces the
  // same -4 here as it would there.
  let script
  try {
    const r = await fetch(body.url)
    if (!r.ok) return send(200, { error: -4 })
    script = await r.text()
  } catch {
    return send(200, { error: -4 })
  }

  if (!script.includes('builder.OpenFile(')) return send(200, { error: -999 })
  // Read the marker out of the capture group rather than slicing a fixed prefix off the match:
  // the slicing version was off by one, which produced a document without the marker and made
  // the tool report a correct "not verified" for a fault that was entirely in this fake.
  let marker = null
  for (const m of script.matchAll(/AddText\("((?:[^"\\]|\\.)*)"\)/g)) {
    if (m[1].startsWith('EURO-MCP-')) marker = m[1]
  }
  lastMarker = marker

  const base = `http://127.0.0.1:${server.address().port}`
  send(200, { key: 'fake_' + crypto.randomBytes(4).toString('hex'), urls: { 'eredmeny.docx': `${base}/out/eredmeny.xml` }, end: true })
})

server.listen(PORT, '127.0.0.1', () => {
  console.log(JSON.stringify({ port: server.address().port, mode: MODE }))
})
