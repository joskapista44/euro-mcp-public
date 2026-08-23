// findPathByFileId (coedit.cjs) -- the fileid -> DAV-path
// resolver the SAVED-package verification depends on. Pure PROPFIND-XML-parsing coverage, no
// network (fetchImpl injected, same pattern as talk-limits.js's resolveChatMaxLength).

const assert = require('assert')
const { findPathByFileId } = require('./coedit.cjs')

let osszes = 0
const hibak = []
async function checkAsync(cimke, fn) {
  osszes += 1
  try { await fn(); console.log(`  ok    ${cimke}`) }
  catch (err) { hibak.push(cimke); console.log(`  BUKAS ${cimke}\n        ${err.message}`) }
}

function fakeFetchOk(xmlBody) {
  return async () => ({ ok: true, status: 200, text: async () => xmlBody })
}

const MULTI_RESPONSE_XML = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:" xmlns:oc="http://owncloud.org/ns">
  <d:response>
    <d:href>/remote.php/dav/files/alpha/</d:href>
    <d:propstat><d:prop><oc:fileid>1000</oc:fileid></d:prop></d:propstat>
  </d:response>
  <d:response>
    <d:href>/remote.php/dav/files/alpha/report.docx</d:href>
    <d:propstat><d:prop><oc:fileid>1225578</oc:fileid></d:prop></d:propstat>
  </d:response>
  <d:response>
    <d:href>/remote.php/dav/files/alpha/other.pptx</d:href>
    <d:propstat><d:prop><oc:fileid>1225616</oc:fileid></d:prop></d:propstat>
  </d:response>
</d:multistatus>`

async function main() {
  console.log('\n[1] GREEN: a keresett fileid megtalalhato tobb bejegyzes kozott')
  await checkAsync('a masodik bejegyzes href-jet adja vissza, URL-decode-olva', async () => {
    const r = await findPathByFileId({
      url: 'https://euro-office.example', user: 'alpha', pass: 'x', fileId: 1225578,
      fetchImpl: fakeFetchOk(MULTI_RESPONSE_XML),
    })
    assert.strictEqual(r.ok, true)
    assert.strictEqual(r.href, '/remote.php/dav/files/alpha/report.docx')
  })

  console.log('\n[2] a fileid szamkent ES sztringkent is illeszkedik (a hivo barmelyik alakot adhatja)')
  await checkAsync('sztring fileid ugyanugy talal', async () => {
    const r = await findPathByFileId({
      url: 'https://euro-office.example', user: 'alpha', pass: 'x', fileId: '1225616',
      fetchImpl: fakeFetchOk(MULTI_RESPONSE_XML),
    })
    assert.strictEqual(r.ok, true)
    assert.strictEqual(r.href, '/remote.php/dav/files/alpha/other.pptx')
  })

  console.log('\n[3] PIROS: nem talalhato fileid -> nevesitett indok, nem dob')
  await checkAsync('ismeretlen fileid -> ok:false, az indok megnevezi a fileid-et es az atnezett bejegyzesszamot', async () => {
    const r = await findPathByFileId({
      url: 'https://euro-office.example', user: 'alpha', pass: 'x', fileId: 999999,
      fetchImpl: fakeFetchOk(MULTI_RESPONSE_XML),
    })
    assert.strictEqual(r.ok, false)
    assert.match(r.indok, /999999/)
    assert.match(r.indok, /3 bejegyzes/)
  })

  console.log('\n[4] PIROS: HTTP hiba -> nevesitve, nem dob')
  await checkAsync('HTTP 401 -> ok:false, a status a indokban', async () => {
    const r = await findPathByFileId({
      url: 'https://euro-office.example', user: 'alpha', pass: 'rossz', fileId: 1,
      fetchImpl: async () => ({ ok: false, status: 401, text: async () => '' }),
    })
    assert.strictEqual(r.ok, false)
    assert.match(r.indok, /401/)
  })

  console.log('\n[5] PIROS: a halozati hivas maga dob -> elkapva, nevesitve')
  await checkAsync('fetch dob -> ok:false, az uzenet atmegy', async () => {
    const r = await findPathByFileId({
      url: 'https://euro-office.example', user: 'alpha', pass: 'x', fileId: 1,
      fetchImpl: async () => { throw new Error('ECONNREFUSED') },
    })
    assert.strictEqual(r.ok, false)
    assert.match(r.indok, /ECONNREFUSED/)
  })

  console.log('\n[6] NEG. KONTROLL: egy blokk, aminek NINCS fileid property-je (mappa-bejegyzes tipikus alakja), nem okoz hamis talalatot')
  await checkAsync('href+fileid csak EGYUTT szamit -- egy csonka blokk nem parosodik a szomszedjaval', async () => {
    const csonka = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:" xmlns:oc="http://owncloud.org/ns">
  <d:response>
    <d:href>/remote.php/dav/files/alpha/nincs-fileid-property/</d:href>
    <d:propstat><d:prop></d:prop></d:propstat>
  </d:response>
  <d:response>
    <d:href>/remote.php/dav/files/alpha/valodi.docx</d:href>
    <d:propstat><d:prop><oc:fileid>42</oc:fileid></d:prop></d:propstat>
  </d:response>
</d:multistatus>`
    const r = await findPathByFileId({
      url: 'https://euro-office.example', user: 'alpha', pass: 'x', fileId: 42,
      fetchImpl: fakeFetchOk(csonka),
    })
    assert.strictEqual(r.ok, true)
    assert.strictEqual(r.href, '/remote.php/dav/files/alpha/valodi.docx')
  })

  console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
  if (hibak.length) { console.log('bukott:', hibak.join(' | ')); process.exitCode = 1 }
}

main().catch((err) => { console.error(err); process.exit(2) })
