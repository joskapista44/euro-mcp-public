// readDocumentContent (coedit.cjs): the office_get_text / office_get_comments
// orchestration -- PROPFIND resolve, WebDAV download, unzip, parse. Network is injected
// (fetchImpl, same pattern as test_coedit_findpath.cjs); the zip itself is REAL (python3
// zipfile, same fixture technique as test_package_consistency.cjs) so unzip actually runs.

const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')
const { execFileSync } = require('child_process')
const { readDocumentContent } = require('./coedit.cjs')

let osszes = 0
const hibak = []
async function checkAsync(cimke, fn) {
  osszes += 1
  try { await fn(); console.log(`  ok    ${cimke}`) }
  catch (err) { hibak.push(cimke); console.log(`  BUKAS ${cimke}\n        ${err.message}`) }
}

const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'euro-coedit03-read-'))
function buildZip(name, files) {
  const zipPath = path.join(TMP_DIR, name)
  const script = `
import zipfile, json, sys
files = json.loads(sys.argv[1])
with zipfile.ZipFile(sys.argv[2], 'w') as z:
    for entry, content in files.items():
        z.writestr(entry, content)
`
  execFileSync('python3', ['-c', script, JSON.stringify(files), zipPath])
  return zipPath
}

const PROPFIND_XML = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:" xmlns:oc="http://owncloud.org/ns">
  <d:response>
    <d:href>/remote.php/dav/files/alpha/dokumentum.docx</d:href>
    <d:propstat><d:prop><oc:fileid>555</oc:fileid></d:prop></d:propstat>
  </d:response>
</d:multistatus>`

// PROPFIND and the actual download share one fetchImpl (same as writeOperationsToDocument's real
// caller would use one `fetch`) -- route on method so PROPFIND gets XML and GET gets the zip bytes.
function fakeFetchRouted(zipPath) {
  return async (url, opts) => {
    if (opts && opts.method === 'PROPFIND') return { ok: true, status: 200, text: async () => PROPFIND_XML }
    const buf = fs.readFileSync(zipPath)
    return { ok: true, status: 200, arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) }
  }
}

async function main() {
  console.log('\n[1] GREEN: docx, van szoveg ES komment -- mindketto helyesen visszaadva')
  const withComment = buildZip('with-comment.docx', {
    'word/document.xml': '<w:p><w:r><w:t>Elso bekezdes.</w:t></w:r></w:p><w:p><w:r><w:t>Masodik bekezdes.</w:t></w:r></w:p>',
    'word/comments.xml': '<w:comments><w:comment w:id="0" w:author="Alpha" w:date="2026-08-16T10:00:00Z"><w:p><w:r><w:t>egy komment</w:t></w:r></w:p></w:comment></w:comments>',
  })
  await checkAsync('ok=true, outcome=olvasva', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'docx', fetchImpl: fakeFetchRouted(withComment) })
    assert.strictEqual(r.ok, true)
    assert.strictEqual(r.outcome, 'olvasva')
  })
  await checkAsync('a ket bekezdes sorrendben, hatarok nelkul osszefuzve', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'docx', fetchImpl: fakeFetchRouted(withComment) })
    assert.deepStrictEqual(r.bekezdesek, ['Elso bekezdes.', 'Masodik bekezdes.'])
  })
  await checkAsync('a komment kiolvasva, mezoi hibatlanok', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'docx', fetchImpl: fakeFetchRouted(withComment) })
    assert.strictEqual(r.kommentek.length, 1)
    assert.strictEqual(r.kommentek[0].author, 'Alpha')
    assert.strictEqual(r.kommentek[0].text, 'egy komment')
  })

  console.log('\n[2] KOMMENT NELKULI dokumentum: bekezdesek megvannak, kommentek URES lista (NEM hiba, NEM dobas)')
  const noComment = buildZip('no-comment.docx', {
    'word/document.xml': '<w:p><w:r><w:t>Csak szoveg, komment nelkul.</w:t></w:r></w:p>',
  })
  await checkAsync('ok=true, kommentek=[]', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'docx', fetchImpl: fakeFetchRouted(noComment) })
    assert.strictEqual(r.ok, true)
    assert.deepStrictEqual(r.kommentek, [])
    assert.deepStrictEqual(r.bekezdesek, ['Csak szoveg, komment nelkul.'])
  })

  console.log('\n[3] PIROS: xlsx -- NEVESITETT megtagadas, MIELOTT barmi letoltes/unzip tortenne (pptx mar TAMOGATOTT, lasd [7]-[9])')
  await checkAsync('xlsx -> core-nem-tamogatott, nem hivja a fetchImpl-t', async () => {
    let fetchHivva = false
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'xlsx', fetchImpl: async () => { fetchHivva = true; return { ok: true } } })
    assert.strictEqual(r.ok, false)
    assert.strictEqual(r.outcome, 'core-nem-tamogatott')
    assert.strictEqual(fetchHivva, false, 'a fetchImpl-nek NEM szabad lefutnia, mielott a core-kapu eldol')
  })

  console.log('\n[7] ZOLD: pptx, 2 dia (a MASODIK dia MEGKEVERT rId-del a rels-ben, a K6 MoveTo-lelet ellen), az elso dian komment')
  const pptxFile = buildZip('with-slides.pptx', {
    'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId id="1" r:id="rIdB"/><p:sldId id="2" r:id="rIdA"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships>' +
      '<Relationship Id="rIdA" Target="slides/slide1.xml"/>' +
      '<Relationship Id="rIdB" Target="slides/slide2.xml"/>' +
      '</Relationships>',
    'ppt/slides/slide1.xml': '<a:p><a:r><a:t>Elso reszfajl szovege (masodikkent jelenik meg).</a:t></a:r></a:p>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="1" name=""/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>',
    'ppt/slides/slide2.xml': '<a:p><a:r><a:t>Masodik reszfajl szovege (elsokent jelenik meg).</a:t></a:r></a:p>',
    'ppt/slides/_rels/slide2.xml.rels': '<Relationships>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>' +
      '</Relationships>',
    'ppt/comments/comment1.xml': '<p:cmLst><p:cm authorId="1" dt="2026-08-17T01:40:26Z" idx="1"><p:text>komment az elso megjelenitett dian</p:text></p:cm></p:cmLst>',
    'ppt/commentAuthors.xml': '<p:cmAuthorLst><p:cmAuthor id="1" name="Beta"></p:cmAuthor></p:cmAuthorLst>',
    'docProps/core.xml': '<cp:coreProperties xmlns:cp="x" xmlns:dc="y"><dc:creator>Beta</dc:creator></cp:coreProperties>',
    'docProps/app.xml': '<Properties xmlns="z"><Application>ONLYOFFICE/2.5.565.0</Application><Slides>2</Slides></Properties>',
    'ppt/slides/_rels/slide1.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
    'ppt/slideLayouts/slideLayout1.xml': '<p:sldLayout type="blank"><p:cSld name="Blank"></p:cSld></p:sldLayout>',
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': '<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>',
    'ppt/theme/theme1.xml': '<a:theme name="Gamma"></a:theme>',
  })
  await checkAsync('ok=true, outcome=olvasva', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxFile) })
    assert.strictEqual(r.ok, true)
    assert.strictEqual(r.outcome, 'olvasva')
  })
  await checkAsync('a diak a MEGJELENITES sorrendjeben (slide2 resz elobb, mert a sldIdLst ugy mondta), NEM a resz-nevek sorrendjeben', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxFile) })
    assert.strictEqual(r.diak.length, 2)
    assert.strictEqual(r.diak[0].bekezdesek[0], 'Masodik reszfajl szovege (elsokent jelenik meg).')
    assert.strictEqual(r.diak[1].bekezdesek[0], 'Elso reszfajl szovege (masodikkent jelenik meg).')
  })
  await checkAsync('a lapos `bekezdesek` mezo UGYANEBBEN a sorrendben all (docx-szal konzisztens alak office_find-nak)', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxFile) })
    assert.deepStrictEqual(r.bekezdesek, ['Masodik reszfajl szovege (elsokent jelenik meg).', 'Elso reszfajl szovege (masodikkent jelenik meg).'])
  })
  await checkAsync('a komment a HELYES megjelenitett dia-indexen (0, mert az a dia jelenik meg elsokent), szerzo-nevvel felodva', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxFile) })
    assert.strictEqual(r.kommentek.length, 1)
    assert.strictEqual(r.kommentek[0].slideIndex, 0)
    assert.strictEqual(r.kommentek[0].authorName, 'Beta')
    assert.strictEqual(r.kommentek[0].text, 'komment az elso megjelenitett dian')
  })
  await checkAsync('a "dia-tartalom visszaolvasasa" (bovitett hatokor): a shape a MEGJELENITETT masodik dian all (diak[1].tartalom), es a tartalomOsszesen ugyanezt sajat slideIndex-szel adja', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxFile) })
    assert.strictEqual(r.diak[1].tartalom.shapes.length, 1)
    assert.strictEqual(r.diak[1].tartalom.shapes[0].shapeType, 'rect')
    assert.strictEqual(r.diak[0].tartalom.shapes.length, 0)
    assert.strictEqual(r.tartalomOsszesen.shapes.length, 1)
    assert.strictEqual(r.tartalomOsszesen.shapes[0].slideIndex, 1)
  })
  await checkAsync('dokumentum-metaadat: docProps/core.xml + app.xml kiolvasva', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxFile) })
    assert.strictEqual(r.metaadat.creator, 'Beta')
    assert.strictEqual(r.metaadat.application, 'ONLYOFFICE/2.5.565.0')
    assert.strictEqual(r.metaadat.slides, 2)
    assert.strictEqual(r.metaadat.title, null)
  })
  await checkAsync('elrendezes/tema: a slide1-resz (masodikkent megjelenitett, diak[1]) teljes lanca feloldva -- slide2-resznek (diak[0]) nincs rels-e, null marad, nem dob', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxFile) })
    assert.strictEqual(r.diak[1].elrendezesEsTema.layoutName, 'Blank')
    assert.strictEqual(r.diak[1].elrendezesEsTema.layoutType, 'blank')
    assert.strictEqual(r.diak[1].elrendezesEsTema.themeName, 'Gamma')
    assert.strictEqual(r.diak[0].elrendezesEsTema.layoutName, null)
  })

  console.log('\n[8] pptx: KOMMENT NELKULI bemutato -- diak/bekezdesek megvannak, kommentek URES lista, NEM dobas')
  const pptxNoComment = buildZip('no-comment.pptx', {
    'ppt/presentation.xml': '<p:sldIdLst><p:sldId id="1" r:id="rId1"/></p:sldIdLst>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>',
    'ppt/slides/slide1.xml': '<a:p><a:r><a:t>Csak szoveg, komment nelkul.</a:t></a:r></a:p>',
  })
  await checkAsync('ok=true, kommentek=[]', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxNoComment) })
    assert.strictEqual(r.ok, true)
    assert.deepStrictEqual(r.kommentek, [])
    assert.deepStrictEqual(r.bekezdesek, ['Csak szoveg, komment nelkul.'])
  })

  console.log('\n[9] PIROS: pptx, ppt/presentation.xml hianyzik -- csomag-olvashatatlan, nem dobott kivetel')
  const pptxHianyos = buildZip('broken.pptx', { 'ppt/slides/slide1.xml': '<a:p><a:r><a:t>x</a:t></a:r></a:p>' })
  await checkAsync('csomag-olvashatatlan', async () => {
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'pptx', fetchImpl: fakeFetchRouted(pptxHianyos) })
    assert.strictEqual(r.ok, false)
    assert.strictEqual(r.outcome, 'csomag-olvashatatlan')
  })

  console.log('\n[4] PIROS: a fileid nem oldhato fel -- nevesitett hiba, nem csendes ures valasz')
  await checkAsync('fileid-nem-oldhato-fel', async () => {
    const ures = async (url, opts) => (opts && opts.method === 'PROPFIND') ? { ok: true, status: 200, text: async () => '<d:multistatus xmlns:d="DAV:"></d:multistatus>' } : { ok: true }
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 999, core: 'docx', fetchImpl: ures })
    assert.strictEqual(r.ok, false)
    assert.strictEqual(r.outcome, 'fileid-nem-oldhato-fel')
  })

  console.log('\n[5] PIROS: a letoltes HTTP-hibat ad')
  await checkAsync('letoltes-sikertelen', async () => {
    const rossz = async (url, opts) => (opts && opts.method === 'PROPFIND') ? { ok: true, status: 200, text: async () => PROPFIND_XML } : { ok: false, status: 503 }
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'docx', fetchImpl: rossz })
    assert.strictEqual(r.ok, false)
    assert.strictEqual(r.outcome, 'letoltes-sikertelen')
  })

  console.log('\n[6] PIROS: a letoltott csomag nem valodi zip (csomag-olvashatatlan, nem dobott kivetel)')
  await checkAsync('csomag-olvashatatlan', async () => {
    const nemZip = async (url, opts) => (opts && opts.method === 'PROPFIND')
      ? { ok: true, status: 200, text: async () => PROPFIND_XML }
      : { ok: true, status: 200, arrayBuffer: async () => Buffer.from('nem zip tartalom').buffer }
    const r = await readDocumentContent({ url: 'https://x', user: 'alpha', pass: 'p', fileId: 555, core: 'docx', fetchImpl: nemZip })
    assert.strictEqual(r.ok, false)
    assert.strictEqual(r.outcome, 'csomag-olvashatatlan')
  })

  console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
  fs.rmSync(TMP_DIR, { recursive: true, force: true })
  if (hibak.length) { console.log('BUKOTT:', hibak.join(' | ')); process.exit(1) }
}
main().catch((e) => { console.error('FATAL', e); process.exit(1) })
