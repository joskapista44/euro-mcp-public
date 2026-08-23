// Teszteli a package-consistency.cjs-t -- a KIÍRT csomag
// önmagában konzisztens-e. Fixture-ök EGYSZER-HASZNÁLATOS, eldobható /tmp zip-ekben, python3
// zipfile-lal építve (ugyanaz az eszköz, amit más egyszerű zip-építő szkriptek is használnak --
// nincs npm zip-függőség ebben a repóban, és egyet felvenni ehhez nem indokolt).
const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')
const { execFileSync } = require('child_process')
const { checkPackageConsistency, checkPptx, checkXlsx, checkDocx } = require('./package-consistency.cjs')

let failed = 0
function check(name, ok, detail = '') {
  if (ok) { console.log(`OK    ${name}`) }
  else { failed++; console.log(`FAIL  ${name}${detail ? '\n      ' + detail : ''}`) }
}

const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'euro-e8-consistency-'))

// Egy fixture-zip felépítése: { entryPath: contentString, ... } -> visszaadja a zip lemez-útját.
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

// --- pptx ------------------------------------------------------------------------------------
console.log('\n[1] pptx -- sldIdLst r:id-k a rels-ben (mag-minimal-pptx.py :: onellenorzes()-evel azonos minta)')

const pptxValid = buildZip('pptx-valid.zip', {
  'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>',
  'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId2" Type="slide" Target="slides/slide1.xml"/></Relationships>',
})
check('GREEN: minden r:id feloldhato -> ok', checkPptx(pptxValid).ok === true, JSON.stringify(checkPptx(pptxValid)))

const pptxBroken = buildZip('pptx-broken.zip', {
  'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>',
  'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId9" Type="slide" Target="slides/slide1.xml"/></Relationships>',
})
{
  const r = checkPptx(pptxBroken)
  check('PIROS: a sldIdLst rId2-re hivatkozik, a rels csak rId9-et deklaral -> ok=false, nevesitve', r.ok === false && /rId2/.test(r.issues.join(' ')), JSON.stringify(r))
}

const pptxEmpty = buildZip('pptx-empty.zip', {
  'ppt/presentation.xml': '<p:presentation><p:sldIdLst></p:sldIdLst></p:presentation>',
  'ppt/_rels/presentation.xml.rels': '<Relationships></Relationships>',
})
check('PIROS: ures sldIdLst -> ok=false, nevesitve (ures prezentacio)', checkPptx(pptxEmpty).ok === false && /ÜRES/.test(checkPptx(pptxEmpty).issues.join(' ')))

{
  const r = checkPptx(buildZip('pptx-missing.zip', { 'ppt/_rels/presentation.xml.rels': '<Relationships/>' }))
  check('hianyzo resz a csomagban (nincs presentation.xml) -> ok=false, nevesitve, nem dob', r.ok === false && /hiányzó rész/.test(r.issues.join(' ')), JSON.stringify(r))
}

// --- xlsx ------------------------------------------------------------------------------------
console.log('\n[2] xlsx -- sheet r:id-k a workbook rels-ben')

const xlsxValid = buildZip('xlsx-valid.zip', {
  'xl/workbook.xml': '<workbook><sheets><sheet name="Munka1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
})
check('GREEN: a sheet r:id feloldhato -> ok', checkXlsx(xlsxValid).ok === true, JSON.stringify(checkXlsx(xlsxValid)))

const xlsxBroken = buildZip('xlsx-broken.zip', {
  'xl/workbook.xml': '<workbook><sheets><sheet name="Munka1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels': '<Relationships></Relationships>',
})
{
  const r = checkXlsx(xlsxBroken)
  check('PIROS: a sheet rId1-re hivatkozik, a rels ures -> ok=false, nevesitve', r.ok === false && /rId1/.test(r.issues.join(' ')), JSON.stringify(r))
}

const xlsxEmpty = buildZip('xlsx-empty.zip', {
  'xl/workbook.xml': '<workbook><sheets></sheets></workbook>',
  'xl/_rels/workbook.xml.rels': '<Relationships/>',
})
check('PIROS: nulla lap -> ok=false, nevesitve', checkXlsx(xlsxEmpty).ok === false && /ÜRES/.test(checkXlsx(xlsxEmpty).issues.join(' ')))

// --- docx ------------------------------------------------------------------------------------
console.log('\n[3] docx -- a rels-ben NEVEZETT reszek tenyleg leteznek-e a csomagban (a MASIK irany)')

const docxValid = buildZip('docx-valid.zip', {
  'word/document.xml': '<w:document/>',
  'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId1" Type="image" Target="media/image1.png"/></Relationships>',
  'word/media/image1.png': 'PNGBYTES',
})
check('GREEN: a rels Targetje letezik a csomagban -> ok', checkDocx(docxValid).ok === true, JSON.stringify(checkDocx(docxValid)))

const docxBroken = buildZip('docx-broken.zip', {
  'word/document.xml': '<w:document/>',
  'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId1" Type="image" Target="media/image1.png"/></Relationships>',
  // media/image1.png SZANDEKOSAN hianyzik
})
{
  const r = checkDocx(docxBroken)
  check('PIROS: a rels media/image1.png-re hivatkozik, a csomagban nincs -> ok=false, nevesitve', r.ok === false && /image1\.png/.test(r.issues.join(' ')), JSON.stringify(r))
}

const docxExternal = buildZip('docx-external.zip', {
  'word/document.xml': '<w:document/>',
  'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId1" Type="hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>',
})
check('NEG. KONTROLL: TargetMode="External" -> NEM szamit hianyzo reszkent (URL, nem csomag-resz)', checkDocx(docxExternal).ok === true, JSON.stringify(checkDocx(docxExternal)))

const docxEmptyRels = buildZip('docx-empty-rels.zip', {
  'word/document.xml': '<w:document/>',
  'word/_rels/document.xml.rels': '<Relationships></Relationships>',
})
check('NEG. KONTROLL: ures rels -> ok=true (nulla kapcsolat egy docx-nek ERVENYES allapot, nem hiba -- eltero a pptx/xlsx-tol, ahol az ures lista maga a hiba)',
  checkDocx(docxEmptyRels).ok === true, JSON.stringify(checkDocx(docxEmptyRels)))

const docxNoRelsAtAll = buildZip('docx-no-rels.zip', {
  'word/document.xml': '<w:document/>',
  // word/_rels/document.xml.rels TELJESEN HIANYZIK -- MERT ALLAPOT (fake-ds.cjs sajat minimal
  // docx-e is igy nez ki): egy kapcsolat nelkuli docx-nek nincs is szuksege erre a reszre.
})
check('NEG. KONTROLL: a rels resz MAGA is hianyozhat -> ok=true, ez a VALODI mert minimal-docx alak, nem hiba',
  checkDocx(docxNoRelsAtAll).ok === true, JSON.stringify(checkDocx(docxNoRelsAtAll)))

// --- diszpecser + mutans-proba -----------------------------------------------------------------
console.log('\n[4] checkPackageConsistency diszpecser + MUTANS-PROBA')

check('checkPackageConsistency a core szerint a helyes al-ellenorzore iranyit (pptx)', checkPackageConsistency(pptxValid, 'pptx').ok === true)
check('  (xlsx)', checkPackageConsistency(xlsxValid, 'xlsx').ok === true)
check('  (docx)', checkPackageConsistency(docxValid, 'docx').ok === true)
{
  let dobott = false, uzenetJo = false
  try { checkPackageConsistency(pptxValid, 'odt') }
  catch (e) { dobott = true; uzenetJo = /unknown core/.test(e.message) }
  check('ismeretlen core -> dob, nevesitve', dobott && uzenetJo)
}

{
  // MUTANS: a pptx-ellenorzes "hianyzo" szamitasat szandekosan elrontjuk (a kulonbseg iranyat
  // megforditjuk), es bizonyitjuk, hogy EMIATT a fenti [1]-es PIROS-teszt zoldre valtana --
  // vagyis a mero tud bukni, nem csak atmenni.
  // A mutaciot a FORRASON hajtjuk vegre, ujra betoltve egy kulon peldanykent -- a
  // checkPackageConsistency sajat, modulon beluli hivasa NEM a module.exports-on at megy, hanem
  // a fajlon beluli nevre, tehat egy egyszeru fuggveny-csere a module.exports-on nem erne el.
  const src = fs.readFileSync(path.join(__dirname, 'package-consistency.cjs'), 'utf8')
  assert.ok(src.includes('const hianyzo = [...hivatkozott].filter((id) => !letezo.has(id)).sort()'),
    'a mutacio horgonya nem talalhato -- a forras elavult, frissiteni kell a probat')
  const mutaltSrc = src.replace(
    'const hianyzo = [...hivatkozott].filter((id) => !letezo.has(id)).sort()\n  const issues = []\n  if (hianyzo.length) issues.push(`a sldIdLst olyan kapcsolatra hivatkozik, ami a rels-ben nincs: ${JSON.stringify(hianyzo)}`)',
    'const hianyzo = [] // MUTACIO (euro-api-e8 red-proof): a hianyzo-lista mindig ures\n  const issues = []\n  if (hianyzo.length) issues.push(`a sldIdLst olyan kapcsolatra hivatkozik, ami a rels-ben nincs: ${JSON.stringify(hianyzo)}`)',
  )
  assert.notStrictEqual(mutaltSrc, src, 'a mutacio cseréje nem valtoztatott a forrason -- a horgony pontatlan')
  const mutaltPath = path.join(TMP_DIR, 'package-consistency.mutant.cjs')
  fs.writeFileSync(mutaltPath, mutaltSrc)
  delete require.cache[require.resolve(mutaltPath)]
  const mutalt = require(mutaltPath)
  const r = mutalt.checkPptx(pptxBroken)
  check('MUTANS ALKALMAZVA, ES A PIROS AG ZOLDRE VALT -- a mero tehat tud bukni, nem csak atmenni', r.ok === true, JSON.stringify(r))
}

fs.rmSync(TMP_DIR, { recursive: true, force: true })
console.log(`\n${failed === 0 ? 'MIND OK' : `${failed} FAILED`}`)
process.exit(failed === 0 ? 0 : 1)
