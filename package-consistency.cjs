// A KIÍRT csomag
// önmagában konzisztens-e -- ugyanaz a minta, mint egy korábban már alkalmazott önellenőrzés.
// A `unzip -t` csak azt mondja meg, hogy a ZIP ép -- azt NEM,
// hogy a belső hivatkozások feloldhatók-e. Ez pontosan azt a hibaosztályt fogja meg, ami neónál
// elrontotta a szerkesztőt: a csomag ép volt (minden rész bent), mégis üres prezentációt adott,
// mert egyetlen ATTRIBÚTUM (egy sldIdLst r:id) nem volt feloldható.
//
// Miért itt, nem lib.cjs-ben: lib.cjs saját fejléce szerint "pure and independently testable: no
// network, no secrets read from disk, no MCP transport" -- ez a modul viszont a KIÍRT fájlt (egy
// valódi zip-et lemezen) olvassa, `unzip`/`zipinfo` külső processzeken át. Az euro-mcp.cjs már
// most is fs-t hív közvetlenül (edit_document, run_builder_script), tehát ez a modul ugyanabba a
// "diszkes I/O" rétegbe tartozik, nem lib.cjs tiszta rétegébe.
const { execFileSync } = require('child_process')
const path = require('path')

// Egy zip-bejegyzés szövege, vagy null ha nincs ilyen bejegyzés -- a hívó nevesíti a hiányt,
// nem ez a függvény dobja el (a "nincs benne a csomagban" MAGA egy mérhető lelet, nem hiba az
// olvasás közben).
function readZipEntry(zipPath, entry) {
  try {
    return execFileSync('unzip', ['-p', zipPath, entry], { encoding: 'utf8' })
  } catch {
    return null
  }
}

function listZipEntries(zipPath) {
  const raw = execFileSync('zipinfo', ['-1', zipPath], { encoding: 'utf8' })
  return new Set(raw.split('\n').map((s) => s.trim()).filter(Boolean))
}

function idsMatching(xml, tagPattern) {
  return new Set([...xml.matchAll(tagPattern)].map((m) => m[1]))
}

function idsDeclaredInRels(relsXml) {
  return new Set([...relsXml.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]))
}

// pptx: minden <p:sldId r:id="..."> szerepel-e a presentation rels-ben. BYTE-AZONOS regex a
// mag-minimal-pptx.py :: onellenorzes()-ével -- ugyanaz a mérő, két nyelven.
function checkPptx(zipPath) {
  const pres = readZipEntry(zipPath, 'ppt/presentation.xml')
  const rels = readZipEntry(zipPath, 'ppt/_rels/presentation.xml.rels')
  if (pres === null || rels === null) {
    return { ok: false, issues: [`hiányzó rész a csomagban: ${pres === null ? 'ppt/presentation.xml' : 'ppt/_rels/presentation.xml.rels'}`] }
  }
  const hivatkozott = idsMatching(pres, /<p:sldId\b[^>]*\br:id="([^"]+)"/g)
  const letezo = idsDeclaredInRels(rels)
  const hianyzo = [...hivatkozott].filter((id) => !letezo.has(id)).sort()
  const issues = []
  if (hianyzo.length) issues.push(`a sldIdLst olyan kapcsolatra hivatkozik, ami a rels-ben nincs: ${JSON.stringify(hianyzo)}`)
  if (!hivatkozott.size) issues.push('a sldIdLst ÜRES -- a szerkesztő üres prezentációt látna')
  return { ok: issues.length === 0, issues }
}

// xlsx: minden <sheet r:id="..."> szerepel-e a workbook rels-ben -- a pptx-ellenőrzés xlsx-
// megfelelője (a csomag TARTALMA hivatkozik a rels-re, ugyanaz az irány).
function checkXlsx(zipPath) {
  const wb = readZipEntry(zipPath, 'xl/workbook.xml')
  const rels = readZipEntry(zipPath, 'xl/_rels/workbook.xml.rels')
  if (wb === null || rels === null) {
    return { ok: false, issues: [`hiányzó rész a csomagban: ${wb === null ? 'xl/workbook.xml' : 'xl/_rels/workbook.xml.rels'}`] }
  }
  const hivatkozott = idsMatching(wb, /<sheet\b[^>]*\br:id="([^"]+)"/g)
  const letezo = idsDeclaredInRels(rels)
  const hianyzo = [...hivatkozott].filter((id) => !letezo.has(id)).sort()
  const issues = []
  if (hianyzo.length) issues.push(`a workbook olyan lap-kapcsolatra hivatkozik, ami a rels-ben nincs: ${JSON.stringify(hianyzo)}`)
  if (!hivatkozott.size) issues.push('a workbook <sheet> listája ÜRES -- egy munkafüzet lap nélkül érvénytelen')
  return { ok: issues.length === 0, issues }
}

// docx: a MÁSIK irány -- a rels-ben NEVEZETT részek tényleg léteznek-e a csomagban. A docx
// tartalom-hivatkozások (drawing r:embed, stb.) száma és alakja műveletenként változik, a rels
// viszont mindig egy zárt lista, aminek minden Target-je egy valódi résznek kell megfeleljen
// (kivéve a TargetMode="External" bejegyzéseket -- azok URL-ek, nem csomag-részek).
//
// *** MÉRVE (a valódi fake Document Server ellen, nem feltételezve): ***
// egy relációt-mentes minimál docx-nek EGYÁLTALÁN NINCS word/_rels/document.xml.rels resze --
// ez ELTÉR a pptx/xlsx mintától (ahol a sldIdLst/sheets lista MINDIG kötelező, és az üresség
// maga a hiba), mert egy docx-nek nem kell legalább egy kapcsolat ahhoz, hogy érvényes legyen.
// A hiányzó vagy üres rels tehát ITT NEM hiba -- csak akkor, ha a rels LÉTEZIK, DE egy nem
// létező részre hivatkozik.
function checkDocx(zipPath) {
  const rels = readZipEntry(zipPath, 'word/_rels/document.xml.rels')
  if (rels === null) return { ok: true, issues: [] } // nincs kapcsolat -> nincs mit ellenorizni
  const entries = listZipEntries(zipPath)
  const relTags = rels.match(/<Relationship\b[^>]*\/>/g) || []
  const issues = []
  for (const tag of relTags) {
    const targetMatch = /\bTarget="([^"]+)"/.exec(tag)
    if (!targetMatch) continue
    const modeMatch = /\bTargetMode="([^"]+)"/.exec(tag)
    if (modeMatch && modeMatch[1] === 'External') continue
    const resolved = path.posix.normalize(path.posix.join('word', targetMatch[1]))
    if (!entries.has(resolved)) {
      issues.push(`a rels egy nem létező részre hivatkozik: ${targetMatch[1]} (várt csomag-út: ${resolved})`)
    }
  }
  return { ok: issues.length === 0, issues: issues.sort() }
}

function checkPackageConsistency(zipPath, core) {
  if (core === 'pptx') return checkPptx(zipPath)
  if (core === 'xlsx') return checkXlsx(zipPath)
  if (core === 'docx') return checkDocx(zipPath)
  throw new Error(`checkPackageConsistency: unknown core: ${core}`)
}

module.exports = { checkPackageConsistency, checkPptx, checkXlsx, checkDocx }
