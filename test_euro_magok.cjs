// E7 -- A MAG-ELOALLITAS, MINDKET IRANYBAN.
//
// A leggyakoribb hiba ennel a fajta kodnal nem az, hogy elszall, hanem hogy egy SZERKEZETILEG
// ERVENYTELEN csomagot ad, ami csak a felhasznalo kepernyojen bukik el. Ezert itt minden allitas
// a KESZ CSOMAGRA vonatkozik: kicsomagoljuk es megszamoljuk, amit vartunk.

const assert = require('assert')
const zlib = require('zlib')
const { xlsxMag, pptxMag, PPTX_STATIKUS } = require('./euro-magok.cjs')

let osszes = 0
const hibak = []
function check(cimke, felteteles, reszlet = '') {
  osszes += 1
  if (felteteles) { console.log(`  ok    ${cimke}`) } else { hibak.push(cimke); console.log(`  BUKAS ${cimke}${reszlet ? ' -- ' + reszlet : ''}`) }
}
function dob(cimke, fn, mintaz) {
  osszes += 1
  try { fn(); hibak.push(cimke); console.log(`  BUKAS ${cimke} -- nem dobott`) }
  catch (err) {
    if (mintaz.test(err.message)) console.log(`  ok    ${cimke}`)
    else { hibak.push(cimke); console.log(`  BUKAS ${cimke} -- rossz uzenet: ${err.message}`) }
  }
}

// Minimalis ZIP-olvaso a KOZPONTI KONYVTARBOL. Szandekosan NEM a sajat iro-kodunkat hasznalja
// visszafele: ha ugyanaz a fuggveny irna es olvasna, egy kozos tevedes mindket iranyban egyezne.
function zipReszek(buf) {
  const reszek = {}
  let i = buf.length - 22
  while (i >= 0 && buf.readUInt32LE(i) !== 0x06054b50) i -= 1
  if (i < 0) throw new Error('nincs ZIP-vegjel')
  let p = buf.readUInt32LE(i + 16)
  const db = buf.readUInt16LE(i + 10)
  for (let k = 0; k < db; k += 1) {
    const nevHossz = buf.readUInt16LE(p + 28)
    const extraHossz = buf.readUInt16LE(p + 30)
    const megjHossz = buf.readUInt16LE(p + 32)
    const nev = buf.toString('utf8', p + 46, p + 46 + nevHossz)
    const helyi = buf.readUInt32LE(p + 42)
    const hNev = buf.readUInt16LE(helyi + 26)
    const hExtra = buf.readUInt16LE(helyi + 28)
    const meret = buf.readUInt32LE(p + 20)
    const kezd = helyi + 30 + hNev + hExtra
    reszek[nev] = zlib.inflateRawSync(buf.subarray(kezd, kezd + meret)).toString('utf8')
    p += 46 + nevHossz + extraHossz + megjHossz
  }
  return reszek
}

console.log('\n[1] XLSX -- a LAPSZAM a magbol jon (futas kozben nem lehet lapot adni)')
const x1 = zipReszek(xlsxMag(['Munka1']))
check('1 lap -> a workbook 1 lapot nevez', (x1['xl/workbook.xml'].match(/<sheet /g) || []).length === 1)
const x3 = zipReszek(xlsxMag(['Adatok', 'Osszesites', 'Jelmagyarazat']))
check('3 lap -> 3 sheet-bejegyzes', (x3['xl/workbook.xml'].match(/<sheet /g) || []).length === 3)
check('  a NEVEK is atmennek', /name="Osszesites"/.test(x3['xl/workbook.xml']), x3['xl/workbook.xml'].slice(0, 120))
check('  minden laphoz van worksheet-resz', [1, 2, 3].every((i) => `xl/worksheets/sheet${i}.xml` in x3))
check('  es minden laphoz Override a Content_Types-ban',
  (x3['[Content_Types].xml'].match(/worksheets\/sheet\d+\.xml/g) || []).length === 3)
check('  a stilus-kapcsolat a lapok UTAN kap rId-t (nem utkozik)',
  /Id="rId4"[^>]*styles\.xml/.test(x3['xl/_rels/workbook.xml.rels']), x3['xl/_rels/workbook.xml.rels'])
// A ' es < karakterek egy lapnevben ervenytelen XML-t adnanak, ha nem escape-elunk.
const xEsc = zipReszek(xlsxMag(['A & B <teszt>']))
check('  a lapnev escape-elve megy a XML-be', /name="A &amp; B &lt;teszt&gt;"/.test(xEsc['xl/workbook.xml']), xEsc['xl/workbook.xml'].slice(0, 160))

console.log('\n[2] PPTX -- a DIASZAM ugyanigy, es a beagyazott reszek is ott vannak')
const p1 = zipReszek(pptxMag(1))
check('1 dia -> 1 sldId', (p1['ppt/presentation.xml'].match(/<p:sldId /g) || []).length === 1)
const p4 = zipReszek(pptxMag(4))
check('4 dia -> 4 sldId', (p4['ppt/presentation.xml'].match(/<p:sldId /g) || []).length === 4)
check('  minden diahoz van resz ES rels', [1, 2, 3, 4].every((i) => `ppt/slides/slide${i}.xml` in p4 && `ppt/slides/_rels/slide${i}.xml.rels` in p4))
check('  a dia-mester rId-je a diak UTAN all', new RegExp(`Id="rId5"[^>]*slideMaster1\\.xml`).test(p4['ppt/_rels/presentation.xml.rels']), p4['ppt/_rels/presentation.xml.rels'])
check('  a diak URESEK (a mag a tipust hozza, nem tartalmat)', !/<a:t>/.test(p4['ppt/slides/slide1.xml']))

console.log('\n[3] *** A PROVENIENCIA-KAPU: a BEAGYAZOTT reszek kicsomagolhatok es a vart listat adjak ***')
// Egy elrontott base64-blob kulonben nem ITT bukna el, hanem a kesz dokumentumon -- es akkor a
// hiba a felhasznalo kepernyojen jelenik meg, nem a bekotesnel.
const vartStatikus = ['ppt/theme/theme1.xml', 'ppt/slideMasters/slideMaster1.xml', 'ppt/slideLayouts/slideLayout1.xml']
check('mind a harom hordozo resz beagyazva van', vartStatikus.every((n) => n in PPTX_STATIKUS), Object.keys(PPTX_STATIKUS).join(', '))
check('  es a csomagba is bekerulnek', vartStatikus.every((n) => n in p1))
check('  a tema tenyleg tema (nem ures blob)', /<a:theme|<a:clrScheme/.test(p1['ppt/theme/theme1.xml']), p1['ppt/theme/theme1.xml'].slice(0, 80))
check('  a dia-mester tenyleg mester', /<p:sldMaster/.test(p1['ppt/slideMasters/slideMaster1.xml']))

console.log('\n[4] MEGTAGADASOK ES HATAROK')
dob('0 dia -> megtagadva (ervenytelen csomag lenne)', () => pptxMag(0), /legalabb 1/)
dob('201 dia -> megtagadva (felso hatar)', () => pptxMag(201), /felso hatar/)
dob('201 lap -> megtagadva', () => xlsxMag(new Array(201).fill('x')), /felso hatar/)
const xUres = zipReszek(xlsxMag([]))
check('ures lap-lista -> egy alapertelmezett lap (nem ures csomag)', (xUres['xl/workbook.xml'].match(/<sheet /g) || []).length === 1)

console.log('\n[5] REPRODUKALHATOSAG -- ugyanaz a bemenet, BAJTRA ugyanaz a csomag')
// A ZIP-fejlecben az idobelyeg FIX; enelkul ket futas kimenete kulonbozne, es egy diff sosem
// mondana meg, hogy a TARTALOM valtozott-e vagy csak az ora.
check('xlsx: ket futas bajtra azonos', Buffer.compare(xlsxMag(['A', 'B']), xlsxMag(['A', 'B'])) === 0)
check('pptx: ket futas bajtra azonos', Buffer.compare(pptxMag(2), pptxMag(2)) === 0)
check('  es KULONBOZO bemenet KULONBOZO kimenet (a mero nem mond mindenre igent)',
  Buffer.compare(xlsxMag(['A']), xlsxMag(['A', 'B'])) !== 0)

console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
assert.strictEqual(hibak.length, 0, `bukott: ${hibak.join(' | ')}`)
