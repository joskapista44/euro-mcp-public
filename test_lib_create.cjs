// E0 -- THE OPERATION SCHEMA AND ITS TRANSLATOR, both directions.
//
// The point of this file is not that a script comes out. It is that a script does NOT come out
// when it should not: an unknown operation type, an operation the core cannot do, an empty table,
// an image without a source. Every one of those was a silent no-op somewhere else today, and a
// silent no-op is reported to the caller as success.

const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')
const { buildCreateScript, OPERATIONS, CREATE_ROUTE } = require('./lib.cjs')

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
    if (mintaz.test(err.message)) { console.log(`  ok    ${cimke}`) }
    else { hibak.push(cimke); console.log(`  BUKAS ${cimke} -- rossz uzenet: ${err.message}`) }
  }
}

console.log('\n[1] A HAROM MUVELET ATMEGY, ES A SCRIPTBEN NYOMOT HAGY')
const harom = buildCreateScript({
  core: 'docx',
  operations: [
    { type: 'text', text: 'cim', bold: true, size: 32 },
    { type: 'table', rows: [['a', 'b'], ['c', 'd']] },
    { type: 'image', src: 'data:image/png;base64,AAAA' },
  ],
})
check('docx: harom muvelet -> script', /OpenFile/.test(harom.script) && /SaveFile\("docx"/.test(harom.script))
check('  a szoveg benne van', /AddText\("cim"\)/.test(harom.script), harom.script.slice(0, 80))
check('  a tabla 2x2-kent all elo', /CreateTable\(2, 2\)/.test(harom.script))
check('  a kep a drawing-uton megy', /CreateImage\("data:image\/png;base64,AAAA"/.test(harom.script))
check('  es az `applied` MINDHAROMRA megmondja az utat',
  harom.applied.length === 3 && harom.applied.every((a) => a.sourceRoute === CREATE_ROUTE),
  JSON.stringify(harom.applied))

console.log('\n[2] MAGONKENT MAS -- es a script is mas')
const x = buildCreateScript({ core: 'xlsx', operations: [{ type: 'table', at: 'B3', rows: [['x', 'y']] }] })
check('xlsx: a tabla CELLAKRA fordul, a megadott helyrol', /GetRange\("B3"\).SetValue\("x"\)/.test(x.script) && /GetRange\("C3"\).SetValue\("y"\)/.test(x.script), x.script)
const p = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'dia' }] })
check('pptx: a szoveg SZOVEGDOBOZ lesz, nem bekezdes', /CreateShape/.test(p.script) && /oSlide.AddObject/.test(p.script))

console.log('\n[3] *** A MEGTAGADASOK -- EZEK A FONTOSABBAK ***')
dob('ismeretlen tipus -> NEVESITETT hiba (a nev is benne)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'nincs-ilyen' }] }), /unknown type "nincs-ilyen"/)
dob('  es felsorolja, mi LETEZIK (kulonben a hivo talalgat)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'nincs-ilyen' }] }), /known: .*text.*table.*image/)
dob('a magban nem letezo muvelet -> nevesitett hiba, nem csendes kihagyas',
  () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'text', text: 'x' }] }), /not available in the xlsx core/)
dob('URES tabla -> megtagadva (egy ures tabla atmeno tesztnek latszik)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [] }] }), /empty/)
dob('kep forras nelkul -> megtagadva',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'image' }] }), /src.*required/)
dob('ismeretlen mag -> megtagadva', () => buildCreateScript({ core: 'odt', operations: [{ type: 'text' }] }), /unknown core/)
dob('NULLA muvelet -> megtagadva (a script no-op lenne)',
  () => buildCreateScript({ core: 'docx', operations: [] }), /no operations/)

console.log('\n[4] A TABLAZAT MAGA IS ADAT: melyik muvelet melyik magban all')
check('minden muvelethez tartozik mag-lista', Object.values(OPERATIONS).every((s) => Array.isArray(s.cores) && s.cores.length))
check('a `text` NEM all az xlsx-en (ez szandekos, nem hiany)', !OPERATIONS.text.cores.includes('xlsx'))

console.log('\n[5] DIAGRAM -- magonkent MAS hivas-alak')
const docxChart = buildCreateScript({ core: 'docx', operations: [{ type: 'chart', chartType: 'bar', series: [[1, 2, 3], [4, 5, 6]], seriesNames: ['A', 'B'], categories: ['X', 'Y', 'Z'] }] })
check('docx: Api.CreateChart + AddDrawing egy bekezdesre', /Api\.CreateChart\("bar"/.test(docxChart.script) && /oChartPara\.AddDrawing\(oChart\)/.test(docxChart.script) && /oDocument\.Push\(oChartPara\)/.test(docxChart.script))
check('  az adatsorok szo szerint a hivasban vannak (literal, nincs cellahivatkozas)', docxChart.script.includes('[[1,2,3],[4,5,6]]'))

const pptxChart = buildCreateScript({ core: 'pptx', operations: [{ type: 'chart', chartType: 'lineNormal', series: [[1, 2]], categories: ['a', 'b'] }] })
check('pptx: Api.CreateChart + SetPosition + oSlide.AddObject', /Api\.CreateChart\("lineNormal"/.test(pptxChart.script) && /oChart\.SetPosition\(/.test(pptxChart.script) && /oSlide\.AddObject\(oChart\)/.test(pptxChart.script))
check('  hianyzo seriesNames eseten a tool sajat, sorszamozott nevet ad (nem hibazik, nem hagyja el)', /\["Sorozat 1"\]/.test(pptxChart.script))

const xlsxChart = buildCreateScript({ core: 'xlsx', operations: [{ type: 'chart', chartType: 'pie', dataRange: 'Munka1!$A$1:$B$3' }] })
check('xlsx: Worksheet.AddChart CELLATARTOMANYBOL, NEM Api.CreateChart (a card sajat leletet -- CreateChart mindig bukik xlsx-en)', /oWorksheet\.AddChart\("Munka1!\$A\$1:\$B\$3", false, "pie"/.test(xlsxChart.script) && !/Api\.CreateChart/.test(xlsxChart.script))

console.log('\n[6] DIAGRAM-MEGTAGADASOK -- nevesitett hiba, NEM nema visszaeses az alapertelmezettre')
dob('ismeretlen chartType -> NEVESITETT hiba, es felsorolja az ismerteket (a tipus MERT allowlist, nem athivas)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'chart', chartType: 'nincs-ilyen-tipus', series: [[1]], categories: ['x'] }] }),
  /unknown chartType "nincs-ilyen-tipus"/)
dob('  -- es kimondja, MIERT allowlist, nem athivas (a mert csendes-noop lelet)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'chart', chartType: 'nincs-ilyen-tipus', series: [[1]], categories: ['x'] }] }),
  /does NOT throw inside the builder, it silently produces no chart/)
dob('xlsx: `dataRange` nelkul -> megtagadva (celladat kell, nem literal tomb)',
  () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'chart', chartType: 'bar' }] }),
  /`dataRange`.*required for xlsx/)
dob('docx/pptx: `series` nelkul -> megtagadva',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'chart', chartType: 'bar', categories: ['x'] }] }),
  /`series`.*required/)
dob('docx/pptx: `categories` nelkul -> megtagadva',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'chart', chartType: 'bar', series: [[1, 2]] }] }),
  /`categories`.*required/)
dob('docx/pptx: `series` ures tombokkel -> megtagadva',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'chart', chartType: 'bar', series: [[]], categories: ['x'] }] }),
  /non-empty array of numbers/)
dob('chartType hianya -> ugyanaz a nevesitett "unknown chartType" ag (ures string sincs az allowlisten)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'chart', series: [[1]], categories: ['x'] }] }),
  /unknown chartType ""/)

console.log('\n[6b] DIAGRAM-FORMAZAS, pptx-only -- MIND opcionalis, semmi nem fut, ha nem kertek')
const chartAlap = buildCreateScript({ core: 'pptx', operations: [{ type: 'chart', chartType: 'bar', series: [[1, 2]], categories: ['a', 'b'] }] })
check('formazasi mezok NELKUL a script NEM tartalmaz egyetlen uj hivast sem (nulla-viselkedesvaltozas)',
  !/SetTitle|SetLegendPos|SetHorAxisTitle|SetVerAxisTitle|SetShowDataLabels/.test(chartAlap.script), chartAlap.script)

const chartTeljes = buildCreateScript({
  core: 'pptx',
  operations: [{
    type: 'chart', chartType: 'bar', series: [[1, 2]], categories: ['a', 'b'],
    title: 'Cim-teszt', legendPos: 'right', horAxisTitle: 'X-tengely', verAxisTitle: 'Y-tengely', showDataLabels: true,
  }],
})
check('title -> oChart.SetTitle(...)', /oChart\.SetTitle\("Cim-teszt"\);/.test(chartTeljes.script), chartTeljes.script)
check('legendPos -> oChart.SetLegendPos(...)', /oChart\.SetLegendPos\("right"\);/.test(chartTeljes.script))
check('horAxisTitle -> oChart.SetHorAxisTitle(...)', /oChart\.SetHorAxisTitle\("X-tengely"\);/.test(chartTeljes.script))
check('verAxisTitle -> oChart.SetVerAxisTitle(...)', /oChart\.SetVerAxisTitle\("Y-tengely"\);/.test(chartTeljes.script))
check('showDataLabels -> oChart.SetShowDataLabels(false, false, true, false, false, false) -- MERT (6-pozicios seprés): a 3. pozicio a `showVal`, a hagyomanyos "adatfelirat" jelentes; a masik ot pozicionak ezen a peldanyon NEM volt megfigyelheto hatasa',
  /oChart\.SetShowDataLabels\(false, false, true, false, false, false\);/.test(chartTeljes.script))

check('a formazo hivasok az AddObject ELOTT allnak (a chart mar all, mire a slide-hoz adjuk)',
  chartTeljes.script.indexOf('SetTitle') < chartTeljes.script.indexOf('oSlide.AddObject(oChart)'))

dob('legendPos: ismeretlen ertek -> NEVESITETT hiba (mert allowlist, nem athivas -- ugyanaz az osztaly, mint a chartType)',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'chart', chartType: 'bar', series: [[1]], categories: ['x'], legendPos: 'kozepen' }] }),
  /unknown legendPos "kozepen"/)
dob('  -- es kimondja, MIERT allowlist', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'chart', chartType: 'bar', series: [[1]], categories: ['x'], legendPos: 'kozepen' }] }),
  /does NOT throw.*it silently produces no legend/)
dob('legendPos: "none" ELUTASITVA -- MERT (sweep), hogy "none" es egy kitalalt nev BAJTRA egyforman viselkedik (nincs <c:legend> egyik esetben sem), tehat NEM bizonyithato kulon tamogatott ertekkent',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'chart', chartType: 'bar', series: [[1]], categories: ['x'], legendPos: 'none' }] }),
  /unknown legendPos "none"/)

check('title/legendPos/axisTitles/showDataLabels docx-en es xlsx-en NEM ertelmezett (a mai `chart` mag-agai valtozatlanok -- NEM dobnak, csak nem hasznaljak fel)',
  (() => {
    const d = buildCreateScript({ core: 'docx', operations: [{ type: 'chart', chartType: 'bar', series: [[1]], categories: ['x'], title: 'X' }] })
    return !/SetTitle/.test(d.script)
  })())

console.log('\n[7] SZOVEG-BEKEZDES BOVITES -- CSAK docx, es CSAK opcionalis mezo')
check('a regi harom mezos hivas EGYETLEN sort sem valtoztat (visszafele-kompatibilitas)',
  buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'cim', bold: true, size: 32 }] }).script ===
  [
    'builder.OpenFile("__DOC_URL__");',
    'var oDocument = Api.GetDocument();',
    'var oParagraph = Api.CreateParagraph();',
    'oParagraph.AddText("cim");',
    'oParagraph.SetBold(true);',
    'oParagraph.SetFontSize(32);',
    'oDocument.Push(oParagraph);',
    'builder.SaveFile("docx", "eredmeny.docx");',
    'builder.CloseFile();',
  ].join('\n') + '\n')

const dus = buildCreateScript({
  core: 'docx',
  operations: [{
    type: 'text', text: 'dus bekezdes', heading: 2, italic: true, underline: true, strikethrough: true,
    color: '1a5fb4', highlight: 'yellow', align: 'center', font: 'Georgia', size: 40,
  }],
})
check('heading -> GetStyle("Heading N") + SetStyle', /oParagraph\.SetStyle\(oDocument\.GetStyle\("Heading 2"\)\);/.test(dus.script))
check('italic', /oParagraph\.SetItalic\(true\);/.test(dus.script))
check('underline', /oParagraph\.SetUnderline\(true\);/.test(dus.script))
check('strikethrough -> SetStrikeout', /oParagraph\.SetStrikeout\(true\);/.test(dus.script))
check('color -> hex parsed to r,g,b', /oParagraph\.SetColor\(26, 95, 180, false\);/.test(dus.script), dus.script)
check('highlight', /oParagraph\.SetHighlight\("yellow"\);/.test(dus.script))
check('align -> SetJc', /oParagraph\.SetJc\("center"\);/.test(dus.script))
check('font family', /oParagraph\.SetFontFamily\("Georgia"\);/.test(dus.script))
check('color with a leading # is accepted the same way',
  buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', color: '#1a5fb4' }] }).script.includes('SetColor(26, 95, 180, false)'))

console.log('\n[8] SZOVEG-BEKEZDES MEGTAGADASOK -- nevesitett hiba, csendes visszaeses nincs')
dob('heading 0 -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', heading: 0 }] }), /heading.*integer 1-9/)
dob('heading 10 -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', heading: 10 }] }), /heading.*integer 1-9/)
dob('heading nem egesz -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', heading: 1.5 }] }), /heading.*integer 1-9/)
dob('ismeretlen align -> NEVESITETT hiba (a mert csendes-noop lelet szerepel a szovegben)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', align: 'nincsilyen' }] }),
  /unknown align "nincsilyen".*does NOT throw inside the builder, it silently produces no alignment/)
dob('ismeretlen highlight -> NEVESITETT hiba', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', highlight: 'nincsilyenszin' }] }), /unknown highlight "nincsilyenszin"/)
dob('rossz szinformatum -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', color: 'nem-egy-szin' }] }), /color.*6-digit hex string/)
dob('rovid hex -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', color: 'fff' }] }), /color.*6-digit hex string/)
// listType MOSTANTOL MUKODIK --
// oDocument.CreateNumbering(...) (NEM Api.CreateNumbering(...)), package-verified on the real
// docbuilder-create route (numbering.xml/document.xml evidence).
// A rossz `listType`/`listLevel` ERTEK tovabbra is NEVESITETT hibat dob -- ez nem valtozott.
dob('listType: ismeretlen ertek -> NEVESITETT hiba', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', listType: 'roman' }] }), /listType.*must be "bullet" or "numbered"/)
dob('listLevel: tartomanyon kivul -> NEVESITETT hiba', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', listType: 'numbered', listLevel: 9 }] }), /listLevel.*must be an integer 0-8/)
{
  const numbered = buildCreateScript({ core: 'docx', operations: [
    { type: 'text', text: 'elso', listType: 'numbered', listLevel: 0 },
    { type: 'text', text: 'masodik', listType: 'numbered', listLevel: 1 },
  ] })
  // A `var oNumbering_numbered = (typeof ... !== 'undefined' && ...) ? ... : oDocument.
  // CreateNumbering(...)` sor OPERANKENT emitalodik (guard, nem source-szintu egyszerihivas) --
  // ket op -> ket sor a forrasban, DE mindket sor UGYANAZT a valtozonevet celozza, es a guard
  // futaskor csak az elsonel hoz letre uj numbering-et (live, package-verified a kartyan: 5
  // "numbered" op -> egyetlen numId a mentett csomagban). Itt a statikus invariáns, amit ez a
  // teszt ellenoriz: mindket op UGYANARRA a valtozonevre hivatkozik, nem ket kulonre.
  check('listType numbered: mindket bekezdes ugyanazt az oNumbering_numbered valtozonevet celozza (egy lista, nem ketto)',
    (numbered.script.match(/var oNumbering_numbered =/g) || []).length === 2 && !/oNumbering_numbered_/.test(numbered.script), numbered.script)
  check('listType numbered: a ket szint ilvl-je 0 es 1 a generalt hivasban', /SetNumPr\(oNumbering_numbered, 0\)/.test(numbered.script) && /SetNumPr\(oNumbering_numbered, 1\)/.test(numbered.script))
  const bullet = buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', listType: 'bullet' }] })
  check('listType bullet: sajat oNumbering_bullet valtozot hasznal, nem a numberedet', /oDocument\.CreateNumbering\("bullet"\)/.test(bullet.script) && !/CreateNumbering\("numbered"\)/.test(bullet.script))
  const sima = buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x' }] })
  check('listType nelkul: nincs CreateNumbering/SetNumPr a generalt scriptben (nulla-viselkedesvaltozas)', !/CreateNumbering|SetNumPr/.test(sima.script))
}

console.log('\n[8c] PPTX listType: bullet ES numbered is MUKODIK, mindketto SetBullet-en at')
{
  const bullet = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', listType: 'bullet' }], slideCount: 1 })
  check('pptx listType bullet: Api.CreateBullet("•") az alap karakter, SetBullet a hivas', /oPara\.SetBullet\(Api\.CreateBullet\("•"\)\);/.test(bullet.script), bullet.script)
  const bulletChar = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', listType: 'bullet', bulletChar: 'o' }], slideCount: 1 })
  check('pptx listType bullet: `bulletChar` felulirja az alapertelmezettet', /oPara\.SetBullet\(Api\.CreateBullet\("o"\)\);/.test(bulletChar.script), bulletChar.script)
  const sima = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x' }], slideCount: 1 })
  check('pptx listType nelkul: nincs SetBullet a generalt scriptben (nulla-viselkedesvaltozas)', !/SetBullet/.test(sima.script), sima.script)

  // Api.CreateNumbering(numType)
  // egy MASODIK ApiBullet-konstruktor, ugyanahhoz a SetBullet hivashoz -- NEM ApiNumbering/SetNumPr,
  // ahogy a korabbi megtagadas feltetelezte. Package-verified elo probaval (<a:buAutoNum
  // type="arabicPeriod">).
  const numbered = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', listType: 'numbered' }], slideCount: 1 })
  check('pptx listType numbered: Api.CreateNumbering("ArabicPeriod") az alapertelmezett tipus, SetBullet a hivas', /oPara\.SetBullet\(Api\.CreateNumbering\("ArabicPeriod"\)\);/.test(numbered.script), numbered.script)
  const numberedType = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', listType: 'numbered', numType: 'RomanUcPeriod' }], slideCount: 1 })
  check('pptx listType numbered: `numType` felulirja az alapertelmezettet', /oPara\.SetBullet\(Api\.CreateNumbering\("RomanUcPeriod"\)\);/.test(numberedType.script), numberedType.script)
  dob('pptx listType numbered: ismeretlen `numType` -> NEVESITETT hiba, nem csendes eltakaras',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', listType: 'numbered', numType: 'roman' }], slideCount: 1 }),
    /numType.*must be one of ArabicPeriod/)
  dob('pptx listType ismeretlen ertek: NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', listType: 'roman' }], slideCount: 1 }), /listType must be "bullet" or "numbered"/)
  const multi = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', paragraphs: [{ text: 'a', listType: 'bullet' }, { text: 'b' }] }], slideCount: 1 })
  check('pptx listType `paragraphs` tombben, PER-BEKEZDES: csak az elso kapja a SetBullet-et', (multi.script.match(/SetBullet/g) || []).length === 1, multi.script)
}

console.log('\n[9] URES `operations` -> a kimenet a maggal byte-azonos marad (a hivas maga meg sem tortenik)')
check('mar a szkript-generalas megtagadva, tehat script SOSEM keszul, a mag valtozatlan marad -- ez a globalis "no operations" kapu (buildCreateScript), amit ez a kartya csak MEGEROSITI, nem ismetel',
  (() => { try { buildCreateScript({ core: 'docx', operations: [] }); return false } catch (e) { return /no operations/.test(e.message) } })())

console.log('\n[10] E2 -- TABLAZAT STILUS (fejlec, szegely, sávozás, egyesites): script-alak,')
console.log('    a csomag-szintu igazolas a receptek-*.md-ben all, elo DS-hivassal, nem itt')
const stilusRows = [['A', 'B'], ['1', '2'], ['3', '4'], ['5', '6']]

const docxAlap = buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: stilusRows }] })
check('docx: fejlec ALAPBOL bekapcsolva -> SetBold a 0. soron', /GetCell\(0, 0\)\.GetContent\(\)\.GetElement\(0\)\.SetBold\(true\)/.test(docxAlap.script))
check('docx: szegely ALAPBOL bekapcsolva -> mind a hat SetTableBorder* hivas', ['Top', 'Bottom', 'Left', 'Right', 'InsideH', 'InsideV'].every((s) => docxAlap.script.includes(`SetTableBorder${s}(`)))
check('docx: zebra ALAPBOL KI -- nincs SetShd a 2. soron', !/GetCell\(2, 0\)\.SetShd/.test(docxAlap.script))
// E2c: a TABLA-szintu SetWidth
// mertekegysegere korabban NEM volt asszercio -- a "percent" -> "dxa" mutans 33/33 zoldet
// hagyott, mert "dxa" ezen a Document Serveren MERT csendes no-op (tcW w:type="auto"
// valtozatlanul jon vissza). A per-oszlop SetWidth-nek MAR van "percent" asszercioja
// ([6] szekcio lent); ez a sor a TABLA-szintu hivast fedi, ugyanazzal a fegyelemmel.
check('docx: a TABLA-szintu SetWidth "percent"-tel megy, ALAPBOL 100-zal', /oTable\.SetWidth\("percent", 100\)/.test(docxAlap.script), docxAlap.script)
check('docx: widthPercent -> a TABLA-szintu SetWidth a MEGADOTT erteket hordozza', /oTable\.SetWidth\("percent", 42\)/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: stilusRows, widthPercent: 42 }] }).script))

const docxStilus = buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: stilusRows, zebra: true, header: false, border: false }] })
check('docx: header:false -> a 0. sor NEM kap SetBold-ot', !/GetCell\(0, 0\)\.GetContent\(\)\.GetElement\(0\)\.SetBold\(true\)/.test(docxStilus.script))
check('docx: border:false -> egy SetTableBorder* hivas sincs', !/SetTableBorder/.test(docxStilus.script))
check('docx: zebra:true, header:false -> a 0. sor (paros index) kap SetShd-t', /GetCell\(0, 0\)\.SetShd/.test(docxStilus.script))

const pptxStilus = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, headerColor: [1, 2, 3] }] })
check('pptx: a fejlec-run SetBold + egyedi RGB SetColor', /oRun_0_0\.SetBold\(true\);/.test(pptxStilus.script) && /oRun_0_0\.SetColor\(255, 255, 255\);/.test(pptxStilus.script))
check('pptx: a cella-szegely (size_mm, ApiFill) alakban megy, NEM Stroke-objektummal', /SetCellBorderTop\(0\.35, Api\.CreateSolidFill/.test(pptxStilus.script) && !/SetCellBorderTop\(Api\.CreateStroke/.test(pptxStilus.script))

const xlsxStilus = buildCreateScript({ core: 'xlsx', operations: [{ type: 'table', rows: stilusRows, merge: ['A1:B1'] }] })
check('xlsx: fejlec-sor SetBold + SetFillColor + SetFontColor Range-szinten', /GetRange\("A1:B1"\)\.SetBold\(true\)/.test(xlsxStilus.script) && /GetRange\("A1:B1"\)\.SetFillColor/.test(xlsxStilus.script))
check('xlsx: szegely a HAT VALODI pozicioval megy, "All"/"Outline" SOHA (mert csendes no-op)', !/SetBorders\("All"/.test(xlsxStilus.script) && !/SetBorders\("Outline"/.test(xlsxStilus.script) && ['Left', 'Top', 'Right', 'Bottom', 'InsideVertical', 'InsideHorizontal'].every((p) => xlsxStilus.script.includes(`SetBorders("${p}"`)))
check('xlsx: a merge a MEGADOTT tartomanyra megy, false-zal (egy-cellas, nem soronkenti)', /GetRange\("A1:B1"\)\.Merge\(false\)/.test(xlsxStilus.script))
check('xlsx: merge nelkul (alap) egyetlen `.Merge(` hivas sincs', !/\.Merge\(/.test(docxAlap.script) && !buildCreateScript({ core: 'xlsx', operations: [{ type: 'table', rows: stilusRows }] }).script.includes('.Merge('))

console.log('\n[10c] pptx CELLA-BEALLITASOK -- margo (twips), fuggoleges igazitas, szovegirany')
console.log('    a hivas-alakok toString()-bol lettek visszafejtve; csomag-szintu igazolas (marT/marB/marL/marR + anchor + vert) elo DS-hivassal, nem itt')

const cellSettingsRows = [['A', 'B'], ['1', '2']]
const pptxMargin = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: cellSettingsRows, cellMargin: { top: 500, bottom: 500, left: 300, right: 300 } }] })
check('pptx: cellMargin MIND A NEGY oldala MINDEN cellara emittalodik (2x2 tablanal 4x4=16 hivas)',
  (pptxMargin.script.match(/SetCellMarginTop\(500\)/g) || []).length === 4 &&
  (pptxMargin.script.match(/SetCellMarginBottom\(500\)/g) || []).length === 4 &&
  (pptxMargin.script.match(/SetCellMarginLeft\(300\)/g) || []).length === 4 &&
  (pptxMargin.script.match(/SetCellMarginRight\(300\)/g) || []).length === 4, pptxMargin.script)

const pptxMarginPartial = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: cellSettingsRows, cellMargin: { top: 200 } }] })
check('pptx: cellMargin CSAK a megadott oldalakra hiv -- top nelkul nincs Bottom/Left/Right hivas',
  /SetCellMarginTop\(200\)/.test(pptxMarginPartial.script) && !/SetCellMarginBottom/.test(pptxMarginPartial.script) && !/SetCellMarginLeft/.test(pptxMarginPartial.script) && !/SetCellMarginRight/.test(pptxMarginPartial.script))

const pptxVAlign = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: cellSettingsRows, cellVAlign: 'center' }] })
check('pptx: cellVAlign MINDEN cellara SetVerticalAlign-et hiv, a megadott ertekkel', (pptxVAlign.script.match(/SetVerticalAlign\("center"\)/g) || []).length === 4)

const pptxTextDir = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: cellSettingsRows, cellTextDirection: 'tbrl' }] })
check('pptx: cellTextDirection MINDEN cellara SetTextDirection-t hiv, a megadott ertekkel', (pptxTextDir.script.match(/SetTextDirection\("tbrl"\)/g) || []).length === 4)

check('pptx: cella-beallitasok NELKUL (alap) egyik uj hivas sem szerepel', !/SetCellMargin|SetVerticalAlign|SetTextDirection/.test(pptxStilus.script))

dob('pptx: cellVAlign ismeretlen ertek -> NEVESITETT hiba (mert allowlist: a builder csendben hagyja allitatlanul)',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: cellSettingsRows, cellVAlign: 'kozepen' }] }),
  /unknown cellVAlign "kozepen"/)
dob('pptx: cellTextDirection ismeretlen ertek -> NEVESITETT hiba',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: cellSettingsRows, cellTextDirection: 'fuggoleges' }] }),
  /unknown cellTextDirection "fuggoleges"/)
dob('pptx: cellMargin negativ szammal -> NEVESITETT hiba',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: cellSettingsRows, cellMargin: { top: -5 } }] }),
  /cellMargin\.top must be a non-negative integer/)

check('docx: cellMargin/cellVAlign/cellTextDirection mezok pptx-specifikusak -- docx-nal jelenletuk semmilyen uj hivast nem valt ki',
  !/SetCellMargin|SetVerticalAlign|SetTextDirection/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: cellSettingsRows, cellMargin: { top: 500 }, cellVAlign: 'center', cellTextDirection: 'tbrl' }] }).script))

console.log('\n[10b] pptx TABLAZAT-EGYESITES -- Table.MergeCells([cell,cell,...])')
console.log('    a valodi hivas-alak toString()-bol lett visszafejtve -- a KEZENFEKVO alak,')
console.log('    MergeCells(startRow,startCol,endRow,endCol), lefutott (nem dobott), DE a mentett csomagban')
console.log('    SEMMIT nem valtoztatott -- csomag-szintu igazolas elo DS-hivassal, nem itt')

const pptxMerge = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, merge: [[0, 0, 0, 1]] }] })
check('pptx: egy [0,0,0,1] tartomany -> MergeCells([oTable.GetRow(0).GetCell(0), oTable.GetRow(0).GetCell(1)])',
  /oTable\.MergeCells\(\[oTable\.GetRow\(0\)\.GetCell\(0\), oTable\.GetRow\(0\)\.GetCell\(1\)\]\);/.test(pptxMerge.script), pptxMerge.script)
check('pptx: a MergeCells hivas a cellak feltoltese ES az oSlide.AddObject KOZOTT all (nem elotte, nem utana)',
  pptxMerge.script.indexOf('oRun_0_1.AddText') < pptxMerge.script.indexOf('MergeCells') &&
  pptxMerge.script.indexOf('MergeCells') < pptxMerge.script.indexOf('oSlide.AddObject(oTable)'))

const pptxMergeRect = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, merge: [[1, 0, 2, 1]] }] })
check('pptx: egy 2x2-es teglalap [1,0,2,1] MIND A NEGY cellajat listazza a MergeCells hivasban, sor-oszlop sorrendben',
  /oTable\.MergeCells\(\[oTable\.GetRow\(1\)\.GetCell\(0\), oTable\.GetRow\(1\)\.GetCell\(1\), oTable\.GetRow\(2\)\.GetCell\(0\), oTable\.GetRow\(2\)\.GetCell\(1\)\]\);/.test(pptxMergeRect.script), pptxMergeRect.script)

const pptxMergeTwo = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, merge: [[0, 0, 0, 1], [2, 0, 3, 0]] }] })
check('pptx: KET fuggetlen tartomany -> KET kulon MergeCells hivas', (pptxMergeTwo.script.match(/MergeCells/g) || []).length === 2)

check('pptx: merge NELKUL (alap) egyetlen MergeCells hivas sincs', !/MergeCells/.test(pptxStilus.script))

dob('pptx: merge[idx] nem 4-elemu tomb -> NEVESITETT hiba',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, merge: [[0, 0, 0]] }] }),
  /merge\[0\] must be a \[startRow, startCol, endRow, endCol\] tuple/)
dob('pptx: merge[idx] negativ szamot tartalmaz -> NEVESITETT hiba',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, merge: [[-1, 0, 0, 1]] }] }),
  /merge\[0\] must be a \[startRow, startCol, endRow, endCol\] tuple/)
dob('pptx: merge[idx] vege a kezdet ELOTT -> NEVESITETT hiba, nem csendben felcsereli',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, merge: [[0, 1, 0, 0]] }] }),
  /merge\[0\] end must not be before start/)
dob('pptx: merge[idx] a tablan KIVULRE mutat -> NEVESITETT hiba (a stilusRows 4 sor x 2 oszlop)',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: stilusRows, merge: [[0, 0, 4, 0]] }] }),
  /merge\[0\] end \(4,0\) is outside the table \(4 rows x 2 cols\)/)

// A korabbi "docx: merge mezo pptx-en/docx-on kivul ... docx-nal NEM okoz MergeCells hivast"
// ellenorzes ELTAVOLITVA: docx-nak SAJAT, fuggetlen merge tamogatasa van (alabb az [35]
// szekcioban fedve UGYANEZZEL a tuple-alakkal) -- a keret allitasa ("merge pptx-specifikus
// API") idokozben mar HAMIS volt.

console.log('\n[11] E2b -- OSZLOPSZELESSEG: docx MUKODIK (percent), pptx NEVESITETT NO')
console.log('    a csomag-szintu igazolas (10/30/60 -> tcW 500/1500/3000) elo DS-hivassal tortent, nem itt')
const owRows = [['A', 'B'], ['1', '2']]

const docxOw = buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: owRows, columnWidths: [30, 70] }] })
check('docx: columnWidths -> SetTableLayout("fixed") a tabla letrehozasa utan', /SetTableLayout\("fixed"\)/.test(docxOw.script))
check('docx: columnWidths -> MINDEN sor MINDEN cellaja kapja a sajat oszlopa szazalekat, "percent"-tel', /GetCell\(0, 0\)\.SetWidth\("percent", 30\)/.test(docxOw.script) && /GetCell\(1, 0\)\.SetWidth\("percent", 30\)/.test(docxOw.script) && /GetCell\(0, 1\)\.SetWidth\("percent", 70\)/.test(docxOw.script))
check('docx: columnWidths NELKUL nincs SetTableLayout es nincs SetWidth cellan', !/SetTableLayout/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: owRows }] }).script))

dob('pptx: columnWidths -> NEVESITETT hiba, nem csendes kihagyas (MERT: nincs ilyen setter ezen a DS-en)',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: owRows, columnWidths: [30, 70] }] }),
  /columnWidths.*not available in the pptx core/)
check('pptx: columnWidths NELKUL a tabla tovabbra is athalad (a tiltas csak a mezo jelenletere all)',
  /CreateTable\(2, 2\)/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: owRows }] }).script))
console.log('\n[12] MUNKAFUZET-MUVELETEK -- a receptek-xlsx.md hivas-alakjai')
const wb = buildCreateScript({
  core: 'xlsx',
  operations: [
    { type: 'numberFormat', at: 'D2', format: '#,##0.00' },
    { type: 'formula', at: 'C2', formula: '=SUM(B2:B4)' },
    { type: 'columnWidth', column: 'B', width: 25 },
    { type: 'autoFilter', range: 'A1:C1' },
    { type: 'sort', range: 'A1:B4', keyCell: 'A1' },
    { type: 'conditionalFormatting', range: 'C2:C10', variant: 'colorScale' },
    { type: 'conditionalFormatting', range: 'D2:D10', variant: 'dataBar' },
    { type: 'conditionalFormatting', range: 'E2:E10', variant: 'iconSet' },
    { type: 'conditionalFormatting', range: 'F2:F10', variant: 'top10' },
    { type: 'conditionalFormatting', range: 'G2:G10', variant: 'aboveAverage' },
    { type: 'conditionalFormatting', range: 'H2:H10', variant: 'uniqueValues' },
    { type: 'border', at: 'A2:C3', position: 'Left', color: [0, 0, 0] },
    { type: 'fillColor', at: 'A1:C1', color: [0x1a, 0x5f, 0xb4] },
    { type: 'fontColor', at: 'A1:C1', color: [0xe0, 0x3c, 0x3c] },
    { type: 'definedName', name: 'HaviOsszesen', ref: 'Munka1!$A$2:$A$10' },
    { type: 'pageSetup', marginLeft: 25.4, marginTop: 10, printHeadings: true, printGridlines: true },
  ],
})
check('numberFormat: a mert hivas-alak', /GetRange\("D2"\).SetNumberFormat\("#,##0.00"\)/.test(wb.script))
check('formula: a mert hivas-alak', /GetRange\("C2"\).SetValue\("=SUM\(B2:B4\)"\)/.test(wb.script))
check('columnWidth: a mert hivas-alak, oszlop 1. sorara horgonyozva', /GetRange\("B1"\).SetColumnWidth\(25\)/.test(wb.script))
check('autoFilter: ARGUMENTUM NELKUL (a bool-os alak silent no-op, receptek-xlsx.md #7)', /var wantedRange = "A1:C1";/.test(wb.script) && /GetRange\(wantedRange\).SetAutoFilter\(\);/.test(wb.script) && !/SetAutoFilter\(true\)/.test(wb.script))
check('sort: Range-kulcs + true (a mert, mukodo alak)', /SetSort\(oWorksheet.GetRange\("A1"\), true\)/.test(wb.script))
check('conditionalFormatting: mind a 6 MUKODO varians a sajat metodusara fordul', [
  /AddColorScale\(\)/, /AddDatabar\(\)/, /AddIconSetCondition\(\)/, /AddTop10\(\)/, /AddAboveAverage\(\)/, /AddUniqueValues\(\)/,
].every((re) => re.test(wb.script)))
check('  es mind a 6 KULON GetFormatConditions()-t hiv a sajat tartomanyara', (wb.script.match(/GetFormatConditions\(\)/g) || []).length === 6)
check('border: a mert hivas-alak, alap "Thin" stilussal', /GetRange\("A2:C3"\).SetBorders\("Left", "Thin", Api.CreateColorFromRGB\(0, 0, 0\)\)/.test(wb.script))
check('fillColor: a mert hivas-alak', /GetRange\("A1:C1"\).SetFillColor\(Api.CreateColorFromRGB\(26, 95, 180\)\)/.test(wb.script))
check('fontColor: a mert hivas-alak', /GetRange\("A1:C1"\).SetFontColor\(Api.CreateColorFromRGB\(224, 60, 60\)\)/.test(wb.script))
check('definedName: a mert hivas-alak (Api.AddDefName, hidden alapertelmezetten false)', /Api\.AddDefName\("HaviOsszesen", "Munka1!\$A\$2:\$A\$10", false\)/.test(wb.script))
check('definedName: hidden:true -> harmadik argumentum true', /Api\.AddDefName\("RejtettNev", "Munka1!\$B\$1", true\)/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'definedName', name: 'RejtettNev', ref: 'Munka1!$B$1', hidden: true }] }).script,
))
check('pageSetup: csak a kert mezok emitalodnak (marginLeft/marginTop/printHeadings/printGridlines), marginRight/marginBottom nem', /oWorksheet\.SetLeftMargin\(25\.4\);/.test(wb.script) && /oWorksheet\.SetTopMargin\(10\);/.test(wb.script) && /oWorksheet\.SetPrintHeadings\(true\);/.test(wb.script) && /oWorksheet\.SetPrintGridlines\(true\);/.test(wb.script) && !/SetRightMargin/.test(wb.script) && !/SetBottomMargin/.test(wb.script))

check('rangeSize: rowHeight -> ApiRange.SetRowHeight, egyetlen argumentummal', /oWorksheet\.GetRange\("A3:C3"\)\.SetRowHeight\(40\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSize', at: 'A3:C3', rowHeight: 40 }] }).script,
))
check('rangeAlign: mindharom mezo, sajat GetRange-hivassal soronkent', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeAlign', at: 'A1', vertical: 'center', horizontal: 'right', readingOrder: 'rtl' }] }).script
  return /oWorksheet\.GetRange\("A1"\)\.SetAlignVertical\("center"\);/.test(s)
    && /oWorksheet\.GetRange\("A1"\)\.SetAlignHorizontal\("right"\);/.test(s)
    && /oWorksheet\.GetRange\("A1"\)\.SetReadingOrder\("rtl"\);/.test(s)
})())
check('rangeAlign: csak vertical -> csak az a sor emitalodik', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeAlign', at: 'A1', vertical: 'top' }] }).script
  return /SetAlignVertical\("top"\);/.test(s) && !/SetAlignHorizontal|SetReadingOrder/.test(s)
})())
check('rangeFontStyle: italic+strikeout+wrap, underline nelkul', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeFontStyle', at: 'A1', italic: true, strikeout: true, wrap: true }] }).script
  return /oWorksheet\.GetRange\("A1"\)\.SetItalic\(true\);/.test(s) && /SetStrikeout\(true\);/.test(s) && /SetWrap\(true\);/.test(s)
})())
check('unmergeRange: a mert hivas-alak', /oWorksheet\.GetRange\("A1:B2"\)\.UnMerge\(\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'unmergeRange', at: 'A1:B2' }] }).script,
))
check('sheet: name -> SetName, SetVisible nem emitalodik', /oWorksheet\.SetName\("Osszefoglalo"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', name: 'Osszefoglalo' }] }).script,
) && !/SetVisible/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', name: 'Osszefoglalo' }] }).script))
check('sheet: visible:false -> SetVisible(false), SetName nem emitalodik', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', visible: false }] }).script
  return /oWorksheet\.SetVisible\(false\);/.test(s) && !/SetName/.test(s)
})())
check('sheet: name+visible egyutt -> mindket sor, SetName eloszor', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', name: 'Adatok', visible: true }] }).script
  const iName = s.indexOf('SetName("Adatok")')
  const iVis = s.indexOf('SetVisible(true)')
  return iName !== -1 && iVis !== -1 && iName < iVis
})())
check('sheet: `target` nev szerint -> Api.GetSheet(nev) egy `var`-ba, es AZON hivodik a SetName', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', target: 'Munka2', name: 'Adatok' }] }).script
  return /var oTargetSheet = Api\.GetSheet\("Munka2"\);/.test(s) && /oTargetSheet\.SetName\("Adatok"\);/.test(s) && !/oWorksheet\.SetName/.test(s)
})())
check('sheet: `target` INDEX szerint -> szam LITERALKENT, nem sztringkent', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', target: 2, visible: false }] }).script
  return /var oTargetSheet = Api\.GetSheet\(2\);/.test(s) && !/GetSheet\("2"\)/.test(s)
})())
check('sheet: `active:true` -> SetActive() a celon', /oTargetSheet\.SetActive\(\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', target: 'Munka3', active: true }] }).script,
))
check('sheet: `active:false` NEM emittal SetActive-ot (csak igazra all)', !/SetActive/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', target: 'Munka3', active: false }] }).script,
))
check('sheet: `delete:true` ONMAGABAN -> Delete() a celon, semmi mas hivas', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', target: 'Munka2', delete: true }] }).script
  return /oTargetSheet\.Delete\(\);/.test(s) && !/SetName|SetVisible|SetActive/.test(s)
})())
check('sheetDisplay: row+rowHeight -> SetRowHeight, a `row` 1-based bemenet -1-re fordul (0-based hivas)', /oWorksheet\.SetRowHeight\(2, 40\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', row: 3, rowHeight: 40 }] }).script,
))
check('sheetDisplay: gridlines:false -> SetDisplayGridlines(false), a tobbi mezo nem emitalodik', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', gridlines: false }] }).script
  return /oWorksheet\.SetDisplayGridlines\(false\);/.test(s) && !/SetRowHeight|SetDisplayHeadings/.test(s)
})())
check('sheetDisplay: headings:false -> SetDisplayHeadings(false)', /oWorksheet\.SetDisplayHeadings\(false\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', headings: false }] }).script,
))
check('sheetDisplay: mindharom egyutt -> harom sor, sorrend row/gridlines/headings', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', row: 1, rowHeight: 20, gridlines: true, headings: true }] }).script
  const iRow = s.indexOf('SetRowHeight(0, 20)')
  const iGrid = s.indexOf('SetDisplayGridlines(true)')
  const iHead = s.indexOf('SetDisplayHeadings(true)')
  return iRow !== -1 && iGrid !== -1 && iHead !== -1 && iRow < iGrid && iGrid < iHead
})())
check('hyperlink: range+address, tooltip nelkul -> 2-argumentumos hivas, STRING range-fel (nem GetRange)', /oWorksheet\.SetHyperlink\("B1", "https:\/\/example\.org"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'hyperlink', range: 'B1', address: 'https://example.org' }] }).script,
))
check('hyperlink: tooltip-tal -> 4-argumentumos hivas, ures 3. argumentum + tooltip 4.-kent (MERT: enelkul csendben eldobodna)', /oWorksheet\.SetHyperlink\("B1", "https:\/\/example\.org", "", "Kattints ide"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'hyperlink', range: 'B1', address: 'https://example.org', tooltip: 'Kattints ide' }] }).script,
))
// az argumentum-sorrend FORDITVA volt --
// a valodi alairas AddProtectedRange(title, range), nem (range, title).
check('protectedRange: a mert hivas-alak, title ELOSZOR, STRING range-fel', /oWorksheet\.AddProtectedRange\("Locked", "A1:B2"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'protectedRange', range: 'A1:B2', title: 'Locked' }] }).script,
))
check('rangeComment: a mert hivas-alak, author ures sztring ha nincs kerve', /oWorksheet\.GetRange\("A1"\)\.AddComment\("Note", ""\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeComment', at: 'A1', text: 'Note' }] }).script,
))
check('rangeComment: author-ral egyutt', /oWorksheet\.GetRange\("A1"\)\.AddComment\("Note", "Alpha"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeComment', at: 'A1', text: 'Note', author: 'Alpha' }] }).script,
))
check('rangeSelect: select+rotation egyutt', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSelect', at: 'B2', select: true, rotation: 45 }] }).script
  return /oWorksheet\.GetRange\("B2"\)\.Select\(\);/.test(s) && /oWorksheet\.GetRange\("B2"\)\.SetOrientation\(45\);/.test(s)
})())
check('rangeSelect: csak rotation -> Select nem emitalodik', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSelect', at: 'A1', rotation: -30 }] }).script
  return /SetOrientation\(-30\);/.test(s) && !/\.Select\(\)/.test(s)
})())
check('rangeEdit: delete+insert+autofit egyutt', (() => {
  const s = buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeEdit', at: 'A2', delete: 'up', insert: 'down', autoFitRows: false, autoFitColumns: true }] }).script
  return /oWorksheet\.GetRange\("A2"\)\.Delete\("up"\);/.test(s) && /oWorksheet\.GetRange\("A2"\)\.Insert\("down"\);/.test(s) && /oWorksheet\.GetRange\("A2"\)\.AutoFit\(false, true\);/.test(s)
})())
check('rangeClipboard: copyTo -> Copy(GetRange(dest))', /oWorksheet\.GetRange\("A1"\)\.Copy\(oWorksheet\.GetRange\("B1"\)\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeClipboard', at: 'A1', copyTo: 'B1' }] }).script,
))
check('rangeClipboard: cutTo -> Cut(GetRange(dest))', /oWorksheet\.GetRange\("A1"\)\.Cut\(oWorksheet\.GetRange\("C1"\)\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeClipboard', at: 'A1', cutTo: 'C1' }] }).script,
))
// `paste` MOSTANTOL MUKODIK -- `at` a CEL,
// `paste` a FORRAS (a copyTo/cutTo iranyahoz kepest FORDITVA), Paste(rangeFrom) alakban.
check('rangeClipboard: paste -> GetRange(at).Paste(GetRange(forras)), FORDITOTT irany a copyTo-hoz kepest', /oWorksheet\.GetRange\("D1"\)\.Paste\(oWorksheet\.GetRange\("A1"\)\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeClipboard', at: 'D1', paste: 'A1' }] }).script,
))
dob('rangeClipboard: `paste` ures string -> NEVESITETT hiba', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeClipboard', at: 'A1', paste: '' }] }), /`paste`.*non-empty/)
check('rangeFormulaArray: a mert hivas-alak', /oWorksheet\.GetRange\("C1:C3"\)\.SetFormulaArray\("=B1:B3\*2"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeFormulaArray', at: 'C1:C3', formula: '=B1:B3*2' }] }).script,
))
check('rangeOffsetWrite: csak ertek (offset/resize nelkul) -> kozvetlen GetRange().SetValue()', /oWorksheet\.GetRange\("A1"\)\.SetValue\("x"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeOffsetWrite', at: 'A1', value: 'x' }] }).script,
))
check('rangeOffsetWrite: offsetRows+offsetCols -> lancolt .Offset(2, 1)', /oWorksheet\.GetRange\("A1"\)\.Offset\(2, 1\)\.SetValue\("offsetted"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeOffsetWrite', at: 'A1', offsetRows: 2, offsetCols: 1, value: 'offsetted' }] }).script,
))
check('sheetDrawing: shape -> AddShape(type, x, y, w, h)', /oWorksheet\.AddShape\("rect", 1, 1, 2000000, 1000000\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDrawing', shape: { type: 'rect' } }] }).script,
))
check('sheetDrawing: shape egyedi meretekkel', /oWorksheet\.AddShape\("ellipse", 5, 6, 3000000, 1500000\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDrawing', shape: { type: 'ellipse', x: 5, y: 6, width: 3000000, height: 1500000 } }] }).script,
))
check('sheetFormatTable: FormatAsTable(range, style) -- RANGE ELSO, STYLE MASODIK', /oWorksheet\.FormatAsTable\("A1:B2", "TableStyleMedium2"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetFormatTable', range: 'A1:B2', style: 'TableStyleMedium2' }] }).script,
))
check('pivotTable: range only -> InsertPivotNewWorksheet', /Api\.InsertPivotNewWorksheet\(oWorksheet\.GetRange\("A1:B4"\)\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'pivotTable', range: 'A1:B4' }] }).script,
))
check('pivotTable: `at` -> InsertPivotExistingWorksheet(source, GetRange(at)) KET Range-objektummal', /Api\.InsertPivotExistingWorksheet\(oWorksheet\.GetRange\("A1:B4"\), oWorksheet\.GetRange\("D5"\)\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'pivotTable', range: 'A1:B4', at: 'D5' }] }).script,
))
check('pivotTable: `refresh` -> RefreshAllPivots() a vegen', /Api\.InsertPivotNewWorksheet\(oWorksheet\.GetRange\("A1:B4"\)\);\nApi\.RefreshAllPivots\(\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'pivotTable', range: 'A1:B4', refresh: true }] }).script,
))
check('rangeOffsetWrite: resizeRows+resizeCols -> lancolt .Resize(2, 2)', /oWorksheet\.GetRange\("A1"\)\.Resize\(2, 2\)\.SetValue\("resized"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeOffsetWrite', at: 'A1', resizeRows: 2, resizeCols: 2, value: 'resized' }] }).script,
))
check('rangeOffsetWrite: offset ES resize egyutt lancolva, sorrendben', /oWorksheet\.GetRange\("A1"\)\.Offset\(1, 0\)\.Resize\(2, 1\)\.SetValue\("both"\);/.test(
  buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeOffsetWrite', at: 'A1', offsetRows: 1, offsetCols: 0, resizeRows: 2, resizeCols: 1, value: 'both' }] }).script,
))

console.log('\n[13] MUNKAFUZET-MEGTAGADASOK -- a "nem-tamogatott" NEVESITETT valasz, nem csendes elhagyas')
dob('numberFormat: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'numberFormat', format: '0' }] }), /`at`.*required/)
dob('numberFormat: `format` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'numberFormat', at: 'A1' }] }), /`format`.*required/)
dob('formula: "=" nelkuli szoveg -> megtagadva (nem csendben ertekke valik)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '10' }] }), /must start with "="/)
dob('columnWidth: negativ szelesseg -> megtagadva (a builder csendben ELREJTENE az oszlopot)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'columnWidth', column: 'B', width: -5 }] }), /silently HIDES the column/)
dob('columnWidth: `column` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'columnWidth', width: 10 }] }), /`column`.*required/)
dob('autoFilter: `range` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'autoFilter' }] }), /`range`.*required/)
dob('border: "All" -> megtagadva (a builder csendben SEMMILYEN szegelyt nem tenne be)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'border', at: 'A1:B2', position: 'All' }] }), /silently add no border/)
dob('border: "Outline" -> megtagadva ugyanazzal az okkal', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'border', at: 'A1:B2', position: 'Outline' }] }), /silently add no border/)
dob('border: ismeretlen pozicio -> megtagadva, a nevesitett listaval', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'border', at: 'A1:B2', position: 'Sideways' }] }), /known:.*Left.*DiagonalUp/)
dob('border: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'border', position: 'Left' }] }), /`at`.*required/)
dob('border: ervenytelen szin (2 elemu tomb) -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'border', at: 'A1:B2', position: 'Left', color: [0, 0] }] }), /`color`.*array of three/)
dob('definedName: `name` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'definedName', ref: 'A1:A2' }] }), /`name`.*required/)
dob('definedName: `ref` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'definedName', name: 'X' }] }), /`ref`.*required/)
dob('pageSetup: egyetlen mezo sem -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'pageSetup' }] }), /at least one of/)
dob('pageSetup: orientation -> megtagadva (LETEZIK a metodus, de MERT INERT ezen a DS-en)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'pageSetup', orientation: 'landscape' }] }), /orientation.*not supported.*confirmed no-op/)
dob('pageSetup: paperSize -> megtagadva (NINCS ilyen metodus, kimerito enumeracioval mert)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'pageSetup', paperSize: 9 }] }), /paperSize.*not supported/)
dob('pageSetup: fitToWidth -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'pageSetup', fitToWidth: 1 }] }), /fitToWidth.*not supported/)
dob('pageSetup: printArea -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'pageSetup', printArea: 'A1:B2' }] }), /printArea.*not supported/)
dob('border: ervenytelen szin (tartomanyon kivul) -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'border', at: 'A1:B2', position: 'Left', color: [0, 0, 300] }] }), /`color`.*values must each be 0-255/)
check('border: `color` nelkul -> alapertelmezett fekete (nem megtagadas)', /Api.CreateColorFromRGB\(0, 0, 0\)/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'border', at: 'A1:B2', position: 'Left' }] }).script))
dob('fillColor: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'fillColor', color: [0, 0, 0] }] }), /`at`.*required/)
dob('fillColor: `color` nelkul -> megtagadva (nincs alapertelmezett szin egy szin-muveletnek)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'fillColor', at: 'A1' }] }), /`color`.*array of three/)
dob('fillColor: ervenytelen szin-komponens -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'fillColor', at: 'A1', color: [-1, 0, 0] }] }), /`color`.*values must each be 0-255/)
dob('fontColor: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'fontColor', color: [0, 0, 0] }] }), /`at`.*required/)
dob('fontColor: `color` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'fontColor', at: 'A1' }] }), /`color`.*array of three/)
dob('sort: `descending:true` -> megtagadva (a builder csendben NOVEKVO sorrendet adna)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sort', range: 'A1:B4', keyCell: 'A1', descending: true }] }), /direction argument to SetSort has no effect/)
dob('sort: `keyCell` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sort', range: 'A1:B4' }] }), /`keyCell`.*required/)
dob('conditionalFormatting: "cellValue" varians -> NEVESITETT nem-tamogatott (FormatConditions.Add inert)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'conditionalFormatting', range: 'A1:A5', variant: 'cellValue' }] }), /FormatConditions\.Add\(\) is inert/)
dob('conditionalFormatting: "expression" varians -> ugyanaz a nevesitett nem-tamogatott', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'conditionalFormatting', range: 'A1:A5', variant: 'expression' }] }), /FormatConditions\.Add\(\) is inert/)
dob('conditionalFormatting: teljesen ismeretlen varians -> megtagadva, felsorolja a mukodoket', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'conditionalFormatting', range: 'A1:A5', variant: 'nincs-ilyen' }] }), /unknown variant "nincs-ilyen"/)
check('freezePanes: mode "row" -> Api.SetFreezePanesType("row"), nincs Select() elotte', /Api\.SetFreezePanesType\("row"\);/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes', mode: 'row' }] }).script) && !/\.Select\(\)/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes', mode: 'row' }] }).script))
check('freezePanes: mode "column" -> Api.SetFreezePanesType("column")', /Api\.SetFreezePanesType\("column"\);/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes', mode: 'column' }] }).script))
check('freezePanes: mode "cell" -> a `cell` cellara Select(), MAJD SetFreezePanesType("cell")', /GetRange\("B2"\)\.Select\(\);[\s\S]*Api\.SetFreezePanesType\("cell"\);/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes', mode: 'cell', cell: 'B2' }] }).script))
dob('freezePanes: `mode` nelkul -> megtagadva, felsorolja az ervenyeseket', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes' }] }), /`mode`.*must be one of row, column, cell/)
dob('freezePanes: ismeretlen `mode` -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes', mode: 'diagonal' }] }), /`mode`.*must be one of row, column, cell/)
dob('freezePanes: mode "cell", `cell` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes', mode: 'cell' }] }), /`cell`.*required when mode is "cell"/)
dob('freezePanes: mode "row" DE `cell` is megadva -> megtagadva (a "row" ugyis figyelmen kivul hagyna, ne hallgassuk el)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'freezePanes', mode: 'row', cell: 'B2' }] }), /`cell`.*only used with mode "cell"/)
dob('rangeSize: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSize', rowHeight: 20 }] }), /`at`.*required/)
dob('rangeSize: `hidden` -> MINDIG megtagadva -- MERT: a setter inert ezen a peldanyon', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSize', at: 'A1', hidden: true }] }), /SetHidden\(\) is inert/)
dob('rangeSize: `rowHeight` nelkul (es hidden sincs) -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSize', at: 'A1' }] }), /`rowHeight`.*required/)
dob('rangeAlign: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeAlign', vertical: 'center' }] }), /`at`.*required/)
dob('rangeAlign: egyetlen mezo sem -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeAlign', at: 'A1' }] }), /at least one of/)
dob('rangeFontStyle: `underline` -> MINDIG megtagadva -- MERT: a setter inert ezen a peldanyon', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeFontStyle', at: 'A1', underline: true }] }), /SetUnderline\(\) is inert/)
dob('rangeFontStyle: egyetlen mezo sem (underline sincs) -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeFontStyle', at: 'A1' }] }), /at least one of/)
dob('unmergeRange: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'unmergeRange' }] }), /`at`.*required/)
dob('sheet: sem `name`, sem `visible` -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet' }] }), /at least one of/)
// `move` MOSTANTOL MUKODIK, a valodi
// Move(before, after) alakkal -- a korabbi megtagadas rossz hivasi alakot (egyetlen szam) mert.
{
  const moveAfter = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', move: { after: 'Munka1' } }] })
  check('sheet: `move: {after}` -> Move(undefined, oMoveTarget), Api.GetSheet a celre', /var oMoveTarget = Api\.GetSheet\("Munka1"\);\s*oWorksheet\.Move\(undefined, oMoveTarget\);/.test(moveAfter.script), moveAfter.script)
  const moveBefore = buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', move: { before: 0 } }] })
  check('sheet: `move: {before}` (szammal) -> Move(oMoveTarget, undefined)', /var oMoveTarget = Api\.GetSheet\(0\);\s*oWorksheet\.Move\(oMoveTarget, undefined\);/.test(moveBefore.script), moveBefore.script)
  dob('sheet: `move` se before se after -> NEVESITETT hiba', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', move: {} }] }), /`move` needs EXACTLY ONE/)
  dob('sheet: `move` MINDKETTO -> NEVESITETT hiba (a valodi Move csak egyet fogad)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', move: { before: 'A', after: 'B' } }] }), /`move` needs EXACTLY ONE/)
  dob('sheet: `move` nem objektum -> NEVESITETT hiba', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', move: 2 }] }), /`move` must be an object/)
}
dob('sheet: `delete` MAS mezovel egyutt -> megtagadva (a lap mar nem letezne azoknak)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheet', target: 'Munka2', delete: true, name: 'X' }] }), /cannot be combined/)
dob('sheetDisplay: sem row/rowHeight, sem gridlines, sem headings -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay' }] }), /at least one of/)
dob('sheetDisplay: `row` `rowHeight` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', row: 2 }] }), /`rowHeight`.*requires.*`row`|`row`.*requires.*`rowHeight`/)
dob('sheetDisplay: `rowHeight` `row` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', rowHeight: 20 }] }), /requires `row`/)
dob('sheetDisplay: `row` 0 -> megtagadva (1-based bemenet, 0 nem ervenyes sor)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', row: 0, rowHeight: 20 }] }), /positive integer/)
dob('sheetDisplay: `row` nem egesz -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDisplay', row: 1.5, rowHeight: 20 }] }), /positive integer/)
dob('hyperlink: `range` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'hyperlink', address: 'https://x.example' }] }), /`range`.*required/)
dob('hyperlink: `address` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'hyperlink', range: 'A1' }] }), /`address`.*required/)
dob('protectedRange: `range` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'protectedRange', title: 'X' }] }), /`range`.*required/)
dob('protectedRange: `title` nelkul -> megtagadva (MERT: a fuggveny sajat guardja hibat dob title/dataRange nelkul)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'protectedRange', range: 'A1:B2' }] }), /`title`.*required.*title or dataRange is invalid/)
dob('rangeComment: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeComment', text: 'x' }] }), /`at`.*required/)
dob('rangeComment: `text` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeComment', at: 'A1' }] }), /`text`.*required/)
dob('rangeSelect: egyetlen mezo sem -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSelect', at: 'A1' }] }), /at least one of/)
dob('rangeEdit: egyetlen mezo sem -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeEdit', at: 'A1' }] }), /at least one of/)
// `paste` MOSTANTOL MUKODIK (lasd fentebb,
// [66] koruli uj tesztek) -- csak `pasteSpecial` marad megtagadva, KULON, csendes no-op miatt.
dob('rangeClipboard: `pasteSpecial` -> MEGTAGADVA -- MERT: csendes no-op minden kiprobalt alakban', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeClipboard', at: 'A1', pasteSpecial: 2 }] }), /`pasteSpecial`.*silent no-op/)
dob('rangeClipboard: egyetlen mezo sem -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeClipboard', at: 'A1' }] }), /at least one of/)
dob('rangeSearch: MINDIG megtagadva -- MERT: Find/Replace inert ezen a peldanyon, nem bemenet-fuggo', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeSearch', at: 'A1', find: 'x' }] }), /Find\(\) always returns null/)
dob('rangeFormulaArray: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeFormulaArray', formula: '=1' }] }), /`at`.*required/)
dob('rangeFormulaArray: `formula` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeFormulaArray', at: 'A1' }] }), /`formula`.*required/)
dob('rangeOffsetWrite: `value` nelkul -> megtagadva (Offset/Resize maga nem ir semmit)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeOffsetWrite', at: 'A1', offsetRows: 1 }] }), /`value`.*required/)
dob('rangeOffsetWrite: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'rangeOffsetWrite', value: 'x' }] }), /`at`.*required/)
dob('sheetDrawing: `wordArt` -> MINDIG megtagadva -- MERT: AddWordArt toresponkius ezen a peldanyon', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDrawing', wordArt: { text: 'x' } }] }), /AddWordArt\(\) always throws/)
dob('sheetDrawing: `ole` -> MINDIG megtagadva -- MERT: csomag-szinten nem valodi OLE-objektum jon letre', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDrawing', ole: { src: 'data:x', name: 'n', progId: 'p' } }] }), /does not create a real OLE object/)
dob('sheetDrawing: `replaceImage` -> MINDIG megtagadva -- MERT: a mentett kep bajtra valtozatlan marad', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDrawing', replaceImage: { index: 0, src: 'data:x' } }] }), /is a silent no-op in the saved package/)
dob('sheetDrawing: egyetlen mezo sem -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDrawing' }] }), /at least one of/)
dob('sheetDrawing.shape: `type` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetDrawing', shape: {} }] }), /`type`.*required/)
dob('sheetFormatTable: `paste` -> MINDIG megtagadva -- MERT: ApiWorksheet.Paste() csendes no-op', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetFormatTable', paste: true }] }), /Paste\(\) is a silent no-op/)
dob('sheetFormatTable: `range` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetFormatTable', style: 'TableStyleMedium2' }] }), /`range`.*required/)
dob('sheetFormatTable: `style` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetFormatTable', range: 'A1:B2' }] }), /`style`.*required/)
dob('pivotTable: `range` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'pivotTable' }] }), /`range`.*required/)

console.log('\n[14] VEGYES KERES: a valasz MINDEN muveletet megnevez')

{
  const vegyes = buildCreateScript({
    core: 'docx',
    operations: [
      { type: 'text', text: 'a' },                 // ALKALMAZVA
      { type: 'text', text: 'x', highlight: 'nincsilyenszin' }, // NEM-TAMOGATOTT (mert leletre hivatkozik -- listType 2026-08-17 ota MAR TAMOGATOTT, lasd [8b], ezert mas peldat kell venni)
      { type: 'table', rows: [['a', 'b']] },        // ALKALMAZVA
    ],
  })
  check('2 tamogatott + 1 nem -> a script MEGVAN (a ket jo muvelet nem veszik el a harmadik miatt)',
    /AddText\("a"\)/.test(vegyes.script) && /CreateTable\(2, 1\)/.test(vegyes.script))
  check('  applied PONTOSAN a ket sikeres muveletet sorolja, index szerint (0 es 2, NEM 1)',
    vegyes.applied.length === 2 && vegyes.applied[0].index === 0 && vegyes.applied[1].index === 2,
    JSON.stringify(vegyes.applied))
  check('  report MIND A HAROM muveletet megnevezi, tetelesen',
    vegyes.report.length === 3, JSON.stringify(vegyes.report))
  check('  report[0] alkalmazva, sourceRoute a letrehozas utja, indok nincs',
    vegyes.report[0].outcome === 'alkalmazva' && vegyes.report[0].sourceRoute === CREATE_ROUTE && vegyes.report[0].reason === null,
    JSON.stringify(vegyes.report[0]))
  check('  report[1] nem-tamogatott, es az indok megnevezi a highlightot',
    vegyes.report[1].outcome === 'nem-tamogatott' && /highlight/.test(vegyes.report[1].reason),
    JSON.stringify(vegyes.report[1]))
  check('  report[2] alkalmazva',
    vegyes.report[2].outcome === 'alkalmazva', JSON.stringify(vegyes.report[2]))
}

{
  const hibas = buildCreateScript({
    core: 'docx',
    operations: [
      { type: 'text', text: 'jo' },                       // ALKALMAZVA
      { type: 'text', text: 'x', heading: 0 },             // HIBA (letezo mezo, ervenytelen ertek)
    ],
  })
  check('ALKALMAZVA + HIBA (nem NEM-TAMOGATOTT): a heading-tartomany egy MEGLEVO mezo rossz erteke, nem hianyzo kepesseg',
    hibas.report[1].outcome === 'hiba', JSON.stringify(hibas.report[1]))
  check('  a HIBA-bejegyzes is megnevezi az okot',
    /heading.*integer 1-9/.test(hibas.report[1].reason), hibas.report[1].reason)
}

console.log('\n[15] A KARTYA SAJAT NEVESITETT PELDAI, VEGIGVIVE A REPORTON (nem csak a mechanizmus, a KONKRET esetek is)')
{
  // a nevesitett peldak negy elemet neveznek: E1 listType (2026-08-17 ota TAMOGATOTT, mar nem tartozik ide -- lasd [8b]) | E2b pptx
  // columnWidths | E3 ismeretlen chartType | E2 xlsx SetBorders("All"). Az utolso STRUKTURALISAN
  // elerhetetlen a hivo szamara (a kod SOHA nem ajanlja fel az "All"/"Outline" erteket, mindig a
  // hat valodi pozicioval ir -- lasd a table emit() sajat kommentjet), tehat nincs mit
  // jelenteni rola; a masik harom viszont VALODI, hivo-inditotta refuzus, es itt all egyutt.
  const negybol_harom = buildCreateScript({
    core: 'pptx',
    operations: [
      { type: 'table', rows: [['a', 'b']], columnWidths: [30, 70] },      // E2b -- NEM-TAMOGATOTT
      { type: 'chart', chartType: 'nincs-ilyen', series: [[1]], categories: ['x'] }, // E3 -- NEM-TAMOGATOTT
      { type: 'text', text: 'ez athalad' },                               // ALKALMAZVA
    ],
  })
  check('E2b (pptx columnWidths) a reportban NEM-TAMOGATOTT, nevesitve',
    negybol_harom.report[0].outcome === 'nem-tamogatott' && /columnWidths.*not available in the pptx core/.test(negybol_harom.report[0].reason),
    JSON.stringify(negybol_harom.report[0]))
  check('E3 (ismeretlen chartType) a reportban NEM-TAMOGATOTT, nevesitve',
    negybol_harom.report[1].outcome === 'nem-tamogatott' && /unknown chartType "nincs-ilyen"/.test(negybol_harom.report[1].reason),
    JSON.stringify(negybol_harom.report[1]))
  check('a harmadik, tamogatott muvelet meg IGY is athalad -- a ket nevesitett NEM nem viszi el',
    negybol_harom.report[2].outcome === 'alkalmazva' && /AddText\("ez athalad"\)/.test(negybol_harom.script),
    JSON.stringify(negybol_harom.report[2]))
}

check('csupa tamogatott keresnel: a report EGYETLEN nem-alkalmazva bejegyzest sem tartalmaz -- URES a hianylista, nem HIANYZIK (kulonben nem tudni, hogy a kapu nezte-e)',
  (() => {
    const r = buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }, { type: 'text', text: 'b' }] })
    return r.report.every((e) => e.outcome === 'alkalmazva') && r.report.filter((e) => e.outcome !== 'alkalmazva').length === 0
  })())

check('ha EGYETLEN muvelet sem alkalmazhato, a dobott hiba MAGAVAL VISZI a teljes reportot (err.report) -- a hivo egy csupa-nem-tamogatott kotegre is tud tetelesen valaszolni, nem csak egy osszevont uzenetet kap',
  (() => {
    try {
      buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', highlight: 'nincsilyenszin' }] })
      return false
    } catch (err) {
      return Array.isArray(err.report) && err.report.length === 1 && err.report[0].outcome === 'nem-tamogatott'
    }
  })())

console.log('\n[16] MUTANS-PROBA (korabbi kikotes): a NEM-TAMOGATOTT/HIBA szetvalasztas SAJAT asszercioja tud pirosat adni')
{
  // A mutacio: ha a notSupportedError() jelzese figyelmen kivul maradna, MINDEN dobott hiba
  // 'hiba'-kent jelenne meg, sose 'nem-tamogatott'-kent. Eloszor bizonyitjuk, hogy a mutacio
  // TENYLEG alkalmazva van (nem egy nem-alkalmazott csere adna hamis zoldet), majd hogy emiatt
  // a fenti [10]-es assercio piros lenne. (2026-08-17: a peldaop listType-rol highlightra
  // valtott, mert a listType 2026-08-17 ota MAR TAMOGATOTT -- lasd [8b] -- es tobbe nem ad
  // notSupportedError-t; a highlight ugyanazt a hibaosztalyt (nevesitett mezo, csendes-noop
  // ertek, notSupportedError) kepviseli, amit ez a mutans-proba merni akar.)
  const eredetiEmit = OPERATIONS.text.emit
  let mutansLefutott = false
  OPERATIONS.text.emit = function mutaltEmit(op, core) {
    try {
      return eredetiEmit(op, core)
    } catch (err) {
      mutansLefutott = true
      err.notSupported = false // MUTACIO: a jelzes eltuntetve
      throw err
    }
  }
  try {
    const r = buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }, { type: 'text', text: 'x', highlight: 'nincsilyenszin' }] })
    check('a mutacio TENYLEG lefutott (a highlight-agon at, nem csak elmeletben)', mutansLefutott === true)
    check('a mutacio HATASA: a report[1] outcome-ja MOST "hiba", NEM "nem-tamogatott" -- a [10]-es assercio ITT piros lenne',
      r.report[1].outcome === 'hiba')
  } finally {
    OPERATIONS.text.emit = eredetiEmit
  }
  const utana = buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }, { type: 'text', text: 'x', highlight: 'nincsilyenszin' }] })
  check('a mutacio visszavonva: a report[1] outcome-ja ismet "nem-tamogatott"', utana.report[1].outcome === 'nem-tamogatott')
}

console.log('\n[17] KEP `path`-ROL -- lokalis fajl, MI kodoljuk')
const LAB = fs.mkdtempSync(path.join(os.tmpdir(), 'e4-kep-teszt-'))
// valodi 1x1 piros PNG, PNG magic byte-okkal
const PNG_1X1 = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64')
const pngPath = path.join(LAB, 'kep.png')
fs.writeFileSync(pngPath, PNG_1X1)
const notImagePath = path.join(LAB, 'nem-kep.png')
fs.writeFileSync(notImagePath, Buffer.from('csak sima szoveg, nem kep-fejlec'))

const kepScript = buildCreateScript({ core: 'docx', operations: [{ type: 'image', path: pngPath }] })
check('path: a fajl base64-kent kerul a scriptbe, MIME a magic byte-bol', kepScript.script.includes(`data:image/png;base64,${PNG_1X1.toString('base64')}`))
check('  es a beagyazott bajtok BYTE-RA egyeznek a bemeneti fajllal', (() => {
  const m = kepScript.script.match(/data:image\/png;base64,([A-Za-z0-9+/=]+)/)
  return m && Buffer.from(m[1], 'base64').equals(PNG_1X1)
})())

dob('path: nem letezo fajl -> NEVESITETT hiba (a muvelet neve + az ut is benne)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'image', path: path.join(LAB, 'nincs-ilyen.png') }] }),
  /image: file not found/)
dob('path: nem kep-tipusu fajl -> NEVESITETT hiba, nem csendes atengedes',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'image', path: notImagePath }] }),
  /image:.*not a recognised image file/)
dob('src: rosszul formazott data: URI -> NEVESITETT hiba',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'image', src: 'nem-data-uri' }] }),
  /image:.*not a well-formed data: URI/)
dob('src ES path nelkul -> NEVESITETT hiba (a mar letezo elutasitas, path bevezetese utan is all)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'image' }] }),
  /image:.*either `src`.*or `path`.*required/)

console.log('\n[18] A TOBBI MUVELET NEM MEGY AT CSENDBEN, HA A KEP MEGTAGAD -- ATIRVA E8 UJ SZERZODESEHEZ (korabbi diagnozis)')
{
  // E8 landolasa ota buildCreateScript muveletenkent halad, es
  // CSAK akkor dob, ha EGYETLEN muvelet sem alkalmazhato -- egy rossz kep-muvelet tehat MAR NEM
  // allitja meg a TELJES hivast, ha van mellette sikeres muvelet is (a masteren ez a szerzodes,
  // nem ez a kartya hozta). Az EREDETI vedelem (ne menjen at csendben) MEGMARAD, csak MASKEPP:
  // a hivo a reportban LATJA a kihagyott muveletet, nem attol vedve, hogy a hivas megallt.
  const vegyesKep = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'ez atmegy' }, { type: 'image', path: path.join(LAB, 'nincs-ilyen.png') }],
  })
  check('a szoveg-muvelet ATMEGY a rossz kep-muvelet MELLETT -- ez MOST a helyes viselkedes, nem regresszio',
    /AddText\("ez atmegy"\)/.test(vegyesKep.script) && vegyesKep.report[0].outcome === 'alkalmazva',
    JSON.stringify(vegyesKep.report))
  check('a rossz kep-muvelet a reportban outcome:"hiba"-val es a SAJAT nevesitett indokaval jelenik meg -- a hivo MEGTUDJA, nem a hivas leallasa vedi',
    vegyesKep.report[1].outcome === 'hiba' && /image: file not found/.test(vegyesKep.report[1].reason),
    JSON.stringify(vegyesKep.report[1]))
}
dob('DE ha CSAK a rossz kep-muvelet van a listan, a hivas TOVABBRA IS dob ("egyetlen muvelet sem alkalmazhato" hatar) -- a vedelem nem tunt el, csak athelyezodott',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'image', path: path.join(LAB, 'nincs-ilyen.png') }] }),
  /image: file not found/)

fs.rmSync(LAB, { recursive: true, force: true })
check('a teszt sajat scratch konyvtara torolve', !fs.existsSync(LAB))

console.log('\n[19] E6 -- ALAKZAT: pozicio, meret, kitoltes, szegely, szoveg')
const alak = buildCreateScript({
  core: 'pptx',
  operations: [{ type: 'shape', shapeType: 'ellipse', x: 111, y: 222, width: 333, height: 444, fill: [0xff, 0, 0], borderColor: [0, 0, 0xff], borderWidth: 12700, text: 'cimke' }],
  slideCount: 1,
})
check('shape: a tipus, meret es pozicio a scriptben all', /CreateShape\("ellipse", 333, 444,/.test(alak.script) && /SetPosition\(111, 222\)/.test(alak.script), alak.script)
check('  a kitoltes SZINE benne van', /CreateRGBColor\(255, 0, 0\)/.test(alak.script))
check('  a szegely SZINE es SZELESSEGE benne van', /CreateStroke\(12700, Api\.CreateSolidFill\(Api\.CreateRGBColor\(0, 0, 255\)\)\)/.test(alak.script))
check('  a szoveg benne van', /AddText\("cimke"\)/.test(alak.script))
check('  a slide-hivatkozas a Presentation-on at megy', /oPresentation\.GetSlideByIndex\(0\)/.test(alak.script))

console.log('\n[20] E6 -- ALAKZAT KITOLTES/SZEGELY NELKUL: NoFill mindket helyen')
const alakUres = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape' }], slideCount: 1 })
check('kitoltes es szegely nelkul ket kulon CreateNoFill() hivas all',
  (alakUres.script.match(/CreateNoFill\(\)/g) || []).length === 2, alakUres.script)

console.log('\n[21] E6 -- DIA-ELRENDEZES: NEVESITETT nem-tamogatott, nem "unknown type"')
dob('layout: megtagadva, MEGNEVEZETT okkal',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'layout' }], slideCount: 1 }),
  /layout: not supported.*ApplyLayout.*blocked/s)
check('  a layout tipus LETEZIK a listaban (nem "unknown type" ut)', 'layout' in OPERATIONS)

console.log('\n[22] E6 -- DIA-INDEX: mindket irany')
const ketDian = buildCreateScript({
  core: 'pptx',
  operations: [{ type: 'shape', slide: 0 }, { type: 'shape', slide: 1 }],
  slideCount: 2,
})
check('ket dia, ket kulon GetSlideByIndex hivas, a MEGADOTT sorrendben',
  ketDian.script.indexOf('GetSlideByIndex(0)') < ketDian.script.indexOf('GetSlideByIndex(1)'))
dob('dia-szamot MEGHALADO index -> NEVESITETT hiba, nem nema no-op',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', slide: 5 }], slideCount: 1 }),
  /slide index 5 is out of range.*1 slide\(s\)/)
dob('negativ dia-index -> szinten megtagadva',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', slide: -1 }], slideCount: 1 }),
  /out of range/)
check('slideCount nelkul (alap: 1) a slide 0 tovabbra is mukodik -- nincs viselkedes-valtozas a regi hivokra',
  /GetSlideByIndex\(0\)/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x' }] }).script))

console.log('\n[22b] E6 -- ALAKZAT DOCX-ON: nincs SetPosition, AddDrawing+Push a mintazata')
const alakDocx = buildCreateScript({
  core: 'docx',
  operations: [{ type: 'shape', shapeType: 'roundRect', width: 333, height: 444, fill: [0xff, 0, 0], borderColor: [0, 0, 0xff], borderWidth: 12700, text: 'cimke' }],
})
check('docx: a tipus es meret a scriptben all', /CreateShape\("roundRect", 333, 444,/.test(alakDocx.script), alakDocx.script)
check('  a kitoltes SZINE benne van', /CreateRGBColor\(255, 0, 0\)/.test(alakDocx.script))
check('  a szegely SZINE es SZELESSEGE benne van', /CreateStroke\(12700, Api\.CreateSolidFill\(Api\.CreateRGBColor\(0, 0, 255\)\)\)/.test(alakDocx.script))
check('  a szoveg benne van', /AddText\("cimke"\)/.test(alakDocx.script))
check('  NINCS SetPosition -- docx-nak nincs dia-abszolut pozicioja', !/SetPosition/.test(alakDocx.script))
check('  a beagyazas AddDrawing + oDocument.Push mintazatot kovet (mint az image docx-agan)',
  /oShapePar\.AddDrawing\(oShape\)/.test(alakDocx.script) && /oDocument\.Push\(oShapePar\)/.test(alakDocx.script))
check('  NINCS oSlide.AddObject (az a pptx-ag)', !/oSlide\.AddObject/.test(alakDocx.script))

console.log('\n[22c] E6 -- ALAKZAT DOCX-ON, KITOLTES/SZEGELY NELKUL: NoFill mindket helyen, meg mindig docx-agban')
const alakDocxUres = buildCreateScript({ core: 'docx', operations: [{ type: 'shape' }] })
check('kitoltes es szegely nelkul ket kulon CreateNoFill() hivas all docx-on is',
  (alakDocxUres.script.match(/CreateNoFill\(\)/g) || []).length === 2, alakDocxUres.script)
check('  docx-on is a Push-mintazat all', /oDocument\.Push\(oShapePar\)/.test(alakDocxUres.script))

console.log('\n[22d] E6 -- ALAKZAT: pptx viselkedese VALTOZATLAN a docx-bekotes utan (NULLA-VISELKEDESVALTOZAS)')
check('pptx-en TOVABBRA IS all a SetPosition (a docx-ag hozzaadasa nem vitte el)',
  /SetPosition\(111, 222\)/.test(alak.script))
check('pptx-en TOVABBRA IS oSlide.AddObject zarja, nem oDocument.Push',
  /oSlide\.AddObject\(oShape\)/.test(alak.script) && !/oDocument\.Push\(oShapePar\)/.test(alak.script))
dob('shape: xlsx tovabbra sincs bekotve (cores: pptx, docx)',
  () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'shape' }] }),
  /type "shape" is not available in the xlsx core.*available: pptx, docx/)

console.log('\n[23] E13 -- BEKEZDES-FORMAZAS (igazitas mar allt: [7]; itt behuzas + sorkoz), receptek-pptx-docx.md #9')
const bekezdes = buildCreateScript({
  core: 'docx',
  operations: [{
    type: 'text', text: 'formazott bekezdes',
    indentFirstLine: 720, indentLeft: 720, indentRight: 720,
    spacingBefore: 200, spacingAfter: 200, spacingLine: 2,
  }],
})
check('indentFirstLine -> SetIndFirstLine', /oParagraph\.SetIndFirstLine\(720\);/.test(bekezdes.script), bekezdes.script)
check('indentLeft -> SetIndLeft', /oParagraph\.SetIndLeft\(720\);/.test(bekezdes.script))
check('indentRight -> SetIndRight', /oParagraph\.SetIndRight\(720\);/.test(bekezdes.script))
check('spacingBefore -> SetSpacingBefore', /oParagraph\.SetSpacingBefore\(200\);/.test(bekezdes.script))
check('spacingAfter -> SetSpacingAfter', /oParagraph\.SetSpacingAfter\(200\);/.test(bekezdes.script))
check('spacingLine RAW ertekkel, "auto" lineRule-lal (a recept EGYETLEN mert alakja -- nincs konvertalas)',
  /oParagraph\.SetSpacingLine\(2, "auto"\);/.test(bekezdes.script), bekezdes.script)
check('a regi haromezos hivas (nincs uj mezo) EGYETLEN sort sem valtoztat -- visszafele-kompatibilitas',
  buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'cim', bold: true, size: 32 }] }).script ===
  [
    'builder.OpenFile("__DOC_URL__");',
    'var oDocument = Api.GetDocument();',
    'var oParagraph = Api.CreateParagraph();',
    'oParagraph.AddText("cim");',
    'oParagraph.SetBold(true);',
    'oParagraph.SetFontSize(32);',
    'oDocument.Push(oParagraph);',
    'builder.SaveFile("docx", "eredmeny.docx");',
    'builder.CloseFile();',
  ].join('\n') + '\n')

console.log('\n[24] E13 -- BEKEZDES-FORMAZAS MEGTAGADASOK -- nevesitett hiba, csendes visszaeses nincs')
dob('indentFirstLine negativ -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', indentFirstLine: -1 }] }), /indentFirstLine.*non-negative integer/)
dob('indentLeft nem egesz -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', indentLeft: 1.5 }] }), /indentLeft.*non-negative integer/)
dob('spacingBefore nem szam -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', spacingBefore: 'sok' }] }), /spacingBefore.*non-negative integer/)
dob('spacingLine negativ -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', spacingLine: -1 }] }), /spacingLine.*non-negative integer/)
dob('spacingLineRule "exact" -> NEVESITETT nem-tamogatott (csak "auto" mert)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', spacingLine: 480, spacingLineRule: 'exact' }] }),
  /`spacingLineRule` "exact" is not supported -- only "auto" is measured/)
dob('spacingLineRule "atLeast" -> ugyanugy megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', spacingLine: 480, spacingLineRule: 'atLeast' }] }), /`spacingLineRule` "atLeast" is not supported/)
check('spacingLineRule ELHAGYASA eseten az alap "auto" -- nem hianyzik nemán', /SetSpacingLine\(480, "auto"\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', spacingLine: 480 }] }).script))

console.log('\n[25] E13 -- NEG. KONTROLL: a muvelet MELLOZESE eseten a script-ben egyik uj elem SEM jelenik meg')
const neg = buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'sima' }] })
check('nincs SetIndFirstLine/SetIndLeft/SetIndRight/SetSpacingBefore/SetSpacingAfter/SetSpacingLine a scriptben',
  !/SetInd|SetSpacing/.test(neg.script), neg.script)

console.log('\n[26] E14 -- TOBB FUTAS EGY BEKEZDESBEN: minden run a SAJAT stilusaval')
const futasok = buildCreateScript({
  core: 'pptx',
  operations: [{
    type: 'runs', x: 111, y: 222, width: 333, height: 444,
    runs: [
      { text: 'Vastag ', bold: true, fontSize: 24, color: 'ff0000' },
      { text: 'dolt-alahuzott', italic: true, underline: true },
      { text: ' strikeout', strikethrough: true },
    ],
  }],
  slideCount: 1,
})
check('harom KULON Api.CreateRun() hivas all, sajat valtozonevvel', (futasok.script.match(/Api\.CreateRun\(\)/g) || []).length === 3, futasok.script)
check('  az elso run: vastag + meret + szin', /oRun0\.SetBold\(true\);/.test(futasok.script) && /oRun0\.SetFontSize\(24\);/.test(futasok.script) && /oRun0\.SetColor\(255, 0, 0\);/.test(futasok.script))
check('  a masodik run: dolt + alahuzott, a SAJAT valtozojan (nem az elsoen)', /oRun1\.SetItalic\(true\);/.test(futasok.script) && /oRun1\.SetUnderline\(true\);/.test(futasok.script) && !/oRun0\.SetItalic/.test(futasok.script))
check('  a harmadik run: athuzott', /oRun2\.SetStrikeout\(true\);/.test(futasok.script))
check('  mindharom run a KOZOS bekezdesbe kerul, a sajat sorrendjeben', futasok.script.indexOf('oRun0') < futasok.script.indexOf('oRun1') && futasok.script.indexOf('oRun1') < futasok.script.indexOf('oRun2'))
check('  a doboz maga hatter/szegely NELKUL (CreateNoFill mindket helyen)', (futasok.script.match(/CreateNoFill\(\)/g) || []).length === 2)

console.log('\n[27] E14 -- MEGTAGADASOK')
dob('runs: ures `runs` tomb -> megtagadva (egy ures bekezdes atmeno tesztnek latszik)',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [] }], slideCount: 1 }),
  /runs.*empty/)
dob('runs: `runs` mezo hianya -> ugyanaz a megtagadas (nem kulon ag)',
  () => buildCreateScript({ core: 'pptx', operations: [{ type: 'runs' }], slideCount: 1 }),
  /runs.*empty/)
dob('runs: ures `runs` tomb docx magban -> UGYANAZ a megtagadas (kozos precondition, nem kulon aganta)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'runs', runs: [] }] }),
  /runs.*empty/)

console.log('\n[28] E14 -- EGYETLEN FUTAS, FORMAZAS NELKUL: nincs feleslegesen kiirt Set* hivas')
const csupaszFutas = buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [{ text: 'sima' }] }], slideCount: 1 })
check('formazas nelkuli run: csak AddText, egyetlen Set* hivas sincs', /oRun0\.AddText\("sima"\);/.test(csupaszFutas.script) && !/oRun0\.Set/.test(csupaszFutas.script))

console.log('\n[29] A `runs` MUVELET DOCX-ON IS: EGY bekezdes, TOBB futam, futamonkent sajat formazas')
// A regi hatar ("runs: docx magban nem all rendelkezesre") megszunt -- ez a kartya sajat targya,
// nem regresszio. Az uj docx-ag alakja a referencia-generatorbol jon (euro-demo-docx.js:84-99, a
// dokumentum sajat 1. fejezetenek demo-mondata), NEM levezetve.
const docxFutasok = buildCreateScript({
  core: 'docx',
  operations: [{
    type: 'runs',
    runs: [
      { text: 'sima ' },
      { text: 'felkover', bold: true },
      { text: 'dolt-alahuzott', italic: true, underline: true },
      { text: 'athuzott', strikethrough: true },
      { text: 'szines', color: 'ff0000' },
      { text: 'kiemelt', highlight: 'yellow' },
      { text: 'also-index', vertAlign: 'subscript' },
      { text: 'felso-index', vertAlign: 'superscript' },
      { text: 'meretezett', size: 32 },
    ],
  }],
})
check('nyolc KULON Api.CreateRun() hivas all, sajat valtozonevvel', (docxFutasok.script.match(/Api\.CreateRun\(\)/g) || []).length === 9, docxFutasok.script)
check('  a bekezdes maga Api.CreateParagraph()-bol jon, NEM shape/textbox (a pptx-agtol elteroen)', /var oParagraph = Api\.CreateParagraph\(\);/.test(docxFutasok.script) && !/CreateShape/.test(docxFutasok.script))
check('  felkover a SAJAT runjan', /oRun1\.SetBold\(true\);/.test(docxFutasok.script) && !/oRun0\.SetBold/.test(docxFutasok.script))
check('  dolt ES alahuzott ugyanazon a runon', /oRun2\.SetItalic\(true\);/.test(docxFutasok.script) && /oRun2\.SetUnderline\(true\);/.test(docxFutasok.script))
check('  athuzott a SetStrikeout hivassal (nem SetStrike)', /oRun3\.SetStrikeout\(true\);/.test(docxFutasok.script))
check('  szin: docx run-szintu SetColor NEGY argumentummal (r,g,b,false) -- MAS alak, mint a pptx-agé', /oRun4\.SetColor\(255, 0, 0, false\);/.test(docxFutasok.script))
check('  kiemelt: SetHighlight("yellow")', /oRun5\.SetHighlight\("yellow"\);/.test(docxFutasok.script))
check('  also index: SetVertAlign("subscript")', /oRun6\.SetVertAlign\("subscript"\);/.test(docxFutasok.script))
check('  felso index: SetVertAlign("superscript")', /oRun7\.SetVertAlign\("superscript"\);/.test(docxFutasok.script))
check('  meret: docx-agon a mezo NEVE `size` (a `text` muvelet docx-mezojevel egyezoen), NEM `fontSize`', /oRun8\.SetFontSize\(32\);/.test(docxFutasok.script))
check('  minden run a KOZOS bekezdesbe kerul, a sajat sorrendjeben', docxFutasok.script.indexOf('oRun0') < docxFutasok.script.indexOf('oRun4') && docxFutasok.script.indexOf('oRun4') < docxFutasok.script.indexOf('oRun8'))
check('  a bekezdes a dokumentumba kerul (oDocument.Push), nem a diara (oSlide.AddObject)', /oDocument\.Push\(oParagraph\);/.test(docxFutasok.script) && !/oSlide\.AddObject/.test(docxFutasok.script))

console.log('\n[30] MEGTAGADASOK, UGYANAZ A FEGYELEM MINT A `text` docx allowlistjeinel')
dob('runs: ismeretlen docx highlight -> megtagadva (a builder csendben eldobna, ha atmenne)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'runs', runs: [{ text: 'x', highlight: 'nincsilyenszin' }] }] }),
  /unknown highlight "nincsilyenszin"/)
dob('runs: ismeretlen docx vertAlign -> megtagadva',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'runs', runs: [{ text: 'x', vertAlign: 'oldalra' }] }] }),
  /unknown vertAlign "oldalra"/)
dob('runs: docx-agon is ervenyes marad a ROSSZ SZIN elutasitasa (parseHexColor kozos)',
  () => buildCreateScript({ core: 'docx', operations: [{ type: 'runs', runs: [{ text: 'x', color: 'zz0000' }] }] }),
  /must be a 6-digit hex string/)

console.log('\n[31] E14 -- NULLA-VISELKEDESVALTOZAS A PPTX-AGON (differencial-futtatas a git HEAD elleni bizonyitek kulon fajlban)')
check('a pptx-ag scriptje a fenti valtoztatas UTAN is pontosan a regi alakot adja a korabbi [26] korpuszra',
  (() => {
    const ujra = buildCreateScript({
      core: 'pptx',
      operations: [{
        type: 'runs', x: 111, y: 222, width: 333, height: 444,
        runs: [
          { text: 'Vastag ', bold: true, fontSize: 24, color: 'ff0000' },
          { text: 'dolt-alahuzott', italic: true, underline: true },
          { text: ' strikeout', strikethrough: true },
        ],
      }],
      slideCount: 1,
    })
    return ujra.script === futasok.script
  })())

console.log('\n[32] SPEAKERNOTES -- eloadoi jegyzet, mert Slide.AddNotesText hivas-alak')
const jegyzet = buildCreateScript({
  core: 'pptx',
  operations: [{ type: 'speakerNotes', slide: 0, text: 'Vegso, regresszios teszt-szoveg.' }],
  slideCount: 1,
})
check('speakerNotes: a mert hivas-alak, a mar bekotott oSlide-ot hasznalva', /oSlide\.AddNotesText\("Vegso, regresszios teszt-szoveg\."\);/.test(jegyzet.script))
dob('speakerNotes: `text` nelkul -> megtagadva', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'speakerNotes', slide: 0 }], slideCount: 1 }), /`text`.*required/)
dob('speakerNotes: docx magban nem all rendelkezesre (pptx-only)', () => buildCreateScript({ core: 'docx', operations: [{ type: 'speakerNotes', text: 'x' }] }), /not available in the docx core/)

console.log('\n[33] pageSetup DOCX-AGON: Section margok/lapmeret/tajolas')
{
  const csakKetMargo = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', marginLeft: 1440, marginTop: 2880 }],
  }).script
  check('pageSetup(docx): a Section a dokumentumbol jon (GetFinalSection)', /var oSection = oDocument\.GetFinalSection\(\);/.test(csakKetMargo))
  check('pageSetup(docx): a KERT ket margo LITERALKENT all a SetPageMargins hivasban',
    /oSection\.SetPageMargins\(1440, 2880, oSection\.GetPageMarginRight\(\), oSection\.GetPageMarginBottom\(\)\);/.test(csakKetMargo), csakKetMargo)
  check('pageSetup(docx): NEM emittal SetPageSize-ot, ha csak margo van kerve', !/SetPageSize/.test(csakKetMargo))

  const mindNegyMargo = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', marginLeft: 1440, marginTop: 2880, marginRight: 4320, marginBottom: 5760 }],
  }).script
  check('pageSetup(docx): mind a negy margo LITERALKENT, egyetlen GetPageMargin* olvasas sem marad a hivasban',
    /oSection\.SetPageMargins\(1440, 2880, 4320, 5760\);/.test(mindNegyMargo) && !/GetPageMargin/.test(mindNegyMargo), mindNegyMargo)

  const explicitMeret = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', pageWidth: 20000, pageHeight: 10000 }],
  }).script
  check('pageSetup(docx): explicit pageWidth/pageHeight -> SetPageSize LITERAL ket erteket kap',
    /oSection\.SetPageSize\(20000, 10000\);/.test(explicitMeret), explicitMeret)
  check('pageSetup(docx): explicit meretnel NEM emittal margo-hivast, ha az nincs kerve', !/SetPageMargins/.test(explicitMeret))

  const csakSzelesseg = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', pageWidth: 20000 }],
  }).script
  check('pageSetup(docx): csak pageWidth -> a magassag futasidoben olvasott (GetPageHeight), nem feltetelezett',
    /oSection\.SetPageSize\(20000, oSection\.GetPageHeight\(\)\);/.test(csakSzelesseg), csakSzelesseg)

  const tajFekvo = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', orientation: 'landscape' }],
  }).script
  check('pageSetup(docx): orientation nincs kulon Set* -- a JELENLEGI meretet olvassa runtime-ban es csak akkor cserel, ha kell',
    /var __sw = oSection\.GetPageWidth\(\), __sh = oSection\.GetPageHeight\(\);/.test(tajFekvo) &&
    /if \(\(__sw > __sh\) !== true\) \{ oSection\.SetPageSize\(__sh, __sw\); \}/.test(tajFekvo), tajFekvo)

  const tajAllo = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', orientation: 'portrait' }],
  }).script
  check('pageSetup(docx): orientation "portrait" -> a feltetel `!== false`-ra fordul (a masik iranyu osszehasonlitas)',
    /if \(\(__sw > __sh\) !== false\) \{ oSection\.SetPageSize\(__sh, __sw\); \}/.test(tajAllo), tajAllo)

  dob('pageSetup(docx): orientation ES explicit pageWidth egyutt -> megtagadva (ketertelmu)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', orientation: 'landscape', pageWidth: 1000 }] }),
    /orientation.*pageWidth.*cannot both be given/)
  dob('pageSetup(docx): ismeretlen orientation ertek -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', orientation: 'sideways' }] }),
    /orientation must be "portrait" or "landscape"/)
  dob('pageSetup(docx): egyetlen mezo sem -> megtagadva, ugyanaz a fegyelem mint xlsx-en',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup' }] }),
    /at least one of marginLeft/)
  dob('pageSetup: pptx magban nem all rendelkezesre (xlsx+docx-only)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'pageSetup', marginLeft: 100 }] }),
    /not available in the pptx core/)

  check('REGRESSZIO: pageSetup(xlsx) valtozatlanul a SAJAT (mm-alapu, kulon Set*Margin) agat futtatja, nem a docx-agat',
    /oWorksheet\.SetLeftMargin\(25\.4\);/.test(wb.script) && !/GetFinalSection/.test(wb.script))
}

console.log('\n[34] pageSetup DOCX-AGON: hasabok, fejlec/lablec-tavolsag, cimlap-flag, kezdo oldalszam, szakasz-tipus')
{
  const csakOszlopszam = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', columns: 3 }],
  }).script
  check('pageSetup(docx): `columns` egyedul -> SetEqualColumns a MERT default tavolsaggal (1701), NEM a builder sajat 0-javal',
    /oSection\.SetEqualColumns\(3, 1701\);/.test(csakOszlopszam), csakOszlopszam)

  const oszlopTavolsaggal = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', columns: 2, columnSpacing: 500 }],
  }).script
  check('pageSetup(docx): `columns`+`columnSpacing` -> mindket ertek LITERALKENT a SetEqualColumns hivasban',
    /oSection\.SetEqualColumns\(2, 500\);/.test(oszlopTavolsaggal), oszlopTavolsaggal)

  dob('pageSetup(docx): `columnSpacing` `columns` nelkul -> megtagadva (nincs mit tavolitani)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columnSpacing: 500 }] }),
    /columnSpacing.*requires.*columns/)
  dob('pageSetup(docx): `columns` nem egesz -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columns: 2.5 }] }),
    /columns.*must be a positive integer/)
  dob('pageSetup(docx): `columns` nulla -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columns: 0 }] }),
    /columns.*must be a positive integer/)

  const tavolsagok = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', headerDistance: 400, footerDistance: 300 }],
  }).script
  check('pageSetup(docx): `headerDistance`+`footerDistance` -> ket kulon hivas, LITERALKENT',
    /oSection\.SetHeaderDistance\(400\);/.test(tavolsagok) && /oSection\.SetFooterDistance\(300\);/.test(tavolsagok), tavolsagok)

  const csakFejlec = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', headerDistance: 400 }],
  }).script
  check('pageSetup(docx): csak `headerDistance` -> NEM emittal SetFooterDistance-t', !/SetFooterDistance/.test(csakFejlec))

  const cimlapBe = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', titlePage: true }],
  }).script
  check('pageSetup(docx): `titlePage: true` -> SetTitlePage(true)', /oSection\.SetTitlePage\(true\);/.test(cimlapBe))

  const cimlapKi = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', titlePage: false }],
  }).script
  check('pageSetup(docx): `titlePage: false` -> SetTitlePage(false) is emittalodik (VALODI kapcsolo, nem csak igazra fut)',
    /oSection\.SetTitlePage\(false\);/.test(cimlapKi))

  const kezdoOldal = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', startPageNumber: 5 }],
  }).script
  check('pageSetup(docx): `startPageNumber` -> SetStartPageNumber LITERAL ertekkel', /oSection\.SetStartPageNumber\(5\);/.test(kezdoOldal))

  dob('pageSetup(docx): `startPageNumber` nulla -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', startPageNumber: 0 }] }),
    /startPageNumber.*must be a positive integer/)

  for (const tipus of ['continuous', 'evenPage', 'nextColumn', 'nextPage', 'oddPage']) {
    const szkript = buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', sectionType: tipus }] }).script
    check(`pageSetup(docx): sectionType "${tipus}" -> SetType(${JSON.stringify(tipus)}) LITERALKENT`,
      szkript.includes(`oSection.SetType(${JSON.stringify(tipus)});`), szkript)
  }
  dob('pageSetup(docx): ismeretlen sectionType -> megtagadva, MERT csendes no-op ellen (nem "unknown type", allowlist-uzenettel)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', sectionType: 'nincsilyen' }] }),
    /unknown sectionType.*known: continuous, evenPage, nextColumn, nextPage, oddPage/)

  const mindEgyutt = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'pageSetup', marginLeft: 1000, columns: 2, headerDistance: 400, titlePage: true, startPageNumber: 3, sectionType: 'continuous' }],
  }).script
  const mindEgyuttMintak = [/SetPageMargins\(1000,/, /SetEqualColumns\(2, 1701\);/, /SetHeaderDistance\(400\);/, /SetTitlePage\(true\);/, /SetStartPageNumber\(3\);/, /SetType\("continuous"\);/]
  check('pageSetup(docx): margo + mind a hat uj mezo EGYUTT egy hivasban -> mindegyik emittalodik, egyik sem nyomja el a masikat',
    mindEgyuttMintak.every((m) => m.test(mindEgyutt)), mindEgyutt)

  check('REGRESSZIO: docx01 mezoi (margo/meret/tajolas) valtozatlanul mukodnek a docx02 valtozas utan',
    /oSection\.SetPageMargins\(1440, 2880, oSection\.GetPageMarginRight\(\), oSection\.GetPageMarginBottom\(\)\);/.test(
      buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', marginLeft: 1440, marginTop: 2880 }] }).script,
    ))
}

console.log('\n[35] pageSetup columnWidths (SetNotEqualColumns, MOSTANTOL MUKODIK) + table.merge on docx (Table.MergeCells)')
{
  // a korabbi "MINDIG megtagadva"
  // meres UJRA-MERVE nem reprodukalodott -- a valodi alairas (aWidths, aSpaces) KET tombbol all,
  // a korabbi bisection egy HARMADIK, nem letezo (count, widths, spaces) alakot is probalt, ami
  // a fuggveny sajat guard-ja miatt (nem hiba, hanem csendes false) tunhetett "blokkolt"-nak. A
  // helyes alakkal package-verified MUKODIK (lasd a lib.cjs sajat kommentjet).
  const cw = buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columnWidths: [4000, 3000, 2500], columnSpacing: 250 }] })
  check('pageSetup(docx): `columnWidths` 3 elemmel -> EGY SetNotEqualColumns hivas, 3 szelesseg + 2 res (N-1)',
    /oSection\.SetNotEqualColumns\(\[4000, 3000, 2500\], \[250, 250\]\);/.test(cw.script), cw.script)
  const cwAlap = buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columnWidths: [4000, 3000] }] })
  check('pageSetup(docx): `columnSpacing` nelkul az alap resertek (1701 twip) all be, ugyanaz mint a columns agnal',
    /oSection\.SetNotEqualColumns\(\[4000, 3000\], \[1701\]\);/.test(cwAlap.script), cwAlap.script)

  dob('pageSetup(docx): `columnWidths` egy elemmel -> NEVESITETT hiba (legalabb 2 kell)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columnWidths: [3000] }] }),
    /columnWidths.*at least 2/)

  dob('pageSetup(docx): `columnWidths` ures tomb is NEVESITETT hiba',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columnWidths: [] }] }),
    /columnWidths.*at least 2/)

  dob('pageSetup(docx): `columnWidths` negativ elemmel -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columnWidths: [3000, -1] }] }),
    /columnWidths.*positive numbers/)

  dob('pageSetup(docx): `columns` es `columnWidths` egyutt -> NEVESITETT hiba (ket ellentmondo mod)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'pageSetup', columns: 2, columnWidths: [3000, 3000] }] }),
    /columns.*columnWidths.*cannot both be given/)

  const egyMerge = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a', 'b', 'c'], ['d', 'e', 'f']], merge: [[0, 0, 0, 1]] }],
  }).script
  check('table(docx): egy [0,0,0,1] merge -> EGY MergeCells hivas ket cellaval, GetCell(0,0)/GetCell(0,1) sorrendben',
    /oTable\.MergeCells\(\[oTable\.GetCell\(0, 0\), oTable\.GetCell\(0, 1\)\]\);/.test(egyMerge), egyMerge)
  check('table(docx): a merge hivas a Push ELOTT all',
    egyMerge.indexOf('oTable.MergeCells(') > -1 && egyMerge.indexOf('oTable.MergeCells(') < egyMerge.indexOf('oDocument.Push(oTable);'))

  const vMerge = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a', 'b'], ['c', 'd']], merge: [[0, 0, 1, 0]] }],
  }).script
  check('table(docx): fuggoleges tartomany [0,0,1,0] -> GetCell(0,0)/GetCell(1,0) a hivasban',
    /oTable\.MergeCells\(\[oTable\.GetCell\(0, 0\), oTable\.GetCell\(1, 0\)\]\);/.test(vMerge), vMerge)

  check('table(docx): merge NELKUL nincs MergeCells hivas',
    !/\.MergeCells\(/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a', 'b']] }] }).script))
  // A korabbi "table(pptx): merge mezo NEM emittal MergeCells-t docx-on kivul" ellenorzes
  // ELTAVOLITVA: idokozben mar HAMIS volt -- a pptx merge idokozben bekotve, sajat szekcioja
  // mereten teljes korben fedi.

  dob('table(docx): merge tuple rossz hosszu -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a', 'b']], merge: [[0, 0, 1]] }] }),
    /must be a \[startRow, startCol, endRow, endCol\] tuple/)

  dob('table(docx): merge vege a tartomanyon kivul -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a', 'b']], merge: [[0, 0, 0, 5]] }] }),
    /outside the table/)

  dob('table(docx): merge vege a kezdet elott -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a', 'b'], ['c', 'd']], merge: [[1, 0, 0, 0]] }] }),
    /end must not be before start/)
}

console.log('\n[36] docx-K2 tablazat-finomitas: ismetlodo fejlec+sormagassag, cella-igazitas/tordeles, cella belso margo, textDirection megtagadva')
{
  const fejlecSzkript = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a', 'b'], ['c', 'd']], repeatHeaderRow: true }],
  }).script
  check('table(docx): repeatHeaderRow -> GetRow(0).SetTableHeader(true) a Push ELOTT', /oTable\.GetRow\(0\)\.SetTableHeader\(true\);/.test(fejlecSzkript) && fejlecSzkript.indexOf('SetTableHeader') < fejlecSzkript.indexOf('oDocument.Push(oTable);'))

  check('table(docx): repeatHeaderRow NELKUL nincs SetTableHeader hivas',
    !/SetTableHeader/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a', 'b']] }] }).script))

  const magassagSzkript = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a'], ['b'], ['c']], rowHeights: [1000, null, 2000] }],
  }).script
  check('table(docx): rowHeights -> soronkent GetRow(i).SetHeight("atLeast", ...), a null KIHAGYVA',
    /oTable\.GetRow\(0\)\.SetHeight\("atLeast", 1000\);/.test(magassagSzkript) &&
    !/GetRow\(1\)\.SetHeight/.test(magassagSzkript) &&
    /oTable\.GetRow\(2\)\.SetHeight\("atLeast", 2000\);/.test(magassagSzkript), magassagSzkript)

  for (const va of ['top', 'center', 'bottom']) {
    const szkript = buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a', 'b']], verticalAlign: va }] }).script
    check(`table(docx): verticalAlign "${va}" -> MINDEN cellan SetVerticalAlign(${JSON.stringify(va)})`,
      (szkript.match(new RegExp(`SetVerticalAlign\\(${JSON.stringify(va)}\\)`, 'g')) || []).length === 2, szkript)
  }
  dob('table(docx): ismeretlen verticalAlign -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a']], verticalAlign: 'middle' }] }),
    /unknown verticalAlign.*known: top, center, bottom/)

  check('table(docx): noWrap -> MINDEN cellan SetNoWrap(true)',
    (buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a', 'b']], noWrap: true }] }).script.match(/SetNoWrap\(true\)/g) || []).length === 2)
  check('table(docx): noWrap NELKUL nincs SetNoWrap hivas',
    !/SetNoWrap/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a']] }] }).script))

  const margoSzkript = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a']], cellMarginTop: 200, cellMarginLeft: 150 }],
  }).script
  check('table(docx): cellMarginTop+cellMarginLeft -> csak a KERT ket oldal emittalodik, LITERALKENT',
    /SetCellMarginTop\(200\);/.test(margoSzkript) && /SetCellMarginLeft\(150\);/.test(margoSzkript) &&
    !/SetCellMarginBottom/.test(margoSzkript) && !/SetCellMarginRight/.test(margoSzkript), margoSzkript)

  dob('table(docx): textDirection -> MINDIG megtagadva (MERT: 6 ertek probalva, egyik sem hagyott nyomot a csomagban)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a']], textDirection: 'lrTb' }] }),
    /textDirection.*SetTextDirection.*refused/)

  check('table(pptx): a docx-K2 mezok NEM emittalnak semmit a pptx agon (verticalAlign/noWrap/cellMargin/repeatHeaderRow/rowHeights)',
    (() => {
      const s = buildCreateScript({
        core: 'pptx',
        operations: [{ type: 'table', rows: [['a', 'b']], verticalAlign: 'top', noWrap: true, cellMarginTop: 1, repeatHeaderRow: true, rowHeights: [500] }],
      }).script
      return !/SetVerticalAlign|SetNoWrap|SetCellMargin|SetTableHeader|GetRow\(0\)\.SetHeight/.test(s)
    })())
}

console.log('\n[37] docx-K3 elso szelete: footnotes/endnotes a `text` mezoin')
{
  const egyLabjegyzet = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'main text', footnotes: ['first note'] }],
  }).script
  check('text(docx): footnotes -> Push UTAN MoveCursorToEnd + AddFootnote, LITERAL szoveggel',
    /oDocument\.Push\(oParagraph\);\noDocument\.MoveCursorToEnd\(\);\noDocument\.AddFootnote\(\)\.GetElement\(0\)\.AddText\("first note"\);/.test(egyLabjegyzet), egyLabjegyzet)

  const ketLabjegyzet = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', footnotes: ['n1', { text: 'n2' }] }],
  }).script
  check('text(docx): tobb footnote -> soronkent kulon MoveCursorToEnd+AddFootnote, sorrendben; objektum-alak (`{text:...}`) is elfogadott',
    (ketLabjegyzet.match(/AddFootnote\(\)/g) || []).length === 2 &&
    ketLabjegyzet.indexOf('"n1"') < ketLabjegyzet.indexOf('"n2"'), ketLabjegyzet)

  const vegjegyzet = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', endnotes: ['e1'] }],
  }).script
  check('text(docx): endnotes -> AddEndnote, KULON hivas a footnote-tol', /oDocument\.AddEndnote\(\)\.GetElement\(0\)\.AddText\("e1"\);/.test(vegjegyzet) && !/AddFootnote/.test(vegjegyzet))

  check('text(docx): footnotes/endnotes NELKUL nincs Add*note hivas',
    !/Add(Foot|End)note/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }] }).script))

  check('text(pptx): footnotes/endnotes mezok NEM emittalnak semmit (docx-only)',
    !/Add(Foot|End)note/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'a', footnotes: ['x'], endnotes: ['y'] }] }).script))
}

console.log('\n[38] docx-K3 masodik szelete: `toc` uj muvelet-tipus')
{
  const alap = buildCreateScript({ core: 'docx', operations: [{ type: 'toc' }] }).script
  check('toc: alapertelmezetten MoveCursorToStart + AddTableOfContents',
    /oDocument\.MoveCursorToStart\(\);\noDocument\.AddTableOfContents\(\);/.test(alap), alap)

  const veg = buildCreateScript({ core: 'docx', operations: [{ type: 'toc', position: 'end' }] }).script
  check('toc: position "end" -> MoveCursorToEnd', /oDocument\.MoveCursorToEnd\(\);\noDocument\.AddTableOfContents\(\);/.test(veg))

  dob('toc: ismeretlen position -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'toc', position: 'middle' }] }),
    /unknown position.*known: start, end/)

  dob('toc: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'toc' }] }), /not available in the pptx core/)

  const sorrendHelyes = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'Chapter One', heading: 1 }, { type: 'toc', position: 'start' }],
  }).script
  check('toc: HELYES sorrend (cim ELOSZOR, toc UTOLJARA a listaban) -> a Push a AddTableOfContents ELOTT all a szkriptben (a MODELL mar tartalmazza a cimet, amikor a toc lefut)',
    sorrendHelyes.indexOf('oDocument.Push(oParagraph);') < sorrendHelyes.indexOf('oDocument.AddTableOfContents();'), sorrendHelyes)
}

console.log('\n[39] docx-K3 harmadik szelete: `comments` a `text` mezoin')
{
  const csakSzoveg = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'main text', comments: ['plain body'] }],
  }).script
  check('text(docx): sima stringes comment -> MoveCursorToEnd + EGYARGUMENTUMOS AddComment (nincs szerzo)',
    /oDocument\.MoveCursorToEnd\(\);\noDocument\.AddComment\("plain body"\);/.test(csakSzoveg), csakSzoveg)

  const szerzos = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', comments: [{ text: 'with author', author: 'beta' }] }],
  }).script
  check('text(docx): {text,author} alak -> KETARGUMENTUMOS AddComment, LITERAL szerzovel',
    /oDocument\.AddComment\("with author", "beta"\);/.test(szerzos), szerzos)

  const tobb = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', comments: ['c1', 'c2'] }],
  }).script
  check('text(docx): tobb comment -> soronkent kulon MoveCursorToEnd+AddComment, sorrendben',
    (tobb.match(/AddComment\(/g) || []).length === 2 && tobb.indexOf('"c1"') < tobb.indexOf('"c2"'))

  check('text(docx): comments NELKUL nincs AddComment hivas',
    !/AddComment/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }] }).script))

  check('text(pptx): comments mezo NEM emittal semmit (docx-only)',
    !/AddComment/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'a', comments: ['x'] }] }).script))

  // MERT SORREND-CSAPDA (2026-08-17): AddComment() a footnote/endnote UTAN ARVA kommentet ad
  // (a hivas sikeres, a komment-torzs bekerul a csomagba, de commentRangeStart/commentReference
  // SOHA nem jon letre -- package-verified). A javitas: comments MINDIG a footnotes/endnotes
  // ELOTT emittalodik, fuggetlenul attol, milyen sorrendben adta oket a hivo az op objektumon.
  const egyuttSzkript = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', footnotes: ['fn'], comments: ['c1'] }],
  }).script
  check('text(docx): comments+footnotes EGYUTT -> az AddComment MINDIG a Push UTAN, DE az ELSO Add*note ELOTT all a szkriptben (a hivo mezo-sorrendjetol fuggetlenul)',
    egyuttSzkript.indexOf('oDocument.Push(oParagraph);') < egyuttSzkript.indexOf('AddComment(') &&
    egyuttSzkript.indexOf('AddComment(') < egyuttSzkript.indexOf('AddFootnote()'), egyuttSzkript)
}

console.log('\n[40] HIPERHIVATKOZAS: shape-szintu (Shape.SetHyperlink) es run-szintu (Run.AddHyperlink), pptx-only')
console.log('    a hivas-alakok toString()-bol lettek visszafejtve; csomag-szintu igazolas (hlinkClick + rels-relationship) elo DS-hivassal, nem itt')
{
  const shapeHl = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', hyperlink: { url: 'https://example.com', tooltip: 'tipp' } }] })
  check('pptx: shape hyperlink -> Api.CreateHyperlink + oShape.SetHyperlink, mindket ertekkel',
    /oShape\.SetHyperlink\(Api\.CreateHyperlink\("https:\/\/example\.com", "tipp"\)\);/.test(shapeHl.script), shapeHl.script)

  const shapeHlNoTooltip = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', hyperlink: { url: 'https://example.com' } }] })
  check('pptx: shape hyperlink tooltip nelkul -> ures string kerul at, nem undefined/null',
    /Api\.CreateHyperlink\("https:\/\/example\.com", ""\)/.test(shapeHlNoTooltip.script))

  check('pptx: shape hyperlink NELKUL (alap) nincs SetHyperlink/CreateHyperlink hivas',
    !/SetHyperlink|CreateHyperlink/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'shape' }] }).script))

  dob('pptx: shape hyperlink url nelkul -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', hyperlink: { tooltip: 'csak tipp' } }] }),
    /hyperlink\.url must be a non-empty string/)
  dob('pptx: shape hyperlink ures url-lel -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', hyperlink: { url: '   ' } }] }),
    /hyperlink\.url must be a non-empty string/)
  dob('docx: shape hyperlink -> NEVESITETT megtagadas, pptx-only ebben az egysegben',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', hyperlink: { url: 'https://example.com' } }] }),
    /shape: hyperlink is pptx-only in this unit/)

  const runsHl = buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [{ text: 'kattints ide', hyperlink: { url: 'example.com/x', tooltip: 'futam-tipp' } }] }] })
  check('pptx: runs hyperlink -> AddHyperlink a MEGFELELO run-valtozon, AZ AddElement UTAN (a hivas elofeltetele: a run mar a bekezdesben van)',
    /oParagraph\.AddElement\(oRun0\);\s*oRun0\.AddHyperlink\("example\.com\/x", "futam-tipp"\);/.test(runsHl.script), runsHl.script)

  const runsHlMulti = buildCreateScript({
    core: 'pptx',
    operations: [{ type: 'runs', runs: [{ text: 'sima' }, { text: 'link', hyperlink: { url: 'example.com/y' } }] }],
  })
  check('pptx: runs -- TOBB futambol csak a SAJAT hyperlinkjevel rendelkezo kapja az AddHyperlink hivast',
    !/oRun0\.AddHyperlink/.test(runsHlMulti.script) && /oRun1\.AddHyperlink\("example\.com\/y", ""\)/.test(runsHlMulti.script), runsHlMulti.script)

  check('pptx: runs hyperlink NELKUL (alap) nincs AddHyperlink hivas',
    !/AddHyperlink/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [{ text: 'sima' }] }] }).script))

  dob('pptx: runs hyperlink url nelkul -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [{ text: 'x', hyperlink: {} }] }] }),
    /hyperlink\.url must be a non-empty string/)
  dob('docx: runs hyperlink -> NEVESITETT megtagadas, pptx-only ebben az egysegben (a TELJES batch elutasitva, nem csak az erintett run)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'runs', runs: [{ text: 'x' }, { text: 'y', hyperlink: { url: 'example.com' } }] }] }),
    /runs: hyperlink is pptx-only in this unit/)
}

console.log('\n[41] FORGATAS, SZOVEG FUGGOLEGES IGAZITAS (shape), FUTAM FUGGOLEGES IGAZITAS (runs, pptx)')
console.log('    a hivas-alakok toString()-bol lettek visszafejtve; csomag-szintu igazolas (rot=/anchor=/baseline=) elo DS-hivassal, nem itt')
{
  const rotated = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', rotation: 45 }] })
  check('pptx: shape rotation -> oShape.SetRotation(45), FOKBAN, valtoztatas nelkul (a builder sajat maga konvertal)',
    /oShape\.SetRotation\(45\);/.test(rotated.script), rotated.script)

  check('pptx: shape rotation negativ szammal is atmegy (arithmetika, nem kulcsszo -- nincs allowlist-elutasitas)',
    /oShape\.SetRotation\(-90\);/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', rotation: -90 }] }).script))

  dob('pptx: shape rotation nem-szammal -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', rotation: 'sok' }] }),
    /rotation must be a finite number/)

  check('pptx: shape rotation NELKUL (alap) nincs SetRotation hivas',
    !/SetRotation/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'shape' }] }).script))

  dob('docx: shape rotation -> NEVESITETT megtagadas, pptx-only ebben az egysegben',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', rotation: 10 }] }),
    /rotation\/verticalTextAlign are pptx-only in this unit/)

  const vAlignShape = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', verticalTextAlign: 'center' }] })
  check('pptx: shape verticalTextAlign -> oShape.SetVerticalTextAlign("center")',
    /oShape\.SetVerticalTextAlign\("center"\);/.test(vAlignShape.script), vAlignShape.script)

  check('pptx: shape verticalTextAlign NELKUL (alap) nincs SetVerticalTextAlign hivas',
    !/SetVerticalTextAlign/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'shape' }] }).script))

  dob('pptx: shape verticalTextAlign ismeretlen ertek -> NEVESITETT hiba (mert: a builder csendben nem csinal semmit, NINCS default ag a switch-ben)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', verticalTextAlign: 'kozepen' }] }),
    /unknown verticalTextAlign "kozepen"/)

  dob('docx: shape verticalTextAlign -> NEVESITETT megtagadas, pptx-only ebben az egysegben',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', verticalTextAlign: 'top' }] }),
    /rotation\/verticalTextAlign are pptx-only in this unit/)

  const runsVertAlign = buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [{ text: 'x' }, { text: 'felso', vertAlign: 'superscript' }] }] })
  check('pptx: runs vertAlign -> CSAK a sajat vertAlign-jevel rendelkezo run kapja a SetVertAlign hivast',
    !/oRun0\.SetVertAlign/.test(runsVertAlign.script) && /oRun1\.SetVertAlign\("superscript"\);/.test(runsVertAlign.script), runsVertAlign.script)

  check('pptx: runs vertAlign NELKUL (alap) nincs SetVertAlign hivas',
    !/SetVertAlign/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [{ text: 'x' }] }] }).script))

  dob('pptx: runs vertAlign ismeretlen ertek -> NEVESITETT hiba, UGYANAZ AZ ALLOWLIST mint docx-nal',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'runs', runs: [{ text: 'x', vertAlign: 'kozep' }] }] }),
    /unknown vertAlign "kozep"/)

  check('REGRESSZIO: docx runs vertAlign valtozatlanul mukodik (emitDocxRun erintetlen)',
    /oRun0\.SetVertAlign\("superscript"\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'runs', runs: [{ text: 'x', vertAlign: 'superscript' }] }] }).script))
}

console.log('\n[42] TABLAZAT-STILUS (SetTableLook) + SOR/OSZLOP-BOVITES LETREHOZAS UTAN (AddRow/AddColumn), pptx-only')
console.log('    a hivas-alakok toString()-bol lettek visszafejtve; csomag-szintu igazolas (firstRow=/bandCol=/sor-szam) elo DS-hivassal, nem itt')
{
  const koteg03Rows = [['A', 'B'], ['1', '2']]

  const withLook = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows, tableLook: { firstColumn: true, lastRow: true, verBand: true } }] })
  check('pptx: tableLook -> SetTableLook a HAT pozicionalis boolean-nal, oszlop ELOBB mint sor a parameterlistaban',
    /oTable\.SetTableLook\(true, false, false, true, false, true\);/.test(withLook.script), withLook.script)

  check('pptx: tableLook mind a hat mezo hianyzik -> mind false',
    /oTable\.SetTableLook\(false, false, false, false, false, false\);/.test(
      buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows, tableLook: {} }] }).script,
    ))

  check('pptx: tableLook NELKUL (alap) nincs SetTableLook hivas',
    !/SetTableLook/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows }] }).script))

  dob('pptx: tableLook nem-objektum -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows, tableLook: 'igen' }] }),
    /tableLook must be an object/)

  dob('docx: tableLook -> NEVESITETT megtagadas, pptx-only ebben az egysegben',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: koteg03Rows, tableLook: { firstRow: true } }] }),
    /tableLook\/extraRows\/extraColumns are pptx-only in this unit/)

  const withExtra = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows, extraColumns: 2, extraRows: 1 }] })
  check('pptx: extraColumns/extraRows -> pontosan ANNYI AddColumn()/AddRow() hivas, argumentum nelkul',
    (withExtra.script.match(/oTable\.AddColumn\(\);/g) || []).length === 2 &&
    (withExtra.script.match(/oTable\.AddRow\(\);/g) || []).length === 1, withExtra.script)

  check('pptx: extraColumns/extraRows NELKUL (alap) nincs AddColumn/AddRow hivas',
    !/AddColumn|AddRow/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows }] }).script))

  dob('pptx: extraColumns negativ szammal -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows, extraColumns: -1 }] }),
    /extraColumns must be a non-negative integer/)
  dob('pptx: extraRows nem-egesz szammal -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: koteg03Rows, extraRows: 1.5 }] }),
    /extraRows must be a non-negative integer/)

  dob('docx: extraColumns -> NEVESITETT megtagadas, pptx-only ebben az egysegben',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: koteg03Rows, extraColumns: 1 }] }),
    /tableLook\/extraRows\/extraColumns are pptx-only in this unit/)
}

console.log('\n[43] UJ MUVELET: comment (Slide.AddComment), pptx-only')
console.log('    a hivas-alak toString()-bol lett visszafejtve; csomag-szintu igazolas (comment1.xml/commentAuthors.xml/rels/Content-Types) elo DS-hivassal, nem itt')
{
  const withComment = buildCreateScript({ core: 'pptx', operations: [{ type: 'comment', text: 'proba', author: 'Gamma' }] })
  check('pptx: comment -> oSlide.AddComment a NEGY argumentummal, alap x/y-nal',
    /oSlide\.AddComment\(500000, 500000, "proba", "Gamma"\);/.test(withComment.script), withComment.script)

  const withXY = buildCreateScript({ core: 'pptx', operations: [{ type: 'comment', text: 'proba', author: 'V', x: 1000000, y: 2000000 }] })
  check('pptx: comment explicit x/y -> LITERALKENT atmegy',
    /oSlide\.AddComment\(1000000, 2000000, "proba", "V"\);/.test(withXY.script))

  dob('pptx: comment `text` nelkul -> NEVESITETT hiba (mert: az ures/hianyzo szoveg a builderben CSENDBEN false-t ad, nem ir megjegyzest)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'comment', author: 'V' }] }),
    /comment: `text` is required/)
  dob('pptx: comment `text` ures string -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'comment', text: '', author: 'V' }] }),
    /comment: `text` is required/)
  dob('pptx: comment `author` nelkul -> NEVESITETT hiba (mert: a builder sajat fallbackje EBBEN a kornyezetben ures nevet ad, nem valodi identitast)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'comment', text: 'proba' }] }),
    /comment: `author` is required/)

  dob('comment: docx magban nem all rendelkezesre (pptx-only)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'comment', text: 'proba', author: 'V' }] }),
    /not available in the docx core/)
}

console.log('\n[44] shape.line (Shape.SetLine LETREHOZAS UTAN), pptx-only')
console.log('    a hivas-alak toString()-bol lett visszafejtve; csomag-szintu igazolas (a:ln w=/srgbClr, poz+neg ugyanazon szkriptben) elo DS-hivassal, nem itt')
{
  const withLine = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', line: { width: 70000, color: [255, 0, 0] } }] })
  check('pptx: shape.line -> oShape.SetLine(Api.CreateStroke(width, solidFill)), a MEGADOTT ertekekkel',
    /oShape\.SetLine\(Api\.CreateStroke\(70000, Api\.CreateSolidFill\(Api\.CreateRGBColor\(255, 0, 0\)\)\)\);/.test(withLine.script), withLine.script)

  const widthOnly = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', line: { width: 30000 } }] })
  check('pptx: shape.line szin nelkul -> NoFill (lathatatlan vonal, ugyanaz a szemantika, mint a letrehozaskori borderColor hianya)',
    /oShape\.SetLine\(Api\.CreateStroke\(30000, Api\.CreateNoFill\(\)\)\);/.test(widthOnly.script))

  check('pptx: shape.line width nelkul -> 0 (alap, ugyanaz mint a letrehozaskori borderWidth alapertelmezese)',
    /oShape\.SetLine\(Api\.CreateStroke\(0, /.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', line: { color: [0, 0, 0] } }] }).script))

  check('pptx: shape.line NELKUL (alap) nincs SetLine hivas',
    !/SetLine/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'shape' }] }).script))

  dob('pptx: shape.line nem-objektum -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', line: 'vastag' }] }),
    /shape: line must be an object/)

  dob('docx: shape.line -> NEVESITETT megtagadas, pptx-only ebben az egysegben',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', line: { width: 1000 } }] }),
    /shape: line is pptx-only in this unit/)
}

console.log('\n[45] replaceText -- a mert hivas-alak (SearchAndReplace)')
const csere = buildCreateScript({ core: 'docx', operations: [{ type: 'replaceText', search: 'RÉGI', replace: 'ÚJ' }], outName: 'x.docx' })
check('replaceText: a mert hivas-alak', /oDocument\.SearchAndReplace\(\{ searchString: "RÉGI", replaceString: "ÚJ" \}\);/.test(csere.script))
const cserePuszta = buildCreateScript({ core: 'docx', operations: [{ type: 'replaceText', search: 'X' }], outName: 'x.docx' })
check('replaceText: `replace` nelkul -> ures sztringre cserel, nem dob', /replaceString: ""/.test(cserePuszta.script))
dob('replaceText: `search` nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'replaceText', replace: 'x' }], outName: 'x.docx' }), /`search`.*required/)
dob('replaceText: ures `search` -> UGYANAZ a megtagadas (nem kuldjuk el mert-nelkul a szervernek)', () => buildCreateScript({ core: 'docx', operations: [{ type: 'replaceText', search: '' }], outName: 'x.docx' }), /`search`.*required/)
dob('replaceText: xlsx magban nem all rendelkezesre (MERVE: oWorksheet.SearchAndReplace hianyzik)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'replaceText', search: 'x' }] }), /not available in the xlsx core/)
dob('replaceText: pptx magban nem all rendelkezesre (docx-only)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'replaceText', search: 'x' }], slideCount: 1 }), /not available in the pptx core/)

console.log('\n[31] addComment -- a mert hivas-alak (Search + AddComment, GRACEFUL SKIP ha nincs talalat)')
const komment = buildCreateScript({ core: 'docx', operations: [{ type: 'addComment', anchorText: 'CEL', text: 'megjegyzes', author: 'Teszt' }], outName: 'x.docx' })
check('addComment: a Search hivas a mert alakban', /oDocument\.Search\("CEL"\);/.test(komment.script))
check('addComment: az AddComment IF-fel korulvéve, nem feltetel nelkul', /if \(oCommentResults && oCommentResults\.length > 0\) \{/.test(komment.script))
check('addComment: az AddComment hivas a mert parameterekkel', /AddComment\("megjegyzes", "Teszt", "TE"\);/.test(komment.script))
const kommentAlapAuthor = buildCreateScript({ core: 'docx', operations: [{ type: 'addComment', anchorText: 'CEL', text: 'x' }], outName: 'x.docx' })
check('addComment: `author` nelkul -> "euro-mcp" alapertelmezett, initials "EU"', /AddComment\("x", "euro-mcp", "EU"\);/.test(kommentAlapAuthor.script))
dob('addComment: `anchorText` nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'addComment', text: 'x' }], outName: 'x.docx' }), /`anchorText`.*required/)
dob('addComment: `text` nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'addComment', anchorText: 'x' }], outName: 'x.docx' }), /`text`.*required/)
dob('addComment: xlsx magban nem all rendelkezesre (docx-only, ugyanaz a hatar mint replaceText)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'addComment', anchorText: 'x', text: 'y' }] }), /not available in the xlsx core/)

console.log('\n[32] trackedChanges -- a mert hivas-alak (SetTrackRevisions)')
const nyomon = buildCreateScript({ core: 'docx', operations: [{ type: 'trackedChanges' }], outName: 'x.docx' })
check('trackedChanges: `enabled` nelkul -> true (a mert alak)', /oDocument\.SetTrackRevisions\(true\);/.test(nyomon.script))
const nyomonKi = buildCreateScript({ core: 'docx', operations: [{ type: 'trackedChanges', enabled: false }], outName: 'x.docx' })
check('trackedChanges: `enabled:false` -> false-t ir, nem a true-t hagyja bent', /oDocument\.SetTrackRevisions\(false\);/.test(nyomonKi.script))
dob('trackedChanges: xlsx magban nem all rendelkezesre (docx-only)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'trackedChanges' }] }), /not available in the xlsx core/)
dob('trackedChanges: pptx magban nem all rendelkezesre (docx-only)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'trackedChanges' }], slideCount: 1 }), /not available in the pptx core/)

console.log('\n[46] UJ MUVELET-TIPUS: slide (dia-szintu muveletek), pptx-only')
console.log('    hivas-alakok toString()-bol visszafejtve; csomag-szintu igazolas (delete/duplicateTo/moveTo/removeObject/SetVisible/background/transition) elo DS-hivassal, nem itt')
{
  dob('slide: docx magban nem all rendelkezesre', () => buildCreateScript({ core: 'docx', operations: [{ type: 'slide', visible: false }], outName: 'x.docx' }), /not available in the docx core/)
  dob('slide: xlsx magban nem all rendelkezesre', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'slide', visible: false }] }), /not available in the xlsx core/)

  dob('slide: mezo nelkul -> megtagadva', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide' }], slideCount: 1 }), /at least one field is required/)

  const bgNone = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', background: 'none' }], slideCount: 1 })
  check('slide: background "none" -> ClearBackground', /oSlide\.ClearBackground\(\);/.test(bgNone.script))
  const bgLayout = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', background: 'layout' }], slideCount: 1 })
  check('slide: background "layout" -> FollowLayoutBackground', /oSlide\.FollowLayoutBackground\(\);/.test(bgLayout.script))
  const bgMaster = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', background: 'master' }], slideCount: 1 })
  check('slide: background "master" -> FollowMasterBackground', /oSlide\.FollowMasterBackground\(\);/.test(bgMaster.script))
  const bgColor = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', background: [10, 20, 30] }], slideCount: 1 })
  check('slide: background [r,g,b] -> SetBackground(SolidFill(RGBColor))', /oSlide\.SetBackground\(Api\.CreateSolidFill\(Api\.CreateRGBColor\(10, 20, 30\)\)\);/.test(bgColor.script))
  dob('slide: background rossz alaku szin -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', background: [999, 0, 0] }], slideCount: 1 }), /color.*values must each be 0-255/)
  dob('slide: background ismeretlen sztring -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', background: 'sky' }], slideCount: 1 }), /background must be "none"\/"layout"\/"master"/)

  const visOn = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', visible: true }], slideCount: 1 })
  check('slide: visible:true -> SetVisible(true)', /oSlide\.SetVisible\(true\);/.test(visOn.script))
  const visOff = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', visible: false }], slideCount: 1 })
  check('slide: visible:false -> SetVisible(false)', /oSlide\.SetVisible\(false\);/.test(visOff.script))

  const remAll = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', removeAllObjects: true }], slideCount: 1 })
  check('slide: removeAllObjects -> RemoveAllObjects()', /oSlide\.RemoveAllObjects\(\);/.test(remAll.script))

  const remObj = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', removeObject: { pos: 1, count: 2 } }], slideCount: 1 })
  check('slide: removeObject {pos,count} -> RemoveObject(pos, count)', /oSlide\.RemoveObject\(1, 2\);/.test(remObj.script))
  const remObjDefaultCount = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', removeObject: { pos: 0 } }], slideCount: 1 })
  check('slide: removeObject count nelkul -> 1 (alap)', /oSlide\.RemoveObject\(0, 1\);/.test(remObjDefaultCount.script))
  dob('slide: removeObject pos nelkul -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', removeObject: {} }], slideCount: 1 }), /removeObject\.pos must be a non-negative integer/)

  const delOp = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', delete: true }], slideCount: 2 })
  check('slide: delete -> oSlide.Delete()', /oSlide\.Delete\(\);/.test(delOp.script))
  const dupOp = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', duplicateTo: 2 }], slideCount: 2 })
  check('slide: duplicateTo -> oSlide.Duplicate(pos)', /oSlide\.Duplicate\(2\);/.test(dupOp.script))
  const moveOp = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', moveTo: 0 }], slideCount: 2 })
  check('slide: moveTo -> oSlide.MoveTo(pos)', /oSlide\.MoveTo\(0\);/.test(moveOp.script))

  dob('slide: delete + duplicateTo egyutt -> NEVESITETT hiba (csak egy strukturalis mezo)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', delete: true, duplicateTo: 1 }], slideCount: 2 }),
    /only one of delete\/duplicateTo\/moveTo/)
  dob('slide: delete + background egyutt -> NEVESITETT hiba (strukturalis + tartalmi nem keverheto)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', delete: true, background: 'none' }], slideCount: 2 }),
    /cannot be combined with content fields/)

  const trans = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', transition: { effect: 'effectFade', speed: 'fast', advanceOnClick: false, advanceOnTime: true, advanceTime: 5000 } }], slideCount: 1 })
  check('slide: transition -> Api.CreateSlideShowTransition + setterek + SetSlideShowTransition',
    /Api\.CreateSlideShowTransition\(\)/.test(trans.script)
    && /oTransition\.SetEntryEffect\("effectFade"\);/.test(trans.script)
    && /oTransition\.SetSpeed\("fast"\);/.test(trans.script)
    && /oTransition\.SetAdvanceOnClick\(false\);/.test(trans.script)
    && /oTransition\.SetAdvanceOnTime\(true\);/.test(trans.script)
    && /oTransition\.SetAdvanceTime\(5000\);/.test(trans.script)
    && /oSlide\.SetSlideShowTransition\(oTransition\);/.test(trans.script),
    trans.script)
  const transDuration = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', transition: { duration: 750 } }], slideCount: 1 })
  check('slide: transition.duration -> SetDuration', /oTransition\.SetDuration\(750\);/.test(transDuration.script))
  dob('slide: transition ismeretlen effect -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', transition: { effect: 'effectWhoosh' } }], slideCount: 1 }), /unknown transition\.effect/)
  dob('slide: transition ismeretlen speed -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', transition: { speed: 'ludicrous' } }], slideCount: 1 }), /unknown transition\.speed/)
  dob('slide: ures transition objektum -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', transition: {} }], slideCount: 1 }), /transition object must set at least one/)

  const group = buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', group: [{ shapeType: 'rect', width: 500000, height: 500000, x: 0, y: 0, fill: [255, 0, 0] }, { shapeType: 'ellipse', width: 500000, height: 500000, x: 600000, y: 0 }] }], slideCount: 1 })
  check('slide: group -> ket alakzat letrehozva sajat valtozonevvel es hozzaadva, majd GroupDrawings a ket valtozoval',
    /var oGroupShape0 = Api\.CreateShape\("rect", 500000, 500000, Api\.CreateSolidFill\(Api\.CreateRGBColor\(255, 0, 0\)\), Api\.CreateStroke\(0, Api\.CreateNoFill\(\)\)\);/.test(group.script)
    && /oGroupShape0\.SetPosition\(0, 0\);/.test(group.script)
    && /oSlide\.AddObject\(oGroupShape0\);/.test(group.script)
    && /var oGroupShape1 = Api\.CreateShape\("ellipse", 500000, 500000, Api\.CreateNoFill\(\), Api\.CreateStroke\(0, Api\.CreateNoFill\(\)\)\);/.test(group.script)
    && /oSlide\.GroupDrawings\(\[oGroupShape0, oGroupShape1\]\);/.test(group.script),
    group.script)
  dob('slide: group egyetlen alakzattal -> NEVESITETT hiba (kevesebb mint 2 nem csoport)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'slide', group: [{ shapeType: 'rect' }] }], slideCount: 1 }), /at least 2 shape descriptors/)
}

console.log('\n[47] docx-K4: urlapok es tartalom-vezerlok (`formField` uj muvelet-tipus)')
{
  const szoveg = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'formField', kind: 'text', key: 'myfield', tip: 'enter text', placeholder: 'type here' }],
  }).script
  check('formField(text): Api.CreateTextForm LITERAL key/tip/placeholder-rel, AddElement+Push',
    /Api\.CreateTextForm\(\{key: "myfield", tip: "enter text", required: false, placeholder: "type here", comb: false\}\);/.test(szoveg) &&
    /oFormParagraph\.AddElement\(oForm\);\noDocument\.Push\(oFormParagraph\);/.test(szoveg), szoveg)

  const jelolo = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'formField', kind: 'checkbox', key: 'mycheck', checked: true, required: true }],
  }).script
  check('formField(checkbox): Api.CreateCheckBoxForm, required+checked LITERALKENT',
    /Api\.CreateCheckBoxForm\(\{key: "mycheck", tip: "", required: true, checked: true\}\);/.test(jelolo), jelolo)

  const legordulo = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'formField', kind: 'combobox', key: 'mycombo', items: ['A', 'B', 'C'] }],
  }).script
  check('formField(combobox): Api.CreateComboBoxForm, items TOMBKENT',
    /Api\.CreateComboBoxForm\(\{key: "mycombo", tip: "", required: false, format: "", items: \["A", "B", "C"\], editable: false\}\);/.test(legordulo), legordulo)

  dob('formField(combobox): ures/hianyzo items -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'formField', kind: 'combobox', key: 'x', items: [] }] }),
    /combobox.*non-empty `items`/)

  const datum = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'formField', kind: 'date', key: 'mydate', format: 'YYYY-MM-DD' }],
  }).script
  check('formField(date): Api.CreateDateForm, format LITERALKENT, alap MM/DD/YYYY ha nincs kerve',
    /Api\.CreateDateForm\(\{key: "mydate", tip: "", required: false, format: "YYYY-MM-DD"\}\);/.test(datum) &&
    /format: "MM\/DD\/YYYY"/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'formField', kind: 'date', key: 'x' }] }).script))

  const kep = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'formField', kind: 'picture', key: 'mypic' }],
  }).script
  check('formField(picture): Api.CreatePictureForm', /Api\.CreatePictureForm\(\{key: "mypic", tip: "", required: false, scaleFlag: "always", lockAspectRatio: true\}\);/.test(kep))

  dob('formField: `key` kotelezo', () => buildCreateScript({ core: 'docx', operations: [{ type: 'formField', kind: 'text' }] }), /`key` is required/)

  dob('formField: ismeretlen kind -> megtagadva',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'formField', kind: 'radio', key: 'x' }] }),
    /unknown kind.*known: text, checkbox, combobox, date, picture/)

  dob('formField: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'formField', kind: 'text', key: 'x' }] }), /not available in the pptx core/)
}

console.log('\n[48] docx-K3/4 negyedik szelete: `bookmark` + `bookmarkRef` a `text` mezoin')
{
  const bmScript = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', bookmark: 'mybm' }],
  }).script
  check('text(docx): bookmark -> Push UTAN, GetRange().AddBookmark(name) LITERALKENT',
    /oDocument\.Push\(oParagraph\);\noParagraph\.GetRange\(\)\.AddBookmark\("mybm"\);/.test(bmScript), bmScript)

  check('text(docx): bookmark NELKUL nincs AddBookmark hivas',
    !/AddBookmark\(/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }] }).script))

  dob('text(docx): bookmark ures string -> megtagadva', () => buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', bookmark: '' }],
  }), /bookmark.*non-empty/)

  const refScript = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'ref: ', bookmarkRef: { to: 'mybm', refTo: 'text', link: true } }],
  }).script
  check('text(docx): bookmarkRef -> AddBookmarkCrossRef(refTo, to, link, aboveBelow, sepWith), MEASURED ARGUMENT ORDER (nem to/displayText/format)',
    /oParagraph\.AddBookmarkCrossRef\("text", "mybm", true, false, ""\);/.test(refScript), refScript)

  check('text(docx): bookmarkRef alapertekek -- link alap true, aboveBelow alap false, sepWith alap ures string',
    /AddBookmarkCrossRef\("pageNum", "mybm", true, false, ""\);/.test(buildCreateScript({
      core: 'docx',
      operations: [{ type: 'text', text: 'a', bookmarkRef: { to: 'mybm', refTo: 'pageNum' } }],
    }).script))

  check('text(docx): bookmarkRef NELKUL nincs AddBookmarkCrossRef hivas',
    !/AddBookmarkCrossRef/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }] }).script))

  dob('text(docx): bookmarkRef.refTo ismeretlen -> megtagadva (a builder csendben false-t adna vissza, nem dob)', () => buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', bookmarkRef: { to: 'mybm', refTo: 'nincsilyen' } }],
  }), /bookmarkRef\.refTo.*unknown/)

  dob('text(docx): bookmarkRef.to hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', bookmarkRef: { refTo: 'text' } }],
  }), /bookmarkRef\.to.*non-empty/)

  // MERT SORREND (2026-08-17, koteg09): a bookmark ugyanabba a csapdaba fut, mint a comments --
  // DE FORDITVA rangsorolva a footnotes-hoz kepest: bookmark<->comments sorrend-fuggetlen,
  // bookmark MINDIG a footnotes/endnotes ELOTT kell (kulonben a bookmarkStart/End marker
  // csendben elmarad, holott a bookmarks-manager es egy kesobbi bookmarkRef meg "sikeresnek"
  // latszik). A generator ezt FIX sorrendben rakja ki, a hivo mezo-sorrendjetol fuggetlenul.
  const sorrendScript = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'a', footnotes: ['fn'], comments: ['c1'], bookmark: 'bm1' }],
  }).script
  check('text(docx): comments+bookmark+footnotes EGYUTT -> AddComment < AddBookmark < AddFootnote a szkriptben (a hivo mezo-sorrendjetol fuggetlenul)',
    sorrendScript.indexOf('AddComment(') < sorrendScript.indexOf('AddBookmark(') &&
    sorrendScript.indexOf('AddBookmark(') < sorrendScript.indexOf('AddFootnote()'), sorrendScript)

  check('text(docx): bookmarkRef a Push+notes UTAN all a szkriptben (sajat paragrafusara hivatkozo self-ref is ervenyes sorrend)',
    (() => {
      const s = buildCreateScript({
        core: 'docx',
        operations: [{ type: 'text', text: 'a', bookmark: 'selfbm', bookmarkRef: { to: 'selfbm', refTo: 'text' } }],
      }).script
      return s.indexOf('AddBookmark(') < s.indexOf('AddBookmarkCrossRef(')
    })())

  check('text(pptx): bookmark/bookmarkRef mezok NEM emittalnak semmit (docx-only)',
    !/AddBookmark/.test(buildCreateScript({
      core: 'pptx',
      operations: [{ type: 'text', text: 'a', bookmark: 'x', bookmarkRef: { to: 'x', refTo: 'text' } }],
    }).script))
}

console.log('\n[49] docx-K5: vizjel (watermark), valtozas-kovetes (trackChanges), kereses-csere (searchReplace)')
{
  const wm = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'watermark', text: 'DRAFT' }],
  }).script
  check('watermark: alapertekekkel -> InsertWatermark type text/isDiagonal true/isAuto true/color 000000/transparent true',
    /oDocument\.InsertWatermark\(\{type: "text", text: "DRAFT", isDiagonal: true, isAuto: true, color: "000000", transparent: true\}\);/.test(wm), wm)

  const wmCustom = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'watermark', text: 'CONFIDENTIAL', diagonal: false, color: 'FF0000', transparent: false }],
  }).script
  check('watermark: diagonal:false/color/transparent:false -> LITERALKENT athelyettesitve',
    /InsertWatermark\(\{type: "text", text: "CONFIDENTIAL", isDiagonal: false, isAuto: true, color: "FF0000", transparent: false\}\);/.test(wmCustom))

  check('watermark: remove:true -> RemoveWatermark, `text` NELKUL is',
    /oDocument\.RemoveWatermark\(\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'watermark', remove: true }] }).script) &&
    !/InsertWatermark/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'watermark', remove: true }] }).script))

  dob('watermark: `text` hianyzik es nincs remove -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'watermark' }] }), /`text` is required unless `remove: true`/)

  check('watermark: pptx-en nem elerheto', (() => {
    try { buildCreateScript({ core: 'pptx', operations: [{ type: 'watermark', text: 'x' }] }); return false }
    catch (e) { return /not available in the pptx core/.test(e.message) }
  })())

  const trOn = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'trackChanges', enabled: true }],
  }).script
  check('trackChanges: enabled:true -> SetTrackRevisions(true)', /oDocument\.SetTrackRevisions\(true\);/.test(trOn))
  const trOff = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'trackChanges', enabled: false }],
  }).script
  check('trackChanges: enabled:false -> SetTrackRevisions(false)', /oDocument\.SetTrackRevisions\(false\);/.test(trOff))
  dob('trackChanges: `enabled` hianyzik -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'trackChanges' }] }), /`enabled` is required/)

  const sr = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'searchReplace', search: 'quick', replace: 'slow', matchCase: true }],
  }).script
  check('searchReplace: search/replace/matchCase -> SearchAndReplace({searchString, replaceString, matchCase}) LITERALKENT',
    /oDocument\.SearchAndReplace\(\{searchString: "quick", replaceString: "slow", matchCase: true\}\);/.test(sr), sr)
  check('searchReplace: matchCase alap false',
    /matchCase: false\}\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'searchReplace', search: 'a', replace: 'b' }] }).script))
  dob('searchReplace: `search` hianyzik -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'searchReplace', replace: 'x' }] }), /`search` is required/)
  dob('searchReplace: `replace` hianyzik -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'searchReplace', search: 'x' }] }), /`replace` is required/)
  check('searchReplace: ures string `replace` (torles) MEGENGEDETT, nem hianyzo mezo',
    /replaceString: ""/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'searchReplace', search: 'x', replace: '' }] }).script))
}

console.log('\n[50] sheetComments (K6) -- a mert hivas-alak (oWorksheet.GetComments(), NEM Api.GetComments())')
check('sheetComments: a query IIFE a mert `at` cellaba ir, oWorksheet.GetComments()-t hasznal',
  /oWorksheet\.GetComments\(\)/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetComments', at: 'Z1' }] }).script) &&
  /oWorksheet\.GetRange\("Z1"\)\.SetValue\(/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetComments', at: 'Z1' }] }).script))
check('sheetComments: NEM Api.GetComments()-t hivja (az mert-en broken, mindig ures)',
  !/Api\.GetComments\(\)/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetComments', at: 'Z1' }] }).script))
dob('sheetComments: `at` nelkul -> megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetComments' }] }), /`at`.*required/)
dob('sheetComments: docx magban nem all rendelkezesre (xlsx-only)', () => buildCreateScript({ core: 'docx', operations: [{ type: 'sheetComments', at: 'Z1' }], outName: 'x.docx' }), /not available in the docx core/)

console.log('\n[51] sheetTheme (K6) -- a mert hivas-alak (Api.SetThemeColors(NAME), NEM szin-tomb)')
check('sheetTheme: `name` -> Api.SetThemeColors("Aspect") STRING-kent, nem szin-tomb',
  /Api\.SetThemeColors\("Aspect"\);/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetTheme', name: 'Aspect' }] }).script))
dob('sheetTheme: `name` nelkul -> megtagadva, felsorolja az ismert neveket', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetTheme' }] }), /`name`.*required.*Aspect/)
dob('sheetTheme: ismeretlen nev -> megtagadva -- MERT: a builder csendben semmit nem valtoztat egy ismeretlen nevre', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'sheetTheme', name: 'NemLetezoTemaNev123' }] }), /unknown theme name.*does not throw inside the builder, it silently changes nothing/)
dob('sheetTheme: pptx magban nem all rendelkezesre (xlsx-only)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'sheetTheme', name: 'Aspect' }], slideCount: 1 }), /not available in the pptx core/)

console.log('\n[52] recalculateFormulas (K6) -- MINDIG megtagadva -- MERT: Api.RecalculateAllFormulas() mert redundans no-op')
dob('recalculateFormulas: MINDIG megtagadva', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'recalculateFormulas' }] }), /confirmed redundant no-op/)
dob('recalculateFormulas: docx magban nem all rendelkezesre (xlsx-only)', () => buildCreateScript({ core: 'docx', operations: [{ type: 'recalculateFormulas' }], outName: 'x.docx' }), /not available in the docx core/)

console.log('\n[53] docx-D3: hiperhivatkozas (`text.hyperlink`), matematikai keplet (`mathEquation`), egyeni bekezdes-stilus (`text.customStyle`)')
{
  const hl = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'click me', hyperlink: { url: 'https://x.invalid/y', tooltip: 'tip' } }],
  }).script
  check('text(docx): hyperlink -> AddText UTAN, Push ELOTT, Paragraph.AddHyperlink(url, tooltip) LITERALKENT',
    hl.indexOf('oParagraph.AddText("click me");') < hl.indexOf('oParagraph.AddHyperlink("https://x.invalid/y", "tip");') &&
    hl.indexOf('oParagraph.AddHyperlink("https://x.invalid/y", "tip");') < hl.indexOf('oDocument.Push(oParagraph);'), hl)

  check('text(docx): hyperlink.tooltip alap ures string',
    /AddHyperlink\("https:\/\/x\.invalid\/y", ""\);/.test(buildCreateScript({
      core: 'docx', operations: [{ type: 'text', text: 'a', hyperlink: { url: 'https://x.invalid/y' } }],
    }).script))

  check('text(docx): hyperlink NELKUL nincs AddHyperlink hivas',
    !/AddHyperlink/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }] }).script))

  dob('text(docx): hyperlink.url hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'text', text: 'a', hyperlink: {} }],
  }), /`hyperlink\.url` is required/)

  check('text(pptx): hyperlink mezo NEM emittal AddHyperlink-et (docx-only, a runs\\[\\]\\.hyperlink/shape\\.hyperlink marad a mar meglevo pptx-uton)',
    !/oParagraph\.AddHyperlink/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'a', hyperlink: { url: 'https://x.invalid/y' } }] }).script))

  const math = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'mathEquation', text: 'x^2+y^2=z^2', format: 'unicode' }],
  }).script
  check('mathEquation: MoveCursorToEnd + AddMathEquation(text, format) LITERALKENT',
    /oDocument\.MoveCursorToEnd\(\);\noDocument\.AddMathEquation\("x\^2\+y\^2=z\^2", "unicode"\);/.test(math), math)

  check('mathEquation: format alap "unicode"',
    /AddMathEquation\("a", "unicode"\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'mathEquation', text: 'a' }] }).script))

  dob('mathEquation: `text` hianyzik -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'mathEquation' }] }), /`text` is required/)

  dob('mathEquation: ismeretlen format -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'mathEquation', text: 'a', format: 'nincsilyen' }],
  }), /unknown format.*known: latex, unicode, mathml/)

  dob('mathEquation: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'mathEquation', text: 'a' }] }), /not available in the pptx core/)

  const style = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'styled', customStyle: { name: 'MyStyle', bold: true, color: 'c81414', size: 32, font: 'Georgia' } }],
  }).script
  check('text(docx): customStyle -> CreateStyle("name","paragraph") + CreateTextPr + Set*(...) + SetTextPr + paragraph.SetStyle, MIND JELEN, PONTOS ERTEKKEL',
    /var oCustomStyle = oDocument\.CreateStyle\("MyStyle", "paragraph"\);/.test(style) &&
    /oCustomTextPr\.SetBold\(true\);/.test(style) &&
    /oCustomTextPr\.SetFontSize\(32\);/.test(style) &&
    /oCustomTextPr\.SetFontFamily\("Georgia"\);/.test(style) &&
    /oCustomTextPr\.SetColor\(200, 20, 20, false\);/.test(style) &&
    /oCustomStyle\.SetTextPr\(oCustomTextPr\);\noParagraph\.SetStyle\(oCustomStyle\);/.test(style), style)

  check('text(docx): customStyle a Push ELOTT all a szkriptben',
    style.indexOf('oParagraph.SetStyle(oCustomStyle);') < style.indexOf('oDocument.Push(oParagraph);'))

  check('text(docx): customStyle NELKUL nincs CreateStyle hivas',
    !/CreateStyle/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'a' }] }).script))

  dob('text(docx): customStyle.name hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'text', text: 'a', customStyle: {} }],
  }), /`customStyle\.name` is required/)

  check('text(pptx): customStyle mezo NEM emittal CreateStyle-t (docx-only)',
    !/CreateStyle/.test(buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'a', customStyle: { name: 'x' } }] }).script))
}

console.log('\n[54] docx-K6: BEAGYAZOTT TABLAZAT egy cella erteken belul (`{ table: {...} }`), rekurzivan')
{
  const flat = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a', 'b'], ['c', 'd']] }],
  }).script
  check('table(docx): REGRESSZIO -- string-ertekes rows valtozatlanul AddText-et emittal, nincs GetContent().Push()',
    /oTable\.GetCell\(0, 0\)\.GetContent\(\)\.GetElement\(0\)\.AddText\("a"\);/.test(flat) &&
    !/GetContent\(\)\.Push\(/.test(flat), flat)

  const nested = buildCreateScript({
    core: 'docx',
    operations: [{
      type: 'table',
      rows: [
        ['outer-a', { table: { rows: [['inner-1', 'inner-2']] } }],
      ],
    }],
  }).script
  check('table(docx): egy cella-ertek `{ table: {...} }` -> masodik Api.CreateTable a cella tartalmaba Push-olva, NEM AddText',
    /var oTable = Api\.CreateTable\(2, 1\);/.test(nested) &&
    /var oNestedTable1 = Api\.CreateTable\(2, 1\);/.test(nested) &&
    /oNestedTable1\.GetCell\(0, 0\)\.GetContent\(\)\.GetElement\(0\)\.AddText\("inner-1"\);/.test(nested) &&
    /oTable\.GetCell\(0, 1\)\.GetContent\(\)\.Push\(oNestedTable1\);/.test(nested) &&
    !/oTable\.GetCell\(0, 1\)\.GetContent\(\)\.GetElement\(0\)\.AddText/.test(nested), nested)

  check('table(docx): a beagyazott tabla var-neve elottebb epul, MIELOTT a Push-sor -- a hivo sorrendjetol fuggetlenul all a szkriptben',
    nested.indexOf('var oNestedTable1 = Api.CreateTable') < nested.indexOf('GetContent().Push(oNestedTable1)'), nested)

  const nestedStyled = buildCreateScript({
    core: 'docx',
    operations: [{
      type: 'table',
      rows: [['h1', 'h2'], ['a', { table: { rows: [['x', 'y'], ['z', 'w'], ['p', 'q']], header: false, zebra: true, border: false } }]],
    }],
  }).script
  check('table(docx): a beagyazott tabla a SAJAT header/zebra/border opcioit kapja, nem az orokolt/kulso tablaet -- kulso tabla hatarolt, belso nem; belso zebra a sajat (header:false) sorindexeles szerint fest (2. sor, i=2)',
    /var oTable = Api\.CreateTable\(2, 2\);/.test(nestedStyled) &&
    /oTable\.SetTableBorderTop/.test(nestedStyled) &&
    /var oNestedTable1 = Api\.CreateTable\(2, 3\);/.test(nestedStyled) &&
    !/oNestedTable1\.SetTableBorderTop/.test(nestedStyled) &&
    !/oNestedTable1\.GetCell\(0, 0\)\.GetContent\(\)\.GetElement\(0\)\.SetBold/.test(nestedStyled) &&
    /oNestedTable1\.GetCell\(2, 0\)\.SetShd/.test(nestedStyled), nestedStyled)

  const doubleNested = buildCreateScript({
    core: 'docx',
    operations: [{
      type: 'table',
      rows: [[{ table: { rows: [[{ table: { rows: [['deepest']] } }]] } }]],
    }],
  }).script
  check('table(docx): KET szintu beagyazas -> HAROM Api.CreateTable hivas, egymasba agyazott Push-okkal (a kod rekurziv, MERT 2 szintig)',
    (doubleNested.match(/Api\.CreateTable\(/g) || []).length === 3 &&
    /oNestedTable2\.GetCell\(0, 0\)\.GetContent\(\)\.GetElement\(0\)\.AddText\("deepest"\);/.test(doubleNested) &&
    /oNestedTable1\.GetCell\(0, 0\)\.GetContent\(\)\.Push\(oNestedTable2\);/.test(doubleNested) &&
    /oTable\.GetCell\(0, 0\)\.GetContent\(\)\.Push\(oNestedTable1\);/.test(doubleNested), doubleNested)

  dob('table(docx): beagyazott tabla `rows` uresen -> megtagadva, UGYANAZ a hibauzenet mint a felso szinten',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [[{ table: { rows: [] } }]] }] }),
    /`rows` is empty/)

  check('table(docx): beagyazott tabla tableLook/extraRows/extraColumns UGYANUGY megtagadva, mint a felso szinten (pptx-only)',
    (() => {
      try {
        buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [[{ table: { rows: [['x']], tableLook: {} } }]] }] })
        return false
      } catch (e) { return /pptx-only/.test(e.message) }
    })())

  check('table(pptx): `{ table: {...} }` cellaertek NEM kap kulon banasmodot -- egyszeruen szoveggé konvertalt cellaertek (pptx-only cella-run logika valtozatlan)',
    (() => {
      const s = buildCreateScript({ core: 'pptx', operations: [{ type: 'table', rows: [[{ table: { rows: [['x']] } }]] }] }).script
      return !/oNestedTable/.test(s)
    })())
}

console.log('\n[55] FEJLEC/LABLEC + OLDALSZAM (`headerFooter` uj muvelet-tipus)')
{
  const hdr = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'headerFooter', target: 'header', text: 'MY HEADER' }],
  }).script
  check('headerFooter: target header, text -> GetHeader("default", true), AddText, oHeader.Push',
    /var oSection = oDocument\.GetFinalSection\(\);\nvar oHeader = oSection\.GetHeader\("default", true\);/.test(hdr) &&
    /oHFParagraph\.AddText\("MY HEADER"\);/.test(hdr) &&
    /oHeader\.Push\(oHFParagraph\);/.test(hdr), hdr)

  const ftr = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'headerFooter', target: 'footer', text: 'MY FOOTER' }],
  }).script
  check('headerFooter: target footer -> GetFooter, oFooter.Push',
    /var oFooter = oSection\.GetFooter\("default", true\);/.test(ftr) &&
    /oFooter\.Push\(oHFParagraph\);/.test(ftr), ftr)

  const variant = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'headerFooter', target: 'header', variant: 'even', text: 'x' }],
  }).script
  check('headerFooter: variant "even" -> GetHeader("even", true) LITERALKENT',
    /GetHeader\("even", true\)/.test(variant))

  dob('headerFooter: ismeretlen variant -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'headerFooter', target: 'header', variant: 'nincsilyen', text: 'x' }],
  }), /unknown variant.*known: default, even, title/)

  const pageNum = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'headerFooter', target: 'footer', parts: ['Page ', { pageNumber: true }, ' of ', { pagesCount: true }] }],
  }).script
  check('headerFooter: parts tomb -> AddText/AddPageNumber/AddText/AddPagesCount, PONTOSAN ebben a sorrendben',
    pageNum.indexOf('oHFParagraph.AddText("Page ");') <
    pageNum.indexOf('oHFParagraph.AddPageNumber();') &&
    pageNum.indexOf('oHFParagraph.AddPageNumber();') <
    pageNum.indexOf('oHFParagraph.AddText(" of ");') &&
    pageNum.indexOf('oHFParagraph.AddText(" of ");') <
    pageNum.indexOf('oHFParagraph.AddPagesCount();'), pageNum)

  dob('headerFooter: `target` hianyzik/ervenytelen -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'headerFooter', text: 'x' }],
  }), /`target` must be "header" or "footer"/)

  dob('headerFooter: sem `parts`, sem `text` -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'headerFooter', target: 'header' }],
  }), /`parts`.*or `text` is required/)

  dob('headerFooter: parts[i] ismeretlen alaku -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'headerFooter', target: 'header', parts: [{ nincsilyen: true }] }],
  }), /parts\[0\] must be a string/)

  dob('headerFooter: pptx-en nem elerheto', () => buildCreateScript({
    core: 'pptx', operations: [{ type: 'headerFooter', target: 'header', text: 'x' }],
  }), /not available in the pptx core/)
}

console.log('\n[56] docx-D1: export (ToMarkdown/ToHtml) + documentStats (GetStatistics/GetCustomProperties), (A1) marker-alak, GetDocumentInfo megtagadva')
{
  const md = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'export', format: 'markdown' }],
  }).script
  check('export: format markdown -> ToMarkdown() + uj bekezdes "__EXPORT_MARKDOWN__:" markerrel',
    /var oExportResult = oDocument\.ToMarkdown\(\);/.test(md) &&
    /oExportParagraph\.AddText\("__EXPORT_MARKDOWN__:" \+ oExportResult\);/.test(md) &&
    /oDocument\.Push\(oExportParagraph\);/.test(md), md)

  const htmlS = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'export', format: 'html' }],
  }).script
  check('export: format html -> ToHtml() + "__EXPORT_HTML__:" marker',
    /var oExportResult = oDocument\.ToHtml\(\);/.test(htmlS) &&
    /"__EXPORT_HTML__:" \+ oExportResult/.test(htmlS), htmlS)

  check('export: format alap "markdown"',
    /ToMarkdown\(\)/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'export' }] }).script))

  dob('export: ismeretlen format -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'export', format: 'pdf' }],
  }), /unknown format.*known: markdown, html/)

  dob('export: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'export' }] }), /not available in the pptx core/)

  const stats = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'documentStats' }],
  }).script
  check('documentStats: GetStatistics() + "__STATS_JSON__:" marker, GetCustomProperties NELKUL alapertelmezetten',
    /var oStatsResult = oDocument\.GetStatistics\(\);/.test(stats) &&
    /oStatsParagraph\.AddText\("__STATS_JSON__:" \+ JSON\.stringify\(oStatsResult\)\);/.test(stats) &&
    !/GetCustomProperties/.test(stats), stats)

  const statsWithProps = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'documentStats', includeCustomProperties: true }],
  }).script
  check('documentStats: includeCustomProperties:true -> MASODIK bekezdes GetCustomProperties()-szel, "__CUSTOM_PROPERTIES_JSON__:" markerrel',
    /var oCustomPropsResult = oDocument\.GetCustomProperties\(\);/.test(statsWithProps) &&
    /"__CUSTOM_PROPERTIES_JSON__:" \+ JSON\.stringify\(oCustomPropsResult\)/.test(statsWithProps), statsWithProps)

  check('documentStats: GetDocumentInfo SOHA nem hivodik (megtagadva, mert throw-ol es kilovi a jobot)',
    !/GetDocumentInfo/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'documentStats', includeCustomProperties: true }] }).script))

  dob('documentStats: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'documentStats' }] }), /not available in the pptx core/)
}

console.log('\n[57] docx-K14: LEZARO/VEGLEGESITO IRO MUVELETEK -- abrajegyzek+caption, cella-szetvalasztas, smartReplace, valtozaskoveto-lezaras (accept/reject), konyvjelzo-torles. UpdateAllTOC NEVESITETT MEGTAGADAS (nincs kod, lasd korabbi kommentek), Table.Split egy korabbi teszt "NULL, hasznalhatatlan" lelete FELULIRVA (rossz argumentum-sorrend volt, nem toresett API)')
{
  // table.caption -- Table.AddCaption(sAdditional, sLabel)
  const cap = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a']], caption: 'My caption text' }],
  }).script
  check('table.caption: string alak -> AddCaption(text, "Table") alap cimkevel',
    /oTable\.AddCaption\("My caption text", "Table"\);/.test(cap), cap)
  const capLabel = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a']], caption: { text: 'Fig text', label: 'Figure' } }],
  }).script
  check('table.caption: objektum alak, caller-megadott label -> AddCaption(text, "Figure")',
    /oTable\.AddCaption\("Fig text", "Figure"\);/.test(capLabel), capLabel)
  check('table.caption: hianyzik -> NINCS AddCaption hivas',
    !/AddCaption/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a']] }] }).script))
  dob('table.caption: objektum alak `text` nelkul -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'table', rows: [['a']], caption: { label: 'Figure' } }],
  }), /`caption\.text` is required/)

  // table.split -- TableCell.Split(nRow, nCol), [row, col, nRow, nCol] tuples
  const split = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'table', rows: [['a', 'b'], ['c', 'd']], split: [[0, 0, 2, 2]] }],
  }).script
  check('table.split: [0,0,2,2] -> oTable.GetCell(0, 0).Split(2, 2)',
    /oTable\.GetCell\(0, 0\)\.Split\(2, 2\);/.test(split), split)
  check('table.split: a Split hivas a MERGE es a SOR-FINOMITASOK UTAN all a scriptben (sorrend-fuggoseg, GetRow(i) meg ervenyes indexekkel fusson)',
    (() => {
      const s = buildCreateScript({
        core: 'docx',
        operations: [{ type: 'table', rows: [['a', 'b'], ['c', 'd']], repeatHeaderRow: true, split: [[1, 1, 2, 1]] }],
      }).script
      return s.indexOf('SetTableHeader(true)') < s.indexOf('.Split(2, 1)')
    })())
  dob('table.split: rossz alaku tuple -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'table', rows: [['a']], split: [[0, 0]] }],
  }), /must be a \[row, col, nRow, nCol\] tuple/)
  dob('table.split: tartomanyon kivuli cella -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'table', rows: [['a']], split: [[5, 0, 2, 2]] }],
  }), /is outside the table/)
  dob('table.split: nRow/nCol 0 vagy negativ -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'table', rows: [['a']], split: [[0, 0, 0, 1]] }],
  }), /nRow\/nCol must both be >= 1/)
  check('table.split: hianyzik -> NINCS Split hivas',
    !/\.Split\(/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'table', rows: [['a']] }] }).script))

  // tableOfFigures -- oDocument.AddTableOfFigures({BuildFrom}, true)
  const tof = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'tableOfFigures' }],
  }).script
  check('tableOfFigures: alapertekekkel -> MoveCursorToStart + AddTableOfFigures({BuildFrom: "Figure"}, true)',
    /oDocument\.MoveCursorToStart\(\);/.test(tof) &&
    /oDocument\.AddTableOfFigures\(\{BuildFrom: "Figure"\}, true\);/.test(tof), tof)
  check('tableOfFigures: position end -> MoveCursorToEnd',
    /oDocument\.MoveCursorToEnd\(\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'tableOfFigures', position: 'end' }] }).script))
  check('tableOfFigures: buildFrom "Table" -> BuildFrom LITERALKENT athelyettesitve',
    /BuildFrom: "Table"/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'tableOfFigures', buildFrom: 'Table' }] }).script))
  dob('tableOfFigures: ismeretlen position -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'tableOfFigures', position: 'middle' }],
  }), /unknown position.*known: start, end/)
  dob('tableOfFigures: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'tableOfFigures' }] }), /not available in the pptx core/)

  // smartReplace -- oParagraph.Select() + Api.ReplaceTextSmart([text])
  const sm = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'text', text: 'original' }, { type: 'smartReplace', text: 'replaced text' }],
  }).script
  check('smartReplace: text UTAN -> oParagraph.Select() + Api.ReplaceTextSmart(["replaced text"])',
    /oParagraph\.Select\(\);/.test(sm) && /Api\.ReplaceTextSmart\(\["replaced text"\]\);/.test(sm), sm)
  check('smartReplace: a Select+ReplaceTextSmart a MEGELOZO text AddText UTAN all a scriptben',
    (() => { const i1 = sm.indexOf('AddText("original")'); const i2 = sm.indexOf('ReplaceTextSmart'); return i1 !== -1 && i2 > i1 })())
  dob('smartReplace: `text` hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'smartReplace' }],
  }), /`text` is required/)
  check('smartReplace: ures string `text` (torles) MEGENGEDETT, nem hianyzo mezo',
    /ReplaceTextSmart\(\[""\]\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x' }, { type: 'smartReplace', text: '' }] }).script))
  dob('smartReplace: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'smartReplace', text: 'x' }] }), /not available in the pptx core/)

  // resolveRevisions -- Accept/RejectAllRevisionChanges
  check('resolveRevisions: action accept -> oDocument.AcceptAllRevisionChanges()',
    /oDocument\.AcceptAllRevisionChanges\(\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'resolveRevisions', action: 'accept' }] }).script))
  check('resolveRevisions: action reject -> oDocument.RejectAllRevisionChanges()',
    /oDocument\.RejectAllRevisionChanges\(\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'resolveRevisions', action: 'reject' }] }).script))
  dob('resolveRevisions: ismeretlen action -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'resolveRevisions', action: 'x' }],
  }), /`action` must be "accept" or "reject"/)
  dob('resolveRevisions: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'resolveRevisions', action: 'accept' }] }), /not available in the pptx core/)

  // deleteBookmark -- oDocument.DeleteBookmark(name)
  check('deleteBookmark: name -> oDocument.DeleteBookmark("myMarker") LITERALKENT',
    /oDocument\.DeleteBookmark\("myMarker"\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'deleteBookmark', name: 'myMarker' }] }).script))
  dob('deleteBookmark: `name` hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'deleteBookmark' }],
  }), /`name` is required/)
  dob('deleteBookmark: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'deleteBookmark', name: 'x' }] }), /not available in the pptx core/)

  // UpdateAllTOC -- named refusal, no operation exists to bind to it
  check('UpdateAllTOC: nincs "updateAllToc"/"updateAllTOC" muvelet-tipus bekotve (NEVESITETT MEGTAGADAS, lasd koteg04/koteg14 dontesi komment)',
    !('updateAllToc' in OPERATIONS) && !('updateAllTOC' in OPERATIONS))
}

console.log('\n[58] wordArt (D2) -- a mert POZICIONALIS hivas-alak (nem objektum-argumentum)')
{
  const wa = buildCreateScript({ core: 'docx', operations: [{ type: 'wordArt', text: 'HELLO' }] }).script
  check('wordArt: alap -> Api.CreateWordArt(null, text, null, NoFill, ..., 0, alap w/h) POZICIONALISAN',
    /Api\.CreateWordArt\(null, "HELLO", null, Api\.CreateNoFill\(\), Api\.CreateStroke\(0, Api\.CreateNoFill\(\)\), 0, 3000000, 1000000\);/.test(wa), wa)
  check('wordArt: AddDrawing+Push -- NEM kozvetlen Push (a mert csapda: kozvetlen Push csendben eldobja)',
    /oWordArtPar\.AddDrawing\(oWordArt\);\noDocument\.Push\(oWordArtPar\);/.test(wa))
  const waColor = buildCreateScript({ core: 'docx', operations: [{ type: 'wordArt', text: 'X', color: [200, 0, 0], width: 5000000, height: 2000000, rotation: 45, transform: 'textArchUp' }] }).script
  check('wordArt: color/width/height/rotation/transform mind LITERALKENT athelyettesitve',
    /Api\.CreateWordArt\(null, "X", "textArchUp", Api\.CreateSolidFill\(Api\.CreateRGBColor\(200, 0, 0\)\), Api\.CreateStroke\(0, Api\.CreateNoFill\(\)\), 45, 5000000, 2000000\);/.test(waColor), waColor)
  dob('wordArt: `text` nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'wordArt' }] }), /`text` is required/)
  // (K8): wordArt is now ALSO bound on pptx (own emit branch,
  // different call shape) -- see section [63] below. Only the docx branch's own fields
  // (`color`, the AddDrawing+Push attach path) are asserted here.
  check('wordArt: pptx-en IS elerheto MOST MAR, sajat aggal (K8) -- nem "not available"',
    (() => { try { buildCreateScript({ core: 'pptx', operations: [{ type: 'wordArt', text: 'x' }], slideCount: 1 }); return true } catch { return false } })())
}

console.log('\n[59] oleObject (D2) -- a mert POZICIONALIS hivas-alak -- ES FUGGETLEN az xlsx sheetDrawing.ole megtagadasatol (mas objektum, mas motor-ag)')
{
  const ole = buildCreateScript({ core: 'docx', operations: [{ type: 'oleObject', imageSrc: 'aW1n', width: 1000000, height: 500000, data: 'PAYLOAD', appId: 'Word.Document' }] }).script
  check('oleObject: Api.CreateOleObject(imageSrc, width, height, data, appId) POZICIONALISAN',
    /Api\.CreateOleObject\("aW1n", 1000000, 500000, "PAYLOAD", "Word\.Document"\);/.test(ole), ole)
  check('oleObject: AddDrawing+Push mintaja', /oOlePar\.AddDrawing\(oOle\);\noDocument\.Push\(oOlePar\);/.test(ole))
  dob('oleObject: `imageSrc` nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'oleObject', data: 'x', appId: 'y' }] }), /`imageSrc`.*required/)
  dob('oleObject: `data` nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'oleObject', imageSrc: 'x', appId: 'y' }] }), /`data`.*required/)
  dob('oleObject: `appId` nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'oleObject', imageSrc: 'x', data: 'y' }] }), /`appId`.*required/)
  dob('oleObject: pptx magban nem all rendelkezesre (docx-only)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'oleObject', imageSrc: 'x', data: 'y', appId: 'z' }], slideCount: 1 }), /not available in the pptx core/)
}

console.log('\n[60] drawingGroup (D2) -- a mert hivas-alak (oDocument.GroupDrawings, NEM Api.CreateGroup -- az MINDIG dob)')
{
  const grp = buildCreateScript({ core: 'docx', operations: [{ type: 'drawingGroup', shapes: [{ shapeType: 'rect', width: 500000, height: 500000, fill: [255, 0, 0] }, { shapeType: 'ellipse', width: 500000, height: 500000 }] }] }).script
  check('drawingGroup: ket alakzat sajat parbekezdessel pusholva, MAJD oDocument.GroupDrawings a ket valtozoval',
    /var oDrawGroupShape0 = Api\.CreateShape\("rect", 500000, 500000, Api\.CreateSolidFill\(Api\.CreateRGBColor\(255, 0, 0\)\), Api\.CreateStroke\(0, Api\.CreateNoFill\(\)\)\);/.test(grp)
    && /oDrawGroupShape0Par\.AddDrawing\(oDrawGroupShape0\);\noDocument\.Push\(oDrawGroupShape0Par\);/.test(grp)
    && /var oDrawGroupShape1 = Api\.CreateShape\("ellipse", 500000, 500000, Api\.CreateNoFill\(\), Api\.CreateStroke\(0, Api\.CreateNoFill\(\)\)\);/.test(grp)
    && /oDocument\.GroupDrawings\(\[oDrawGroupShape0, oDrawGroupShape1\]\);/.test(grp),
    grp)
  check('drawingGroup: NEM Api.CreateGroup-ot hivja (az mert-en mindig dob "All drawings must be in document")',
    !/Api\.CreateGroup\(/.test(grp))
  dob('drawingGroup: egyetlen alakzattal -> megtagadva (kevesebb mint 2 nem csoport)', () => buildCreateScript({ core: 'docx', operations: [{ type: 'drawingGroup', shapes: [{ shapeType: 'rect' }] }] }), /at least 2 shape descriptors/)
  dob('drawingGroup: pptx magban nem all rendelkezesre (docx-only -- a pptx sajat utja a mar landolt slide.group)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'drawingGroup', shapes: [{}, {}] }], slideCount: 1 }), /not available in the pptx core/)
}

console.log('\n[61] shape.fillGradient / fillPattern / geometry (D2) -- DOCX-ONLY kiterjesztes a mar landolt `shape` op-on')
{
  const grad = buildCreateScript({ core: 'docx', operations: [{ type: 'shape', fillGradient: { stops: [{ color: [255, 0, 0], pos: 0 }, { color: [0, 0, 255], pos: 100000 }], angle: 5400000 } }] }).script
  check('shape.fillGradient: CreateGradientStop x2 + CreateLinearGradientFill(stops, angle) POZICIONALISAN',
    /Api\.CreateLinearGradientFill\(\[Api\.CreateGradientStop\(Api\.CreateRGBColor\(255, 0, 0\), 0\), Api\.CreateGradientStop\(Api\.CreateRGBColor\(0, 0, 255\), 100000\)\], 5400000\)/.test(grad), grad)
  dob('shape.fillGradient: pptx-en nem all rendelkezesre (docx-only)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fillGradient: { stops: [{ color: [0, 0, 0], pos: 0 }, { color: [1, 1, 1], pos: 1 }] } }], slideCount: 1 }), /fillGradient\/fillPattern are docx-only/)
  dob('shape.fillGradient: 1 stop -> megtagadva (legalabb 2 kell)', () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', fillGradient: { stops: [{ color: [0, 0, 0], pos: 0 }] } }] }), /at least 2/)

  const pat = buildCreateScript({ core: 'docx', operations: [{ type: 'shape', fillPattern: { patternType: 'pct25', bgColor: [255, 0, 0], fgColor: [255, 255, 255] } }] }).script
  check('shape.fillPattern: CreatePatternFill(patternType, bgColor, fgColor) POZICIONALISAN',
    /Api\.CreatePatternFill\("pct25", Api\.CreateRGBColor\(255, 0, 0\), Api\.CreateRGBColor\(255, 255, 255\)\)/.test(pat), pat)
  dob('shape.fillPattern: patternType nelkul -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', fillPattern: { bgColor: [0, 0, 0], fgColor: [1, 1, 1] } }] }), /patternType is required/)

  dob('shape: fill + fillGradient egyutt -> megtagadva (csak egy adhato)', () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', fill: [0, 0, 0], fillGradient: { stops: [{ color: [0, 0, 0], pos: 0 }, { color: [1, 1, 1], pos: 1 }] } }] }), /only one of fill\/fillGradient\/fillPattern/)

  const geom = buildCreateScript({ core: 'docx', operations: [{ type: 'shape', geometry: { path: [{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 1000000, y: 0 }, { cmd: 'lineTo', x: 500000, y: 1000000 }, { cmd: 'close' }] } }] }).script
  check('shape.geometry: CreateCustomGeometry + AddPath + MoveTo/LineTo/Close + SetGeometry, SORRENDBEN',
    /var oShapeGeom = Api\.CreateCustomGeometry\(\);\nvar oShapePath = oShapeGeom\.AddPath\(\);\noShapePath\.MoveTo\(0, 0\);\noShapePath\.LineTo\(1000000, 0\);\noShapePath\.LineTo\(500000, 1000000\);\noShapePath\.Close\(\);\noShape\.SetGeometry\(oShapeGeom\);/.test(geom), geom)
  dob('shape.geometry: ures path -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', geometry: { path: [] } }] }), /non-empty array/)
  dob('shape.geometry: ismeretlen cmd -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', geometry: { path: [{ cmd: 'curveTo', x: 1, y: 1 }] } }] }), /unknown.*known: moveTo, lineTo, close/)
}

console.log('\n[65] PPTX-K9: shape.geometry KITERJESZTVE pptx-re + UJ shape.placeholder')
{
  // shape.geometry (D2) MAR landolt docx-en; K9 UGYANAZT a
  // hivas-alakot package-verified UJRAMERTE pptx-en is (<a:custGeom> package-verified) --
  // a codegen SORRENDJE azonos, csak a core valtozott.
  const geomPptx = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', geometry: { path: [{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 1000000, y: 0 }, { cmd: 'lineTo', x: 500000, y: 1000000 }, { cmd: 'close' }] } }], slideCount: 1 }).script
  check('shape.geometry pptx-en: CreateCustomGeometry + AddPath + MoveTo/LineTo/Close + SetGeometry, SORRENDBEN, SetPosition ELOTT',
    /var oShapeGeom = Api\.CreateCustomGeometry\(\);\nvar oShapePath = oShapeGeom\.AddPath\(\);\noShapePath\.MoveTo\(0, 0\);\noShapePath\.LineTo\(1000000, 0\);\noShapePath\.LineTo\(500000, 1000000\);\noShapePath\.Close\(\);\noShape\.SetGeometry\(oShapeGeom\);\noShape\.SetPosition/.test(geomPptx), geomPptx)
  dob('shape.geometry: xlsx-en nem all rendelkezesre (a `shape` tipus maga sincs bekotve xlsx-en)', () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'shape', geometry: { path: [{ cmd: 'close' }] } }] }), /type "shape" is not available in the xlsx core/)

  // shape.placeholder -- UJ K9-mezo, pptx-only. Api.CreatePlaceholder(type) -> SetPlaceholder.
  const ph = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', placeholder: 'title' }], slideCount: 1 }).script
  check('shape.placeholder: oShape.SetPlaceholder(Api.CreatePlaceholder("title"))',
    /oShape\.SetPlaceholder\(Api\.CreatePlaceholder\("title"\)\);/.test(ph), ph)
  for (const t of ['title', 'body', 'ctrTitle', 'subTitle', 'chart', 'clipArt', 'media']) {
    const s = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', placeholder: t }], slideCount: 1 }).script
    check(`shape.placeholder: ismert tipus ${t} elfogadva`, new RegExp(`Api\\.CreatePlaceholder\\("${t}"\\)`).test(s), s)
  }
  dob('shape.placeholder: docx-en nem all rendelkezesre (nem probalt, es placeholder fogalmilag dia-layout-kotott)',
    () => buildCreateScript({ core: 'docx', operations: [{ type: 'shape', placeholder: 'title' }] }),
    /placeholder is pptx-only/)
  dob('shape.placeholder: ismeretlen tipus -> NEVESITETT megtagadas (a motor csendben "body"-ra esne vissza, nem dobna)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', placeholder: 'sldNum' }], slideCount: 1 }),
    /unknown placeholder type "sldNum".*known: title, body, ctrTitle, subTitle, chart, clipArt, media/)
  dob('shape.placeholder: OOXML-ben letezo, de a motor altal csendben body-ra vetitett tipus is megtagadva (pic)',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', placeholder: 'pic' }], slideCount: 1 }),
    /unknown placeholder type "pic"/)
}

console.log('\n[62] docx-K15: KURZOR-POZICIO ALAPU CELZAS (enterText/insertContent/replaceCurrent/insertBlankPage/clearAllFields, goToPage NEVESITETT MEGTAGADAS)')
{
  // enterText -- position kotelezo, MoveCursorToStart/End + EnterText
  const enterEnd = buildCreateScript({ core: 'docx', operations: [{ type: 'enterText', text: 'hello', position: 'end' }] }).script
  check('enterText: position end -> MoveCursorToEnd + EnterText',
    /oDocument\.MoveCursorToEnd\(\);/.test(enterEnd) && /oDocument\.EnterText\("hello"\);/.test(enterEnd), enterEnd)
  const enterStart = buildCreateScript({ core: 'docx', operations: [{ type: 'enterText', text: 'hi', position: 'start' }] }).script
  check('enterText: position start -> MoveCursorToStart', /oDocument\.MoveCursorToStart\(\);/.test(enterStart))
  dob('enterText: `text` hianyzik -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'enterText', position: 'end' }] }), /`text` is required/)
  dob('enterText: `position` hianyzik -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'enterText', text: 'x' }] }), /`position` must be "start" or "end"/)
  dob('enterText: ismeretlen position -> megtagadva', () => buildCreateScript({ core: 'docx', operations: [{ type: 'enterText', text: 'x', position: 'middle' }] }), /`position` must be "start" or "end"/)
  check('enterText: ures string `text` (torles) MEGENGEDETT',
    /EnterText\(""\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'enterText', text: '', position: 'end' }] }).script))
  dob('enterText: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'enterText', text: 'x', position: 'end' }] }), /not available in the pptx core/)

  // insertContent -- tobb bekezdes, sorrendben
  const insertMulti = buildCreateScript({
    core: 'docx',
    operations: [{ type: 'insertContent', paragraphs: ['first', 'second'], position: 'end' }],
  }).script
  check('insertContent: ket bekezdes -> ket Api.CreateParagraph, sorrendben az InsertContent tombben',
    /oInsertContentPara0\.AddText\("first"\);/.test(insertMulti) &&
    /oInsertContentPara1\.AddText\("second"\);/.test(insertMulti) &&
    /oDocument\.InsertContent\(\[oInsertContentPara0, oInsertContentPara1\], false\);/.test(insertMulti), insertMulti)
  check('insertContent: MoveCursorToEnd elobb fut, mint az InsertContent',
    (() => { const i1 = insertMulti.indexOf('MoveCursorToEnd'); const i2 = insertMulti.indexOf('InsertContent(['); return i1 !== -1 && i2 > i1 })())
  dob('insertContent: `paragraphs` ures tomb -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'insertContent', paragraphs: [], position: 'end' }],
  }), /`paragraphs` must be a non-empty array/)
  dob('insertContent: `paragraphs` hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'insertContent', position: 'end' }],
  }), /`paragraphs` must be a non-empty array/)
  dob('insertContent: `position` hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'insertContent', paragraphs: ['x'] }],
  }), /`position` must be "start" or "end"/)

  // replaceCurrent -- scope word/sentence, part opcionalis
  const replWord = buildCreateScript({
    core: 'docx', operations: [{ type: 'replaceCurrent', scope: 'word', replace: 'NEW', position: 'start' }],
  }).script
  check('replaceCurrent: scope word -> ReplaceCurrentWord("NEW")',
    /oDocument\.MoveCursorToStart\(\);/.test(replWord) && /oDocument\.ReplaceCurrentWord\("NEW"\);/.test(replWord), replWord)
  const replSentence = buildCreateScript({
    core: 'docx', operations: [{ type: 'replaceCurrent', scope: 'sentence', replace: 'NEW', position: 'end', part: 'after' }],
  }).script
  check('replaceCurrent: scope sentence + part after -> ReplaceCurrentSentence("NEW", "after")',
    /oDocument\.ReplaceCurrentSentence\("NEW", "after"\);/.test(replSentence), replSentence)
  dob('replaceCurrent: ismeretlen scope -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'replaceCurrent', scope: 'paragraph', replace: 'x', position: 'end' }],
  }), /`scope` must be "word" or "sentence"/)
  dob('replaceCurrent: `replace` hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'replaceCurrent', scope: 'word', position: 'end' }],
  }), /`replace` is required/)
  dob('replaceCurrent: `position` hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'replaceCurrent', scope: 'word', replace: 'x' }],
  }), /`position` must be "start" or "end"/)
  dob('replaceCurrent: ismeretlen `part` -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'replaceCurrent', scope: 'word', replace: 'x', position: 'end', part: 'middle' }],
  }), /`part` must be "before" or "after"/)

  // insertBlankPage
  const blankEnd = buildCreateScript({ core: 'docx', operations: [{ type: 'insertBlankPage', position: 'end' }] }).script
  check('insertBlankPage: position end -> MoveCursorToEnd + InsertBlankPage',
    /oDocument\.MoveCursorToEnd\(\);/.test(blankEnd) && /oDocument\.InsertBlankPage\(\);/.test(blankEnd), blankEnd)
  dob('insertBlankPage: `position` hianyzik -> megtagadva', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'insertBlankPage' }],
  }), /`position` must be "start" or "end"/)

  // clearAllFields -- nulla-argumentumu, nincs position
  check('clearAllFields: oDocument.ClearAllFields() LITERALKENT',
    /oDocument\.ClearAllFields\(\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'clearAllFields' }] }).script))
  dob('clearAllFields: pptx-en nem elerheto', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'clearAllFields' }] }), /not available in the pptx core/)

  // goToPage -- MINDIG megtagadva
  dob('goToPage: MINDIG megtagadva (MEASURED no-op)', () => buildCreateScript({
    core: 'docx', operations: [{ type: 'goToPage', index: 0 }],
  }), /GoToPage\(\).*no-op on this route/)
}

console.log('\n[63] UJ MUVELET-TIPUS: wordArt (Api.CreateWordArt), pptx-only')
console.log('    hivas-alak toString()-bol visszafejtve; csomag-szintu igazolas elo DS-hivassal, nem itt')
{
  const alap = buildCreateScript({ core: 'pptx', operations: [{ type: 'wordArt', text: 'Cim' }], slideCount: 1 })
  check('wordArt: alap hivas -- Api.CreateWordArt(null, szoveg, "textNoShape", NoFill, Stroke(0,NoFill), 0, 1828800, 1828800), pozicio NELKUL (auto-kozepre)',
    /var oArt = Api\.CreateWordArt\(null, "Cim", "textNoShape", Api\.CreateNoFill\(\), Api\.CreateStroke\(0, Api\.CreateNoFill\(\)\), 0, 1828800, 1828800\);/.test(alap.script), alap.script)
  check('wordArt: AddObject a letrehozott alakzatra', /oSlide\.AddObject\(oArt\);/.test(alap.script))

  dob('wordArt: `text` hianyzik -> megtagadva', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'wordArt' }], slideCount: 1 }), /`text` is required/)

  const teljes = buildCreateScript({ core: 'pptx', operations: [{ type: 'wordArt', text: 'X', transform: 'textArchUp', fill: [255, 0, 0], lineColor: [0, 0, 0], lineWidth: 10000, rotation: 45, width: 2000000, height: 1000000, x: 300000, y: 400000 }], slideCount: 1 })
  check('wordArt: minden mezo a sajat erteket adja, pozicioval EGYUTT',
    /Api\.CreateWordArt\(null, "X", "textArchUp", Api\.CreateSolidFill\(Api\.CreateRGBColor\(255, 0, 0\)\), Api\.CreateStroke\(10000, Api\.CreateSolidFill\(Api\.CreateRGBColor\(0, 0, 0\)\)\), 45, 2000000, 1000000, 300000, 400000\);/.test(teljes.script), teljes.script)

  dob('wordArt: ismeretlen transform -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'wordArt', text: 'x', transform: 'textWhoosh' }], slideCount: 1 }), /unknown transform/)
  dob('wordArt: csak `x` MEGADVA, `y` nelkul -> NEVESITETT hiba (auto-kozepre csak MINDKETTO hianyaban mukodik)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'wordArt', text: 'x', x: 100 }], slideCount: 1 }), /`x` and `y` must both be given or both omitted/)
  // (D2): wordArt is ALSO bound on docx (own emit branch, D2's
  // positional/`color`/AddDrawing+Push call shape) -- see section [58] above. Not "not
  // available" on docx anymore, so no refusal check here.
  check('wordArt: docx-en IS elerheto (D2) -- nem "not available"',
    (() => { try { buildCreateScript({ core: 'docx', operations: [{ type: 'wordArt', text: 'x' }] }); return true } catch { return false } })())
}

console.log('\n[64] shape.fill BOVITVE: gradiens/mintas kitoltes (buildFillExpression), a regi [r,g,b] alak VALTOZATLAN')
{
  const legacy = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: [1, 2, 3] }], slideCount: 1 })
  check('shape.fill: a REGI bare [r,g,b] alak valtozatlanul solid fill-t ad', /Api\.CreateShape\("rect", 2000000, 2000000, Api\.CreateSolidFill\(Api\.CreateRGBColor\(1, 2, 3\)\)/.test(legacy.script))

  const explicitSolid = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'solid', color: [9, 9, 9] } }], slideCount: 1 })
  check('shape.fill: {type:"solid",color} ugyanazt adja, mint a bare tomb', /Api\.CreateSolidFill\(Api\.CreateRGBColor\(9, 9, 9\)\)/.test(explicitSolid.script))

  const linear = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'gradient', shape: 'linear', angle: 90, stops: [{ color: [255, 0, 0], pos: 0 }, { color: [0, 0, 255], pos: 100 }] } }], slideCount: 1 })
  check('shape.fill: linearis gradiens -- CreateLinearGradientFill([stopok], szog*60000), a pos 0-100 -> 0-100000',
    /Api\.CreateLinearGradientFill\(\[Api\.CreateGradientStop\(Api\.CreateRGBColor\(255, 0, 0\), 0\), Api\.CreateGradientStop\(Api\.CreateRGBColor\(0, 0, 255\), 100000\)\], 5400000\)/.test(linear.script), linear.script)

  const radial = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'gradient', shape: 'radial', stops: [{ color: [1, 1, 1], pos: 0 }, { color: [2, 2, 2], pos: 50 }] } }], slideCount: 1 })
  check('shape.fill: sugaras gradiens -- CreateRadialGradientFill([stopok]), szog nelkul', /Api\.CreateRadialGradientFill\(\[Api\.CreateGradientStop/.test(radial.script))

  const pattern = buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'pattern', patternType: 'pct50', bgColor: [255, 255, 255], fgColor: [0, 128, 0] } }], slideCount: 1 })
  check('shape.fill: mintas kitoltes -- CreatePatternFill(preset, bg, fg)',
    /Api\.CreatePatternFill\("pct50", Api\.CreateRGBColor\(255, 255, 255\), Api\.CreateRGBColor\(0, 128, 0\)\)/.test(pattern.script))

  dob('shape.fill: ismeretlen patternType -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'pattern', patternType: 'nemLetezoMinta', bgColor: [0, 0, 0], fgColor: [0, 0, 0] } }], slideCount: 1 }), /unknown fill\.patternType/)
  dob('shape.fill: gradiens 1 stop-pal -> NEVESITETT hiba (legalabb 2 kell)', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'gradient', stops: [{ color: [0, 0, 0], pos: 0 }] } }], slideCount: 1 }), /at least 2/)
  dob('shape.fill: gradiens stop pos tartomanyon kivul -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'gradient', stops: [{ color: [0, 0, 0], pos: -1 }, { color: [1, 1, 1], pos: 50 }] } }], slideCount: 1 }), /pos must be 0-100/)
  dob('shape.fill: ismeretlen fill.type -> NEVESITETT hiba', () => buildCreateScript({ core: 'pptx', operations: [{ type: 'shape', fill: { type: 'whoosh' } }], slideCount: 1 }), /fill\.type must be/)
}

console.log('\n[65] XLSX LAP-INDEX: mindket irany, ugyanaz a minta mint a pptx E6 dia-indexe')
{
  const ketLapon = buildCreateScript({
    core: 'xlsx',
    operations: [{ type: 'formula', at: 'A1', formula: '=1+1', sheet: 0 }, { type: 'formula', at: 'A1', formula: '=2+2', sheet: 1 }],
    sheetCount: 2,
  })
  check('ket lap, ket kulon Api.GetSheet hivas, a MEGADOTT sorrendben',
    ketLapon.script.indexOf('Api.GetSheet(0)') < ketLapon.script.indexOf('Api.GetSheet(1)'), ketLapon.script)

  dob('lap-szamot MEGHALADO index -> NEVESITETT hiba, nem nema no-op (visszaadja az E2 lelet targyat)',
    () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1', sheet: 5 }], sheetCount: 1 }),
    /sheet index 5 is out of range.*1 sheet\(s\)/)
  dob('negativ lap-index -> szinten megtagadva',
    () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1', sheet: -1 }], sheetCount: 1 }),
    /out of range/)

  // a lapNEV (amit a
  // co-editing ut elfogad) ezen az uton NaN-ra szamitodik, es korabban ez is "out of range"-kent
  // volt jelentve -- helyes symptoma, felrevezeto ok. Az uj hiba a VALODI okot nevezi meg.
  dob('lap-NEV (nem index) -> NEVESITETT hiba a VALODI okrol, nem "out of range" (a hiba korabban ezt jelentette)',
    () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1', sheet: 'Osszefoglalo' }], sheetCount: 1 }),
    /sheet "Osszefoglalo" is not a non-negative integer.*cannot resolve a sheet name/)
  check('lap-NEV hibaja NEM tartalmazza az "out of range" szoveget (a ket ok mostantol megkulonboztetve)',
    (() => {
      try {
        buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1', sheet: 'Osszefoglalo' }], sheetCount: 1 })
        return false
      } catch (err) { return !/out of range/.test(err.message) }
    })())
  dob('lap-NEV sheetCount NELKUL is a nevesitett hibat adja, nem az altalanos "sheet count is not known" szoveget',
    () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1', sheet: 'Osszefoglalo' }] }),
    /sheet "Osszefoglalo" is not a non-negative integer/)
  dob('POZ. KONTROLL: hatarserto INDEX-szel (nem nev) tovabbra is a range-hibat adja, valtozatlanul',
    () => buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1', sheet: 5 }], sheetCount: 1 }),
    /sheet index 5 is out of range.*1 sheet\(s\)/)

  check('sheetCount nelkul (a co-editing/regi hivok esete) a lap-index tovabbra is ellenorizetlen, GetSheet(0)-ra all',
    /Api\.GetSheet\(0\)/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1' }] }).script))

  check('sheet mezo NELKUL (a regi, egy-lapos hivok esete) -- GetSheet(0), NEM GetActiveSheet -- de UGYANARRA a lapra mutat egy 1-lapos dokumentumon, tehat viselkedes-azonos',
    /Api\.GetSheet\(0\)/.test(buildCreateScript({ core: 'xlsx', operations: [{ type: 'formula', at: 'A1', formula: '=1' }], sheetCount: 1 }).script))

  const harmasIsIsmetelve = buildCreateScript({
    core: 'xlsx',
    operations: [{ type: 'formula', at: 'A1', formula: '=1', sheet: 1 }, { type: 'formula', at: 'A1', formula: '=2' }],
    sheetCount: 2,
  })
  check('az op.sheet NEM SZIVAROG at a kovetkezo, sheet-mezo NELKULI operacioba -- a masodik op ujra GetSheet(0)-ra all (a var-ujra-kotes vedelme)',
    harmasIsIsmetelve.script.indexOf('Api.GetSheet(1)') < harmasIsIsmetelve.script.lastIndexOf('Api.GetSheet(0)'),
    harmasIsIsmetelve.script)
}

console.log('\n[66] pptx `text`: `paragraphs` tomb TOBB bekezdest enged EGY szovegdobozban')
{
  const legacy = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'dia', size: 24, bold: true }], slideCount: 1 })
  check('`paragraphs` NELKUL a script pontosan a regi alak (nincs CreateParagraph, nincs Push)',
    !/Api\.CreateParagraph/.test(legacy.script) && !/oContent\.Push/.test(legacy.script) && /oContent\.GetElement\(0\)/.test(legacy.script), legacy.script)

  const multi = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', paragraphs: [
    { text: 'elso', bold: true },
    { text: 'masodik', size: 18 },
    { text: 'harmadik' },
  ] }], slideCount: 1 })
  check('harom bekezdes -- EGY oShape/oContent, harom AddText, ket CreateParagraph (az elso GetElement(0)-t hasznalja)',
    (multi.script.match(/Api\.CreateShape/g) || []).length === 1 &&
    (multi.script.match(/Api\.CreateParagraph\(\)/g) || []).length === 2 &&
    (multi.script.match(/oContent\.Push\(oPara\)/g) || []).length === 2 &&
    /oContent\.GetElement\(0\)/.test(multi.script), multi.script)
  check('  a harom szoveg sorrendben, a sajat formazasukkal', /AddText\("elso"\)/.test(multi.script) && /AddText\("masodik"\)/.test(multi.script) && /AddText\("harmadik"\)/.test(multi.script))
  check('  az elso bekezdes bold, a masodik size=18, a harmadik formazatlan',
    multi.script.indexOf('AddText("elso")') < multi.script.indexOf('SetBold(true)') &&
    multi.script.indexOf('SetBold(true)') < multi.script.indexOf('AddText("masodik")') &&
    /AddText\("masodik"\);\n\s*oRun\.SetFontSize\(18\);/.test(multi.script) &&
    !/AddText\("harmadik"\);\n\s*oRun\.Set/.test(multi.script), multi.script)
  check('  csak EGY oSlide.AddObject(oShape) -- egyetlen szovegdoboz, nem harom kulon', (multi.script.match(/oSlide\.AddObject\(oShape\)/g) || []).length === 1)

  const emptyArr = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'visszaesik', paragraphs: [] }], slideCount: 1 })
  check('`paragraphs: []` (ures tomb) -- visszaesik az op.text-es egy-bekezdeses alapviselkedesre, NEM hibazik es NEM ad 0 bekezdest',
    /AddText\("visszaesik"\)/.test(emptyArr.script) && !/Api\.CreateParagraph/.test(emptyArr.script), emptyArr.script)

  check('docx-n a `paragraphs` mezo jelenlete semmilyen uj hivast nem valt ki (pptx-specifikus)',
    !/Api\.CreateParagraph\(\);\n.*oContent/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', paragraphs: [{ text: 'y' }] }] }).script))
}

console.log('\n[67] pptx `text`/`paragraphs`: `indentLeft` (TWIP) -> oPara.SetIndLeft, csomag-szinten package-verified')
{
  const noIndent = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'sima' }], slideCount: 1 })
  check('`indentLeft` NELKUL nincs SetIndLeft hivas a scriptben', !/SetIndLeft/.test(noIndent.script), noIndent.script)

  const single = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'behuzott', indentLeft: 720 }], slideCount: 1 })
  check('egy-bekezdeses `text` op indentLeft-tel -- SetIndLeft(720) a RAW erteket kapja (nincs elore-szorzas: package-mert, hogy a DocBuilder maga konvertal TWIP->EMU)',
    /oPara\.SetIndLeft\(720\);/.test(single.script), single.script)

  const multi = buildCreateScript({ core: 'pptx', operations: [{ type: 'text', paragraphs: [
    { text: 'level0' },
    { text: 'level1', indentLeft: 720 },
  ] }] , slideCount: 1 })
  check('`paragraphs` tombon csak a behuzott bekezdes kap SetIndLeft-et, a masik nem', (() => {
    const levelSplit = multi.script.split('AddText("level1")')
    return (multi.script.match(/SetIndLeft/g) || []).length === 1 &&
      !/SetIndLeft/.test(levelSplit[0]) && /SetIndLeft\(720\)/.test(levelSplit[1])
  })(), multi.script)

  dob('negativ indentLeft -> NEVESITETT hiba, nem csendes no-op',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', indentLeft: -5 }], slideCount: 1 }),
    /indentLeft.*non-negative integer/)
  dob('nem-egesz indentLeft -> NEVESITETT hiba',
    () => buildCreateScript({ core: 'pptx', operations: [{ type: 'text', text: 'x', indentLeft: 1.5 }], slideCount: 1 }),
    /indentLeft.*non-negative integer/)

  check('docx-n az op-szintu `indentLeft` a MEGLEVO applyParagraphIndentSpacing utjan megy, nem a pptx uton (SetIndLeft ott is megjelenik, de a docx sajat, korabban landolt utja -- nem ez a kartya targya, csak nem-utkozes bizonyitasa)',
    /oParagraph\.SetIndLeft\(720\);/.test(buildCreateScript({ core: 'docx', operations: [{ type: 'text', text: 'x', indentLeft: 720 }] }).script))
}

console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
assert.strictEqual(hibak.length, 0, `bukott: ${hibak.join(' | ')}`)
