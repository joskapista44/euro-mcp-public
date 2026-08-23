// paragraphTextsFromDocumentXml / commentsFromCommentsXml, the pure XML
// parsing office_get_text / office_get_comments are built on. Pure-function coverage (no
// browser, no Document Server, no network) -- fixtures below are minimal, hand-written OOXML
// fragments shaped after real saved output (from diagnostic rounds against a live Document
// Server, docx core), not invented from documentation alone.

const {
  paragraphTextsFromDocumentXml, commentsFromCommentsXml, findMatchesInParagraphs,
  paragraphTextsFromSlideXml, pptxSlideOrderFromPresentationXml,
  pptxCommentsPartFromSlideRelsXml, commentsFromPptxCommentXml,
  slideContentSummaryFromSlideXml, documentMetadataFromCoreAndAppXml,
  resolveOoxmlRelativeTarget, ooxmlRelationshipTargetByType, pptxLayoutAndThemeFromSlideLayoutXml,
  footnotesFromFootnotesXml, endnotesFromEndnotesXml, bookmarksFromDocumentXml, tablesFromDocumentXml,
} = require('./lib.cjs')

let osszes = 0
const hibak = []
function check(cimke, felteteles, reszlet = '') {
  osszes += 1
  if (felteteles) { console.log(`  ok    ${cimke}`) } else { hibak.push(cimke); console.log(`  BUKAS ${cimke}${reszlet ? ' -- ' + reszlet : ''}`) }
}

console.log('\n[1] paragraphTextsFromDocumentXml -- EGY bekezdes, TOBB futam: a futamok osszefuznek')
{
  const xml = '<w:body><w:p w14:paraId="1"><w:r><w:rPr><w:b/></w:rPr><w:t>Elso </w:t></w:r><w:r><w:t>resz.</w:t></w:r></w:p></w:body>'
  const out = paragraphTextsFromDocumentXml(xml)
  check('  egy bekezdes all a listaban', out.length === 1)
  check('  a ket futam osszefuzve, sorrendben', out[0] === 'Elso resz.')
}

console.log('\n[2] paragraphTextsFromDocumentXml -- TOBB bekezdes: a hataron NEM fuzodik ossze')
{
  const xml = '<w:p><w:r><w:t>Elso bekezdes.</w:t></w:r></w:p><w:p><w:r><w:t>Masodik bekezdes.</w:t></w:r></w:p>'
  const out = paragraphTextsFromDocumentXml(xml)
  check('  ket kulon bekezdes', out.length === 2)
  check('  az elso tartalma sajat', out[0] === 'Elso bekezdes.')
  check('  a masodik tartalma sajat, NEM az elsoevel osszefuzve', out[1] === 'Masodik bekezdes.')
}

console.log('\n[3] paragraphTextsFromDocumentXml -- URES bekezdes (nincs <w:t>): ures string, NEM esik ki a listabol')
{
  const xml = '<w:p><w:r><w:t>Elotte.</w:t></w:r></w:p><w:p><w:pPr/></w:p><w:p><w:r><w:t>Utana.</w:t></w:r></w:p>'
  const out = paragraphTextsFromDocumentXml(xml)
  check('  harom bekezdes (a kozepso URES, nem hianyzik)', out.length === 3 && out[1] === '')
}

console.log('\n[4] paragraphTextsFromDocumentXml -- XML-ENTITASOK oldva, MERT alak (attributumbelul is `w:t` egyezik)')
{
  const xml = '<w:p><w:r><w:t>&lt;cimke&gt; &amp; &quot;idezet&quot; &apos;aposztrof&apos;</w:t></w:r></w:p>'
  const out = paragraphTextsFromDocumentXml(xml)
  check('  mind az ot entitas oldva', out[0] === '<cimke> & "idezet" \'aposztrof\'')
}

console.log('\n[5] paragraphTextsFromDocumentXml -- TOROLT (tracked-change) tartalom NEM kerul bele (delText, nem t)')
{
  // SetTrackRevisions(true) alatt egy torles <w:del><w:r><w:delText> alakot ir, nem <w:t>-t --
  // a sima <w:t> mintaillesztes ezt strukturalisan kihagyja, semmi extra szures nem kell hozza.
  // Ez a teszt ROGZITI ezt a viselkedest, nem csak allitja.
  const xml = '<w:p><w:del><w:r><w:delText>torolt resz</w:delText></w:r></w:del><w:ins><w:r><w:t>uj resz</w:t></w:r></w:ins></w:p>'
  const out = paragraphTextsFromDocumentXml(xml)
  check('  a torolt resz NEM szerepel, csak a beszurt', out[0] === 'uj resz')
}

console.log('\n[6] paragraphTextsFromDocumentXml -- NULLA/URES bemenet nem dob, ures listat ad')
{
  check('  null -> []', Array.isArray(paragraphTextsFromDocumentXml(null)) && paragraphTextsFromDocumentXml(null).length === 0)
  check('  undefined -> []', paragraphTextsFromDocumentXml(undefined).length === 0)
  check('  ures string -> []', paragraphTextsFromDocumentXml('').length === 0)
}

console.log('\n[7] commentsFromCommentsXml -- EGY komment, MINDEN mezo (id/author/date/text) kiolvasva, egy valodi, elo Document Server-fixtura alakja szerint')
{
  const xml = '<w:comments><w:comment w:id="0" w:author="Alpha" w:date="2026-08-16T11:52:04Z"><w:p><w:r><w:t xml:space="preserve">mp komment</w:t></w:r></w:p></w:comment></w:comments>'
  const out = commentsFromCommentsXml(xml)
  check('  egy komment', out.length === 1)
  check('  id a w:id attributumbol (STRING, nem a live editor GetId()-je -- lasd fuggveny-fejlec)', out[0].id === '0')
  check('  szerzo', out[0].author === 'Alpha')
  check('  datum', out[0].date === '2026-08-16T11:52:04Z')
  check('  szoveg', out[0].text === 'mp komment')
}

console.log('\n[8] commentsFromCommentsXml -- TOBB komment, sorrendben, es a szoveg TOBB futambol is osszeall')
{
  const xml =
    '<w:comments>' +
    '<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>elso </w:t></w:r><w:r><w:t>komment</w:t></w:r></w:p></w:comment>' +
    '<w:comment w:id="1" w:author="B" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>masodik</w:t></w:r></w:p></w:comment>' +
    '</w:comments>'
  const out = commentsFromCommentsXml(xml)
  check('  ket komment', out.length === 2)
  check('  az elso szovege TOBB futambol all ossze', out[0].text === 'elso komment')
  check('  a masodik komment sajat mezoi', out[1].id === '1' && out[1].author === 'B' && out[1].text === 'masodik')
}

console.log('\n[9] commentsFromCommentsXml -- HIANYZO comments.xml (a dokumentumnak nincs kommentje): ures lista, NEM dobas')
{
  check('  null -> []', commentsFromCommentsXml(null).length === 0)
  check('  ures string -> []', commentsFromCommentsXml('').length === 0)
}

console.log('\n[10] findMatchesInParagraphs -- egy talalat egy bekezdesben')
{
  const out = findMatchesInParagraphs(['elso bekezdes.', 'ebben van MINTA egyszer.', 'harmadik.'], 'MINTA')
  check('  osszesitett szam 1', out.totalCount === 1)
  check('  a talalat a helyes bekezdes-indexen', out.matches.length === 1 && out.matches[0].paragraphIndex === 1)
  check('  a bekezdes-szoveg is athozva (a hivo ne kelljen ujra lekernie)', out.matches[0].paragraphText === 'ebben van MINTA egyszer.')
}

console.log('\n[11] findMatchesInParagraphs -- TOBBSZOR egy bekezdesben, ES tobb bekezdesben is')
{
  const out = findMatchesInParagraphs(['x x x', 'nincs itt semmi', 'x es meg x'], 'x')
  check('  osszesitett szam 5 (3+0+2, nem csak a bekezdesek szama)', out.totalCount === 5)
  check('  ket bekezdesen erintett, a masodik kimarad', out.matches.length === 2 &&
    out.matches[0].paragraphIndex === 0 && out.matches[0].count === 3 &&
    out.matches[1].paragraphIndex === 2 && out.matches[1].count === 2)
}

console.log('\n[12] findMatchesInParagraphs -- KISBETU/NAGYBETU SZAMIT (szandekosan nem case-insensitive)')
{
  const out = findMatchesInParagraphs(['Alma', 'alma', 'ALMA'], 'alma')
  check('  csak a pontos kis-nagybetus egyezes szamit', out.totalCount === 1 && out.matches[0].paragraphIndex === 1)
}

console.log('\n[13] findMatchesInParagraphs -- URES kereses vagy URES lista: nulla talalat, nem dobas')
{
  check('  ures query -> 0, nem minden bekezdes "egyezik"', findMatchesInParagraphs(['a', 'b'], '').totalCount === 0)
  check('  ures bekezdes-lista -> 0', findMatchesInParagraphs([], 'x').totalCount === 0)
  check('  null query -> 0, nem dob', findMatchesInParagraphs(['a'], null).totalCount === 0)
}

console.log('\n[14] paragraphTextsFromSlideXml -- MERT alak (elo Document Server), a:p/a:t drawingml nevter')
{
  const xml = '<p:sp><p:txBody><a:p><a:r><a:t>Elso dia szovege, ekezettel: arvizturo</a:t></a:r></a:p></p:txBody></p:sp>'
  const out = paragraphTextsFromSlideXml(xml)
  check('  egy bekezdes', out.length === 1)
  check('  a szoveg pontosan a mert alak', out[0] === 'Elso dia szovege, ekezettel: arvizturo')
}

console.log('\n[15] paragraphTextsFromSlideXml -- TOBB bekezdes NEM fuzodik ossze, tablazat-cella bekezdese IS bekerul (docx-szal konzisztens flatten)')
{
  const xml = '<a:p><a:r><a:t>Cim.</a:t></a:r></a:p>' +
    '<p:graphicFrame><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>Cella szovege.</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></p:graphicFrame>'
  const out = paragraphTextsFromSlideXml(xml)
  check('  ket bekezdes, a cella-szoveg is bekerult', out.length === 2 && out[0] === 'Cim.' && out[1] === 'Cella szovege.')
}

console.log('\n[16] paragraphTextsFromSlideXml -- NULLA/URES bemenet nem dob')
{
  check('  null -> []', paragraphTextsFromSlideXml(null).length === 0)
  check('  ures string -> []', paragraphTextsFromSlideXml('').length === 0)
}

console.log('\n[17] pptxSlideOrderFromPresentationXml -- a sldIdLst SORRENDJE dont, NEM a resz-nevek szama (MoveTo-lelet: egy athelyezett dia a SAJAT reszneven marad)')
{
  // Szandekosan MEGKEVERT rId->resz terkep: a sldIdLst rId9-et mond ELOSZOR, a rels viszont
  // rId9-et slide3.xml-hez rendeli -- ha a fuggveny resz-nevszam szerint rendezne, ez a teszt
  // buknia kellene.
  const presentationXml = '<p:presentation><p:sldIdLst><p:sldId id="1" r:id="rId9"/><p:sldId id="2" r:id="rId4"/></p:sldIdLst></p:presentation>'
  const relsXml = '<Relationships>' +
    '<Relationship Id="rId4" Type=".../slide" Target="slides/slide1.xml"/>' +
    '<Relationship Id="rId9" Type=".../slide" Target="slides/slide3.xml"/>' +
    '</Relationships>'
  const out = pptxSlideOrderFromPresentationXml(presentationXml, relsXml)
  check('  ket dia, a MEGKEVERT sorrendben (slide3 elobb, slide1 utana)', out.length === 2 && out[0] === 'ppt/slides/slide3.xml' && out[1] === 'ppt/slides/slide1.xml')
}

console.log('\n[18] pptxSlideOrderFromPresentationXml -- MERT alak, elo Document Server (5 dia, egyenkent)')
{
  const presentationXml = '<p:sldIdLst><p:sldId id="256" r:id="rId4"></p:sldId><p:sldId id="257" r:id="rId5"></p:sldId><p:sldId id="258" r:id="rId6"></p:sldId><p:sldId id="259" r:id="rId7"></p:sldId><p:sldId id="260" r:id="rId8"></p:sldId></p:sldIdLst>'
  const relsXml = '<Relationships>' +
    '<Relationship Id="rId4" Target="slides/slide1.xml"/><Relationship Id="rId5" Target="slides/slide2.xml"/>' +
    '<Relationship Id="rId6" Target="slides/slide3.xml"/><Relationship Id="rId7" Target="slides/slide4.xml"/>' +
    '<Relationship Id="rId8" Target="slides/slide5.xml"/></Relationships>'
  const out = pptxSlideOrderFromPresentationXml(presentationXml, relsXml)
  check('  5 dia, sorrendben', out.length === 5, JSON.stringify(out))
  check('  mind az ot resz-nev helyes', ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml', 'ppt/slides/slide4.xml', 'ppt/slides/slide5.xml'].every((x, i) => out[i] === x))
}

console.log('\n[19] pptxSlideOrderFromPresentationXml -- NULLA/URES bemenet nem dob')
{
  check('  null,null -> []', pptxSlideOrderFromPresentationXml(null, null).length === 0)
  check('  ures,ures -> []', pptxSlideOrderFromPresentationXml('', '').length === 0)
}

console.log('\n[20] pptxCommentsPartFromSlideRelsXml -- MERT alak (elo Document Server): "…/relationships/comments" tipusu kapcsolat -> ppt/comments/commentN.xml')
{
  const relsXml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml" />' +
    '</Relationships>'
  check('  a comments-tipusu kapcsolat resz-neve, a tobbi (layout/notes) figyelmen kivul', pptxCommentsPartFromSlideRelsXml(relsXml) === 'ppt/comments/comment1.xml')
}

console.log('\n[21] pptxCommentsPartFromSlideRelsXml -- HIANYZO/NINCS comments-kapcsolat: null, NEM kitalalt utvonal')
{
  const relsXml = '<Relationships><Relationship Id="rId1" Type=".../slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>'
  check('  nincs comments-kapcsolat -> null', pptxCommentsPartFromSlideRelsXml(relsXml) === null)
  check('  null bemenet -> null', pptxCommentsPartFromSlideRelsXml(null) === null)
}

console.log('\n[22] commentsFromPptxCommentXml -- MERT alak (elo Document Server): idx/authorId/date/text + a szerzo nev a KULON commentAuthors.xml-bol')
{
  const commentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmLst><p:cm authorId="1" dt="2026-08-17T01:40:26Z" idx="1"><p:pos x="1888" y="1888"/><p:text>Elso komment szovege</p:text></p:cm></p:cmLst>'
  const authorsXml = '<p:cmAuthorLst><p:cmAuthor id="1" name="Proba Szerzo" initials="PS" lastIdx="1" clrIdx="0"></p:cmAuthor></p:cmAuthorLst>'
  const out = commentsFromPptxCommentXml(commentXml, authorsXml)
  check('  egy komment', out.length === 1)
  check('  idx/authorId a mert alakbol', out[0].idx === '1' && out[0].authorId === '1')
  check('  datum', out[0].date === '2026-08-17T01:40:26Z')
  check('  szoveg (a <p:text> egyetlen csomopontja, nincs futam-szerkezet mint docx-nel)', out[0].text === 'Elso komment szovege')
  check('  szerzo nev a KULON authors XML-bol felodva', out[0].authorName === 'Proba Szerzo')
}

console.log('\n[23] commentsFromPptxCommentXml -- HIANYZO authorsXml: a komment SZOVEGE megvan, csak a nev null (nem az egesz olvasas bukik el miatta)')
{
  const commentXml = '<p:cmLst><p:cm authorId="1" dt="2026-01-01T00:00:00Z" idx="1"><p:text>szoveg szerzo nelkul</p:text></p:cm></p:cmLst>'
  const out = commentsFromPptxCommentXml(commentXml, null)
  check('  a komment megvan, szovege ep', out.length === 1 && out[0].text === 'szoveg szerzo nelkul')
  check('  a nev null, nem ures string es nem dob', out[0].authorName === null)
}

console.log('\n[24] commentsFromPptxCommentXml -- TOBB komment, es a szerzo-terkep TOBB szerzot is felold')
{
  const commentXml = '<p:cmLst>' +
    '<p:cm authorId="1" dt="2026-01-01T00:00:00Z" idx="1"><p:text>elso</p:text></p:cm>' +
    '<p:cm authorId="2" dt="2026-01-02T00:00:00Z" idx="2"><p:text>masodik</p:text></p:cm>' +
    '</p:cmLst>'
  const authorsXml = '<p:cmAuthorLst><p:cmAuthor id="1" name="A"></p:cmAuthor><p:cmAuthor id="2" name="B"></p:cmAuthor></p:cmAuthorLst>'
  const out = commentsFromPptxCommentXml(commentXml, authorsXml)
  check('  ket komment, sajat szerzovel', out.length === 2 && out[0].authorName === 'A' && out[1].authorName === 'B')
}

console.log('\n[25] commentsFromPptxCommentXml -- NULLA/URES bemenet nem dob')
{
  check('  null,null -> []', commentsFromPptxCommentXml(null, null).length === 0)
  check('  ures,ures -> []', commentsFromPptxCommentXml('', '').length === 0)
}

console.log('\n[26] slideContentSummaryFromSlideXml -- EGY dia, MIND A NEGY tipus (shape/kep/tablazat/diagram), MERT alak (elo Document Server)')
{
  const xml = '<p:sp><p:nvSpPr><p:cNvPr id="137006514" name=""/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="800000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>' +
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="1746553365" name=""/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="700000"/><a:ext cx="8128000" cy="538435"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl></a:tbl></a:graphicData></a:graphic></p:graphicFrame>' +
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="1443495553" name=""/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="1400000"/><a:ext cx="2000000" cy="1500000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"></c:chart></a:graphicData></a:graphic></p:graphicFrame>' +
    '<p:pic><p:nvPicPr><p:cNvPr id="992130440" name=""/></p:nvPicPr><p:blipFill><a:blip r:embed="rId4"/></p:blipFill><p:spPr><a:xfrm><a:off x="2500000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm></p:spPr></p:pic>'
  const out = slideContentSummaryFromSlideXml(xml)
  check('  1 shape, helyes mezokkel (id/pozicio/meret/tipus)', out.shapes.length === 1 &&
    out.shapes[0].id === '137006514' && out.shapes[0].x === 100000 && out.shapes[0].y === 100000 &&
    out.shapes[0].cx === 800000 && out.shapes[0].cy === 500000 && out.shapes[0].shapeType === 'rect')
  check('  1 kep', out.images.length === 1 && out.images[0].id === '992130440' && out.images[0].x === 2500000)
  check('  1 tablazat (graphicFrame, table URI)', out.tables.length === 1 && out.tables[0].id === '1746553365')
  check('  1 diagram (graphicFrame, chart URI) -- MEGKULONBOZTETVE a tablazattol, NEM a tablazat listajaba kerult', out.charts.length === 1 && out.charts[0].id === '1443495553')
  check('  0 csoport', out.groups.length === 0)
}

console.log('\n[27] slideContentSummaryFromSlideXml -- CSOPORT: a NESTED alakzatok IS bekerulnek a shapes listaba (nem csak a csoport onmagaban), MERT alak (K7 sajat group-probaja)')
{
  const xml = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="132748601" name=""/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="900000" cy="400000"/><a:chOff x="0" y="0"/><a:chExt cx="900000" cy="400000"/></a:xfrm></p:grpSpPr>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="1201675296" name=""/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="1552308183" name=""/></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="0"/><a:ext cx="400000" cy="400000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:sp>' +
    '</p:grpSp>'
  const out = slideContentSummaryFromSlideXml(xml)
  check('  1 csoport, sajat id-vel', out.groups.length === 1 && out.groups[0].id === '132748601')
  check('  a KET alakzat a csoporton BELUL IS bekerult a shapes listaba (dokumentalt egyszerusites, nem hiba)',
    out.shapes.length === 2 && out.shapes[0].shapeType === 'rect' && out.shapes[1].shapeType === 'ellipse')
}

console.log('\n[28] slideContentSummaryFromSlideXml -- URES dia (nincs semmi rajta): minden lista ures, NEM dobas')
{
  const out = slideContentSummaryFromSlideXml('<p:cSld><p:spTree></p:spTree></p:cSld>')
  check('  minden kategoria ures', out.shapes.length === 0 && out.images.length === 0 && out.tables.length === 0 && out.charts.length === 0 && out.groups.length === 0)
}

console.log('\n[29] slideContentSummaryFromSlideXml -- NULLA/URES bemenet nem dob')
{
  const out = slideContentSummaryFromSlideXml(null)
  check('  null -> minden lista ures tomb, nem dobas', Array.isArray(out.shapes) && out.shapes.length === 0 && Array.isArray(out.tables) && out.tables.length === 0)
}

console.log('\n[30] documentMetadataFromCoreAndAppXml -- MERT alak, VALODI elo Document Server kimenet (K7 sajat probaja): egy DocBuilder-mentett dokumentum docProps-a szinte URES')
{
  const coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><cp:lastModifiedBy></cp:lastModifiedBy></cp:coreProperties>'
  const appXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><TotalTime>0</TotalTime><Words>0</Words><Application>ONLYOFFICE/2.5.565.0</Application><Slides>1</Slides><Notes>1</Notes><HiddenSlides>0</HiddenSlides></Properties>'
  const out = documentMetadataFromCoreAndAppXml(coreXml, appXml)
  check('  hianyzo core-mezok (title/creator/created/...) NULL, nem ures string', out.title === null && out.creator === null && out.created === null)
  check('  JELENLEVO, de URES cp:lastModifiedBy -- URES STRING, nem null (a ketto kulonbozo teny)', out.lastModifiedBy === '')
  check('  app.xml szammezoi SZAMKENT (nem sztringkent) allnak', out.slides === 1 && out.words === 0 && out.hiddenSlides === 0 && typeof out.slides === 'number')
  check('  application sztring', out.application === 'ONLYOFFICE/2.5.565.0')
}

console.log('\n[31] documentMetadataFromCoreAndAppXml -- POZITIV UT: minden nevesitett core-mezo JELEN van es kiolvasva (kezzel irt fixture, mert a valodi probafajl vegig ures volt ezekre)')
{
  const coreXml = '<cp:coreProperties xmlns:cp="x" xmlns:dc="y" xmlns:dcterms="z"><dc:title>Cim</dc:title><dc:subject>Targy</dc:subject><dc:creator>Alpha</dc:creator><cp:lastModifiedBy>Beta</cp:lastModifiedBy><dcterms:created>2026-08-17T01:00:00Z</dcterms:created><dcterms:modified>2026-08-17T02:00:00Z</dcterms:modified><cp:revision>3</cp:revision><cp:category>Proba</cp:category></cp:coreProperties>'
  const out = documentMetadataFromCoreAndAppXml(coreXml, null)
  check('  mind a nyolc core-mezo a sajat erteket adja', out.title === 'Cim' && out.subject === 'Targy' && out.creator === 'Alpha' &&
    out.lastModifiedBy === 'Beta' && out.created === '2026-08-17T01:00:00Z' && out.modified === '2026-08-17T02:00:00Z' &&
    out.revision === '3' && out.category === 'Proba')
  check('  hianyzo appXml -> az app-mezok NULL, nem dobas', out.application === null && out.slides === null)
}

console.log('\n[32] documentMetadataFromCoreAndAppXml -- XML-ENTITASOK oldva a core-mezokben')
{
  const coreXml = '<cp:coreProperties xmlns:dc="y"><dc:title>Cim &amp; alcim</dc:title></cp:coreProperties>'
  check('  entitas oldva', documentMetadataFromCoreAndAppXml(coreXml, null).title === 'Cim & alcim')
}

console.log('\n[33] documentMetadataFromCoreAndAppXml -- NULLA/URES bemenet nem dob, minden mezo null')
{
  const out = documentMetadataFromCoreAndAppXml(null, null)
  check('  minden mezo null', Object.values(out).every((v) => v === null))
  const out2 = documentMetadataFromCoreAndAppXml('', '')
  check('  ures stringek is', Object.values(out2).every((v) => v === null))
}

console.log('\n[34] resolveOoxmlRelativeTarget -- MERT alak (K7 sajat probaja): "../x" egy szinttel felfele lep, sima nev csak hozzafuz')
{
  check('  "../slideLayouts/slideLayout1.xml" ppt/slides-bol -> ppt/slideLayouts/slideLayout1.xml',
    resolveOoxmlRelativeTarget('ppt/slides', '../slideLayouts/slideLayout1.xml') === 'ppt/slideLayouts/slideLayout1.xml')
  check('  "slides/slide1.xml" ppt-bol -> ppt/slides/slide1.xml (a presentation.xml.rels sajat alakja, nincs "../")',
    resolveOoxmlRelativeTarget('ppt', 'slides/slide1.xml') === 'ppt/slides/slide1.xml')
  check('  ket szinttel felfele is mukodik', resolveOoxmlRelativeTarget('ppt/a/b', '../../c.xml') === 'ppt/c.xml')
}

console.log('\n[35] ooxmlRelationshipTargetByType -- a MEGADOTT tipus-vegzodesu kapcsolat Target-je, a tobbi figyelmen kivul')
{
  const relsXml = '<Relationships><Relationship Id="rId1" Type=".../slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type=".../notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>'
  check('  a slideLayout kapcsolat celja', ooxmlRelationshipTargetByType(relsXml, '/slideLayout') === '../slideLayouts/slideLayout1.xml')
  check('  nincs ilyen tipusu kapcsolat -> null', ooxmlRelationshipTargetByType(relsXml, '/theme') === null)
  check('  null bemenet -> null, nem dob', ooxmlRelationshipTargetByType(null, '/theme') === null)
}

console.log('\n[36] pptxLayoutAndThemeFromSlideLayoutXml -- MERT alak, VALODI elo Document Server kimenet (K7 sajat probaja, teljes lanc: dia -> elrendezes -> mester -> tema)')
{
  const layoutXml = '<p:sldLayout xmlns:p="x" matchingName="" preserve="0" type="blank" userDrawn="1"><p:cSld name="Blank"></p:cSld></p:sldLayout>'
  const themeXml = '<a:theme xmlns:a="y" name="SampleTheme"></a:theme>'
  const out = pptxLayoutAndThemeFromSlideLayoutXml(layoutXml, themeXml)
  check('  elrendezes neve es tipusa', out.layoutName === 'Blank' && out.layoutType === 'blank')
  check('  tema neve (a KULON mester-lancon at felodva)', out.themeName === 'SampleTheme')
}

console.log('\n[37] pptxLayoutAndThemeFromSlideLayoutXml -- a mester-lanc megszakadt (nincs tema-kapcsolat): themeName null, a layout mezok megvannak')
{
  const out = pptxLayoutAndThemeFromSlideLayoutXml('<p:sldLayout type="title"><p:cSld name="Cim"></p:cSld></p:sldLayout>', null)
  check('  a layout mezok fuggetlenek a tema hianyatol', out.layoutName === 'Cim' && out.layoutType === 'title')
  check('  themeName null, nem ures string, nem dob', out.themeName === null)
}

console.log('\n[38] pptxLayoutAndThemeFromSlideLayoutXml -- NULLA/URES bemenet nem dob')
{
  const out = pptxLayoutAndThemeFromSlideLayoutXml(null, null)
  check('  layoutName ures string (a mezo letezik, csak nincs benne nev), layoutType/themeName null', out.layoutName === '' && out.layoutType === null && out.themeName === null)
}

console.log('\n[39] footnotesFromFootnotesXml -- valodi Document Server-fixtura, boilerplate kiszurve')
{
  // Elo futasbol: ket AddFootnote() hivas ELO Document Server-en, a mentett word/footnotes.xml
  // pontos alakja (rovidítve a nem-lenyeges rPr/pBdr/ind ures elemekre, a szoveg-hordozo resz
  // VALTOZATLAN).
  const xml = '<w:footnotes>' +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="2"><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>first footnote content</w:t></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="3"><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>second footnote content</w:t></w:r></w:p></w:footnote>' +
    '</w:footnotes>'
  const out = footnotesFromFootnotesXml(xml)
  check('  a KET boilerplate (separator/continuationSeparator) NEM szerepel', out.length === 2)
  check('  a valodi footnote-ok id-vel es szoveggel', out[0].id === '2' && out[0].text === ' first footnote content' && out[1].id === '3')
}

console.log('\n[40] footnotesFromFootnotesXml / endnotesFromEndnotesXml -- NULLA/URES bemenet nem dob')
{
  check('  footnotesFromFootnotesXml(null) -> ures tomb', Array.isArray(footnotesFromFootnotesXml(null)) && footnotesFromFootnotesXml(null).length === 0)
  check('  endnotesFromEndnotesXml(null) -> ures tomb', Array.isArray(endnotesFromEndnotesXml(null)) && endnotesFromEndnotesXml(null).length === 0)
}

console.log('\n[41] endnotesFromEndnotesXml -- ugyanaz a szerkezet, <w:endnote> cimkevel')
{
  const xml = '<w:endnotes>' +
    '<w:endnote w:type="separator" w:id="-1"><w:p></w:p></w:endnote>' +
    '<w:endnote w:id="1"><w:p><w:r><w:t>an endnote</w:t></w:r></w:p></w:endnote>' +
    '</w:endnotes>'
  const out = endnotesFromEndnotesXml(xml)
  check('  a boilerplate kiszurve, egy valodi endnote marad', out.length === 1 && out[0].id === '1' && out[0].text === 'an endnote')
}

console.log('\n[42] bookmarksFromDocumentXml -- valodi fixtura (koteg09 sajat probaja), nev+szoveg')
{
  // Elo futasbol: Range.AddBookmark("mybm") egy "hello world" szovegu bekezdesen.
  const xml = '<w:body><w:p><w:bookmarkStart w:id="0" w:name="mybm"/><w:r><w:t>hello world</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p></w:body>'
  const out = bookmarksFromDocumentXml(xml)
  check('  egy konyvjelzo, nevvel es a wrap-elt szoveggel', out.length === 1 && out[0].name === 'mybm' && out[0].text === 'hello world')
}

console.log('\n[43] bookmarksFromDocumentXml -- a w:id PAROSITAS szamit, NEM a dokumentum-sorrend')
{
  // Ket konyvjelzo, at nem fedo tartomanyokkal, MAS sorrendben zarva (b elobb zarul, mint a nyilik).
  const xml = '<w:body>' +
    '<w:p><w:bookmarkStart w:id="5" w:name="alpha"/><w:r><w:t>alpha-text</w:t></w:r><w:bookmarkEnd w:id="5"/>' +
    '<w:bookmarkStart w:id="7" w:name="beta"/><w:r><w:t>beta-text</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>' +
    '</w:body>'
  const out = bookmarksFromDocumentXml(xml)
  check('  ket konyvjelzo, MINDEGYIK a SAJAT id-parjahoz tartozo szoveggel (nem keveredik ossze)',
    out.length === 2 && out[0].name === 'alpha' && out[0].text === 'alpha-text' && out[1].name === 'beta' && out[1].text === 'beta-text')
}

console.log('\n[44] bookmarksFromDocumentXml -- hianyzo bookmarkEnd -> text:null, nem dob')
{
  const xml = '<w:body><w:p><w:bookmarkStart w:id="0" w:name="orphan"/><w:r><w:t>x</w:t></w:r></w:p></w:body>'
  const out = bookmarksFromDocumentXml(xml)
  check('  a nev megvan, a szoveg null (nincs zaro jel)', out.length === 1 && out[0].name === 'orphan' && out[0].text === null)
}

console.log('\n[45] tablesFromDocumentXml -- LAPOS lista, valodi 2-tablazatos fixtura (koteg07 sajat probaja)')
{
  // Elo futasbol: kulso tablazat (2x2: name/detail fejlec, outer-row + beagyazott tablazat),
  // a beagyazott tablazat 2x2 (inner-a/b/c/d).
  const xml = '<w:body><w:tbl>' +
    '<w:tr><w:tc><w:p><w:r><w:t>name</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>detail</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>outer-row</w:t></w:r></w:p></w:tc><w:tc><w:tbl>' +
    '<w:tr><w:tc><w:p><w:r><w:t>inner-a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>inner-b</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>inner-c</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>inner-d</w:t></w:r></w:p></w:tc></w:tr>' +
    '</w:tbl></w:tc></w:tr>' +
    '</w:tbl></w:body>'
  const out = tablesFromDocumentXml(xml)
  check('  KET tablazat-bejegyzes (kulso + beagyazott)', out.length === 2)
  check('  a kulso tablazat parentIndex null, 2 sor, a masodik sor MASODIK cellaja URES (a beagyazott tabla NEM szivarog bele)',
    out[0].parentIndex === null && out[0].sorok.length === 2 && out[0].sorok[1][0] === 'outer-row' && out[0].sorok[1][1] === '')
  check('  a beagyazott tablazat parentIndex a kulso index-ere mutat, sajat 2x2 tartalommal',
    out[1].parentIndex === 0 && out[1].sorok.length === 2 && out[1].sorok[0][0] === 'inner-a' && out[1].sorok[1][1] === 'inner-d')
}

console.log('\n[46] tablesFromDocumentXml -- HAROM szintu beagyazas (koteg07 mert melyseg-hatara)')
{
  const xml = '<w:body><w:tbl><w:tr><w:tc>' +
    '<w:tbl><w:tr><w:tc>' +
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>deepest</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    '</w:tc></w:tr></w:tbl>' +
    '</w:tc></w:tr></w:tbl></w:body>'
  const out = tablesFromDocumentXml(xml)
  check('  HAROM tablazat-bejegyzes, dokumentum-sorrendben (kulso elobb)', out.length === 3)
  check('  lancolt szulo-viszony: 0<-1<-2, es CSAK a legbelso hordoz szoveget',
    out[0].parentIndex === null && out[1].parentIndex === 0 && out[2].parentIndex === 1 &&
    out[0].sorok[0][0] === '' && out[1].sorok[0][0] === '' && out[2].sorok[0][0] === 'deepest')
}

console.log('\n[47] tablesFromDocumentXml -- NULLA tablazat / NULLA bemenet nem dob')
{
  check('  ures dokumentum -> ures tomb', tablesFromDocumentXml('<w:body><w:p><w:r><w:t>no tables here</w:t></w:r></w:p></w:body>').length === 0)
  check('  null bemenet -> ures tomb, nem dob', Array.isArray(tablesFromDocumentXml(null)) && tablesFromDocumentXml(null).length === 0)
}

console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
if (hibak.length) { console.log('BUKOTT:', hibak.join(' | ')); process.exit(1) }
