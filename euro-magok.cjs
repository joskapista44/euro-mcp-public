// EURO-MCP: MAG-ELOALLITAS (E7) -- a keszulo fajl TIPUSAT es SZERKEZETET adja, tartalmat nem.
//
// MIERT LETEZIK: a DocBuilder `run_builder_script` kotelezoen `builder.OpenFile()`-t var (a
// `CreateFile` szandekosan nincs kinalva, upstream #321), ezert egy UJ fajlhoz is kell tipushelyes
// bemenet. *** ES FUTAS KOZBEN NEM LEHET LAPOT/DIAT HOZZAADNI: *** az `Api.AddSheet()` es az
// `Api.CreateSlide()+AddSlide()` MEGOLI a jobot (merve 2026-08-15: outcome "blocked", nulla kimenet,
// nulla hibauzenet). A lap-/diaszam tehat CSAK innen johet.
//
// A FELOSZTAS, ES MIERT IGY:
//   a VALTOZO reszek  -> itt generalodnak (lapok/diak + a rajuk hivatkozo lista-jellegu reszek)
//   a STATIKUS reszek -> BEAGYAZVA, base64-ben (a pptx temaja, dia-mestere, elrendezese)
//   a csomagolas      -> sajat ZIP-iro, `zlib`-bel: a Node-ban VAN deflate ES crc32, uj fuggoseg nincs
//
// *** PROVENIENCIA (a beagyazott reszek nem olvashatok ranezesre, ezert ez kotelezo): ***
//   forras:  a demo-pptx-bol csupaszitott minimalis mag (mag-minimal-pptx.py), 2026-08-15
//   ujra-eloallitas:  python3 mag-minimal-pptx.py <demo.pptx> <kimenet.pptx> 1
//                     majd a statikus reszek base64-kent kiemelve
//   *** ES A TESZT, AMI EZT ORZI: a beagyazott reszeket KICSOMAGOLJA es a resz-listat ellenorzi.
//       Egy elrontott blob kulonben nem itt bukik el, hanem a kesz dokumentumon. ***

const zlib = require('zlib')

const PPTX_STATIKUS = {
  "_rels/.rels": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxSZWxhdGlvbnNoaXBzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0aW9uc2hpcHMiPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDEiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0icHB0L3ByZXNlbnRhdGlvbi54bWwiLz48UmVsYXRpb25zaGlwIElkPSJySWQyIiBUeXBlPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0aW9uc2hpcHMvbWV0YWRhdGEvY29yZS1wcm9wZXJ0aWVzIiBUYXJnZXQ9ImRvY1Byb3BzL2NvcmUueG1sIi8+PFJlbGF0aW9uc2hpcCBJZD0icklkMyIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9leHRlbmRlZC1wcm9wZXJ0aWVzIiBUYXJnZXQ9ImRvY1Byb3BzL2FwcC54bWwiLz48L1JlbGF0aW9uc2hpcHM+',
  "docProps/core.xml": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxjcDpjb3JlUHJvcGVydGllcyB4bWxuczpjcD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9tZXRhZGF0YS9jb3JlLXByb3BlcnRpZXMiIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyI+PGRjOnRpdGxlPktlcGVzc2VnLWJlbXV0YXRvPC9kYzp0aXRsZT48ZGM6Y3JlYXRvcj5FdXJvLU9mZmljZTwvZGM6Y3JlYXRvcj48L2NwOmNvcmVQcm9wZXJ0aWVzPg==',
  "docProps/app.xml": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxQcm9wZXJ0aWVzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9leHRlbmRlZC1wcm9wZXJ0aWVzIj48QXBwbGljYXRpb24+RXVyby1PZmZpY2U8L0FwcGxpY2F0aW9uPjwvUHJvcGVydGllcz4=',
  "ppt/slideMasters/slideMaster1.xml": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxwOnNsZE1hc3RlciB4bWxuczphPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvZHJhd2luZ21sLzIwMDYvbWFpbiIgeG1sbnM6cD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3ByZXNlbnRhdGlvbm1sLzIwMDYvbWFpbiIgeG1sbnM6cj0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcyI+PHA6Y1NsZD48cDpzcFRyZWU+PHA6bnZHcnBTcFByPjxwOmNOdlByIGlkPSIxIiBuYW1lPSIiLz48cDpjTnZHcnBTcFByLz48cDpudlByLz48L3A6bnZHcnBTcFByPjxwOmdycFNwUHI+PGE6eGZybT48YTpvZmYgeD0iMCIgeT0iMCIvPjxhOmV4dCBjeD0iMCIgY3k9IjAiLz48YTpjaE9mZiB4PSIwIiB5PSIwIi8+PGE6Y2hFeHQgY3g9IjAiIGN5PSIwIi8+PC9hOnhmcm0+PC9wOmdycFNwUHI+PC9wOnNwVHJlZT48L3A6Y1NsZD48cDpjbHJNYXAgYmcxPSJsdDEiIHR4MT0iZGsxIiBiZzI9Imx0MiIgdHgyPSJkazIiIGFjY2VudDE9ImFjY2VudDEiIGFjY2VudDI9ImFjY2VudDIiIGFjY2VudDM9ImFjY2VudDMiIGFjY2VudDQ9ImFjY2VudDQiIGFjY2VudDU9ImFjY2VudDUiIGFjY2VudDY9ImFjY2VudDYiIGhsaW5rPSJobGluayIgZm9sSGxpbms9ImZvbEhsaW5rIi8+PHA6c2xkTGF5b3V0SWRMc3Q+PHA6c2xkTGF5b3V0SWQgaWQ9IjIxNDc0ODM2NDkiIHI6aWQ9InJJZDEiLz48L3A6c2xkTGF5b3V0SWRMc3Q+PC9wOnNsZE1hc3Rlcj4=',
  "ppt/slideMasters/_rels/slideMaster1.xml.rels": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxSZWxhdGlvbnNoaXBzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0aW9uc2hpcHMiPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDEiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvc2xpZGVMYXlvdXQiIFRhcmdldD0iLi4vc2xpZGVMYXlvdXRzL3NsaWRlTGF5b3V0MS54bWwiLz48UmVsYXRpb25zaGlwIElkPSJySWQyIiBUeXBlPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzL3RoZW1lIiBUYXJnZXQ9Ii4uL3RoZW1lL3RoZW1lMS54bWwiLz48L1JlbGF0aW9uc2hpcHM+',
  "ppt/slideLayouts/slideLayout1.xml": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxwOnNsZExheW91dCB4bWxuczphPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvZHJhd2luZ21sLzIwMDYvbWFpbiIgeG1sbnM6cD0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3ByZXNlbnRhdGlvbm1sLzIwMDYvbWFpbiIgeG1sbnM6cj0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcyIgdHlwZT0iYmxhbmsiIHByZXNlcnZlPSIxIj48cDpjU2xkPjxwOnNwVHJlZT48cDpudkdycFNwUHI+PHA6Y052UHIgaWQ9IjEiIG5hbWU9IiIvPjxwOmNOdkdycFNwUHIvPjxwOm52UHIvPjwvcDpudkdycFNwUHI+PHA6Z3JwU3BQcj48YTp4ZnJtPjxhOm9mZiB4PSIwIiB5PSIwIi8+PGE6ZXh0IGN4PSIwIiBjeT0iMCIvPjxhOmNoT2ZmIHg9IjAiIHk9IjAiLz48YTpjaEV4dCBjeD0iMCIgY3k9IjAiLz48L2E6eGZybT48L3A6Z3JwU3BQcj48L3A6c3BUcmVlPjwvcDpjU2xkPjxwOmNsck1hcE92cj48YTptYXN0ZXJDbHJNYXBwaW5nLz48L3A6Y2xyTWFwT3ZyPjwvcDpzbGRMYXlvdXQ+',
  "ppt/slideLayouts/_rels/slideLayout1.xml.rels": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/PjxSZWxhdGlvbnNoaXBzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0aW9uc2hpcHMiPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDEiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvc2xpZGVNYXN0ZXIiIFRhcmdldD0iLi4vc2xpZGVNYXN0ZXJzL3NsaWRlTWFzdGVyMS54bWwiLz48L1JlbGF0aW9uc2hpcHM+',
  "ppt/theme/theme1.xml": 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pgo8YTp0aGVtZSB4bWxuczphPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvZHJhd2luZ21sLzIwMDYvbWFpbiIgbmFtZT0iRXVyby1PZmZpY2UiPjxhOnRoZW1lRWxlbWVudHM+CjxhOmNsclNjaGVtZSBuYW1lPSJFdXJvLU9mZmljZSI+PGE6ZGsxPjxhOnNyZ2JDbHIgdmFsPSIwMDAwMDAiLz48L2E6ZGsxPjxhOmx0MT48YTpzcmdiQ2xyIHZhbD0iRkZGRkZGIi8+PC9hOmx0MT48YTpkazI+PGE6c3JnYkNsciB2YWw9IjFGMzg2NCIvPjwvYTpkazI+PGE6bHQyPjxhOnNyZ2JDbHIgdmFsPSJFRUYyRjgiLz48L2E6bHQyPjxhOmFjY2VudDE+PGE6c3JnYkNsciB2YWw9IjJFNzRCNSIvPjwvYTphY2NlbnQxPjxhOmFjY2VudDI+PGE6c3JnYkNsciB2YWw9IjlEQzNFNiIvPjwvYTphY2NlbnQyPjxhOmFjY2VudDM+PGE6c3JnYkNsciB2YWw9IjAwODA2MCIvPjwvYTphY2NlbnQzPjxhOmFjY2VudDQ+PGE6c3JnYkNsciB2YWw9IkMwMDAwMCIvPjwvYTphY2NlbnQ0PjxhOmFjY2VudDU+PGE6c3JnYkNsciB2YWw9IkVEN0QzMSIvPjwvYTphY2NlbnQ1PjxhOmFjY2VudDY+PGE6c3JnYkNsciB2YWw9IjcwMzBBMCIvPjwvYTphY2NlbnQ2PjxhOmhsaW5rPjxhOnNyZ2JDbHIgdmFsPSIwNTYzQzEiLz48L2E6aGxpbms+PGE6Zm9sSGxpbms+PGE6c3JnYkNsciB2YWw9Ijk1NEY3MiIvPjwvYTpmb2xIbGluaz48L2E6Y2xyU2NoZW1lPgo8YTpmb250U2NoZW1lIG5hbWU9IkV1cm8tT2ZmaWNlIj48YTptYWpvckZvbnQ+PGE6bGF0aW4gdHlwZWZhY2U9IkNhbGlicmkgTGlnaHQiLz48YTplYSB0eXBlZmFjZT0iIi8+PGE6Y3MgdHlwZWZhY2U9IiIvPjwvYTptYWpvckZvbnQ+CjxhOm1pbm9yRm9udD48YTpsYXRpbiB0eXBlZmFjZT0iQ2FsaWJyaSIvPjxhOmVhIHR5cGVmYWNlPSIiLz48YTpjcyB0eXBlZmFjZT0iIi8+PC9hOm1pbm9yRm9udD48L2E6Zm9udFNjaGVtZT4KPGE6Zm10U2NoZW1lIG5hbWU9IkV1cm8tT2ZmaWNlIj4KPGE6ZmlsbFN0eWxlTHN0PjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIi8+PC9hOnNvbGlkRmlsbD48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6c29saWRGaWxsPjwvYTpmaWxsU3R5bGVMc3Q+CjxhOmxuU3R5bGVMc3Q+PGE6bG4gdz0iNjM1MCI+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIi8+PC9hOnNvbGlkRmlsbD48L2E6bG4+PGE6bG4gdz0iMTI3MDAiPjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PC9hOmxuPjxhOmxuIHc9IjE5MDUwIj48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6c29saWRGaWxsPjwvYTpsbj48L2E6bG5TdHlsZUxzdD4KPGE6ZWZmZWN0U3R5bGVMc3Q+PGE6ZWZmZWN0U3R5bGU+PGE6ZWZmZWN0THN0Lz48L2E6ZWZmZWN0U3R5bGU+PGE6ZWZmZWN0U3R5bGU+PGE6ZWZmZWN0THN0Lz48L2E6ZWZmZWN0U3R5bGU+PGE6ZWZmZWN0U3R5bGU+PGE6ZWZmZWN0THN0Lz48L2E6ZWZmZWN0U3R5bGU+PC9hOmVmZmVjdFN0eWxlTHN0Pgo8YTpiZ0ZpbGxTdHlsZUxzdD48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6c29saWRGaWxsPjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIi8+PC9hOnNvbGlkRmlsbD48L2E6YmdGaWxsU3R5bGVMc3Q+CjwvYTpmbXRTY2hlbWU+PC9hOnRoZW1lRWxlbWVudHM+PC9hOnRoZW1lPg==',
}

const PPTX_DIA_RELS = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide\" Target=\"../notesSlides/notesSlide1.xml\"/></Relationships>"

// --- A ZIP-IRO ---------------------------------------------------------------------------------
// Egy OOXML csomag egy sima ZIP. A Node-ban minden megvan hozza (`deflateRawSync`, `crc32`), tehat
// nincs uj fuggoseg -- es ez szandekos: egy dokumentum-eloallito, ami egy kulso csomagon all, egy
// verzio-emelesnel tud csendben mast csinalni.
function zipCsomag(reszek) {
  const helyi = []
  const kozponti = []
  let eltolas = 0
  for (const [nev, tartalom] of Object.entries(reszek)) {
    const adat = Buffer.isBuffer(tartalom) ? tartalom : Buffer.from(tartalom, 'utf8')
    const tomor = zlib.deflateRawSync(adat)
    const crc = zlib.crc32(adat)
    const nevBuf = Buffer.from(nev, 'utf8')
    const fej = Buffer.alloc(30)
    fej.writeUInt32LE(0x04034b50, 0)
    fej.writeUInt16LE(20, 4)          // verzio
    fej.writeUInt16LE(0, 6)           // jelzok
    fej.writeUInt16LE(8, 8)           // deflate
    fej.writeUInt32LE(0, 10)          // ido/datum: FIX, hogy a kimenet REPRODUKALHATO legyen
    fej.writeUInt32LE(crc, 14)
    fej.writeUInt32LE(tomor.length, 18)
    fej.writeUInt32LE(adat.length, 22)
    fej.writeUInt16LE(nevBuf.length, 26)
    fej.writeUInt16LE(0, 28)
    helyi.push(fej, nevBuf, tomor)
    const kfej = Buffer.alloc(46)
    kfej.writeUInt32LE(0x02014b50, 0)
    kfej.writeUInt16LE(20, 4); kfej.writeUInt16LE(20, 6)
    kfej.writeUInt16LE(0, 8); kfej.writeUInt16LE(8, 10)
    kfej.writeUInt32LE(0, 12)
    kfej.writeUInt32LE(crc, 16)
    kfej.writeUInt32LE(tomor.length, 20)
    kfej.writeUInt32LE(adat.length, 24)
    kfej.writeUInt16LE(nevBuf.length, 28)
    kfej.writeUInt32LE(eltolas, 42)
    kozponti.push(kfej, nevBuf)
    eltolas += fej.length + nevBuf.length + tomor.length
  }
  const kozpontiBuf = Buffer.concat(kozponti)
  const veg = Buffer.alloc(22)
  veg.writeUInt32LE(0x06054b50, 0)
  veg.writeUInt16LE(Object.keys(reszek).length, 8)
  veg.writeUInt16LE(Object.keys(reszek).length, 10)
  veg.writeUInt32LE(kozpontiBuf.length, 12)
  veg.writeUInt32LE(eltolas, 16)
  return Buffer.concat([...helyi, kozpontiBuf, veg])
}

const NS_SS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

// --- xlsx: MINDEN RESZ GENERALT -- nulla beagyazott bajt ---------------------------------------
function xlsxMag(lapnevek) {
  const nevek = (Array.isArray(lapnevek) && lapnevek.length ? lapnevek : ['Munka1']).map(String)
  if (nevek.length > 200) throw new Error(`xlsxMag: ${nevek.length} lap -- a felso hatar 200 (ennel a mag maga lenne a szuk keresztmetszet)`)
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const reszek = {
    '[Content_Types].xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + nevek.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '</Types>',
    '_rels/.rels':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}">`
      + `<Relationship Id="rId1" Type="${OD}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ${NS_SS} ${NS_R}><sheets>`
      + nevek.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
      + '</sheets></workbook>',
    'xl/_rels/workbook.xml.rels':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}">`
      + nevek.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${OD}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
      + `<Relationship Id="rId${nevek.length + 1}" Type="${OD}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet ${NS_SS}>`
      + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
      + '<borders count="1"><border/></borders>'
      + '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>',
  }
  nevek.forEach((_, i) => {
    reszek[`xl/worksheets/sheet${i + 1}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${NS_SS}><sheetData/></worksheet>`
  })
  return zipCsomag(reszek)
}

// --- pptx: a tema/mester/elrendezes BEAGYAZVA, a diak generalva --------------------------------
const URES_DIA =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
  + '<p:cSld><p:spTree>'
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
  + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
  + '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
  + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'

function pptxMag(diaszam = 1) {
  // *** A `Number(x) || 1` ALAK ITT HIBA VOLNA: a 0 hamis, tehat egy EXPLICIT 0 csendben 1-re
  // valtana -- vagyis a hivo nulla diat ker, es kap egyet. Sajat teszt fogta meg. ***
  const n = diaszam === undefined ? 1 : Number(diaszam)
  if (!Number.isFinite(n) || n < 1) throw new Error(`pptxMag: a diaszam legalabb 1 -- egy nulla dias bemutato ervenytelen csomag (kapott: ${JSON.stringify(diaszam)})`)
  if (n > 200) throw new Error(`pptxMag: ${n} dia -- a felso hatar 200`)
  const reszek = {}
  for (const [nev, b64] of Object.entries(PPTX_STATIKUS)) reszek[nev] = Buffer.from(b64, 'base64')
  reszek['[Content_Types].xml'] =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    + '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    + '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    + '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    + Array.from({ length: n }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')
    + '</Types>'
  reszek['ppt/presentation.xml'] =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    + `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId${n + 1}"/></p:sldMasterIdLst>`
    + '<p:sldIdLst>' + Array.from({ length: n }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join('') + '</p:sldIdLst>'
    + '<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
  reszek['ppt/_rels/presentation.xml.rels'] =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}">`
    + Array.from({ length: n }, (_, i) => `<Relationship Id="rId${i + 1}" Type="${OD}/slide" Target="slides/slide${i + 1}.xml"/>`).join('')
    + `<Relationship Id="rId${n + 1}" Type="${OD}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`
    + `<Relationship Id="rId${n + 2}" Type="${OD}/theme" Target="theme/theme1.xml"/></Relationships>`
  for (let i = 1; i <= n; i += 1) {
    reszek[`ppt/slides/slide${i}.xml`] = URES_DIA
    reszek[`ppt/slides/_rels/slide${i}.xml.rels`] = PPTX_DIA_RELS
  }
  return zipCsomag(reszek)
}

module.exports = { xlsxMag, pptxMag, zipCsomag, PPTX_STATIKUS }
