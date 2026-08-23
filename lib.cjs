const crypto = require('crypto')
const fs = require('fs')

// This module's own mtime, captured ONCE at require() time -- the moment this exact code started
// running in this process. A long-lived MCP child process keeps running whatever was on disk when
// it started; if a later `git pull`/merge/checkout changes this file, the RUNNING process never
// notices on its own (a measured case produced two interleaved sheets written into one
// sheet1.xml by a stale lib.cjs). checkFreshness() re-stats the file on every call and compares
// against this constant -- if the disk copy is now newer than what this process loaded, the code
// that is ABOUT to run is not the code on disk. Deliberately a SIGNAL, not a fix: this module does
// not reload or restart itself -- a caller decides what to do with the answer.
const __LIB_LOAD_MTIME_MS = fs.statSync(__filename).mtimeMs
function checkFreshness() {
  const diskMtimeMs = fs.statSync(__filename).mtimeMs
  if (diskMtimeMs <= __LIB_LOAD_MTIME_MS) return { fresh: true }
  return {
    fresh: false,
    file: 'lib.cjs',
    message:
      `lib.cjs: a betoltott kod regebbi, mint a lemezen levo (betoltve: ${new Date(__LIB_LOAD_MTIME_MS).toISOString()}, ` +
      `lemezen: ${new Date(diskMtimeMs).toISOString()}) -- inditsd ujra a sessiont`,
  }
}

// EURO-MCP core. Everything here is pure and independently testable: no network, no secrets
// read from disk, no MCP transport. The server (euro-mcp.js) wires these into tools; the
// tests (test-lib.js) prove them without a Document Server.
//
// WHY THERE IS NO open/edit/save TRIPLE: the DocBuilder API has no session that survives an HTTP
// call. One POST carries one
// whole script, which the server runs start to finish and then discards. Three separate tools
// would each have to open the file again, so `open_document` could not hand anything to
// `edit_document` - the name would promise state the protocol does not keep. The honest split
// is by what a caller actually decides between: ask the service what it can do
// (`service_status`), do a document edit (`edit_document`), or drive the builder directly for
// something the convenience tool does not cover (`run_builder_script`).

function b64url(buf) {
  return Buffer.from(buf).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_')
}

// HS256, per the Document Server contract as measured on our experimental instance: the same
// token goes in the request body as `token` AND in the Authorization header. The secret is a
// parameter rather than module state so a test can sign with a known key, and so no code path
// can accidentally reach for a global.
function signJwt(secret, payload) {
  if (!secret) throw new Error('signJwt: missing secret')
  const header = b64url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }))
  const body = b64url(JSON.stringify(payload))
  const signature = b64url(crypto.createHmac('sha256', secret).update(`${header}.${body}`).digest())
  return `${header}.${body}.${signature}`
}

// The marker proves the document actually changed, and it is per-call for a reason. With a
// constant marker the check would pass in two cases where nothing happened: when the input
// document already contains it (feeding a previous output back in), and when the service
// returns a cached result of an earlier run. Both look exactly like a successful edit.
function makeMarker() {
  return `EURO-MCP-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`
}

// Text destined for a .docbuilder script is JS source on the other side, so it goes through
// JSON.stringify rather than manual quoting: a caller-supplied string containing a quote or a
// newline would otherwise end the statement early and produce a script that fails in a way
// that reads like a Document Server fault.
function jsString(value) {
  return JSON.stringify(String(value))
}

// `hyperlink: { url, tooltip? }` -- shared shape by the
// `shape` op (whole-shape click target, Api.CreateHyperlink + Shape.SetHyperlink) and the `runs`
// op's per-run entries (text-run click target, Run.AddHyperlink). Both call shapes recovered via
// toString() on the live DocBuilder instance, package-verified: a shape-level hyperlink lands as
// `<a:hlinkClick>` inside the shape's `<p:cNvPr>`; a run-level one lands inside that run's
// `<a:rPr>`, and the builder auto-adds `u="sng"` (underline) to the run -- not requested, the
// engine's own doing. Both register a real External relationship in slide1.xml.rels. A bare
// domain without a protocol (e.g. "example.com/x") is NOT rejected -- Run.AddHyperlink's own
// source silently prepends "http://" (package-verified) -- so this validation only requires a
// non-empty string, it does not itself enforce a URL shape.
function validateHyperlink(hyperlink, context) {
  if (!hyperlink || typeof hyperlink !== 'object') {
    throw new Error(`${context}: hyperlink must be an object with a \`url\` field`)
  }
  if (typeof hyperlink.url !== 'string' || !hyperlink.url.trim()) {
    throw new Error(`${context}: hyperlink.url must be a non-empty string`)
  }
  const tooltip = hyperlink.tooltip !== undefined && hyperlink.tooltip !== null ? String(hyperlink.tooltip) : ''
  return { url: hyperlink.url, tooltip }
}

// `Api.CreateColorFromRGB(r, g, b)` argument list, shared by the `fillColor`/`fontColor`
// operations below. `color` is validated separately (validateRgbColor) before this ever runs --
// this only formats.
function rgbArg(color) {
  return color.map(Number).join(', ')
}

// A bad color does not throw inside the builder either -- it either coerces to something
// unrelated or is silently dropped, so this is refused at the boundary this tool owns, same
// shape as columnWidth's negative-width refusal.
function validateRgbColor(opName, color) {
  if (!Array.isArray(color) || color.length !== 3) {
    throw new Error(`${opName}: \`color\` must be an [r, g, b] array of three numbers -- got ${JSON.stringify(color)}`)
  }
  for (const c of color) {
    if (!Number.isFinite(c) || c < 0 || c > 255) {
      throw new Error(`${opName}: \`color\` values must each be 0-255 -- got ${JSON.stringify(color)}`)
    }
  }
}

// The complete OOXML ST_PresetPatternVal set (DrawingML spec -- a fixed enum, not an
// implementation detail like the pptx transition-effect names elsewhere in this file). MEASURED
// that an unrecognised value does NOT throw (Api.CreatePatternFill's own
// source has no validation): the saved package still gets an <a:pattFill> element, just with NO
// `prst` attribute at all -- a silently different (and likely blank/invalid) fill, not an error.
// Refused client-side for the same reason as every other silent-drop field in this file.
const KNOWN_PATTERN_TYPES = ['pct5', 'pct10', 'pct20', 'pct25', 'pct30', 'pct40', 'pct50', 'pct60', 'pct70', 'pct75', 'pct80', 'pct90', 'horz', 'vert', 'ltHorz', 'ltVert', 'dkHorz', 'dkVert', 'narHorz', 'narVert', 'dashHorz', 'dashVert', 'cross', 'dnDiag', 'upDiag', 'ltDnDiag', 'ltUpDiag', 'dkDnDiag', 'dkUpDiag', 'wdDnDiag', 'wdUpDiag', 'dashDnDiag', 'dashUpDiag', 'diagCross', 'smCheck', 'lgCheck', 'smGrid', 'lgGrid', 'dotGrid', 'smConfetti', 'lgConfetti', 'horzBrick', 'diagBrick', 'solidDmnd', 'openDmnd', 'dotDmnd', 'plaid', 'sphere', 'weave', 'divot', 'shingle', 'wave', 'trellis', 'zigZag']

// A single fill-expression builder shared by every operation that needs one -- `shape`'s own
// `fill` field (extended here, backward compatible: a bare [r,g,b] array still means solid
// color, exactly as before) and the new `wordArt` operation below.
// MEASURED live: CreateLinearGradientFill/CreateRadialGradientFill/CreatePatternFill
// all package-verify correctly (a:gradFill with the right stops/angle, a:pattFill with the right
// prst) -- not taken from documentation.
//
// `spec` shapes accepted:
//   undefined/null                                    -> Api.CreateNoFill()
//   [r, g, b]                                          -> Api.CreateSolidFill (legacy/default form)
//   { type: 'solid', color: [r,g,b] }                  -> same as above, explicit
//   { type: 'gradient', shape: 'linear'|'radial', stops: [{color:[r,g,b], pos}], angle? }
//     `pos` is 0-100 (percent along the gradient) -- Api.CreateGradientStop's own pos is
//     0-100000 (measured: it clamps rather than throws on out-of-range), this multiplies by 1000
//     so callers use the more familiar percent scale. `angle` is DEGREES for the caller
//     (converted to the API's 60000ths-of-a-degree unit here) -- linear only, MEASURED absent
//     (unused) on radial in this unit's own probe.
//   { type: 'pattern', patternType, bgColor: [r,g,b], fgColor: [r,g,b] }
function buildGradientStopExpr(stop, idx, opName) {
  if (!stop || typeof stop !== 'object') throw new Error(`${opName}: fill.stops[${idx}] must be an object`)
  validateRgbColor(`${opName}: fill.stops[${idx}]`, stop.color)
  const pos = Number(stop.pos)
  if (!Number.isFinite(pos) || pos < 0 || pos > 100) throw new Error(`${opName}: fill.stops[${idx}].pos must be 0-100, got ${JSON.stringify(stop.pos)}`)
  return `Api.CreateGradientStop(Api.CreateRGBColor(${rgbArg(stop.color)}), ${Math.round(pos * 1000)})`
}

function buildGradientFillExpression(spec, opName) {
  if (!Array.isArray(spec.stops) || spec.stops.length < 2) {
    throw new Error(`${opName}: fill.stops must be an array of at least 2 {color, pos} entries, got ${JSON.stringify(spec.stops)}`)
  }
  const stopsExpr = `[${spec.stops.map((stop, idx) => buildGradientStopExpr(stop, idx, opName)).join(', ')}]`
  if (spec.shape === 'radial') return `Api.CreateRadialGradientFill(${stopsExpr})`
  if (spec.shape !== 'linear' && spec.shape !== undefined && spec.shape !== null) {
    throw new Error(`${opName}: fill.shape must be "linear" or "radial", got ${JSON.stringify(spec.shape)}`)
  }
  const angleDeg = Number(spec.angle ?? 0)
  if (!Number.isFinite(angleDeg)) throw new Error(`${opName}: fill.angle must be a number (degrees), got ${JSON.stringify(spec.angle)}`)
  return `Api.CreateLinearGradientFill(${stopsExpr}, ${Math.round(angleDeg * 60000)})`
}

function buildPatternFillExpression(spec, opName) {
  if (!KNOWN_PATTERN_TYPES.includes(spec.patternType)) {
    throw new Error(`${opName}: unknown fill.patternType ${JSON.stringify(spec.patternType)} (known: ${KNOWN_PATTERN_TYPES.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently produces a pattern fill with no preset, so this tool refuses it here instead`)
  }
  validateRgbColor(`${opName}: fill.bgColor`, spec.bgColor)
  validateRgbColor(`${opName}: fill.fgColor`, spec.fgColor)
  return `Api.CreatePatternFill(${jsString(spec.patternType)}, Api.CreateRGBColor(${rgbArg(spec.bgColor)}), Api.CreateRGBColor(${rgbArg(spec.fgColor)}))`
}

function buildFillExpression(spec, opName) {
  if (spec === undefined || spec === null) return 'Api.CreateNoFill()'
  if (Array.isArray(spec)) {
    validateRgbColor(opName, spec)
    return `Api.CreateSolidFill(Api.CreateRGBColor(${rgbArg(spec)}))`
  }
  if (typeof spec !== 'object') {
    throw new Error(`${opName}: fill must be an [r,g,b] array or a {type:...} object, got ${JSON.stringify(spec)}`)
  }
  if (spec.type === 'solid') {
    validateRgbColor(opName, spec.color)
    return `Api.CreateSolidFill(Api.CreateRGBColor(${rgbArg(spec.color)}))`
  }
  if (spec.type === 'gradient') return buildGradientFillExpression(spec, opName)
  if (spec.type === 'pattern') return buildPatternFillExpression(spec, opName)
  throw new Error(`${opName}: fill.type must be "solid"/"gradient"/"pattern", got ${JSON.stringify(spec.type)}`)
}

// A tagged Error, not a new class: needed to tell apart a
// capability that genuinely does not exist here (NEM-TAMOGATOTT -- an unknown operation type, an
// operation type this core cannot do, or a value outside a MEASURED allowlist like chartType/
// align/highlight/listType) from a malformed call to a capability that DOES exist (HIBA -- a
// missing required field, an out-of-range number, a bad string format). The message text is
// UNCHANGED at every site that uses this -- only a `notSupported` flag is added, so nothing that
// already matches these messages (existing tests, callers) sees a different string.
function notSupportedError(message) {
  const err = new Error(message)
  err.notSupported = true
  return err
}

// The exact 24 names Api.GetThemesColors() returns on this
// DocBuilder instance, measured live -- not the OOXML spec's list, not guessed. `sheetTheme`
// validates against this so an unrecognized name is a NAMED error, not the silent no-op it is
// inside the builder itself (see that entry's own comment).
const KNOWN_THEME_NAMES = [
  'Aspect', 'Blue Green', 'Blue II', 'Blue Warm', 'Blue', 'Grayscale', 'Green Yellow', 'Green',
  'Marquee', 'Median', 'Office 2007 - 2010', 'Office 2013 - 2022', 'Office', 'Orange Red',
  'Orange', 'Paper', 'Red Orange', 'Red Violet', 'Red', 'Slipstream', 'Violet II', 'Violet',
  'Yellow Orange', 'Yellow',
]

// Builds the script for an edit round-trip: open the existing document, apply the operations,
// save under a new name, close. OpenFile (not CreateFile) is deliberate - upstream #321 reports
// CreateFile failing on missing templates, and that is not the path this tool offers.
function buildEditScript({ docUrl, operations, outName = 'eredmeny.docx', marker }) {
  if (!docUrl) throw new Error('buildEditScript: docUrl required')
  const ops = Array.isArray(operations) ? operations : []
  if (!ops.length && !marker) throw new Error('buildEditScript: no operations and no marker - the script would be a no-op')

  const lines = [
    `builder.OpenFile(${jsString(docUrl)}, "docx");`,
    'var oDocument = Api.GetDocument();',
  ]

  for (const op of ops) {
    if (op.type === 'append_paragraph') {
      lines.push(
        'var oParagraph = Api.CreateParagraph();',
        `oParagraph.AddText(${jsString(op.text ?? '')});`,
        'oDocument.Push(oParagraph);',
      )
    } else if (op.type === 'replace_text') {
      // SearchAndReplace is a document-level call in the builder API; it needs no paragraph.
      lines.push(`oDocument.SearchAndReplace({ searchString: ${jsString(op.search ?? '')}, replaceString: ${jsString(op.replace ?? '')} });`)
    } else {
      throw new Error(`buildEditScript: unknown operation type: ${op.type}`)
    }
  }

  // The marker paragraph is what the content check looks for afterwards. It is appended last so
  // that a replace operation cannot rewrite it.
  if (marker) {
    lines.push(
      'var oMarkerParagraph = Api.CreateParagraph();',
      `oMarkerParagraph.AddText(${jsString(marker)});`,
      'oDocument.Push(oMarkerParagraph);',
    )
  }

  lines.push(`builder.SaveFile("docx", ${jsString(outName)});`, 'builder.CloseFile();')
  return lines.join('\n') + '\n'
}

// Four outcomes, and none of them collapses into another. A caller that cannot tell "the
// licence gate is shut" from "my script was wrong" from "the service never reached my file"
// will debug the wrong layer; that is exactly what cost the first live run its morning.
const OUTCOME = {
  OK: 'ok',
  BLOCKED: 'blocked',
  AUTH: 'auth',
  FETCH: 'fetch',
  UNKNOWN: 'unknown',
}

function classifyDsResponse(body) {
  if (!body || typeof body !== 'object') {
    return { outcome: OUTCOME.UNKNOWN, detail: 'the response was not a JSON object' }
  }

  // A successful docbuilder run answers with the output urls; `error` is absent or 0.
  if (body.urls && typeof body.urls === 'object') {
    return { outcome: OUTCOME.OK, urls: body.urls, key: body.key ?? null }
  }

  const code = body.error
  if (code === -3) {
    // -3 has TWO causes and the response does not say which. Measured live
    // 2026-08-14: a single non-existent method (SetTextSpacing) returned -3 while the service
    // was working perfectly two minutes either side -- nine API calls tried one by one, EIGHT
    // went through and one gave -3. A shut advanced_api gate would have failed all nine.
    //
    // Naming only the licence cause is worse than naming neither: it sends the caller to the
    // instance configuration, which is the one layer they cannot fix and where they will find
    // nothing wrong. The two causes are separable by the caller, so the message says how.
    return {
      outcome: OUTCOME.BLOCKED,
      detail:
        'error -3 has two possible causes and the response does not distinguish them: either the ' +
        'docbuilder API is not available on this instance (advanced_api gate), or the submitted ' +
        'script itself is faulty -- a single unknown method is enough to produce this. ' +
        'To tell them apart, run a MINIMAL known-good script: if that one succeeds, the gate is ' +
        'open and the fault is in the script.',
    }
  }
  if (code === -8 || code === 6) {
    return { outcome: OUTCOME.AUTH, detail: `the Document Server rejected the token (error ${code})` }
  }
  if (code === -4) {
    return {
      outcome: OUTCOME.FETCH,
      detail:
        'the Document Server could not download the script or the input document (error -4). ' +
        'It fetches both itself, so both URLs must be reachable FROM the server.',
    }
  }
  if (code === 0) {
    // error:0 with no urls means it ran and produced nothing we can fetch - not a success.
    return { outcome: OUTCOME.UNKNOWN, detail: 'the service reported success but returned no output urls' }
  }
  return { outcome: OUTCOME.UNKNOWN, detail: `unrecognised docbuilder error code: ${JSON.stringify(code)}` }
}

// A docx is a zip; the text lives in word/document.xml. Checking for the marker there rather
// than in the raw bytes matters because the raw file is compressed - a substring search over
// the container would miss a marker that is genuinely present.
function markerInDocumentXml(documentXml, marker) {
  if (!documentXml || !marker) return false
  // The builder may split a run across XML elements, so compare on text content with tags
  // stripped rather than on the literal serialisation.
  const text = String(documentXml).replace(/<[^>]*>/g, '')
  return text.includes(marker)
}

// The five named XML entities OOXML text can carry. Deliberately NOT a general
// entity/numeric-character-reference decoder (no &#160; / &amp;#x2019; handling) -- the builder's
// own output only ever escapes these five (measured on live save output across several
// diagnostic rounds), and a general decoder invites a false sense of completeness for content this
// function never actually has to handle: text typed by a THIRD PARTY inside the live editor
// (not this tool's own writes) could carry other entities, and a caller reading office.get_text
// output must not assume they were resolved.
function decodeXmlEntities(s) {
  return String(s ?? '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

// Splits word/document.xml into paragraphs and returns each paragraph's own
// text (its own <w:t> runs joined, tags/entities stripped) -- the building block office_get_text
// needs, and NOT the same job as markerInDocumentXml above (that one deliberately throws every
// paragraph break away, because "is this marker anywhere in the file" doesn't care where). A
// docx paragraph boundary is <w:p ...>...</w:p>; a run's text is the FIRST-level <w:t> inside it,
// but this also has to skip <w:t> runs that sit inside DELETED tracked-change content
// (<w:del>...<w:delText>) -- SetTrackRevisions(true) writes deletions as <w:delText>, not <w:t>,
// so a plain global <w:t> match already excludes them; nothing extra to filter (measured: a
// tracked deletion produced zero <w:t> matches inside its span).
function paragraphTextsFromDocumentXml(documentXml) {
  const xml = String(documentXml ?? '')
  const paragraphs = xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || []
  return paragraphs.map((p) => {
    const runs = p.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || []
    return runs
      .map((r) => decodeXmlEntities(r.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '')))
      .join('')
  })
}

// word/comments.xml -> [{id, author, date, text}], the shape office_get_comments
// reports. `id` here is the OOXML w:id attribute (a small sequential integer PER DOCUMENT) --
// this is a DIFFERENT number from what the live editor's own Comment.GetId() returns (measured:
// w:id="0" in the saved XML, GetId() answered "1928914938" in the same session).
// The two are not interchangeable; a caller matching a comment across the two paths must not
// assume either numbering is the other's.
function commentsFromCommentsXml(commentsXml) {
  const xml = String(commentsXml ?? '')
  const items = xml.match(/<w:comment\s[^>]*>[\s\S]*?<\/w:comment>/g) || []
  return items.map((c) => {
    const id = (c.match(/\bw:id="([^"]*)"/) || [])[1] ?? null
    const author = decodeXmlEntities((c.match(/\bw:author="([^"]*)"/) || [])[1] ?? '')
    const date = (c.match(/\bw:date="([^"]*)"/) || [])[1] ?? null
    const runs = c.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || []
    const text = runs
      .map((r) => decodeXmlEntities(r.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '')))
      .join('')
    return { id, author, date, text }
  })
}

// word/footnotes.xml -> [{id, text}], a sibling of
// commentsFromCommentsXml above -- same "extract runs, join, decode" body, but with an EXTRA
// filter step that comments.xml does not need: EVERY docx has two BOILERPLATE footnote entries
// (`<w:footnote w:type="separator" w:id="-1">` and `<w:footnote w:type="continuationSeparator"
// w:id="0">`), present even in a document with zero real footnotes -- measured live (a fixture
// with two AddFootnote() calls landed real content at w:id="2"/"3", never at -1/0). A REAL
// footnote's own `<w:footnote>` tag carries NO `w:type` attribute at all; that absence is the
// filter, not an id-number heuristic (id numbering is not guaranteed to start at any fixed
// value once separators exist). Endnotes share the identical structure one level up
// (word/endnotes.xml, `<w:endnote>`) -- endnotesFromEndnotesXml below is this same body with
// the two tag names swapped, not a differently-designed sibling.
function footnotesFromFootnotesXml(footnotesXml) {
  const xml = String(footnotesXml ?? '')
  const items = xml.match(/<w:footnote\b[^>]*>[\s\S]*?<\/w:footnote>/g) || []
  return items
    .filter((f) => !/\bw:type="/.test(f.match(/<w:footnote\b[^>]*>/)[0]))
    .map((f) => {
      const id = (f.match(/\bw:id="([^"]*)"/) || [])[1] ?? null
      const runs = f.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || []
      const text = runs
        .map((r) => decodeXmlEntities(r.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '')))
        .join('')
      return { id, text }
    })
}

// word/endnotes.xml -> [{id, text}] -- see footnotesFromFootnotesXml above for the shared design
// note (the boilerplate-filter reasoning applies identically here).
function endnotesFromEndnotesXml(endnotesXml) {
  const xml = String(endnotesXml ?? '')
  const items = xml.match(/<w:endnote\b[^>]*>[\s\S]*?<\/w:endnote>/g) || []
  return items
    .filter((e) => !/\bw:type="/.test(e.match(/<w:endnote\b[^>]*>/)[0]))
    .map((e) => {
      const id = (e.match(/\bw:id="([^"]*)"/) || [])[1] ?? null
      const runs = e.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || []
      const text = runs
        .map((r) => decodeXmlEntities(r.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '')))
        .join('')
      return { id, text }
    })
}

// word/document.xml -> [{name, text}], the read side of the
// `bookmark`/`bookmarkRef` write fields. `name` comes straight off
// `<w:bookmarkStart w:name="...">`. `text` is the SPAN between that bookmarkStart and its
// matching `<w:bookmarkEnd w:id="...">` (matched by the shared `w:id`, NOT by document order --
// measured live: a bookmark can wrap a run that is not the immediately
// next element) -- `null` if the matching end tag is not found (a malformed or externally-edited
// package, not this reader's failure to try). This is the read equivalent of
// `Api.GetAllBookmarksNames()` (the name list) folded together with a single `GetBookmark(name)`
// per entry (the text) -- one pass over the XML rather than a name-list call followed by N
// separate per-bookmark reads, since the source data (one linear document.xml) does not actually
// support "look up bookmark by name" any cheaper than "extract every bookmark's span while
// already scanning".
function bookmarksFromDocumentXml(documentXml) {
  const xml = String(documentXml ?? '')
  const starts = xml.match(/<w:bookmarkStart\b[^>]*\/?>/g) || []
  return starts
    .map((startTag) => {
      const id = (startTag.match(/\bw:id="([^"]*)"/) || [])[1]
      const name = (startTag.match(/\bw:name="([^"]*)"/) || [])[1]
      if (id === undefined || name === undefined) return null
      const startIdx = xml.indexOf(startTag)
      const endTagPattern = new RegExp(`<w:bookmarkEnd\\b[^>]*\\bw:id="${id}"[^>]*/?>`)
      const rest = xml.slice(startIdx + startTag.length)
      const endMatch = rest.match(endTagPattern)
      if (!endMatch) return { name: decodeXmlEntities(name), text: null }
      const span = rest.slice(0, endMatch.index)
      const runs = span.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || []
      const text = runs
        .map((r) => decodeXmlEntities(r.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '')))
        .join('')
      return { name: decodeXmlEntities(name), text }
    })
    .filter(Boolean)
}

// word/document.xml -> a FLAT, depth-aware list of every `<w:tbl>`
// (top-level AND nested -- the write-side mechanism this reads back). Depth-aware because
// a naive non-greedy `<w:tbl>...</w:tbl>` regex breaks on nesting (it would close at the FIRST
// `</w:tbl>` it sees, which belongs to the innermost table, not the one the regex thinks it
// matched) -- this walks the string with an explicit open/close counter instead, so each table's
// own span is bounded correctly regardless of how deep it nests. `parentIndex` is the enclosing
// table's own index in this SAME returned array (null for a top-level table) -- a caller wanting
// "just this table's own rows, not a nested one's" filters by `parentIndex === null` OR by depth.
// Cell text is EVERY `<w:t>` inside that cell's own `<w:tc>...</w:tc>` span, EXCLUDING any nested
// `<w:tbl>` -- checked live: without stripping the inner table first, an outer cell containing
// only a nested table (no text of its own) reported the INNER table's own cell text as if it
// were the outer cell's, because a flat `<w:t>` scan cannot tell "this run is mine" from "this
// run belongs to a table nested three levels down inside me". A nested table's own text is
// available at its own array entry (linked via `parentIndex`) -- reporting it twice, once on the
// parent and once on the child, would not be a richer read, it would be the SAME data at two
// indices with no way to tell that from the shape alone.
function tablesFromDocumentXml(documentXml) {
  const xml = String(documentXml ?? '')
  const tagPattern = /<w:tbl\b[^>]*>|<\/w:tbl>/g
  const spans = []
  const openStarts = []
  let m
  while ((m = tagPattern.exec(xml))) {
    if (m[0].startsWith('<w:tbl')) {
      openStarts.push(m.index)
    } else {
      const start = openStarts.pop()
      if (start !== undefined) spans.push({ start, end: m.index + m[0].length })
    }
  }
  // Closed innermost-first (stack pop order); sort by START so a table's index reflects DOCUMENT
  // order, and so parent-lookup below (smallest span that strictly contains this one) has a
  // stable array to search.
  spans.sort((a, b) => a.start - b.start)
  // Precompute every span's parent FIRST (needed below to find DIRECT children only -- removing
  // just the direct children, each a self-contained range that already carries its own nested
  // descendants inside it, avoids the multi-level offset bookkeeping that removing "any strictly
  // contained span" at once would require).
  const parentIndexOf = spans.map((s, i) => {
    let parentIndex = null
    let tightest = Infinity
    spans.forEach((candidate, j) => {
      if (j === i) return
      const contains = candidate.start < s.start && candidate.end > s.end
      if (contains && candidate.end - candidate.start < tightest) {
        tightest = candidate.end - candidate.start
        parentIndex = j
      }
    })
    return parentIndex
  })
  return spans.map((s, i) => {
    const directChildren = spans.filter((_, j) => parentIndexOf[j] === i)
    let tblXml = xml.slice(s.start, s.end)
    // Remove highest-offset first so an earlier (lower-offset, not-yet-processed) child's
    // position is never shifted by a later removal.
    directChildren
      .slice()
      .sort((a, b) => b.start - a.start)
      .forEach((child) => {
        tblXml = tblXml.slice(0, child.start - s.start) + tblXml.slice(child.end - s.start)
      })
    const rows = tblXml.match(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g) || []
    const sorok = rows.map((row) => {
      const cells = row.match(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g) || []
      return cells.map((cell) => {
        const runs = cell.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g) || []
        return runs
          .map((r) => decodeXmlEntities(r.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '')))
          .join('')
      })
    })
    return { index: i, parentIndex: parentIndexOf[i], sorok }
  })
}

// ppt/slides/slideN.xml -> paragraph texts, the pptx sibling of
// paragraphTextsFromDocumentXml above. Same flattening choice as that function (every <a:p> in
// the slide's XML, INCLUDING ones inside a table cell's <a:txBody> -- not just top-level shape
// text) for the same reason: docx's own regex already does not distinguish table-cell paragraphs
// from body paragraphs, so this keeps the two cores' "read all the text" semantics consistent
// rather than inventing a narrower rule for pptx alone. Deliberately does NOT read
// ppt/notesSlides/notesSlideN.xml -- that is a SEPARATE part with its own <a:p>/<a:t> structure,
// speaker notes are not "the slide's text" by this function's own scope.
function paragraphTextsFromSlideXml(slideXml) {
  const xml = String(slideXml ?? '')
  const paragraphs = xml.match(/<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g) || []
  return paragraphs.map((p) => {
    const runs = p.match(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g) || []
    return runs
      .map((r) => decodeXmlEntities(r.replace(/^<a:t(?:\s[^>]*)?>/, '').replace(/<\/a:t>$/, '')))
      .join('')
  })
}

// ppt/presentation.xml's own <p:sldIdLst> gives the DISPLAY order as
// a sequence of r:id references, NOT the slideN.xml part numbers -- measured live via MoveTo
// verification: a moved slide keeps its ORIGINAL part
// filename, only its position in this list changes. Reading ppt/slides/slide1.xml, slide2.xml...
// in NAME order would silently report the pre-reorder sequence. This function resolves the
// r:id -> part-name mapping through ppt/_rels/presentation.xml.rels and returns the ordered list
// of part paths (e.g. "ppt/slides/slide3.xml") a caller can then fetch one by one.
function pptxSlideOrderFromPresentationXml(presentationXml, relsXml) {
  const pres = String(presentationXml ?? '')
  const rels = String(relsXml ?? '')
  const rIds = (pres.match(/<p:sldId\b[^>]*\br:id="([^"]*)"/g) || [])
    .map((tag) => (tag.match(/\br:id="([^"]*)"/) || [])[1])
    .filter(Boolean)
  const relEntries = rels.match(/<Relationship\b[^>]*\/?>/g) || []
  const targetById = new Map()
  relEntries.forEach((entry) => {
    const id = (entry.match(/\bId="([^"]*)"/) || [])[1]
    const target = (entry.match(/\bTarget="([^"]*)"/) || [])[1]
    if (id && target) targetById.set(id, target)
  })
  return rIds
    .map((rId) => targetById.get(rId))
    .filter(Boolean)
    .map((target) => `ppt/${target.replace(/^\.?\/?/, '')}`)
}

// ppt/slides/_rels/slideN.xml.rels -> the comments part that slide
// owns, if any (measured live: a "…/relationships/comments" relationship pointing at
// "../comments/commentM.xml", resolved relative to ppt/slides/ -- see this function's own path
// join). A slide with no comments has no such relationship at all (not an empty one) -- returns
// null in that case, not a made-up path.
function pptxCommentsPartFromSlideRelsXml(slideRelsXml) {
  const xml = String(slideRelsXml ?? '')
  const entries = xml.match(/<Relationship\b[^>]*\/?>/g) || []
  for (const entry of entries) {
    const type = (entry.match(/\bType="([^"]*)"/) || [])[1] ?? ''
    if (!/\/relationships\/comments$/.test(type)) continue
    const target = (entry.match(/\bTarget="([^"]*)"/) || [])[1]
    if (!target) continue
    // Target is relative to ppt/slides/ (e.g. "../comments/comment1.xml") -- this route only
    // ever measured the one-level-up form, so a bare `../` strip is enough here; a target that
    // does not start with it is returned unresolved rather than guessed at.
    if (!target.startsWith('../')) return null
    return `ppt/${target.slice(3)}`
  }
  return null
}

// ppt/comments/commentN.xml + ppt/commentAuthors.xml ->
// [{idx, authorId, authorName, date, text}] -- the pptx sibling of commentsFromCommentsXml above,
// but a DIFFERENT schema, measured live:
// pptx comments carry NO w:id-equivalent identity of their own beyond `idx` (a per-SLIDE sequence
// number, not a per-document one like docx's w:id), the author is a NUMERIC `authorId` pointing
// into the SEPARATE commentAuthors.xml part (not an inline name attribute), and the comment text
// is a single <p:text> plain-text node -- no run structure to join, unlike docx's <w:t> runs.
// `authorsXml` is optional: if it could not be read, `authorName` is null for every comment
// rather than the whole read failing (the comment TEXT is still real information on its own).
function commentsFromPptxCommentXml(commentXml, authorsXml) {
  const xml = String(commentXml ?? '')
  const authorNameById = new Map()
  const authorEntries = String(authorsXml ?? '').match(/<p:cmAuthor\b[^>]*\/?>/g) || []
  authorEntries.forEach((entry) => {
    const id = (entry.match(/\bid="([^"]*)"/) || [])[1]
    const name = (entry.match(/\bname="([^"]*)"/) || [])[1]
    if (id !== undefined) authorNameById.set(id, decodeXmlEntities(name ?? ''))
  })
  const items = xml.match(/<p:cm\b[^>]*>[\s\S]*?<\/p:cm>/g) || []
  return items.map((c) => {
    const authorId = (c.match(/\bauthorId="([^"]*)"/) || [])[1] ?? null
    const idx = (c.match(/\bidx="([^"]*)"/) || [])[1] ?? null
    const date = (c.match(/\bdt="([^"]*)"/) || [])[1] ?? null
    const textMatch = c.match(/<p:text>([\s\S]*?)<\/p:text>/)
    const text = textMatch ? decodeXmlEntities(textMatch[1]) : ''
    return {
      idx,
      authorId,
      authorName: authorId !== null && authorNameById.has(authorId) ? authorNameById.get(authorId) : null,
      date,
      text,
    }
  })
}

// The shared id/name/position/size a shape, picture, and graphicFrame (table or
// chart) all carry in the SAME xfrm shape, MEASURED live: <p:sp>/<p:pic> keep it
// under <p:spPr><a:xfrm>, <p:graphicFrame> keeps the SAME <a:off>/<a:ext> children directly under
// its own <p:xfrm> (no spPr wrapper) -- both forms handled here by matching the innermost
// <a:off>/<a:ext> pair regardless of which parent tag wraps it, since this function is always
// called with a snippet that already IS one such element (never the whole slide).
function pptxObjectXfrmFields(elementXml) {
  const id = (elementXml.match(/<p:cNvPr\s[^>]*\bid="([^"]*)"/) || [])[1] ?? null
  const name = decodeXmlEntities((elementXml.match(/<p:cNvPr\s[^>]*\bname="([^"]*)"/) || [])[1] ?? '')
  const off = elementXml.match(/<a:off\s[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/)
  const ext = elementXml.match(/<a:ext\s[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/)
  return {
    id,
    name,
    x: off ? Number(off[1]) : null,
    y: off ? Number(off[2]) : null,
    cx: ext ? Number(ext[1]) : null,
    cy: ext ? Number(ext[2]) : null,
  }
}

// ppt/slides/slideN.xml -> what is ALREADY on the slide -- a slide-content read-back covering
// what Slide.GetAllShapes/GetAllCharts/GetAllTables/GetAllImages/GetAllOleObjects would report,
// MEASURED as a shared xfrm shape rather than called live. Read from the
// SAVED PACKAGE, same architecture as paragraphTextsFromSlideXml above, not a live DocBuilder
// call -- a live GetAllShapes() etc. would return API object handles that still need their own
// introspection calls to get id/position/size, and this route already has that data in the XML.
//
// A <p:graphicFrame> is EITHER a table or a chart, discriminated by its <a:graphicData uri=...>
// (MEASURED: table -> ".../drawingml/2006/table", chart -> ".../drawingml/2006/chart") -- one
// single regex over all graphicFrame elements, then split by that URI, rather than two separate
// element-level regexes that would each have to re-derive "is this the OTHER kind, skip it".
//
// *** NAMED, NOT ATTEMPTED: OLE-object detection. *** This unit has no bound operation that
// creates an OLE object (Api.CreateOleObject is K8 scope, unbound as of this commit), so there
// was nothing to measure the <p:graphicFrame> OLE uri against -- returning an empty `oleObjects`
// array here would look exactly like a MEASURED zero, which it is not. The caller-facing tool is
// expected to report this field as null with a nemMertIndok, not silently omit the caveat.
//
// A <p:grpSp> (group) is reported by id/name only -- objects NESTED inside it also match this
// function's own shape/pic/graphicFrame patterns (this function does not respect XML nesting,
// same simplification paragraphTextsFromSlideXml already makes for table-cell paragraphs): a
// shape inside a group is still counted as a real shape on the slide, just not marked as grouped.
//
// *** NAMED, NOT CORRECTED: a shape's x/y INSIDE a group is in the GROUP'S OWN child coordinate
// space (<p:grpSpPr><a:xfrm>'s <a:chOff>/<a:chExt>), not the slide's absolute space -- MEASURED
// live (group-probe with nested GroupDrawings): a group at slide
// position (100000,100000) with chOff (0,0) contained a shape reporting x/y (0,0), which is
// nowhere near where it actually renders. This function does not walk parent groups to compute a
// true absolute position -- a caller comparing a grouped shape's x/y against a slide-level
// position (e.g. "is this shape inside the header area?") will get the WRONG answer for grouped
// shapes without correcting for the enclosing group's own xfrm. Not attempted here (chOff/chExt
// can differ in scale from off/ext, i.e. groups can be non-uniformly scaled, which would need a
// real coordinate transform, not just an offset add) -- ungrouped shapes are unaffected. ***
function slideContentSummaryFromSlideXml(slideXml) {
  const xml = String(slideXml ?? '')
  const shapes = (xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []).map((el) => ({
    ...pptxObjectXfrmFields(el),
    shapeType: (el.match(/<a:prstGeom\s[^>]*\bprst="([^"]*)"/) || [])[1] ?? null,
  }))
  const images = (xml.match(/<p:pic>[\s\S]*?<\/p:pic>/g) || []).map(pptxObjectXfrmFields)
  const groups = (xml.match(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g) || []).map((el) => {
    const f = pptxObjectXfrmFields(el)
    return { id: f.id, name: f.name }
  })
  const frames = xml.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g) || []
  const tables = []
  const charts = []
  frames.forEach((el) => {
    const uri = (el.match(/<a:graphicData\s[^>]*\buri="([^"]*)"/) || [])[1] ?? ''
    const fields = pptxObjectXfrmFields(el)
    if (uri.endsWith('/table')) tables.push(fields)
    else if (uri.endsWith('/chart')) charts.push(fields)
    // Any other graphicData uri (OLE included) is neither counted nor dropped silently -- it
    // simply does not appear in `tables` or `charts`, which is correct (it is genuinely neither).
  })
  return { shapes, images, tables, charts, groups }
}

// docProps/core.xml + docProps/app.xml -> the "dokumentum-
// metaadat" item (Presentation.GetDocumentInfo/GetCustomProperties). *** THE LIVE API ROUTE WAS
// MEASURED BROKEN, NOT SKIPPED: *** `oPresentation.GetDocumentInfo()` THROWS on this Document
// Server instance ("Cannot read property 'asc_getApplication' of null", K7 own probe) -- the
// static-package route below is not a shortcut around a working live call, it is the ONLY
// working route measured for this data. `docProps/` is a package-level path shared by docx/xlsx/
// pptx alike (not under word/ppt/xl/) -- this function itself is core-agnostic, callable for any
// of the three if a future unit wires it up there; THIS unit only calls it for pptx (its own
// card's scope).
//
// A field absent from the XML (no such tag at all) is null; a field present but EMPTY
// (`<cp:lastModifiedBy></cp:lastModifiedBy>`, MEASURED -- the default state of a document this
// Document Server saves without ever setting an author) is an empty string, not null -- the two
// are different facts ("never set" vs "explicitly cleared") and this function does not collapse
// them into one.
//
// *** NAMED, NOT COVERED: GetCustomProperties/GetCustomXmlParts. *** Reflected live (K7 probe):
// GetCustomProperties returns a live wrapper object ({CustomProperties:{properties:[]}}), not raw
// data, and the fixture available at implementation time had zero custom properties set to
// measure the POPULATED shape against -- there is no `Api.Set*`-style call on Presentation to
// create one from a DocBuilder script either (reflected: no such method exists). GetCustomXmlParts
// is a rarer OOXML feature (arbitrary embedded XML parts) not attempted at all. Both remain
// genuinely open, not silently folded into this function's "null means absent" convention.
function documentMetadataFromCoreAndAppXml(coreXml, appXml) {
  const core = String(coreXml ?? '')
  const app = String(appXml ?? '')
  const coreField = (tag) => {
    const m = core.match(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`))
    return m ? decodeXmlEntities(m[1]) : null
  }
  const appNumberField = (tag) => {
    const m = app.match(new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`))
    return m ? Number(m[1]) : null
  }
  return {
    title: coreField('dc:title'),
    subject: coreField('dc:subject'),
    creator: coreField('dc:creator'),
    lastModifiedBy: coreField('cp:lastModifiedBy'),
    created: coreField('dcterms:created'),
    modified: coreField('dcterms:modified'),
    revision: coreField('cp:revision'),
    category: coreField('cp:category'),
    application: (app.match(/<Application>([\s\S]*?)<\/Application>/) || [])[1] ?? null,
    slides: appNumberField('Slides'),
    hiddenSlides: appNumberField('HiddenSlides'),
    notes: appNumberField('Notes'),
    words: appNumberField('Words'),
    totalTime: appNumberField('TotalTime'),
  }
}

// A generic OOXML zip-relative path resolver -- "../x/y.z" or
// "x/y.z" against the DIRECTORY the rels file's OWN part lives in. Needed because the layout/
// theme resolution chain below crosses THREE different rels files, each relative to a DIFFERENT
// directory (ppt/slides -> ppt/slideLayouts -> ppt/slideMasters -> ppt/theme) -- MEASURED live,
// not assumed: a slide's rels use "../slideLayouts/slideLayout1.xml" (one level up), while
// presentation.xml.rels (pptxSlideOrderFromPresentationXml above) uses bare "slides/slide1.xml"
// (no leading "../") for the SAME kind of reference, because IT already lives at ppt/ itself.
function resolveOoxmlRelativeTarget(fromDir, target) {
  const stack = String(fromDir ?? '').split('/').filter(Boolean)
  for (const part of String(target ?? '').split('/')) {
    if (part === '..') stack.pop()
    else if (part === '.' || part === '') continue
    else stack.push(part)
  }
  return stack.join('/')
}

// The Target of the FIRST relationship in a rels XML whose Type ends
// in `typeSuffix` (e.g. "/slideLayout", "/slideMaster", "/theme") -- the same discrimination
// pptxCommentsPartFromSlideRelsXml already does for "/comments", generalised so the layout/theme
// chain below does not need three near-duplicate single-purpose functions.
function ooxmlRelationshipTargetByType(relsXml, typeSuffix) {
  const entries = String(relsXml ?? '').match(/<Relationship\b[^>]*\/?>/g) || []
  for (const entry of entries) {
    const type = (entry.match(/\bType="([^"]*)"/) || [])[1] ?? ''
    if (!type.endsWith(typeSuffix)) continue
    return (entry.match(/\bTarget="([^"]*)"/) || [])[1] ?? null
  }
  return null
}

// "elrendezes/tema lekerdezese" (Slide.GetLayout/GetTheme). MEASURED: the live calls do NOT
// throw and return real ApiLayout/ApiTheme handles, but
// those are live object handles needing further introspection -- the SAME data is already present
// as static XML, reachable through a rels chain: slide -> layout (own type/name attributes) ->
// (layout's rels) -> master -> (master's rels) -> theme (own name attribute). A layout does NOT
// link directly to its theme -- MEASURED: the theme relationship lives on the MASTER, not the
// layout, so this cannot skip the master hop even though nothing about the master's own CONTENT
// is used here (only its rels).
function pptxLayoutAndThemeFromSlideLayoutXml(layoutXml, masterThemeXml) {
  const xml = String(layoutXml ?? '')
  const layoutName = decodeXmlEntities((xml.match(/<p:cSld\s[^>]*\bname="([^"]*)"/) || [])[1] ?? '')
  const layoutType = (xml.match(/<p:sldLayout\s[^>]*\btype="([^"]*)"/) || [])[1] ?? null
  const themeName = masterThemeXml !== null && masterThemeXml !== undefined
    ? decodeXmlEntities((String(masterThemeXml).match(/<a:theme\s[^>]*\bname="([^"]*)"/) || [])[1] ?? '')
    : null
  return { layoutName, layoutType, themeName }
}

// xl/worksheets/sheet1.xml -> the set of cell refs
// ("F2", "E3", ...) that ACTUALLY GOT A <c> ELEMENT in the saved package. A cell whose write was
// silently swallowed (measured root cause: `formula`'s SetValue call with a syntactically-"="
// but semantically-broken formula, lib.cjs's own `formula` entry below) produces NO <c> element
// at all -- not an empty one, not an error value. This is therefore a PRESENCE check, not a
// value-read: the caller already knows what it asked to write, it needs to know whether the
// slot exists at all. Self-closing (`<c r="X1" .../>`) and content-bearing (`<c r="X1">...</c>`)
// forms are both cells that landed -- OOXML uses the self-closing form for a cell that has
// styling/type set but no cached value, still a real cell, not a missing one.
function cellRefsPresentInSheetXml(sheetXml) {
  const xml = String(sheetXml ?? '')
  const refs = new Set()
  const re = /<c\s[^>]*?\br="([A-Z]+[0-9]+)"[^>]*?(?:\/>|>)/g
  let m
  while ((m = re.exec(xml))) refs.add(m[1])
  return refs
}

// Unlike cellRefsPresentInSheetXml above (presence-only), this
// resolves a SPECIFIC cell's actual text VALUE -- needed to read back a marker string this module
// itself wrote via SetValue, which the Document Server stores as a shared-string reference
// (`t="s"`, `<v>` holding an INDEX into sharedStrings.xml), not inline text. Pure string parsing
// only, no filesystem/zip access (this module stays side-effect-free per its own header comment)
// -- the caller unzips the two XML parts and hands them in as strings. Returns null when the
// cell is absent, has no value, or is not a string cell (a numeric/boolean cell has no shared-
// string index to resolve, and this helper's one job is string cells).
function resolveXlsxCellText(sheetXml, sharedStringsXml, cellRef) {
  const sheet = String(sheetXml ?? '')
  const m = new RegExp(`<c\\s[^>]*?\\br="${cellRef}"[^>]*?>(.*?)</c>`).exec(sheet)
  if (!m) return null
  const cellTag = sheet.slice(m.index, m.index + m[0].length)
  const isSharedString = /\bt="s"/.test(cellTag.split('>')[0] + '>')
  const vMatch = /<v>([^<]*)<\/v>/.exec(m[1])
  if (!vMatch) return null
  if (!isSharedString) return null
  const idx = Number(vMatch[1])
  if (!Number.isInteger(idx) || idx < 0) return null
  const sst = String(sharedStringsXml ?? '')
  const items = [...sst.matchAll(/<si>(?:<t[^>]*>([^<]*)<\/t>|.*?)<\/si>/gs)].map((x) => x[1] ?? '')
  return idx < items.length ? items[idx] : null
}

// The SAME single-letter-column range math the `table`
// xlsx branch below uses to place each cell (see that entry's own `colLetter`/`startCol` -- this
// is a deliberate extraction, not a re-derivation, so the verification side can never drift from
// what emit() actually wrote: both call this one function for the ref list). Kept to
// single-letter start columns, same limitation `table`'s own emit() already has (a two-letter
// start column like "AA1" would compute a garbage `colLetter` there too) -- not this fix's job to
// widen that, only to not silently disagree with it.
function xlsxTableCellRefs(op) {
  const rows = Array.isArray(op.rows) ? op.rows : []
  if (!rows.length) return []
  const at = String(op.at ?? 'A1')
  const startCol = at.replace(/\d+/g, '')
  const startRow = Number(at.replace(/\D+/g, '') || 1)
  const colLetter = (j) => String.fromCharCode(startCol.charCodeAt(0) + j)
  const refs = []
  rows.forEach((sor, i) => {
    const cells = Array.isArray(sor) ? sor : [sor]
    cells.forEach((_cella, j) => refs.push(`${colLetter(j)}${startRow + i}`))
  })
  return refs
}

// Which cell refs a GIVEN operation, on the xlsx core,
// asked to land -- the request side of the presence check above. Only `table` and `formula` are
// named here (the two operation types known to silently drop a cell write): every other operation type
// returns an empty list, meaning "this fix does not cover it yet", NOT "verified empty" -- a
// caller must not read an empty list here as a pass.
function xlsxRequestedCellRefs(op) {
  if (!op || typeof op !== 'object') return []
  if (op.type === 'table') return xlsxTableCellRefs(op)
  if (op.type === 'formula' && op.at) return [String(op.at)]
  return []
}

// Resolves an operation's `sheet` field (0-based
// index, sheet name, or absent -> sheet 0) to the SAVED PACKAGE part name that operation actually
// targeted -- e.g. "xl/worksheets/sheet2.xml". Same document-order convention `Api.GetSheet(n)`
// uses (position in workbook.xml's own <sheet> list), so the READ side can never disagree with
// the WRITE side about what index N means (resolveXlsxSheetLine, the coedit write-side resolver,
// relies on the identical convention). Returns null on anything unresolvable (malformed workbook/
// rels XML, an out-of-range index, an unknown name, a <sheet> with no r:id, or a relationship Id
// missing from the rels part) -- the caller treats null as "cannot verify", never as sheet 0.
function resolveXlsxSheetFile(workbookXml, relsXml, sheetTarget) {
  if (typeof workbookXml !== 'string' || typeof relsXml !== 'string') return null
  const sheets = (workbookXml.match(/<sheet\b[^>]*\/>/g) || []).map((tag) => ({
    name: (/\bname="([^"]*)"/.exec(tag) || [])[1] ?? '',
    rid: (/\br:id="([^"]+)"/.exec(tag) || [])[1] ?? null,
  }))
  if (!sheets.length) return null
  const entry = sheetTarget === undefined || sheetTarget === null ? sheets[0]
    : typeof sheetTarget === 'number' ? sheets[sheetTarget]
    : sheets.find((s) => s.name === String(sheetTarget))
  if (!entry || !entry.rid) return null
  const relTag = (relsXml.match(new RegExp(`<Relationship\\b[^>]*\\bId="${entry.rid}"[^>]*/?>`)) || [])[0]
  if (!relTag) return null
  const target = (/\bTarget="([^"]+)"/.exec(relTag) || [])[1]
  if (!target) return null
  const norm = target.replace(/^\.?\//, '')
  return norm.startsWith('xl/') ? norm : `xl/${norm}`
}

// The PURE report-building step -- given the operations a batch asked for and the SAVED sheet
// XML, which of the xlsx-cell-targeting operations actually landed. Split out of coedit.cjs's
// writeOperationsToDocument specifically so this can be unit-tested without a browser/Document
// Server: coedit.cjs only does the I/O (unzip the saved file, catch a read failure), this
// function does the deciding.
//
// `sheetXmlOrPerOp` accepts TWO shapes, kept both for backward compatibility:
//   a string (or null) -> the SAME sheet XML applies to every operation (legacy, single-sheet
//     batches)
//   an array, same length/order as `operations` -> sheetXmlOrPerOp[index] is THAT operation's own
//     resolved sheet XML (or null) -- a batch
//     that targets DIFFERENT sheets via `op.sheet` no longer gets checked against a single,
//     possibly-wrong sheet (MEASURED root cause of a false `hianyzoCellak`: this function used to
//     always receive sheet1.xml regardless of which sheet an operation actually wrote to).
//
// `sheetXml === null` (for an operation) means "the read/resolution itself failed" (a read-back
// is never the condition for a write) -- that entry gets `hianyzoCellak: null,
// mindLetrejott: null` (NEM-MERT, not a failure verdict) rather than being silently omitted or
// reported as missing.
//
// Mixed-batch behaviour: one entry PER cell-targeting operation, not one
// collapsed bool for the whole call -- 3 operations that landed and 1 that did not report as
// three `mindLetrejott:true` and one `mindLetrejott:false`, never a single "partial failure".
function xlsxCellVerificationReport(operations, sheetXmlOrPerOp) {
  const erintettMuveletek = (Array.isArray(operations) ? operations : [])
    .map((op, index) => ({ index, op, kertCellak: xlsxRequestedCellRefs(op) }))
    .filter((m) => m.kertCellak.length)
  if (!erintettMuveletek.length) return null
  const perOp = Array.isArray(sheetXmlOrPerOp)
  return erintettMuveletek.map(({ index, op, kertCellak }) => {
    const sheetXml = perOp ? sheetXmlOrPerOp[index] : sheetXmlOrPerOp
    if (sheetXml === null || sheetXml === undefined) {
      return {
        index, type: op.type, kertCellak, hianyzoCellak: null, mindLetrejott: null,
        nemMertIndok: 'a mentett csomag munkalap-resze nem volt kicsomagolhato -- ez NEM azt jelenti, hogy az iras sikertelen volt, csak hogy a cella-jelenlet nincs visszaigazolva',
      }
    }
    const jelenLevo = cellRefsPresentInSheetXml(sheetXml)
    const hianyzoCellak = kertCellak.filter((ref) => !jelenLevo.has(ref))
    return { index, type: op.type, kertCellak, hianyzoCellak, mindLetrejott: hianyzoCellak.length === 0, nemMertIndok: null }
  })
}

// xlsxCellVerificationReport's verdict previously existed only as a SIDE-CHANNEL
// (`cellaEllenorzes` in the tool response)
// -- a caller has to know to check it separately; `report[i].outcome` stayed 'alkalmazva' even when
// `mindLetrejott` was false. This closes that gap the same way autoFilter's own equivalent function
// does: mutate the report entry itself when verification PROVES a cell never landed, so a caller
// checking `report[i].outcome` alone (the field explicitly documented as the per-operation verdict)
// is not misled into "alkalmazva" for an operation that silently wrote nothing.
//
// No 'kihagyva-idempotens' branch here, unlike autoFilter: a `table`/`formula` write has no runtime
// guard with a distinct "already there, safe no-op" branch to detect -- `SetValue` either lands the
// cell or (per the silent-no-op class this whole check exists for) it does not. Only two outcomes
// apply: 'vegrehajtva' (cells present, unchanged from 'alkalmazva' in spirit but the word now means
// "verified", not "validated before running") or 'megtagadva' (cells missing). `mindLetrejott: null`
// (re-read failed) leaves the entry untouched, same NEM-MERT-is-not-a-verdict rule as autoFilter's.
function applyCellVerificationToReport(report, verification) {
  if (!Array.isArray(report) || !Array.isArray(verification)) return report
  for (const v of verification) {
    const entry = report[v.index]
    if (!entry || entry.outcome !== 'elkuldve-nem-verifikalt') continue
    if (v.mindLetrejott === true) {
      entry.outcome = 'vegrehajtva'
    } else if (v.mindLetrejott === false) {
      entry.outcome = 'megtagadva'
      entry.reason = `${v.type}: a kert cella(k) nem talalhato(k) a mentett munkalapon: ${v.hianyzoCellak.join(', ')} -- a hivas nem dobott kivetelt, de a dokumentum-szerver csendben nem irta be az erteket`
    }
    // mindLetrejott===null: nincs visszaigazolas (a lap nem volt kicsomagolhato) -- 'elkuldve-nem-verifikalt' marad
  }
  return report
}

// The read side of a "third case": an autoFilter
// runtime guard throws
// INSIDE the generated Document Server script when a DIFFERENT range already carries a filter --
// but that throw happens on the server, AFTER buildCoeditScript's own client-side report already
// marked the operation 'alkalmazva' (validated fine structurally, before anything ran). Same rule
// as xlsxCellVerificationReport above (only the saved package is proof), applied to the
// <autoFilter ref="..."> element instead of a cell.
function xlsxAutoFilterRangeInSheetXml(sheetXml) {
  if (typeof sheetXml !== 'string') return null
  const m = /<autoFilter\b[^>]*\bref="([^"]+)"/.exec(sheetXml)
  return m ? m[1] : null
}

// The dispatcher xlsxRequestedCellRefs has for `table`/`formula`, for autoFilter: does THIS
// operation name a range worth verifying. A separate predicate (not folded into the report
// function below) so the caller can cheaply ask "does this batch need the re-read at all" without
// building report entries for an I/O step that has not happened yet -- same division as
// xlsxRequestedCellRefs vs. xlsxCellVerificationReport.
function xlsxRequestsAutoFilterVerification(op) {
  return Boolean(op && op.type === 'autoFilter' && typeof op.range === 'string' && op.range)
}

// Mirrors xlsxCellVerificationReport's shape: one entry per autoFilter operation that named a
// `range`, comparing what was ASKED (`kertRange`) against what the saved package actually HAS
// (`tenylegesRange`). `egyezik: null` (not false) means the sheet part itself could not be
// re-read -- NEM-MERT, not a failure verdict (5. pont: a visszaolvasas hibaja nem az iras hibaja).
//
// `beforeSheetXmlOrPerOp` (OPTIONAL, same string-or-per-op-
// array shape as the AFTER argument): the SAME sheet part read from the package BEFORE the write.
// Without it, a matching AFTER state cannot tell "the runtime guard's safe no-op branch fired
// because the wanted range already existed" apart from "this call is what created it" -- both look
// identical from the AFTER snapshot alone. `korabbanMarOtt` carries that distinction: true (the
// wanted range's autoFilter already existed pre-write -- a genuine idempotent no-op) | false (it
// did not -- this write is what created it) | null (no before-snapshot supplied, or its own re-read
// failed -- NEM-MERT on this axis specifically, independent of whether `egyezik` succeeded).
function xlsxAutoFilterVerificationReport(operations, sheetXmlOrPerOp, beforeSheetXmlOrPerOp) {
  const erintettMuveletek = (Array.isArray(operations) ? operations : [])
    .map((op, index) => ({ index, op }))
    .filter(({ op }) => xlsxRequestsAutoFilterVerification(op))
  if (!erintettMuveletek.length) return null
  const perOp = Array.isArray(sheetXmlOrPerOp)
  const beforePerOp = Array.isArray(beforeSheetXmlOrPerOp)
  return erintettMuveletek.map(({ index, op }) => {
    const sheetXml = perOp ? sheetXmlOrPerOp[index] : sheetXmlOrPerOp
    if (sheetXml === null || sheetXml === undefined) {
      return {
        index, type: op.type, kertRange: op.range, tenylegesRange: null, egyezik: null, korabbanMarOtt: null,
        nemMertIndok: 'a mentett csomag munkalap-resze nem volt kicsomagolhato -- ez NEM azt jelenti, hogy az iras sikertelen volt, csak hogy az autoFilter-allapot nincs visszaigazolva',
      }
    }
    const tenylegesRange = xlsxAutoFilterRangeInSheetXml(sheetXml)
    const egyezik = tenylegesRange === op.range
    let korabbanMarOtt = null
    if (beforeSheetXmlOrPerOp !== undefined) {
      const beforeXml = beforePerOp ? beforeSheetXmlOrPerOp[index] : beforeSheetXmlOrPerOp
      if (typeof beforeXml === 'string') korabbanMarOtt = xlsxAutoFilterRangeInSheetXml(beforeXml) === op.range
    }
    return { index, type: op.type, kertRange: op.range, tenylegesRange, egyezik, korabbanMarOtt, nemMertIndok: null }
  })
}

// Applies the verdict above (the mismatch case, plus the
// three-way outcome split) back onto buildCoeditScript's own client-side
// `report`. Mutates in place, same style as rejectCoeditBatch's own post-hoc relabelling in this
// file -- both correct an earlier 'alkalmazva' once later information (there: a sibling operation's
// failure; here: what the saved package actually has) proves it wrong.
//
// THREE OUTCOMES, replacing the old binary 'alkalmazva'/'hiba' for autoFilter specifically (this
// three-way split applies ONLY to operation types that already have post-save verification wired
// -- autoFilter here, table/formula cells separately below -- not a blanket rename of every
// operation's outcome vocabulary):
//   'vegrehajtva'          -- egyezik AND korabbanMarOtt===false: the write is what created this state
//   'kihagyva-idempotens'  -- egyezik AND korabbanMarOtt===true: the wanted state already existed;
//                             the runtime guard's safe no-op branch fired, nothing needed to change
//   'megtagadva'           -- egyezik===false: the document-server runtime refused the write (a
//                             different range was already filtered) -- was 'hiba', renamed to match
//                             the new three-term vocabulary (megtagadva is a MORE ACCURATE word here
//                             than 'hiba': nothing malfunctioned, the runtime correctly refused)
// If `korabbanMarOtt` is null (no before-snapshot given, or its own re-read failed), egyezik===true
// entries are left as 'elkuldve-nem-verifikalt' -- a caller not supplying the before-state gets the
// OLD two-outcome behaviour, not a fabricated vegrehajtva/kihagyva-idempotens guess.
function applyAutoFilterVerificationToReport(report, verification) {
  if (!Array.isArray(report) || !Array.isArray(verification)) return report
  for (const v of verification) {
    const entry = report[v.index]
    if (!entry || entry.outcome !== 'elkuldve-nem-verifikalt') continue
    if (v.egyezik === false) {
      entry.outcome = 'megtagadva'
      entry.reason = v.tenylegesRange === null
        ? `autoFilter: a mentett munkalapon nincs autoFilter a kert "${v.kertRange}" tartomanyon -- a muvelet a dokumentum-szerveren nem hajtodott vegre`
        : `autoFilter: a mentett munkalapon a "${v.tenylegesRange}" tartomany van szurve, nem a kert "${v.kertRange}" -- a muvelet a dokumentum-szerveren elutasitva (mas tartomanyon mar volt szuro, es a SetAutoFilter() ujra-hivasa toggle-kent torolte volna azt)`
    } else if (v.egyezik === true && v.korabbanMarOtt === true) {
      entry.outcome = 'kihagyva-idempotens'
      entry.reason = `autoFilter: a "${v.kertRange}" tartomanyon MAR volt szuro a muvelet ELOTT is -- a hivas biztonsagos no-op volt, a kert vegallapot mar fennallt es nem valtozott`
    } else if (v.egyezik === true && v.korabbanMarOtt === false) {
      entry.outcome = 'vegrehajtva'
    }
    // egyezik===true, korabbanMarOtt===null: nincs elotte-adat -- 'elkuldve-nem-verifikalt' marad
    // (regi, ket-allapotu viselkedes, nem talalt ki harmadik allapotot bizonyitek nelkul)
  }
  return report
}

// office_find / findText: a plain substring search over ALREADY-PARSED
// paragraphs (paragraphTextsFromDocumentXml's own output) -- no new document-reading mechanism,
// reuses the one office_get_text already has. Deliberately case-SENSITIVE and literal (no regex,
// no case-folding): a caller who wants "find text" almost always means the exact string they are
// about to hand to a replace call next, and a case-insensitive default would silently match more
// than the caller typed. Returns per-paragraph match counts, not global offsets -- there is no
// stable, cross-call character offset to report (the document can be re-saved, re-paginated,
// re-flowed between a find and any later operation), so this reports WHERE (which paragraph,
// how many times in it), not a position a later call could seek to directly.
function findMatchesInParagraphs(paragraphs, query) {
  const q = String(query ?? '')
  if (!q) return { totalCount: 0, matches: [] }
  const list = Array.isArray(paragraphs) ? paragraphs : []
  const matches = []
  let totalCount = 0
  list.forEach((text, index) => {
    let count = 0
    let from = 0
    while (true) {
      const at = text.indexOf(q, from)
      if (at === -1) break
      count++
      from = at + q.length
    }
    if (count > 0) {
      matches.push({ paragraphIndex: index, count, paragraphText: text })
      totalCount += count
    }
  })
  return { totalCount, matches }
}

// --- E0: THE OPERATION SCHEMA AND ITS TRANSLATOR ---------------------------------------------
// Owner's scope (2026-08-15): put everything the Office API can do into euro-mcp. A caller should
// describe WHAT they want, not write DocBuilder JS -- so the tool takes a list of operations and
// this translator turns it into a script.
//
// *** WHICH ROUTE EACH OPERATION LIVES ON, AND WHY IT MATTERS: *** every capability below was
// measured on the DocBuilder route, which the protocol reserves for CREATING a file. On a file
// that already exists the same operation is not reachable today: the co-editing route makes a
// single call (PasteHtml), and the editor frame has no Api object at all (measured in all three
// cores). The caller learns this from the ANSWER, not from a document -- hence `sourceRoute`.
const CREATE_ROUTE = 'docbuilder-create'

// --- image `path`/`src` resolution ---------------------------------------------------------
// Magic-byte sniffing, not the file extension: an extension is a claim the caller makes about
// the file, magic bytes are what the file actually is. A caller passing `foo.png` that is
// really a text file should get a named "not an image" error, not a broken embed three steps
// downstream inside the DocBuilder run.
const IMAGE_MAGIC = [
  { mime: 'image/png', test: (b) => b.length >= 8 && b.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) },
  { mime: 'image/jpeg', test: (b) => b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff },
  { mime: 'image/gif', test: (b) => b.length >= 6 && ['GIF87a', 'GIF89a'].includes(b.subarray(0, 6).toString('ascii')) },
  { mime: 'image/bmp', test: (b) => b.length >= 2 && b.subarray(0, 2).toString('ascii') === 'BM' },
]
function sniffImageMime(buf) {
  const hit = IMAGE_MAGIC.find((m) => m.test(buf))
  return hit ? hit.mime : null
}
const DATA_URI_RE = /^data:[a-zA-Z0-9.+-]+\/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/]+=*$/

function resolveImageSrc(op) {
  if (op.path) {
    // Named, not silent: a missing/wrong-type file must stop THIS operation with an error that
    // names it, not fall through to a script that DocBuilder would run with nothing embedded.
    if (!fs.existsSync(op.path)) throw new Error(`image: file not found: ${op.path}`)
    const buf = fs.readFileSync(op.path)
    const mime = sniffImageMime(buf)
    if (!mime) throw new Error(`image: \`path\` (${op.path}) is not a recognised image file (checked PNG/JPEG/GIF/BMP magic bytes)`)
    return `data:${mime};base64,${buf.toString('base64')}`
  }
  if (!op.src) throw new Error('image: either `src` (a data: URI) or `path` (a local file) is required')
  if (!DATA_URI_RE.test(op.src)) {
    throw new Error('image: `src` is not a well-formed data: URI (expected "data:<mime>/<subtype>;base64,<data>")')
  }
  return op.src
}

// Indentation + spacing for `text`'s docx branch, split out of emit() so a paragraph property
// does not keep growing that function's own complexity (qlty: emit() already carried a
// pre-existing high-complexity smell before this addition -- see git blame on the docx branch
// above).
function applyParagraphIndentSpacing(lines, op) {
  for (const [field, setter, label] of [
    ['indentFirstLine', 'SetIndFirstLine', '`indentFirstLine`'],
    ['indentLeft', 'SetIndLeft', '`indentLeft`'],
    ['indentRight', 'SetIndRight', '`indentRight`'],
    ['spacingBefore', 'SetSpacingBefore', '`spacingBefore`'],
    ['spacingAfter', 'SetSpacingAfter', '`spacingAfter`'],
  ]) {
    if (op[field] === undefined || op[field] === null) continue
    const twip = Number(op[field])
    if (!Number.isInteger(twip) || twip < 0) {
      throw new Error(`text: ${label} must be a non-negative integer (TWIP), got ${JSON.stringify(op[field])}`)
    }
    lines.push(`oParagraph.${setter}(${twip});`)
  }
  // `spacingLine` is DELIBERATELY NOT converted to a friendly unit ("single"/"double"/
  // "1.5") -- recipe #9 measured SetSpacingLine(2, "auto") producing the RAW value in the
  // package (<w:spacing w:line="2" w:lineRule="auto"/>), not a Word-standard 240-multiple
  // (a true double spacing would be w:line="480"). Whether the Document Server's renderer
  // treats a raw "2" as 2 TWIP, 2/240 of a line, or something else is NOT measured here --
  // guessing a conversion formula would be exactly the "csendben atengedett nem-mert
  // egyseg" this card forbids. So this passes the caller's number straight through,
  // unconverted, and ONLY accepts the one lineRule value the recipe actually proved
  // ("auto") -- any other rule is refused by name rather than attempted blind.
  if (op.spacingLine !== undefined && op.spacingLine !== null) {
    const rawLine = Number(op.spacingLine)
    if (!Number.isInteger(rawLine) || rawLine < 0) {
      throw new Error(`text: \`spacingLine\` must be a non-negative integer, got ${JSON.stringify(op.spacingLine)} -- passed RAW to SetSpacingLine, no unit conversion (see the code comment above this check)`)
    }
    const rule = op.spacingLineRule ?? 'auto'
    if (rule !== 'auto') {
      throw notSupportedError(`text: \`spacingLineRule\` ${JSON.stringify(rule)} is not supported -- only "auto" is measured (receptek-pptx-docx.md #9); other SetSpacingLine lineRule values were never run against this Document Server instance, so this tool refuses them here instead of guessing`)
    }
    lines.push(`oParagraph.SetSpacingLine(${rawLine}, ${jsString(rule)});`)
  }
}

// The `table` operation's docx-K2 per-cell refinements
// (verticalAlign/noWrap/cellMargin*), pulled into its own function for the same reason
// applyDocxTableMerge was: this pushed table.emit()'s own complexity past the qlty threshold, and
// a smell that only exists because unrelated fields share one function body is about where the
// brace goes, not the logic. Called once per docx cell, so `cell` is already the resolved
// `oTable.GetCell(i, j)` expression string, not the row/col indices.
function applyDocxTableCellRefinements(lines, op, cell) {
  if (op.verticalAlign) lines.push(`${cell}.SetVerticalAlign(${jsString(op.verticalAlign)});`)
  if (op.noWrap) lines.push(`${cell}.SetNoWrap(true);`)
  for (const side of ['Top', 'Bottom', 'Left', 'Right']) {
    const key = `cellMargin${side}`
    if (op[key] !== undefined && op[key] !== null) lines.push(`${cell}.SetCellMargin${side}(${Number(op[key])});`)
  }
}

// Validation half of the docx-K2 refinements above -- separated from the per-cell emitter (which
// runs once per cell) because this only needs to run ONCE per table, and separated from emit()
// itself for the same complexity reason as applyDocxTableCellRefinements.
function validateDocxTableRefinements(op) {
  // textDirection is NOT offered: every OOXML value tried (lrTb/tbRl/btLr/tbRlV/lrTbV/tbLrV) ran
  // to completion with NO <w:textDirection/> in the saved package (MEASURED 2026-08-17) -- the
  // same silent-no-op class as `align`/`highlight` elsewhere in this file, except here there was
  // no value at all that landed, so there is nothing to allowlist.
  if (op.textDirection !== undefined && op.textDirection !== null) {
    throw notSupportedError('table: `textDirection` (docx cell text direction, Cell.SetTextDirection) is refused -- MEASURED 2026-08-17: every OOXML direction value tried (lrTb, tbRl, btLr, tbRlV, lrTbV, tbLrV) ran to completion but produced no <w:textDirection/> in the saved package, on a cell where a sibling SetVerticalAlign call on the SAME cell in the SAME script did land -- so this is not a whole-job failure, just this one setter having no observable effect on this DocBuilder instance')
  }
  if (op.verticalAlign !== undefined && op.verticalAlign !== null && !KNOWN_CELL_VALIGNS.includes(op.verticalAlign)) {
    throw notSupportedError(`table: unknown verticalAlign ${JSON.stringify(op.verticalAlign)} (known: ${KNOWN_CELL_VALIGNS.join(', ')})`)
  }
}

// Row-level half (SetTableHeader/SetHeight run on the ROW, not per cell) -- applied once after
// all cells are built, docx only.
function applyDocxTableRowRefinements(lines, op, varName = 'oTable') {
  if (op.repeatHeaderRow) lines.push(`${varName}.GetRow(0).SetTableHeader(true);`)
  if (Array.isArray(op.rowHeights)) {
    // "atLeast" is the ONLY hRule value that landed in measurement -- "exact"/"Exact"/"EXACT"/
    // "auto"/"Auto" and both argument orders all ran to completion with no <w:trHeight/> at all
    // in the saved package. Not offered as a caller-facing choice, since there is nothing else to
    // choose from that has been shown to work.
    op.rowHeights.forEach((h, i) => {
      if (h !== null && h !== undefined) lines.push(`${varName}.GetRow(${i}).SetHeight("atLeast", ${Number(h)});`)
    })
  }
}

// Footnotes/endnotes attach to the paragraph `text` just built, so
// they are fields on `text` rather than a separate operation type (there is nothing else for them
// to attach to). Split out of that operation's emit() for the same complexity reason as
// emitDocxRun below. MEASURED 2026-08-17: `oDocument.AddFootnote()`/`AddEndnote()` insert the
// reference mark at the document's CURRENT CURSOR position, which by default is the START of the
// document, not wherever a caller last pushed a paragraph -- a first attempt without
// `MoveCursorToEnd()` put the reference on the wrong paragraph (the input marker text, not the
// one just built). `MoveCursorToEnd()` immediately before each Add* call fixes it, package-
// verified: footnote/endnote reference lands as the LAST run of THIS paragraph, the note body
// lands in footnotes.xml/endnotes.xml, and a script with no Add*note call has zero
// <w:footnoteReference/>/<w:endnoteReference/> anywhere.
function applyDocxParagraphNotes(lines, op) {
  // *** ORDER TRAP, MEASURED 2026-08-17: `comments` MUST be applied BEFORE footnotes/endnotes on
  // the same paragraph, not after. *** A footnote or endnote added first leaves the document
  // cursor in a state where the FOLLOWING AddComment() still returns a truthy comment object and
  // still writes the comment body into the saved package -- but with NO <w:commentRangeStart>/
  // <w:commentReference> anywhere in document.xml, i.e. a genuinely ORPHANED comment (present in
  // the package, invisible and unreachable in Word). Comment-then-footnote and comment-then-
  // endnote both package-verified correct (commentRangeStart + footnoteReference/endnoteReference
  // both present); footnote-then-comment and endnote-then-comment both package-verified broken.
  applyDocxParagraphComments(lines, op)
  // `bookmark` -- SAME order trap, same class (measured 2026-08-17). A
  // footnote added to this paragraph BEFORE the bookmark leaves `oParagraph.GetRange().
  // AddBookmark(name)` returning `true` and the name showing up in `oDocument.
  // GetAllBookmarksNames()` -- but with ZERO <w:bookmarkStart>/<w:bookmarkEnd> anywhere in the
  // saved document.xml. The bookmark exists in the document's bookmarks MANAGER (so a later
  // AddBookmarkCrossRef targeting it still returns true and still emits a REF/PAGEREF field --
  // pointing at a name that was never materialised) but is otherwise a phantom. Package-verified
  // both ways: bookmark-then-footnote and bookmark-then-comment both correct (bookmarkStart/End
  // present); footnote-then-bookmark loses the markers, comment-then-bookmark does not (comments
  // do not share footnote's failure mode here). Bookmark is therefore placed here, after comments
  // (order-independent against comments) and before footnotes/endnotes (order-DEPENDENT).
  if (op.bookmark !== undefined && op.bookmark !== null) {
    const name = String(op.bookmark)
    if (!name) throw new Error('text: `bookmark` must be a non-empty string')
    lines.push('oParagraph.GetRange().AddBookmark(' + jsString(name) + ');')
  }
  for (const note of Array.isArray(op.footnotes) ? op.footnotes : []) {
    const body = typeof note === 'string' ? note : note?.text
    lines.push('oDocument.MoveCursorToEnd();', 'oDocument.AddFootnote().GetElement(0).AddText(' + jsString(String(body ?? '')) + ');')
  }
  for (const note of Array.isArray(op.endnotes) ? op.endnotes : []) {
    const body = typeof note === 'string' ? note : note?.text
    lines.push('oDocument.MoveCursorToEnd();', 'oDocument.AddEndnote().GetElement(0).AddText(' + jsString(String(body ?? '')) + ');')
  }
}

// `bookmarkRef` -- `Paragraph.AddBookmarkCrossRef(sRefTo,
// sBookmarkName, bLink, bAboveBelow, sSepWith)` -- MEASURED via the live function's own
// `.toString()` (2026-08-17), because a plausible-looking guess at the call shape
// (bookmarkName, displayText, format) compiled and ran with NO throw but returned `false` on
// every attempt: the ARGUMENT ORDER is `(refTo, bookmarkName, ...)`, not the other way round,
// and there is no display-text parameter at all -- the field's displayed text is generated by
// Word itself from `refTo` when the field is updated, the same way a native Word cross-reference
// works. `refTo` is validated against the engine's own switch (read out of the same
// `.toString()`): an unrecognised value leaves the engine's internal `nRefTo` at -1 and the call
// returns `false` without throwing -- the established "silent no-op" class this tool refuses
// eagerly elsewhere (`align`, `highlight`), so it is refused here too rather than shipped through
// to a false a caller has to notice on their own. A `to` naming a bookmark that does not exist
// (wrong name, or the bookmark's own operation has not run yet in the caller's `operations`
// array) is NOT caught here -- the engine's bookmarks-manager lookup happens at DocBuilder
// runtime, this tool has no static view of it, and the call still returns `false` rather than
// throwing, same as `toc`'s own documented ordering risk. Package-verified: `refTo: "text"` with
// `link: true` produces ` REF <name>  \h`; `refTo: "pageNum"` with `link: false` produces `
// PAGEREF <name>`. Also package-verified in combination with `comments`+`bookmark`+`footnotes`
// all on the SAME paragraph (the mandatory order-trap gate the bookmark finding above
// required): comment, then bookmark, then footnote, then a self-referencing crossref -- all four
// artifacts (commentRangeStart, bookmarkStart/End, footnoteReference, the REF field) landed
// together in one saved package.
// `customStyle` -- oDocument.CreateStyle(name, "paragraph") +
// Api.CreateTextPr() + Style.SetTextPr() + Paragraph.SetStyle(), the SAME mechanism the built-in
// `heading` field on `text` already uses (oDocument.GetStyle("Heading N")), just with a
// caller-named style instead of a built-in one. Package-verified: the style lands in
// word/styles.xml with its own numeric w:styleId, a <w:name> holding the CALLER's string, and the
// requested bold/color/size in its rPr; the paragraph references that id via <w:pStyle>. Scope,
// named explicitly: this tool does NOT deduplicate or reuse a style by name across multiple
// `text` operations -- every `text` op carrying `customStyle` creates its OWN style object, even
// if two ops pass the same `name`. Whether the DocBuilder engine itself merges same-named styles
// into one styles.xml entry, or keeps duplicates, is NOT measured here (the fixtures used here
// only ever used one `text` op per style) -- a caller reusing one style name across many paragraphs
// should treat the result as unverified until someone measures that specific case.
function applyDocxParagraphCustomStyle(lines, op) {
  if (op.customStyle === undefined || op.customStyle === null) return
  const cs = op.customStyle
  if (!cs.name) throw new Error('text: `customStyle.name` is required')
  lines.push(
    `var oCustomStyle = oDocument.CreateStyle(${jsString(String(cs.name))}, "paragraph");`,
    'var oCustomTextPr = Api.CreateTextPr();',
  )
  if (cs.bold) lines.push('oCustomTextPr.SetBold(true);')
  if (cs.italic) lines.push('oCustomTextPr.SetItalic(true);')
  if (cs.underline) lines.push('oCustomTextPr.SetUnderline(true);')
  if (cs.size) lines.push(`oCustomTextPr.SetFontSize(${Number(cs.size)});`)
  if (cs.font) lines.push(`oCustomTextPr.SetFontFamily(${jsString(String(cs.font))});`)
  if (cs.color !== undefined && cs.color !== null) {
    const { r, g, b } = parseHexColor(cs.color)
    lines.push(`oCustomTextPr.SetColor(${r}, ${g}, ${b}, false);`)
  }
  lines.push('oCustomStyle.SetTextPr(oCustomTextPr);', 'oParagraph.SetStyle(oCustomStyle);')
}

// `hyperlink` -- Paragraph.AddHyperlink(sLink, sScreenTipText), MEASURED via the live function's
// own .toString(): a THIRD argument (sBookmarkName) exists but is mutually exclusive with sLink
// (the function returns null if both are given, or if neither is) -- not exposed here, since this
// field is for an external URL, not an internal bookmark jump (that use case belongs to
// `bookmarkRef` above, a different mechanism entirely). AddHyperlink wraps the paragraph's OWN
// CURRENT TEXT (internally: `this.Paragraph.SelectAll(1)`) -- there is no separate "display text"
// parameter, matching the same "the call reads its own paragraph content" convention
// `WrapInMailMergeField` and `AddBookmarkCrossRef` already established elsewhere in this file.
// Package-verified: a real <w:hyperlink r:id="..."> wraps the run, with the External relationship
// registered in word/_rels/document.xml.rels -- both checked, not just the element. Two early
// measurement runs came back the whole-job "blocked" outcome; five identical re-runs immediately
// after all succeeded with the hyperlink correctly present -- read as Document Server flakiness
// (the same class already noted on `headerFooter`+comments), not a deterministic incompatibility,
// so this is NOT refused.
//
// *** THIS IS PARAGRAPH-LEVEL, NOT THE SAME AS `runs[].hyperlink`/`shape.hyperlink` ELSEWHERE IN
// THIS FILE, WHICH REMAIN REFUSED FOR DOCX. *** Run.AddHyperlink() on docx (the mechanism those
// refusals are about) was re-measured for this card: it now returns `null` and produces no XML
// change at all -- confirming, not just carrying forward, that refusal. Paragraph.AddHyperlink()
// is a genuinely different call on a different object and was not covered by that earlier
// refusal's own measurement.
function applyDocxParagraphHyperlink(lines, op) {
  if (op.hyperlink === undefined || op.hyperlink === null) return
  const url = op.hyperlink.url
  if (!url) throw new Error('text: `hyperlink.url` is required')
  const tooltip = op.hyperlink.tooltip !== undefined ? String(op.hyperlink.tooltip) : ''
  lines.push(`oParagraph.AddHyperlink(${jsString(String(url))}, ${jsString(tooltip)});`)
}

const KNOWN_BOOKMARK_REF_TYPES = ['text', 'pageNum', 'paraNum', 'noCtxParaNum', 'fullCtxParaNum', 'aboveBelow']
function applyDocxParagraphBookmarkRef(lines, op) {
  if (op.bookmarkRef === undefined || op.bookmarkRef === null) return
  const ref = op.bookmarkRef
  const to = ref?.to !== undefined ? String(ref.to) : ''
  if (!to) throw new Error('text: `bookmarkRef.to` must be a non-empty bookmark name')
  const refTo = ref?.refTo
  if (!KNOWN_BOOKMARK_REF_TYPES.includes(refTo)) {
    throw notSupportedError(`text: \`bookmarkRef.refTo\` unknown ${JSON.stringify(refTo)} (known: ${KNOWN_BOOKMARK_REF_TYPES.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently returns false with no field inserted at all, so this tool refuses it here instead`)
  }
  const link = ref?.link !== undefined ? Boolean(ref.link) : true
  const aboveBelow = ref?.aboveBelow !== undefined ? Boolean(ref.aboveBelow) : false
  const sepWith = ref?.sepWith !== undefined ? String(ref.sepWith) : ''
  lines.push(`oParagraph.AddBookmarkCrossRef(${jsString(refTo)}, ${jsString(to)}, ${link}, ${aboveBelow}, ${jsString(sepWith)});`)
}

// a pptx `text` op ide van kiemelve a text.emit()-bol -- a
// docx-ag mellett a hurkos logika a mar high-complexity-vel jelolt (qlty smells: count=37,
// javitas elott) fuggvenyt tovabb hizlalta volna. Sajat, nevesitett fuggvenyben a novekedes
// csak ezt a fuggvenyt terheli, nem az egesz emit()-et.
//
// `paragraphs` (opcionalis tomb) tobb bekezdest enged EGYETLEN szovegdobozban -- nelkule
// (a default eset) ez PONTOSAN azt a hat sort adja, amit a regi kod adott: az elso (es
// egyetlen) ag byte-azonos a valtoztatas elotti alakkal, tehat egy meglevo hivo scriptje
// nem valtozik. Az elso bekezdes a doboz sajat alap-bekezdeset hasznalja fel
// (oContent.GetElement(0)), minden tovabbi UJ, es oContent.Push()-sal kerul be -- ugyanaz
// a modszer, mint oDocument.Push(oParagraph) a docx-torzsnel (fent) es
// cell.GetContent().Push() egy beagyazott tablanal (table op) -- mindharom ugyanolyan
// DocContent-csaladbeli objektumon. VALODI DocBuilder-futtatassal igazolva (nem csak a
// script-szoveg regex-illesztesevel): a szovegdoboz csomagban PONTOSAN annyi <a:p> all elo,
// ahany bekezdest kertek, es EGYETLEN <p:sp> (szovegdoboz) jon letre -- nem tobb kulon shape.
//
// A per-bekezdes mezok szandekosan a mai run-szintu keszletre szukulnek (text/size/bold) --
// a docx-ag gazdagabb keszlete (italic/underline/color/align/...) itt NEM lett portolva: ezt
// senki nem kerte, es teszteletlen felulet lenne.
// `listType` (measured 2026-08-17): `bullet` WORKS, via `Api.CreateBullet
// (sSymbol)` + `oPara.SetBullet(oBullet)` -- package-verified: a real `<a:pPr><a:buChar
// char="..."/></a:pPr>` lands on the paragraph.
//
// `numbered` ALSO WORKS (measured 2026-08-17) -- the
// EARLIER refusal was measured against the WRONG method pair. The premise ("pptx's ApiParagraph
// shares SetNumPr/SetNumbering with docx, but there is no presentation-scoped numbering source")
// was correct as far as it went, but `Api.CreateNumbering(numType, startAt)` was never meant to
// feed `SetNumPr` on this core: reflected off the live instance (`String(Api.CreateNumbering)`),
// on pptx it builds an `AscFormat.CBullet` with `bulletType.type = BULLET_TYPE_BULLET_AUTONUM` and
// returns `new ApiBullet(oBullet)` -- an ApiBullet, the SAME class `Api.CreateBullet` returns, not
// an ApiNumbering. It was never going to pass `SetNumPr`'s `instanceof ApiNumbering` guard, because
// it isn't trying to be one -- it's a second ApiBullet constructor, meant for `SetBullet`, exactly
// like the already-working bullet path. Package-verified: `Api.CreateNumbering("ArabicPeriod")` +
// `oPara.SetBullet(...)` lands a real `<a:pPr><a:buAutoNum type="arabicPeriod"/>...` on the
// paragraph (byte-identical mechanism to the bullet char case, just a different CBullet shape).
// `numType` accepts the DocBuilder-native type names (measured live, from the reflected switch
// statement) -- these are NOT the OOXML attribute spelling 1:1 for the Roman/Alpha cases (DocBuilder
// spells `RomanUcPeriod`, the emitted XML says `type="romanUcPeriod"` -- same string, just the
// DocBuilder side is the one this field's values are validated against, not the raw XML).
const PPTX_NUM_TYPES = ['ArabicPeriod', 'ArabicParenR', 'RomanUcPeriod', 'RomanLcPeriod', 'AlphaLcParenR', 'AlphaLcPeriod', 'AlphaUcParenR', 'AlphaUcPeriod']
function applyPptxParagraphListType(lines, p) {
  if (p.listType === undefined || p.listType === null) return
  if (p.listType === 'bullet') {
    lines.push(`oPara.SetBullet(Api.CreateBullet(${jsString(String(p.bulletChar ?? '•'))}));`)
    return
  }
  if (p.listType === 'numbered') {
    const numType = p.numType ?? 'ArabicPeriod'
    if (!PPTX_NUM_TYPES.includes(numType)) {
      throw new Error(`text: \`numType\` must be one of ${PPTX_NUM_TYPES.join(', ')}, got ${JSON.stringify(numType)}`)
    }
    lines.push(`oPara.SetBullet(Api.CreateNumbering(${jsString(numType)}));`)
    return
  }
  throw new Error(`text: listType must be "bullet" or "numbered", got ${JSON.stringify(p.listType)}`)
}

// REFLEXIO-VERIFIKALT hogy
// SetIndLeft LETEZIK a pptx paragrafuson is (typeof-fal, a FULL prototype-lancon at -- egy
// getOwnPropertyNames egy szinten ezt NEM latta, mert oroklott tag), de a letezes onmagaban NEM
// bizonyitja hogy IR a csomagba (lasd a `numbered` megtagadasat, ugyanez az alak). MEASURE-FIRST:
// az egyseg csak azutan all be MUKODO agkent, hogy egy elo DocBuilder-futtatas a mentett pptx-ben
// tenylegesen `<a:pPr marL="...">`-t mutat a kert bekezdesen, ES egy behuzas nelkulin NINCS ott.
function applyPptxParagraphIndent(lines, p) {
  if (p.indentLeft === undefined || p.indentLeft === null) return
  const twip = Number(p.indentLeft)
  if (!Number.isInteger(twip) || twip < 0) {
    throw new Error(`text: \`indentLeft\` must be a non-negative integer (TWIP), got ${JSON.stringify(p.indentLeft)}`)
  }
  // MERTEN, NEM FELTETELEZVE: a RAW twip
  // erteket adjuk at, KONVERZIO NELKUL -- ez EL VAN DONTVE, nem nyitott kerdes. Csomag-szintu
  // meres igazolta: 720 twip -> `<a:pPr marL="457200">`, azaz a DocBuilder MAGA vegzi a
  // TWIP->EMU szorzast erre a metodusra (720*635=457200). A cellamargo-eset (SetCellMarginLeft,
  // szinten 635-os szorzo) MASIK metodus volt -- ez a sor NEM abbol oroklodott, hanem KULON
  // meresbol all. NE szorozz itt kezzel 635-tel: az a behuzast 635-szorosra torzitana.
  lines.push(`oPara.SetIndLeft(${twip});`)
}

function buildPptxTextShapeLines(op) {
  const paragraphSpecs = Array.isArray(op.paragraphs) && op.paragraphs.length > 0
    ? op.paragraphs
    : [{ text: op.text, size: op.size, bold: op.bold, listType: op.listType, bulletChar: op.bulletChar, numType: op.numType, indentLeft: op.indentLeft }]
  const shapeLines = [
    `var oShape = Api.CreateShape("rect", ${Number(op.width ?? 9000000)}, ${Number(op.height ?? 800000)}, Api.CreateNoFill(), Api.CreateStroke(0, Api.CreateNoFill()));`,
    `oShape.SetPosition(${Number(op.x ?? 800000)}, ${Number(op.y ?? 800000)});`,
    'var oContent = oShape.GetDocContent();',
  ]
  paragraphSpecs.forEach((p, i) => {
    const pValue = jsString(p.text ?? '')
    if (i === 0) {
      shapeLines.push('var oPara = oContent.GetElement(0);', 'oPara.RemoveAllElements();')
    } else {
      shapeLines.push('var oPara = Api.CreateParagraph();')
    }
    shapeLines.push(
      'var oRun = Api.CreateRun();',
      `oRun.AddText(${pValue});`,
      ...(p.size ? [`oRun.SetFontSize(${Number(p.size)});`] : []),
      ...(p.bold ? ['oRun.SetBold(true);'] : []),
      'oPara.AddElement(oRun);',
    )
    applyPptxParagraphListType(shapeLines, p)
    applyPptxParagraphIndent(shapeLines, p)
    // Push AFTER the paragraph is fully built (matches the docx branch's own order --
    // oDocument.Push(oParagraph) is its last line too), not before: the first paragraph
    // needs no push at all (GetElement(0) already sits in the shape's content tree).
    if (i > 0) shapeLines.push('oContent.Push(oPara);')
  })
  shapeLines.push('oSlide.AddObject(oShape);')
  return shapeLines
}

// `comments` -- same MoveCursorToEnd() dependency as applyDocxParagraphNotes above, split into
// its own function for the same complexity reason. The shape differs from footnotes/endnotes in
// one way worth calling out: AddComment(text, author, initials) anchors its
// <w:commentRangeStart>/<w:commentRangeEnd> around ONLY THE LAST WORD before the cursor, not the
// whole paragraph (package-verified: two-word paragraph "main text", the range wrapped "text"
// alone). `author`/`initials` are optional -- a call with just the body text omits `w:author`
// from the saved comments.xml entirely rather than writing an empty string.
//
// MUST run before footnotes/endnotes are added to the same paragraph -- see the order-trap note
// at the top of applyDocxParagraphNotes above.
function applyDocxParagraphComments(lines, op) {
  for (const c of Array.isArray(op.comments) ? op.comments : []) {
    const body = typeof c === 'string' ? c : c?.text
    const author = typeof c === 'object' && c?.author !== undefined ? c.author : undefined
    const args = [jsString(String(body ?? ''))]
    if (author !== undefined) args.push(jsString(String(author)))
    lines.push('oDocument.MoveCursorToEnd();', `oDocument.AddComment(${args.join(', ')});`)
  }
}

// The `table` operation's pptx-only cell-formatting fields -- split out of that operation's emit() from the start,
// same reasoning as emitDocxRun below (this file's table.emit is already its largest function).
// Applied UNIFORMLY to every cell, same "table-wide" shape as `header`/`zebra`/`border` above --
// a per-cell override schema was explicitly out of scope for this unit (see the card).
//
// All three call shapes recovered via toString() on the live DocBuilder instance (this tool's
// own established methodology), NOT assumed from the method name:
//   SetCellMarginTop/Bottom/Left/Right(nValue) -- nValue is TWIPS (the function's own source
//     calls private_GetTableMeasure("twips", nValue)); package-verified exact: 300 twips ->
//     marL/marR="190499"/"190500" EMU, 500 twips -> marT/marB="317500" EMU (300*635=190500,
//     500*635=317500 -- the standard twip->EMU constant, confirmed, not assumed).
//   SetVerticalAlign(sType) -- an unrecognised value does NOT throw or fall back, it silently
//     leaves VAlign unset (the source's own if/else-if chain has no default branch) -- allowlist
//     of the three it recognises: "top" | "bottom" | "center". Package-verified: "center" ->
//     anchor="ctr".
//   SetTextDirection(sType) -- same silent-fallthrough shape; allowlist "lrtb" | "tbrl" | "btlr".
//     Package-verified: "tbrl" -> vert="eaVert" (an internal OOXML enum name, not the input
//     string literal -- expected, not a bug).
//
// NAMED OUT (measured, not guessed): TableRow.SetHeight exists and runs, but its stored value
// does not track the requested one by any unit conversion this tool could establish -- 500 and
// 1500 (EMU-scale inputs) both landed as h="36000", while 900000 landed as h="783160" (LESS than
// requested, not a simple clamp-up-to-a-floor either). This looks like a content-/table-size-
// driven computed minimum, not a settable value -- refused here rather than exposed as a field
// that would silently ignore most of its own input. Follow-up measurement, not this card.
const KNOWN_PPTX_CELL_VALIGNS = ['top', 'bottom', 'center']
const KNOWN_CELL_TEXT_DIRECTIONS = ['lrtb', 'tbrl', 'btlr']

function forEachCell(rowCount, colCount, fn) {
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) fn(r, c)
  }
}

function applyPptxCellMargin(lines, cellMargin, rowCount, colCount, cella) {
  if (!cellMargin || typeof cellMargin !== 'object') return
  for (const [side, setter] of [['top', 'SetCellMarginTop'], ['bottom', 'SetCellMarginBottom'], ['left', 'SetCellMarginLeft'], ['right', 'SetCellMarginRight']]) {
    const value = cellMargin[side]
    if (value === undefined || value === null) continue
    const twips = Number(value)
    if (!Number.isInteger(twips) || twips < 0) {
      throw new Error(`table: cellMargin.${side} must be a non-negative integer (TWIP), got ${JSON.stringify(value)}`)
    }
    forEachCell(rowCount, colCount, (r, c) => lines.push(`${cella(r, c)}.${setter}(${twips});`))
  }
}

function applyPptxCellVAlign(lines, cellVAlign, rowCount, colCount, cella) {
  if (cellVAlign === undefined || cellVAlign === null) return
  if (!KNOWN_PPTX_CELL_VALIGNS.includes(cellVAlign)) {
    throw notSupportedError(`table: unknown cellVAlign ${JSON.stringify(cellVAlign)} (known: ${KNOWN_PPTX_CELL_VALIGNS.join(', ')}) -- an unrecognised value does NOT throw, it silently leaves the alignment unset, so this tool refuses it here instead`)
  }
  forEachCell(rowCount, colCount, (r, c) => lines.push(`${cella(r, c)}.SetVerticalAlign(${jsString(cellVAlign)});`))
}

function applyPptxCellTextDirection(lines, cellTextDirection, rowCount, colCount, cella) {
  if (cellTextDirection === undefined || cellTextDirection === null) return
  if (!KNOWN_CELL_TEXT_DIRECTIONS.includes(cellTextDirection)) {
    throw notSupportedError(`table: unknown cellTextDirection ${JSON.stringify(cellTextDirection)} (known: ${KNOWN_CELL_TEXT_DIRECTIONS.join(', ')}) -- an unrecognised value does NOT throw, it silently leaves the direction unset, so this tool refuses it here instead`)
  }
  forEachCell(rowCount, colCount, (r, c) => lines.push(`${cella(r, c)}.SetTextDirection(${jsString(cellTextDirection)});`))
}

function applyPptxTableCellSettings(lines, op, rowCount, colCount, cella) {
  applyPptxCellMargin(lines, op.cellMargin, rowCount, colCount, cella)
  applyPptxCellVAlign(lines, op.cellVAlign, rowCount, colCount, cella)
  applyPptxCellTextDirection(lines, op.cellTextDirection, rowCount, colCount, cella)
}

// The `table` operation's pptx-only `merge` field. Split out
// of that operation's emit() -- qlty smells already flagged it (already the file's largest
// function before this addition, at count=50) and this addition alone pushed it to 65; split keeps
// it off the smells report entirely, same reasoning as emitDocxRun below.
//
// `Table.MergeCells` was already reflected as existing (this tool's own pptx API inventory) --
// the SCOPE decision above ("MI NEM TARTOZIK BELE for docx/pptx") was never an API limit. The
// real call shape had to be recovered via toString() on this DocBuilder instance: it is NOT
// MergeCells(startRow, startCol, endRow, endCol) (measured: runs, returns no error, writes
// NOTHING to the saved package -- the exact "answers green and changes nothing" class this
// file's own header comment warns about). The real shape: MergeCells([cell1, cell2, ...]), one
// ApiTableCell per cell to merge, returning the merged ApiTableCell on success or null on
// failure. Package-verified: <a:tc gridSpan="2" .../><a:tc hMerge="1"/> present after a 2-cell
// merge, absent before.
//
// Rectangular ranges only (a caller listing a non-rectangular cell set gets whatever
// MergeTableCells does with it, which this tool has not separately measured) -- each
// [startRow, startCol, endRow, endCol] tuple expands to every cell in that rectangle.
function applyPptxTableMerge(lines, op, rowCount, colCount, cella) {
  if (!Array.isArray(op.merge) || !op.merge.length) return
  op.merge.forEach((range, idx) => {
    if (!Array.isArray(range) || range.length !== 4 || !range.every((n) => Number.isInteger(n) && n >= 0)) {
      throw new Error(`table: merge[${idx}] must be a [startRow, startCol, endRow, endCol] tuple of non-negative integers for pptx, got ${JSON.stringify(range)}`)
    }
    const [r1, c1, r2, c2] = range
    if (r2 < r1 || c2 < c1) throw new Error(`table: merge[${idx}] end must not be before start, got ${JSON.stringify(range)}`)
    if (r2 >= rowCount || c2 >= colCount) throw new Error(`table: merge[${idx}] end (${r2},${c2}) is outside the table (${rowCount} rows x ${colCount} cols)`)
    const cells = []
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) cells.push(cella(r, c))
    }
    lines.push(`oTable.MergeCells([${cells.join(', ')}]);`)
  })
}


// The `chart` operation's pptx-only formatting fields (identified via this tool's own pptx API
// inventory audit) -- split out of that
// operation's emit() for the same reason emitDocxRun below was: keeps emit()'s own branching
// complexity down (qlty smells flagged it at count=28 before this split). All five fields
// optional, none fires unless the caller asks -- same discipline as
// applyParagraphIndentSpacing above. Package-verified against the live DocBuilder instance
// (chart1.xml, ppt/charts/): each element present when set, absent on an unset/untouched fixture.
function applyChartFormatting(lines, op) {
  if (op.title !== undefined && op.title !== null) {
    lines.push(`oChart.SetTitle(${jsString(String(op.title))});`)
  }
  if (op.legendPos !== undefined && op.legendPos !== null) {
    if (!KNOWN_LEGEND_POSITIONS.includes(op.legendPos)) {
      throw notSupportedError(`chart: unknown legendPos ${JSON.stringify(op.legendPos)} (known: ${KNOWN_LEGEND_POSITIONS.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently produces no legend at all, so this tool refuses it here instead`)
    }
    lines.push(`oChart.SetLegendPos(${jsString(op.legendPos)});`)
  }
  if (op.horAxisTitle !== undefined && op.horAxisTitle !== null) {
    lines.push(`oChart.SetHorAxisTitle(${jsString(String(op.horAxisTitle))});`)
  }
  if (op.verAxisTitle !== undefined && op.verAxisTitle !== null) {
    lines.push(`oChart.SetVerAxisTitle(${jsString(String(op.verAxisTitle))});`)
  }
  if (op.showDataLabels) {
    // Measured (6-position sweep, one chart per position, package-checked): the real signature
    // is SetShowDataLabels(showSerName, showCatName, showVal, ...3 more with NO observed effect
    // on this DocBuilder instance). Position index 2 is `showVal` -- the conventional meaning of
    // "data labels" (the number on each point), so that is the one this boolean turns on. The
    // other five positions are left named-but-unexposed for a future unit rather than guessed at.
    lines.push('oChart.SetShowDataLabels(false, false, true, false, false, false);')
  }
}

// One run's worth of the `runs` operation's docx branch: create it, add its text, apply whichever
// optional fields this entry sets, append it to the paragraph. Split out of that operation's
// emit() purely to keep emit()'s own branching complexity
// down -- this is not a general-purpose run builder, it is that one branch's body moved one level
// out; the pptx branch (a different call shape per field, see that branch's own comments) stays
// inline, unchanged.
function emitDocxRun(v, r, lines) {
  lines.push(`var ${v} = Api.CreateRun();`, `${v}.AddText(${jsString(String(r.text ?? ''))});`)
  if (r.bold) lines.push(`${v}.SetBold(true);`)
  if (r.italic) lines.push(`${v}.SetItalic(true);`)
  if (r.underline) lines.push(`${v}.SetUnderline(true);`)
  if (r.strikethrough) lines.push(`${v}.SetStrikeout(true);`)
  if (r.size) lines.push(`${v}.SetFontSize(${Number(r.size)});`)
  if (r.color !== undefined && r.color !== null) {
    const { r: rr, g, b } = parseHexColor(r.color)
    // Run-level SetColor on docx takes a fourth (bool) argument -- measured, euro-demo-docx.js:94
    // -- unlike the pptx branch's three-byte run-level call.
    lines.push(`${v}.SetColor(${rr}, ${g}, ${b}, false);`)
  }
  if (r.highlight !== undefined && r.highlight !== null) {
    if (!KNOWN_HIGHLIGHT_COLORS.includes(r.highlight)) {
      throw notSupportedError(`runs: unknown highlight ${JSON.stringify(r.highlight)} (known: ${KNOWN_HIGHLIGHT_COLORS.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently produces no highlight at all, so this tool refuses it here instead (same allowlist as \`text\`'s docx highlight field)`)
    }
    lines.push(`${v}.SetHighlight(${jsString(r.highlight)});`)
  }
  if (r.vertAlign !== undefined && r.vertAlign !== null) {
    if (!KNOWN_VERT_ALIGNS.includes(r.vertAlign)) {
      throw notSupportedError(`runs: unknown vertAlign ${JSON.stringify(r.vertAlign)} (known: ${KNOWN_VERT_ALIGNS.join(', ')})`)
    }
    lines.push(`${v}.SetVertAlign(${jsString(r.vertAlign)});`)
  }
  lines.push(`oParagraph.AddElement(${v});`)
}

// The `table` operation's `merge` field, docx branch. The header
// comment on OPERATIONS.table used to call `merge` xlsx-only; that was a SCOPE decision (documented
// as "not included, for docx/pptx"), not an API limit -- `Table.MergeCells` was
// already visible in this tool's own reflection inventory.
//
// The real call shape had to be recovered by trial, not read off a doc page (the same caution
// this file's header applies everywhere): NOT MergeCells(startRow, startCol, endRow, endCol) --
// that shape was never tried against a working case here, so it is simply unmeasured, not ruled
// out. What IS measured and package-verified (2026-08-17, unzipped word/document.xml): a
// rectangular cell selection merges via `oTable.MergeCells([cell1, cell2, ...])`, one
// ApiDocumentTableCell per cell to merge. A horizontal merge (two cells, same row) produced
// `<w:gridSpan w:val="2"/>` on the surviving cell and removed the second `<w:tc>`; a vertical
// merge (two cells, same column, two rows) produced `<w:vMerge w:val="restart"/>` on the first
// row's cell and `<w:vMerge w:val="continue"/>` on the second. A negative-control table built the
// same way with no MergeCells call had zero gridSpan/vMerge occurrences in either case.
//
// Rectangular ranges only -- each [startRow, startCol, endRow, endCol] tuple expands to every
// cell in that rectangle before the single MergeCells call. A non-rectangular cell list's
// behaviour on this route is unmeasured, so it is not offered as an input shape here.
function applyDocxTableMerge(lines, op, rowCount, colCount, cella, varName = 'oTable') {
  if (!Array.isArray(op.merge) || !op.merge.length) return
  op.merge.forEach((range, idx) => {
    if (!Array.isArray(range) || range.length !== 4 || !range.every((n) => Number.isInteger(n) && n >= 0)) {
      throw new Error(`table: merge[${idx}] must be a [startRow, startCol, endRow, endCol] tuple of non-negative integers for docx, got ${JSON.stringify(range)}`)
    }
    const [r1, c1, r2, c2] = range
    if (r2 < r1 || c2 < c1) throw new Error(`table: merge[${idx}] end must not be before start, got ${JSON.stringify(range)}`)
    if (r2 >= rowCount || c2 >= colCount) throw new Error(`table: merge[${idx}] end (${r2},${c2}) is outside the table (${rowCount} rows x ${colCount} cols)`)
    const cells = []
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) cells.push(cella(r, c))
    }
    lines.push(`${varName}.MergeCells([${cells.join(', ')}]);`)
  })
}

// The `table` operation's `split` field, docx only. Real call
// shape MEASURED via `.toString()` on TWO objects (2026-08-17): `Table.Split(oCell, nRow, nCol)`
// AND `TableCell.Split(nRow, nCol)` (the cell method is a thin wrapper: `this.GetParentTable().
// Split(this, nRow, nCol)`), so the cell-level call is used here -- no separate table reference
// needed. *** THIS OVERTURNS koteg03's "Table.Split(): NULL, unusable" finding -- that call was
// tried with a ROW INDEX as the first argument (matching the field's own name, "split the table
// at this row"), which fails the engine's own `oCell instanceof ApiTableCell` guard and returns
// null before doing anything -- a wrong-arg-order silent-null, the same trap class as
// `AddBookmarkCrossRef`/`Table.AddCaption` earlier in this file, not a broken API. *** The real
// operation SPLITS ONE CELL into an nRow x nCol grid of new cells (Word's own "Split Cells"
// feature), not the table into two tables -- package-verified (2026-08-17): a 2-row table with a
// single cell split 2x2 came back with 4 `<w:tr>` and 10 `<w:tc>` total (started at 2x2 = 4
// `<w:tc>`), a real structural change, not a same-count no-op.
function applyDocxTableSplit(lines, op, rowCount, colCount, cella) {
  if (!Array.isArray(op.split) || !op.split.length) return
  // Multiple entries in one call are UNMEASURED for index drift: the first split changes the
  // table's own row/cell counts, so a second entry's (row, col) may no longer name the cell the
  // caller meant -- the same class of gap this file already names rather than solves for pptx
  // slide-structure operations (see OPERATIONS.slide's own comment). Bounds are checked against
  // the table's ORIGINAL shape for every entry regardless, which is still correct for the common
  // single-split call and at least catches an out-of-range request for later ones.
  op.split.forEach((entry, idx) => {
    if (!Array.isArray(entry) || entry.length !== 4 || !entry.every((n) => Number.isInteger(n))) {
      throw new Error(`table: split[${idx}] must be a [row, col, nRow, nCol] tuple of integers, got ${JSON.stringify(entry)}`)
    }
    const [row, col, nRow, nCol] = entry
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) {
      throw new Error(`table: split[${idx}] cell (${row},${col}) is outside the table (${rowCount} rows x ${colCount} cols)`)
    }
    if (nRow < 1 || nCol < 1) throw new Error(`table: split[${idx}] nRow/nCol must both be >= 1, got ${JSON.stringify(entry)}`)
    lines.push(`${cella(row, col)}.Split(${nRow}, ${nCol});`)
  })
}

// The `table` operation's `caption` field, docx only. Real call
// shape MEASURED via `.toString()`: `Table.AddCaption(sAdditional, sLabel, bExcludeLabel,
// sNumberingFormat, bBefore, nHeadingLvl, sCaptionSep)`. A first probe guessed `(sLabel,
// sAdditional)` -- wrong order, produced a garbled caption and left `AddTableOfFigures` unable to
// find it ("No table of figures entries found"). The correct order
// (`oTable.AddCaption("My caption text", "Figure")`) is package-verified: a real
// `SEQ Figure \* Arabic 1` caption field, later found by `tableOfFigures` below. Only the two
// measured positional arguments are emitted -- the other five (bExcludeLabel/sNumberingFormat/
// bBefore/nHeadingLvl/sCaptionSep) were never tried here, so this leaves them at the builder's
// own default rather than guessing a value for them.
function applyDocxTableCaption(lines, op, varName = 'oTable') {
  if (op.caption === undefined || op.caption === null) return
  const spec = typeof op.caption === 'string' ? { text: op.caption } : op.caption
  if (!spec || !spec.text) throw new Error('table: `caption.text` is required when `caption` is set')
  const label = spec.label ?? 'Table'
  lines.push(`${varName}.AddCaption(${jsString(String(spec.text))}, ${jsString(String(label))});`)
}

// A `table` op's docx-only cell-value extension: a `rows`
// cell entry may be a primitive (unchanged: AddText, byte-identical to before this addition) OR an
// object `{ table: {...} }`, marking that cell's content as a NESTED table built with this SAME
// function, recursively -- full feature parity with the top level (header/zebra/border/
// columnWidths/merge/rowHeights/cellMargins etc), not a narrower nested-only subset. The `table`
// key (not "is it an object") is the marker, so a future differently-shaped cell object (an
// image, say) does not collide with this one by accident.
//
// MEASURED 2026-08-17: three Api.CreateTable() calls nested via cell.GetContent().Push(), one
// inside the next, package-verified against a real DocBuilder round-trip -- 3 <w:tbl> in the
// saved word/document.xml, the innermost cell's own text intact. The code below is recursive and
// therefore has no hardcoded depth limit, but only 2 nesting levels (3 tables total) are
// package-verified so far -- deeper is NOT measured, not "does not work".
//
// Pulled out of `table`'s own emit() (which still handles xlsx/pptx inline) because a fully
// recursive nested-table path inside that already-large function would have pushed its
// complexity well past the qlty threshold; this is the entire docx table-construction body,
// parameterised on its own script variable name so a nested call can build `oNestedTableN`
// instead of colliding with the outer `oTable`.
function buildDocxTableLines(varName, spec, counterRef) {
  const rows = Array.isArray(spec.rows) ? spec.rows : []
  if (!rows.length) throw new Error('table: `rows` is empty -- an empty table cannot be told apart from a successful one')
  const header = spec.header !== false
  const zebra = Boolean(spec.zebra)
  const border = spec.border !== false
  const headerFill = spec.headerColor ?? [0x1f, 0x38, 0x64]
  const headerText = spec.headerTextColor ?? [0xff, 0xff, 0xff]
  const zebraFill = spec.zebraColor ?? [0xf7, 0xf9, 0xfc]
  const borderColor = spec.borderColor ?? [0x1a, 0x5f, 0xb4]
  const rgb = (c) => c.map(Number).join(', ')
  const cols = Math.max(...rows.map((r) => (Array.isArray(r) ? r.length : 1)))
  const cella = (i, j) => `${varName}.GetCell(${i}, ${j})`
  const lines = [`var ${varName} = Api.CreateTable(${cols}, ${rows.length});`]
  lines.push(`${varName}.SetWidth("percent", ${Number(spec.widthPercent ?? 100)});`)
  if (border) {
    for (const side of ['Top', 'Bottom', 'Left', 'Right', 'InsideH', 'InsideV']) {
      lines.push(`${varName}.SetTableBorder${side}("single", 4, 0, ${rgb(borderColor)});`)
    }
  }
  if (Array.isArray(spec.columnWidths) && spec.columnWidths.length) {
    lines.push(`${varName}.SetTableLayout("fixed");`)
  }
  const colWidths = Array.isArray(spec.columnWidths) ? spec.columnWidths : null
  validateDocxTableRefinements(spec)
  applyPptxTableLookAndStructuralExtras(lines, spec, 'docx')
  rows.forEach((sor, i) => {
    const cells = Array.isArray(sor) ? sor : [sor]
    const isHeader = header && i === 0
    const isZebra = !isHeader && zebra && i % 2 === 0
    cells.forEach((ertek, j) => {
      const cell = cella(i, j)
      if (ertek !== null && typeof ertek === 'object' && !Array.isArray(ertek) && ertek.table) {
        counterRef.n += 1
        const nestedVar = `oNestedTable${counterRef.n}`
        lines.push(...buildDocxTableLines(nestedVar, ertek.table, counterRef))
        lines.push(`${cell}.GetContent().Push(${nestedVar});`)
      } else {
        // Measured (receptek-pptx-docx.md #11 + live E2 probe): SetBold/SetColor/SetShd all
        // apply straight to the CELL PARAGRAPH here -- no separate Run object needed, unlike pptx.
        lines.push(`${cell}.GetContent().GetElement(0).AddText(${jsString(ertek)});`)
        if (isHeader) {
          lines.push(
            `${cell}.GetContent().GetElement(0).SetBold(true);`,
            `${cell}.GetContent().GetElement(0).SetColor(${rgb(headerText)}, false);`,
            `${cell}.SetShd("clear", ${rgb(headerFill)}, false);`,
          )
        } else if (isZebra) {
          lines.push(`${cell}.SetShd("clear", ${rgb(zebraFill)}, false);`)
        }
      }
      // E2b -- set on EVERY row's cell for the column, not just the header: a fixed-layout
      // table reads tcW per cell, and leaving later rows at their default risks an
      // inconsistent grid.
      if (colWidths && colWidths[j] != null) {
        lines.push(`${cell}.SetWidth("percent", ${Number(colWidths[j])});`)
      }
      applyDocxTableCellRefinements(lines, spec, cell)
    })
  })
  applyDocxTableMerge(lines, spec, rows.length, cols, cella, varName)
  applyDocxTableRowRefinements(lines, spec, varName)
  // caption is table-level (no cell indices to disturb); split runs LAST because it changes the
  // table's own row/cell counts, and the row refinements above still need the ORIGINAL row
  // indices (GetRow(0)/GetRow(i)) to be valid when they run. Applies the same way to a nested
  // table (this function's recursive branch above) as to the top-level one -- not independently
  // package-verified on a NESTED table specifically, only on the top-level `oTable` case.
  applyDocxTableCaption(lines, spec, varName)
  applyDocxTableSplit(lines, spec, rows.length, cols, cella)
  return lines
}

// The docx branch of `pageSetup`, pulled out to its own function
// for the same reason emitDocxRun above was: keeping it inline pushed pageSetup's emit() past the
// qlty complexity threshold, and a smell that only exists because two unrelated branches share one
// function body is not a smell about the LOGIC, it is about where the brace goes.
//
// MEASURED via a real DocBuilder round-trip (margins 1440/2880/4320/5760 + size 20000x10000),
// unzipped and read back from the saved word/document.xml -- not guessed from the OOXML schema:
//   SetPageMargins(left, top, right, bottom) -- FOUR args, this exact order, units are TWIPS
//     directly (1440 in -> saved pgMar left="1440"), no mm/inch conversion unlike the xlsx margin
//     setters in the sibling branch (those ARE in millimetres -- two different units on two cores
//     of the same conceptual operation, each matching what its own Set* method actually takes).
//   SetPageSize(width, height) -- TWO args, TWIPS. The saved pgSz's `w:orient` attribute is
//     DERIVED by the Document Server from width-vs-height (20000x10000 -> "landscape"
//     automatically) -- there is no separate SetOrientation on Section, unlike xlsx's confirmed
//     no-op SetPageOrientation in the sibling branch. This is why `orientation` below computes a
//     swap instead of calling a setter that does not exist here.
function pageSetupDocx(op) {
  const MARGIN_KEYS = ['marginLeft', 'marginTop', 'marginRight', 'marginBottom']
  const hasMargin = MARGIN_KEYS.some((k) => op[k] !== undefined)
  const hasExplicitSize = op.pageWidth !== undefined || op.pageHeight !== undefined
  const hasOrientation = op.orientation !== undefined
  if (hasExplicitSize && hasOrientation) {
    throw new Error('pageSetup (docx): `orientation` and `pageWidth`/`pageHeight` cannot both be given in one call -- orientation is defined relative to the CURRENT size, pass one or the other')
  }
  if (hasOrientation && op.orientation !== 'portrait' && op.orientation !== 'landscape') {
    throw new Error(`pageSetup (docx): orientation must be "portrait" or "landscape" -- got ${JSON.stringify(op.orientation)}`)
  }
  validatePageSetupDocxSection2(op)
  const hasSection2 = SECTION2_DOCX_KEYS.some((k) => op[k] !== undefined)
  if (!hasMargin && !hasExplicitSize && !hasOrientation && !hasSection2) {
    throw new Error(`pageSetup (docx): at least one of marginLeft, marginTop, marginRight, marginBottom, pageWidth, pageHeight, orientation, ${SECTION2_DOCX_KEYS.join(', ')} is required`)
  }
  const lines = ['var oSection = oDocument.GetFinalSection();']
  if (hasMargin) {
    // Section.SetPageMargins takes all four sides in one call (unlike xlsx's four separate
    // Set*Margin methods) -- a side the caller did NOT ask to change is read back from the
    // section at RUNTIME (GetPageMargin*()) and passed straight through, so "set only the left
    // margin" does not silently reset the other three to zero.
    const side = (key, getter) => (op[key] !== undefined ? Number(op[key]) : `oSection.${getter}()`)
    lines.push(
      `oSection.SetPageMargins(${side('marginLeft', 'GetPageMarginLeft')}, ${side('marginTop', 'GetPageMarginTop')}, ${side('marginRight', 'GetPageMarginRight')}, ${side('marginBottom', 'GetPageMarginBottom')});`,
    )
  }
  if (hasExplicitSize) {
    const w = op.pageWidth !== undefined ? Number(op.pageWidth) : 'oSection.GetPageWidth()'
    const h = op.pageHeight !== undefined ? Number(op.pageHeight) : 'oSection.GetPageHeight()'
    lines.push(`oSection.SetPageSize(${w}, ${h});`)
  } else if (hasOrientation) {
    // No SetOrientation exists on Section (see the block comment above) -- orientation is the
    // Document Server's own derived read of width-vs-height, so achieving it means reading the
    // CURRENT size at runtime and swapping the two numbers only if the current shape does not
    // already match what was asked.
    lines.push(
      'var __sw = oSection.GetPageWidth(), __sh = oSection.GetPageHeight();',
      `if ((__sw > __sh) !== ${op.orientation === 'landscape'}) { oSection.SetPageSize(__sh, __sw); }`,
    )
  }
  emitPageSetupDocxSection2(op, lines)
  return lines
}

// The SECOND round on Section, the six fields docx01 named and
// left for "the next unit": columns, header/footer distance, the title-page flag, start page
// number and section type. Split into its own validate+emit pair (not folded into the block
// above) for the same complexity reason docx01 was split out of emitDocxRun in the first place --
// six more independent fields inline would have pushed pageSetupDocx itself past the threshold.
//
// MEASURED via a real DocBuilder round-trip on a fresh single-section document (unzipped, read
// back from word/document.xml), not guessed from the OOXML schema:
//   SetHeaderDistance(n) / SetFooterDistance(n) -- TWIPS, one arg each, independent of each other
//     and of SetPageMargins (measured: header=400/footer=300 landed as pgMar's own w:header/
//     w:footer attributes, the four SetPageMargins sides untouched by these two calls).
//   SetTitlePage(bool) -- a genuine toggle, not "only true does anything": SetTitlePage(true)
//     followed by SetTitlePage(false) in the same script left NO <w:titlePg> in the saved package
//     (its presence/absence IS the flag; there is no w:val).
//   SetStartPageNumber(n) -- one arg, landed as <w:pgNumType w:start="n"/>.
//   SetType(sType) -- one arg, a STRING (not a number/enum constant): "continuous", "evenPage",
//     "oddPage", "nextColumn" and "nextPage" all individually round-tripped to the matching
//     <w:type w:val="..."/>. An unrecognised value is the SAME silent-no-op class as `align`/
//     `highlight`/`chartType` elsewhere in this file: SetType("nincsilyentipus") ran to
//     completion and the saved section kept w:val="nextPage" (the untouched default) -- no
//     exception, no error field in the docbuilder response. Hence the allowlist below rather than
//     passing the string straight through.
//   SetEqualColumns(count, space) -- TWO args, TWIPS for the second. The trap here is the
//     single-arg call: SetEqualColumns(3) alone (no space) saved w:space="0" -- columns flush
//     against each other, NOT the space a document with no column setting at all already has
//     (measured on the same fixture, no pageSetup call: w:space="1701"). Omitting the second arg
//     is therefore its own silent surprise, distinct from the allowlist class above -- the fix is
//     the same shape as the margin sides: default the space explicitly rather than ever calling
//     SetEqualColumns with one argument.
//   SetNotEqualColumns(aWidths, aSpaces) -- WORKS (re-measured
//     2026-08-17 against the real Document Server via run_builder_script). The
//     REAL signature, reflected off the live instance (`String(oSection.SetNotEqualColumns)`),
//     is TWO array arguments -- `aWidths` (twips, length >= 2) and `aSpaces` (twips, length ===
//     aWidths.length - 1, one gap per pair of adjacent columns). The function's own guard
//     (`if(!aWidths||!aWidths.length||aWidths.length<=1||aSpaces.length!==aWidths.length-1)
//     return false`) silently no-ops on a malformed shape rather than throwing -- it does NOT
//     kill the job. Package-verified with the correct shape: `SetNotEqualColumns([4000,4000],
//     [200])` returns `true` and the saved package carries a real `<w:cols w:num="2"
//     w:equalWidth="0"><w:col w:w="4000" w:space="200"/><w:col w:w="4000" w:space="0"/></w:cols>`.
//     *** THIS CONTRADICTS A PRIOR MEASUREMENT (dated the SAME DAY, which
//     reported all three shapes it tried -- including a `(count, [widths], [spaces])` 3-argument
//     shape that does not match the real 2-argument signature -- as outcome:"blocked", no saved
//     output, same failure class as Api.CreateNumbering(). *** A cross-check TODAY (the
//     same route, same tool) also found `Api.CreateNumbering()` alone no longer kills the job
//     either, where it reliably did earlier in the day across many independent measurements.
//     Whether this is the earlier bisection having used the wrong call shape, or the Document
//     Server's behaviour changing mid-day on this shared, actively-used instance, is NOT
//     determined here -- flagged as an open anomaly. What
//     IS re-measured and package-verified, today, right now, on the real instance: the correct
//     call shape works. See `columnWidths` below for how this is wired into the operation.
const KNOWN_SECTION_TYPES = ['continuous', 'evenPage', 'nextColumn', 'nextPage', 'oddPage']
const DEFAULT_COLUMN_SPACE_TWIPS = 1701 // measured: what a section with no column setting already has
const SECTION2_DOCX_KEYS = ['columns', 'columnSpacing', 'headerDistance', 'footerDistance', 'titlePage', 'startPageNumber', 'sectionType', 'columnWidths']

function validatePageSetupDocxSection2(op) {
  // `columnWidths` (asymmetric multi-column via Section.SetNotEqualColumns) -- see the MEASURED
  // note above the constants block for the real (aWidths, aSpaces) signature and evidence.
  if (op.columnWidths !== undefined) {
    if (op.columns !== undefined) {
      throw new Error('pageSetup (docx): `columns` and `columnWidths` cannot both be given -- `columns` asks for EQUAL-width columns (SetEqualColumns), `columnWidths` asks for explicit, possibly unequal widths (SetNotEqualColumns); pick one')
    }
    if (!Array.isArray(op.columnWidths) || op.columnWidths.length < 2) {
      throw new Error(`pageSetup (docx): \`columnWidths\` must be an array of at least 2 twip widths, got ${JSON.stringify(op.columnWidths)}`)
    }
    if (!op.columnWidths.every((w) => Number.isFinite(Number(w)) && Number(w) > 0)) {
      throw new Error(`pageSetup (docx): \`columnWidths\` entries must all be positive numbers (twips), got ${JSON.stringify(op.columnWidths)}`)
    }
  }
  if (op.columnSpacing !== undefined && op.columns === undefined && op.columnWidths === undefined) {
    throw new Error('pageSetup (docx): `columnSpacing` requires `columns` or `columnWidths` -- it sets the gap between columns, there is nothing to space without a column count')
  }
  if (op.columns !== undefined && (!Number.isInteger(op.columns) || op.columns < 1)) {
    throw new Error(`pageSetup (docx): \`columns\` must be a positive integer -- got ${JSON.stringify(op.columns)}`)
  }
  if (op.startPageNumber !== undefined && (!Number.isInteger(op.startPageNumber) || op.startPageNumber < 1)) {
    throw new Error(`pageSetup (docx): \`startPageNumber\` must be a positive integer -- got ${JSON.stringify(op.startPageNumber)}`)
  }
  if (op.sectionType !== undefined && !KNOWN_SECTION_TYPES.includes(op.sectionType)) {
    throw notSupportedError(`pageSetup (docx): unknown sectionType ${JSON.stringify(op.sectionType)} (known: ${KNOWN_SECTION_TYPES.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently leaves the section type at whatever it already was (measured 2026-08-16: SetType("nincsilyentipus") ran to completion, saved section kept w:val="nextPage")`)
  }
}

function emitPageSetupDocxSection2(op, lines) {
  if (op.columns !== undefined) {
    const space = op.columnSpacing !== undefined ? Number(op.columnSpacing) : DEFAULT_COLUMN_SPACE_TWIPS
    lines.push(`oSection.SetEqualColumns(${Number(op.columns)}, ${space});`)
  }
  if (op.columnWidths !== undefined) {
    // `aSpaces` needs exactly `aWidths.length - 1` entries (one gap per adjacent pair) -- the
    // SAME `columnSpacing` value is repeated for every gap, same "one number, applied uniformly"
    // shape `columns`+`columnSpacing` already uses for the equal-width path.
    const space = op.columnSpacing !== undefined ? Number(op.columnSpacing) : DEFAULT_COLUMN_SPACE_TWIPS
    const widths = op.columnWidths.map(Number)
    const spaces = new Array(widths.length - 1).fill(space)
    lines.push(`oSection.SetNotEqualColumns([${widths.join(', ')}], [${spaces.join(', ')}]);`)
  }
  if (op.headerDistance !== undefined) lines.push(`oSection.SetHeaderDistance(${Number(op.headerDistance)});`)
  if (op.footerDistance !== undefined) lines.push(`oSection.SetFooterDistance(${Number(op.footerDistance)});`)
  if (op.titlePage !== undefined) lines.push(`oSection.SetTitlePage(${op.titlePage ? 'true' : 'false'});`)
  if (op.startPageNumber !== undefined) lines.push(`oSection.SetStartPageNumber(${Number(op.startPageNumber)});`)
  if (op.sectionType !== undefined) lines.push(`oSection.SetType(${jsString(op.sectionType)});`)
}

// pptx-only in this unit (not probed against docx -- refused there
// rather than applied unverified, same discipline as the earlier hyperlink/rotation fields
// tonight). Split out of `table`'s emit() from the start (that function is already this file's
// largest, per its own comment) -- both fields recovered via toString(), package-verified. Two
// functions, not one -- the first draft (one function for both fields) still hit qlty smells.
function applyPptxTableLook(lines, tableLook) {
  // Table.SetTableLook(isFirstColumn, isFirstRow, isLastColumn, isLastRow, isHorBand,
  // isVerBand) -- POSITIONAL, column-before-row (toString()-recovered, package-verified:
  // {firstRow:true} alone -> firstRow="1", the other five stay "0"/default, confirmed against a
  // same-fixture table WITHOUT this call, whose own CreateTable defaults differ -- so this call
  // genuinely changes the packaged tblPr, not a coincidental match). Every sub-field defaults to
  // false when omitted -- no partial/inherit-builtin-default behaviour, same predictability as
  // the tablacella unit's cellMargin sides.
  if (typeof tableLook !== 'object' || tableLook === null) {
    throw new Error('table: tableLook must be an object (e.g. { firstRow: true })')
  }
  const flag = (key) => (tableLook[key] ? 'true' : 'false')
  lines.push(`oTable.SetTableLook(${flag('firstColumn')}, ${flag('firstRow')}, ${flag('lastColumn')}, ${flag('lastRow')}, ${flag('horBand')}, ${flag('verBand')});`)
}

function applyPptxTableStructuralExtras(lines, op) {
  // Table.AddColumn()/AddRow() called with NO arguments append at the END of the grid
  // (toString()-recovered: with no oCell given, the call falls back to the table's own last cell
  // and isBefore=false) -- new cells are blank, this tool does not (yet) accept content for them.
  // Package-verified: a 2x2 table with extraColumns:1, extraRows:1 saves as 3 rows / 3 cells-per-row.
  if (op.extraColumns !== undefined && op.extraColumns !== null) {
    const n = Number(op.extraColumns)
    if (!Number.isInteger(n) || n < 0) throw new Error(`table: extraColumns must be a non-negative integer, got ${JSON.stringify(op.extraColumns)}`)
    for (let k = 0; k < n; k++) lines.push('oTable.AddColumn();')
  }
  if (op.extraRows !== undefined && op.extraRows !== null) {
    const n = Number(op.extraRows)
    if (!Number.isInteger(n) || n < 0) throw new Error(`table: extraRows must be a non-negative integer, got ${JSON.stringify(op.extraRows)}`)
    for (let k = 0; k < n; k++) lines.push('oTable.AddRow();')
  }
}

function applyPptxTableLookAndStructuralExtras(lines, op, core) {
  const hasTableLook = op.tableLook !== undefined && op.tableLook !== null
  const hasExtraColumns = op.extraColumns !== undefined && op.extraColumns !== null
  const hasExtraRows = op.extraRows !== undefined && op.extraRows !== null
  if (core !== 'pptx' && (hasTableLook || hasExtraColumns || hasExtraRows)) {
    throw notSupportedError('table: tableLook/extraRows/extraColumns are pptx-only in this unit -- not verified for docx, refused rather than applied unverified')
  }
  if (hasTableLook) applyPptxTableLook(lines, op.tableLook)
  applyPptxTableStructuralExtras(lines, op)
}


// `shape` op fields that modify the shape AFTER Api.CreateShape (as opposed to `fill`/
// `borderWidth`/`borderColor`, which only apply AT creation) -- split out once this grew to four
// independent fields (hyperlink, rotation, verticalTextAlign, line) landing into the same spot;
// qlty smells flagged shape.emit itself once all four were
// inline together, this extraction is the fix, not a stylistic preference.
function applyPptxShapePostCreationFields(lines, op, core) {
  const rgb = (c) => c.map(Number).join(', ')
  if (op.hyperlink !== undefined && op.hyperlink !== null) {
    const { url, tooltip } = validateHyperlink(op.hyperlink, 'shape')
    lines.push(`oShape.SetHyperlink(Api.CreateHyperlink(${jsString(url)}, ${jsString(tooltip)}));`)
  }
  // SetRotation takes DEGREES (package-verified: 45 -> rot="2700000", the OOXML 60000ths-of-
  // a-degree unit -- the builder does that conversion itself, this tool passes plain degrees).
  // Arithmetic, not a keyword lookup -- any finite number is accepted, same reasoning as
  // parseHexColor's own comment on why hyperlink/color fields don't need an allowlist here.
  // KNOWN LIMITATION, named not silently accepted: SetRotation returns false if the shape
  // cannot be rotated (`!this.Drawing.canRotate()`) -- this tool's generated script is
  // fire-and-forget (matching every other Set* call in this file) and cannot observe that
  // return value, so a shape state that refuses rotation would silently produce no rotation
  // with no error surfaced. Not probed further -- out of this unit's scope.
  if (op.rotation !== undefined && op.rotation !== null) {
    const rotation = Number(op.rotation)
    if (!Number.isFinite(rotation)) throw new Error(`shape: rotation must be a finite number (degrees), got ${JSON.stringify(op.rotation)}`)
    lines.push(`oShape.SetRotation(${rotation});`)
  }
  if (op.verticalTextAlign !== undefined && op.verticalTextAlign !== null) {
    lines.push(`oShape.SetVerticalTextAlign(${jsString(op.verticalTextAlign)});`)
  }
  // `line` updates the border AFTER creation -- distinct from
  // `borderWidth`/`borderColor` above, which only apply at CreateShape time. toString()-
  // recovered (Shape.SetLine(oStroke)): silently returns false on a non-Stroke argument or a
  // shape with no spPr, but this tool always constructs a real Stroke via Api.CreateStroke,
  // so that failure mode does not apply here. Package-verified against a same-fixture shape
  // WITHOUT this call (kept its own CreateStroke-time border unchanged): width/color both
  // land exactly as given, distinct from the creation-time value.
  if (op.line !== undefined && op.line !== null) {
    if (core !== 'pptx') {
      throw notSupportedError('shape: line is pptx-only in this unit -- not verified for docx, refused rather than applied unverified')
    }
    if (typeof op.line !== 'object') throw new Error('shape: line must be an object (e.g. { width: 20000, color: [255,0,0] })')
    const lineWidth = Number(op.line.width ?? 0)
    const lineFill = op.line.color ? `Api.CreateSolidFill(Api.CreateRGBColor(${rgb(op.line.color)}))` : 'Api.CreateNoFill()'
    lines.push(`oShape.SetLine(Api.CreateStroke(${lineWidth}, ${lineFill}));`)
  }
  applyShapePlaceholder(lines, op, core)
}

// Split out of applyPptxShapePostCreationFields for the same reason as resolveShapeFill/
// buildGradientStop: keeps that function's own branch count from absorbing this field's
// validation (qlty-smells discipline).
//
// `placeholder` -- pptx-only in this unit (placeholders are a slide-layout concept, no docx
// call shape was probed or is expected to apply). Api.CreatePlaceholder(sType) ->
// Shape.SetPlaceholder(oPh), package-verified: <p:ph type="..." idx="N"/> lands in nvSpPr.
// MEASURED, NOT ASSUMED, that CreatePlaceholder accepts an allowlist smaller than the full
// OOXML placeholder-type vocabulary -- its own source (toString()-recovered) does `if (typeof
// sType !== "string") sType = "body"`, which reads as "unknown strings fall back to body" but
// does NOT say WHICH strings that private inner-type mapper actually recognises. Empirically
// round-tripped all 14 OOXML placeholder type strings through Create+Set+save-and-inspect-the-
// XML: only 7 survive as themselves (title, body, ctrTitle, subTitle, chart, clipArt, media) --
// the other 7 (dt, ftr, sldNum, pic, tbl, obj, dgm, sldImg) SetPlaceholder still returns true
// for (no error, no signal) and silently land as type="body" instead. Same failure shape as
// verticalTextAlign above: a wrong value does not throw, so this tool refuses it here rather
// than emit a call that silently produces a different result than asked.
function applyShapePlaceholder(lines, op, core) {
  if (op.placeholder === undefined || op.placeholder === null) return
  if (core !== 'pptx') {
    throw notSupportedError('shape: placeholder is pptx-only in this unit -- not applicable to docx (placeholders are a slide-layout concept), refused rather than applied unverified')
  }
  if (!KNOWN_PLACEHOLDER_TYPES.includes(op.placeholder)) {
    throw notSupportedError(`shape: unknown placeholder type ${JSON.stringify(op.placeholder)} (known: ${KNOWN_PLACEHOLDER_TYPES.join(', ')}) -- other OOXML placeholder-type strings do NOT throw inside the builder, they silently fall back to "body", so this tool refuses them here instead`)
  }
  lines.push(`oShape.SetPlaceholder(Api.CreatePlaceholder(${jsString(op.placeholder)}));`)
}

// (D2): the `shape` op's fill argument -- solid (existing), or
// gradient/pattern (NEW, DOCX-ONLY IN THIS UNIT). Split out of `shape.emit()` for the same
// complexity reason as `applyPptxShapePostCreationFields` above (qlty: emit() crossed the
// high-complexity threshold once this branching was added inline). Both new call shapes are
// MEASURED via toString() against a docx seed, package-verified (<a:gradFill>/<a:pattFill>
// present) -- never probed against a pptx seed, so refused there rather than applied unverified,
// same discipline as `hyperlink`/`rotation`/`verticalTextAlign` being pptx-only above.
// (K8): `fill` gained a SECOND way to ask for gradient/pattern,
// independently of the `fillGradient`/`fillPattern` fields above (D2, docx-only) -- an object
// form, `{type:'gradient'|'pattern', ...}` (see buildFillExpression's own header comment),
// MEASURED against a pptx seed only. Deliberately NOT merged into one field/one syntax across
// cores: D2's fillGradient/fillPattern are already shipped and tested for docx, and this unit
// never re-measured THAT call shape against a pptx seed (or vice versa) -- so each stays gated to
// the core it was actually verified against, same discipline as `hyperlink`/`rotation`/`line`
// above. A caller wanting gradient/pattern on docx uses fillGradient/fillPattern; on pptx uses
// `fill: {type:...}`; crossing the two is refused, not guessed at.
function resolveShapeObjectFill(op, core) {
  if (core !== 'pptx') {
    throw notSupportedError('shape: fill as a {type:...} object is pptx-only in this unit -- docx has its own fillGradient/fillPattern fields instead (measured separately, D2), refused rather than applied unverified')
  }
  return buildFillExpression(op.fill, 'shape')
}

// Extracted from resolveShapeFill's fillGradient branch so the nested-map return doesn't count
// toward that function's own "many returns" smell (qlty counts a closure's return against its
// enclosing function -- see nested-function-still-counts-toward-parent-complexity skill).
function buildGradientStop(s, i, rgb) {
  validateRgbColor(`shape: fillGradient.stops[${i}].color`, s.color)
  const pos = Number(s.pos ?? 0)
  return `Api.CreateGradientStop(Api.CreateRGBColor(${rgb(s.color)}), ${pos})`
}

function resolveShapeFill(op, core, rgb) {
  const fillFields = ['fill', 'fillGradient', 'fillPattern'].filter((k) => op[k] !== undefined && op[k] !== null)
  if (fillFields.length > 1) {
    throw new Error(`shape: only one of fill/fillGradient/fillPattern may be given -- got ${fillFields.join(', ')}`)
  }
  if ((op.fillGradient || op.fillPattern) && core !== 'docx') {
    throw notSupportedError('shape: fillGradient/fillPattern are docx-only in this unit -- not verified for pptx (measured against a docx seed only), refused rather than applied unverified')
  }
  if (op.fill !== undefined && op.fill !== null && typeof op.fill === 'object' && !Array.isArray(op.fill)) {
    return resolveShapeObjectFill(op, core)
  }
  if (op.fill) return `Api.CreateSolidFill(Api.CreateRGBColor(${rgb(op.fill)}))`
  if (op.fillGradient) {
    if (!Array.isArray(op.fillGradient.stops) || op.fillGradient.stops.length < 2) {
      throw new Error('shape: fillGradient.stops must be an array of at least 2 {color, pos} entries')
    }
    const stopVars = op.fillGradient.stops.map((s, i) => buildGradientStop(s, i, rgb))
    const angle = Number(op.fillGradient.angle ?? 0)
    return `Api.CreateLinearGradientFill([${stopVars.join(', ')}], ${angle})`
  }
  if (op.fillPattern) {
    if (!op.fillPattern.patternType) throw new Error('shape: fillPattern.patternType is required (e.g. "pct25")')
    validateRgbColor('shape: fillPattern.bgColor', op.fillPattern.bgColor)
    validateRgbColor('shape: fillPattern.fgColor', op.fillPattern.fgColor)
    return `Api.CreatePatternFill(${jsString(op.fillPattern.patternType)}, Api.CreateRGBColor(${rgb(op.fillPattern.bgColor)}), Api.CreateRGBColor(${rgb(op.fillPattern.fgColor)}))`
  }
  return 'Api.CreateNoFill()'
}

// `geometry` (custom point-path
// outline) was measured working against a docx seed first (D2) and, separately, against a pptx
// seed here (K9) -- same call shape on both cores: Api.CreateCustomGeometry() ->
// oGeom.AddPath() -> MoveTo/LineTo/Close -> oShape.SetGeometry(oGeom), package-verified
// <a:custGeom>/<a:pathLst>/<a:moveTo>/<a:lnTo>/<a:close> on EACH core independently (not assumed
// transferable -- see the resolveShapeFill-era note on why one core's result never carries over
// to another in this file). xlsx has no shape geometry surface to probe against. Split out of
// `shape.emit()` for the same complexity reason as `resolveShapeFill` above.
function applyShapeGeometry(lines, op, core) {
  if (op.geometry === undefined || op.geometry === null) return
  if (core !== 'docx' && core !== 'pptx') {
    throw notSupportedError('shape: geometry is docx/pptx-only in this unit -- not verified for xlsx, refused rather than applied unverified')
  }
  if (!Array.isArray(op.geometry.path) || op.geometry.path.length === 0) {
    throw new Error('shape: geometry.path must be a non-empty array of {cmd, x, y} path commands')
  }
  lines.push('var oShapeGeom = Api.CreateCustomGeometry();', 'var oShapePath = oShapeGeom.AddPath();')
  op.geometry.path.forEach((cmd, i) => {
    if (!cmd || typeof cmd !== 'object' || !cmd.cmd) throw new Error(`shape: geometry.path[${i}] must be an object with a \`cmd\` field`)
    if (cmd.cmd === 'moveTo') lines.push(`oShapePath.MoveTo(${Number(cmd.x)}, ${Number(cmd.y)});`)
    else if (cmd.cmd === 'lineTo') lines.push(`oShapePath.LineTo(${Number(cmd.x)}, ${Number(cmd.y)});`)
    else if (cmd.cmd === 'close') lines.push('oShapePath.Close();')
    else throw new Error(`shape: geometry.path[${i}].cmd unknown ${JSON.stringify(cmd.cmd)} (known: moveTo, lineTo, close)`)
  })
  lines.push('oShape.SetGeometry(oShapeGeom);')
}

// The operation table is data, not a switch: adding a capability means adding a row, and the
// "which cores support it" question has exactly one place to be answered.


function resolveSlideBackgroundLine(bg) {
  if (bg === 'none') return 'oSlide.ClearBackground();'
  if (bg === 'layout') return 'oSlide.FollowLayoutBackground();'
  if (bg === 'master') return 'oSlide.FollowMasterBackground();'
  if (Array.isArray(bg)) {
    validateRgbColor('slide: background', bg)
    return `oSlide.SetBackground(Api.CreateSolidFill(Api.CreateRGBColor(${rgbArg(bg)})));`
  }
  throw new Error(`slide: background must be "none"/"layout"/"master" or an [r,g,b] array, got ${JSON.stringify(bg)}`)
}

function buildSlideRemoveObjectLine(spec) {
  if (!spec || typeof spec !== 'object') throw new Error('slide: removeObject must be an object with a `pos` field')
  const pos = Number(spec.pos)
  if (!Number.isInteger(pos) || pos < 0) throw new Error(`slide: removeObject.pos must be a non-negative integer, got ${JSON.stringify(spec.pos)}`)
  const count = spec.count === undefined || spec.count === null ? 1 : Number(spec.count)
  if (!Number.isInteger(count) || count < 1) throw new Error(`slide: removeObject.count must be a positive integer, got ${JSON.stringify(spec.count)}`)
  return `oSlide.RemoveObject(${pos}, ${count});`
}

// `transition` -- SetEntryEffect/SetSpeed both silently no-op on an unrecognised value (measured
// from the ApiSlideShowTransition source: SetEntryEffect returns false when the name is not a key
// of ENTRY_EFFECT_MAP, SetSpeed via _getSpeedValue's `default: return undefined`), so both are
// allowlisted here rather than passed through -- same discipline as `shape.verticalTextAlign`.
function buildSlideTransitionLines(t) {
  if (!t || typeof t !== 'object') throw new Error('slide: transition must be an object')
  const lines = ['var oTransition = Api.CreateSlideShowTransition();']
  let any = false
  if (t.effect !== undefined && t.effect !== null) {
    if (!KNOWN_TRANSITION_EFFECTS.includes(t.effect)) {
      throw new Error(`slide: unknown transition.effect ${JSON.stringify(t.effect)} (known: ${KNOWN_TRANSITION_EFFECTS.join(', ')}) -- an unrecognised value does NOT throw inside the builder, SetEntryEffect silently returns false and leaves the transition without an effect, so this tool refuses it here instead`)
    }
    lines.push(`oTransition.SetEntryEffect(${jsString(t.effect)});`)
    any = true
  }
  if (t.speed !== undefined && t.speed !== null) {
    if (!KNOWN_TRANSITION_SPEEDS.includes(t.speed)) {
      throw new Error(`slide: unknown transition.speed ${JSON.stringify(t.speed)} (known: ${KNOWN_TRANSITION_SPEEDS.join(', ')})`)
    }
    lines.push(`oTransition.SetSpeed(${jsString(t.speed)});`)
    any = true
  }
  if (t.duration !== undefined && t.duration !== null) {
    const d = Number(t.duration)
    if (!Number.isFinite(d) || d < 0) throw new Error(`slide: transition.duration must be a non-negative number (milliseconds), got ${JSON.stringify(t.duration)}`)
    lines.push(`oTransition.SetDuration(${d});`)
    any = true
  }
  if (t.advanceOnClick !== undefined && t.advanceOnClick !== null) {
    lines.push(`oTransition.SetAdvanceOnClick(${Boolean(t.advanceOnClick)});`)
    any = true
  }
  if (t.advanceOnTime !== undefined && t.advanceOnTime !== null) {
    lines.push(`oTransition.SetAdvanceOnTime(${Boolean(t.advanceOnTime)});`)
    any = true
  }
  if (t.advanceTime !== undefined && t.advanceTime !== null) {
    const at = Number(t.advanceTime)
    if (!Number.isFinite(at) || at < 0) throw new Error(`slide: transition.advanceTime must be a non-negative number (milliseconds), got ${JSON.stringify(t.advanceTime)}`)
    lines.push(`oTransition.SetAdvanceTime(${at});`)
    any = true
  }
  if (!any) throw new Error('slide: transition object must set at least one of effect/speed/duration/advanceOnClick/advanceOnTime/advanceTime')
  lines.push('oSlide.SetSlideShowTransition(oTransition);')
  return lines
}

function buildSlideGroupShapeLines(shapeSpec, idx) {
  if (!shapeSpec || typeof shapeSpec !== 'object') throw new Error(`slide: group[${idx}] must be an object`)
  const shapeType = String(shapeSpec.shapeType ?? 'rect')
  const w = Number(shapeSpec.width ?? 1000000)
  const h = Number(shapeSpec.height ?? 1000000)
  const x = Number(shapeSpec.x ?? 0)
  const y = Number(shapeSpec.y ?? 0)
  let fill = 'Api.CreateNoFill()'
  if (shapeSpec.fill) {
    validateRgbColor(`slide: group[${idx}].fill`, shapeSpec.fill)
    fill = `Api.CreateSolidFill(Api.CreateRGBColor(${rgbArg(shapeSpec.fill)}))`
  }
  const v = `oGroupShape${idx}`
  return {
    varName: v,
    lines: [
      `var ${v} = Api.CreateShape(${jsString(shapeType)}, ${w}, ${h}, ${fill}, Api.CreateStroke(0, Api.CreateNoFill()));`,
      `${v}.SetPosition(${x}, ${y});`,
      `oSlide.AddObject(${v});`,
    ],
  }
}

function buildSlideGroupLines(group) {
  if (!Array.isArray(group) || group.length < 2) {
    throw new Error(`slide: group must be an array of at least 2 shape descriptors (grouping fewer than 2 shapes is not a group) -- got ${JSON.stringify(group)}`)
  }
  const lines = []
  const varNames = []
  group.forEach((shapeSpec, idx) => {
    const built = buildSlideGroupShapeLines(shapeSpec, idx)
    varNames.push(built.varName)
    lines.push(...built.lines)
  })
  lines.push(`oSlide.GroupDrawings([${varNames.join(', ')}]);`)
  return lines
}

// Measured allowlist for the `slide` operation's `transition.effect` field -- the complete
// ApiSlideShowTransition.ENTRY_EFFECT_MAP key set, toString()-recovered from the live Document
// Server, not typed from documentation. SetEntryEffect silently returns false on an unrecognised
// name (measured from its own source), so this must be an allowlist, not a passthrough.
// Measured allowlist for the `wordArt` operation's `transform` field
// -- the complete OOXML ST_TextShapeType set (DrawingML spec, a fixed enum). Api.CreateWordArt's
// own source silently substitutes "textNoShape" for anything that is not a non-empty string
// (measured from its toString()), so an unrecognised name would not throw -- it would silently
// produce a DIFFERENT WordArt shape than requested. This unit package-verified only "textNoShape"
// live; the rest of the list is the formal spec set, same discipline as KNOWN_PATTERN_TYPES above.
const KNOWN_WORDART_TRANSFORMS = ['textNoShape', 'textPlain', 'textStop', 'textTriangle', 'textTriangleInverted', 'textChevron', 'textChevronInverted', 'textRingInside', 'textRingOutside', 'textArchUp', 'textArchDown', 'textCircle', 'textButton', 'textArchUpPour', 'textArchDownPour', 'textCirclePour', 'textButtonPour', 'textCurveUp', 'textCurveDown', 'textCanUp', 'textCanDown', 'textWave1', 'textWave2', 'textDoubleWave1', 'textWave4', 'textInflate', 'textDeflate', 'textInflateBottom', 'textDeflateBottom', 'textInflateTop', 'textDeflateTop', 'textDeflateInflate', 'textDeflateInflateDeflate', 'textFadeRight', 'textFadeLeft', 'textFadeUp', 'textFadeDown', 'textSlantUp', 'textSlantDown', 'textCascadeUp', 'textCascadeDown']

const KNOWN_TRANSITION_EFFECTS = ['effectAppear', 'effectCut', 'effectCutThroughBlack', 'effectBlindsHorizontal', 'effectBlindsVertical', 'effectBoxDown', 'effectBoxLeft', 'effectBoxRight', 'effectBoxUp', 'effectCubeDown', 'effectCubeLeft', 'effectCubeRight', 'effectCubeUp', 'effectOrbitDown', 'effectOrbitLeft', 'effectOrbitRight', 'effectOrbitUp', 'effectRotateDown', 'effectRotateLeft', 'effectRotateRight', 'effectRotateUp', 'effectBoxIn', 'effectBoxOut', 'effectCheckerboardAcross', 'effectCheckerboardDown', 'effectCombHorizontal', 'effectCombVertical', 'effectConveyorLeft', 'effectConveyorRight', 'effectCoverDown', 'effectCoverLeft', 'effectCoverLeftDown', 'effectCoverLeftUp', 'effectCoverRight', 'effectCoverRightDown', 'effectCoverRightUp', 'effectCoverUp', 'effectCircleOut', 'effectDiamondOut', 'effectDissolve', 'effectFlashbulb', 'effectHoneycomb', 'effectNewsflash', 'effectPlusOut', 'effectRandom', 'effectWedge', 'effectDoorsHorizontal', 'effectDoorsVertical', 'effectFade', 'effectFadeSmoothly', 'effectFerrisWheelLeft', 'effectFerrisWheelRight', 'effectFlipDown', 'effectFlipLeft', 'effectFlipRight', 'effectFlipUp', 'effectFlyThroughIn', 'effectFlyThroughInBounce', 'effectFlyThroughOut', 'effectFlyThroughOutBounce', 'effectGalleryLeft', 'effectGalleryRight', 'effectGlitterDiamondDown', 'effectGlitterDiamondLeft', 'effectGlitterDiamondRight', 'effectGlitterDiamondUp', 'effectGlitterHexagonDown', 'effectGlitterHexagonLeft', 'effectGlitterHexagonRight', 'effectGlitterHexagonUp', 'effectPanDown', 'effectPanLeft', 'effectPanRight', 'effectPanUp', 'effectPushDown', 'effectPushLeft', 'effectPushRight', 'effectPushUp', 'effectRandomBarsHorizontal', 'effectRandomBarsVertical', 'effectRevealBlackLeft', 'effectRevealBlackRight', 'effectRevealSmoothLeft', 'effectRevealSmoothRight', 'effectRippleCenter', 'effectRippleLeftDown', 'effectRippleLeftUp', 'effectRippleRightDown', 'effectRippleRightUp', 'effectShredRectangleIn', 'effectShredRectangleOut', 'effectShredStripsIn', 'effectShredStripsOut', 'effectSplitHorizontalIn', 'effectSplitHorizontalOut', 'effectSplitVerticalIn', 'effectSplitVerticalOut', 'effectStripsDownLeft', 'effectStripsDownRight', 'effectStripsLeftDown', 'effectStripsLeftUp', 'effectStripsRightDown', 'effectStripsRightUp', 'effectStripsUpLeft', 'effectStripsUpRight', 'effectSwitchDown', 'effectSwitchLeft', 'effectSwitchRight', 'effectSwitchUp', 'effectUncoverDown', 'effectUncoverLeft', 'effectUncoverLeftDown', 'effectUncoverLeftUp', 'effectUncoverRight', 'effectUncoverRightDown', 'effectUncoverRightUp', 'effectUncoverUp', 'effectVortexDown', 'effectVortexLeft', 'effectVortexRight', 'effectVortexUp', 'effectWarpIn', 'effectWarpOut', 'effectWheel1Spoke', 'effectWheel2Spokes', 'effectWheel3Spokes', 'effectWheel4Spokes', 'effectWheel8Spokes', 'effectWheelReverse1Spoke', 'effectWindowHorizontal', 'effectWindowVertical', 'effectWipeDown', 'effectWipeLeft', 'effectWipeRight', 'effectWipeUp', 'effectMorphByObject', 'effectMorphByWord', 'effectMorphByChar', 'effectNone']

// Measured allowlist for the `slide` operation's `transition.speed` field
// -- ApiSlideShowTransition._getSpeedValue's own switch (toString()-
// recovered): any other string hits its `default: return undefined`, silently leaving
// SetSpeed's duration unset.
const KNOWN_TRANSITION_SPEEDS = ['fast', 'medium', 'slow']

// Measured allowlist for the `chart` operation -- see that
// operation's own comment for why this must be an allowlist, not a passthrough.
const KNOWN_CHART_TYPES = ['bar', 'barStacked', 'bar3D', 'lineNormal', 'lineStacked', 'pie', 'pie3D', 'area', 'scatter', 'doughnut', 'radar']

// Measured allowlist for the `chart` operation's pptx-only `legendPos` field.
// Each of these four was individually run and its saved package
// checked for a `<c:legendPos val="...">`. "none" and a deliberately made-up value produced BYTE-
// IDENTICAL output (no `<c:legend>` element at all) -- there is no way to tell them apart on this
// route, so "none" is refused here rather than assumed to be a supported request for "no legend"
// (a caller who wants no legend gets that by simply not setting `legendPos`).
const KNOWN_LEGEND_POSITIONS = ['right', 'left', 'top', 'bottom']

// Measured allowlists for the `text` operation's docx-only `align`/`highlight` fields
// -- see that operation's own comment for why these must be
// allowlists: an unrecognised value does not throw, it silently drops the whole formatting
// element. `align` is every value SetJc accepted on this route; `highlight` is the complete
// ST_HighlightColor set (OOXML has no others) -- both were run one-by-one against the live
// server and confirmed present in the saved package, not taken from documentation alone.
const KNOWN_ALIGNMENTS = ['left', 'right', 'center', 'both']
const KNOWN_HIGHLIGHT_COLORS = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'red', 'white', 'yellow', 'none']

// Table's docx-only `verticalAlign` field (TableCell.
// SetVerticalAlign). All three package-verified individually (2026-08-17): top/bottom/center
// each round-tripped to the matching <w:vAlign w:val="..."/>.
const KNOWN_CELL_VALIGNS = ['top', 'center', 'bottom']

// Measured allowlist for the `runs` operation's `vertAlign` field (docx originally) -- same
// discipline as KNOWN_HIGHLIGHT_COLORS
// above: an unrecognised value must not reach SetVertAlign silently. The complete
// ST_VerticalAlignRun set (OOXML has no others); all three run one-by-one against the live server
// and confirmed present in the saved package's word/document.xml (`<w:vertAlign w:val="...">`),
// not taken from documentation alone.
// This EXTENDS to the pptx branch too: `Run.SetVertAlign` on pptx
// delegates to `this.GetTextPr().SetVertAlign(sType)` -- the EXACT same call `emitDocxRun` below
// already makes on docx (Run/Paragraph are the same shared class on both cores, byte-identical
// method count per the pptx/docx leltár cross-check) -- reused with confidence, and re-verified
// package-side on pptx anyway rather than assumed: "superscript" -> `<a:rPr baseline="30000">`.
const KNOWN_VERT_ALIGNS = ['baseline', 'subscript', 'superscript']

// Measured allowlist for the `shape` operation's pptx-only `verticalTextAlign` field
// -- `Shape.SetVerticalTextAlign`'s own source (toString()-recovered) is a
// switch on exactly these three strings with NO default case: an unrecognised value does not
// throw, it silently does nothing. Package-verified: "center" -> `<a:bodyPr ... anchor="ctr">`.
const KNOWN_VERTICAL_TEXT_ALIGNS = ['top', 'center', 'bottom']
// The subset of OOXML placeholder-type strings that
// Api.CreatePlaceholder actually preserves -- see applyPptxShapePostCreationFields's comment
// for the round-trip measurement. NOT the full OOXML vocabulary (dt/ftr/sldNum/pic/tbl/obj/
// dgm/sldImg are excluded because they silently degrade to "body" in this engine).
const KNOWN_PLACEHOLDER_TYPES = ['title', 'body', 'ctrTitle', 'subTitle', 'chart', 'clipArt', 'media']

// `color` takes a plain 6-digit hex string rather than an allowlist -- SetColor(r, g, b) is
// arithmetic, not a keyword lookup, so there is no string value the builder could silently
// reject; the only failure mode is a caller-supplied string that is not a colour at all, which
// this function catches before any script line is generated.
function parseHexColor(value) {
  const hex = String(value).replace(/^#/, '')
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) {
    throw new Error(`text: \`color\` must be a 6-digit hex string (e.g. "1a5fb4" or "#1a5fb4"), got ${JSON.stringify(value)}`)
  }
  return { r: parseInt(hex.slice(0, 2), 16), g: parseInt(hex.slice(2, 4), 16), b: parseInt(hex.slice(4, 6), 16) }
}

// The xlsx sibling of checkCoeditSlideIndex below -- split
// out for the same reason (keeps buildCreateScript's own branching complexity down; that
// function was already flagged by qlty smells before this addition). Returns either
// `{ line }` (the `var oWorksheet = Api.GetSheet(n);` line to splice in) or `{ error }` (a
// NAMED reason to report, never a silent fall-back to sheet 0 -- see the call site's own
// comment on why that specific failure mode matters here).
function resolveXlsxSheetLine(op, i, xlsxSheetCount) {
  const sheetIdx = Number(op.sheet ?? 0)
  if (Number.isInteger(sheetIdx) && sheetIdx >= 0 && (xlsxSheetCount === null || sheetIdx < xlsxSheetCount)) {
    return { line: `var oWorksheet = Api.GetSheet(${sheetIdx});` }
  }
  // A non-integer op.sheet (e.g. a
  // sheet NAME, which the co-editing route accepts) used to fall through to the range message
  // below and report "out of range" -- true of the symptom (Number.isInteger(NaN) is false) but
  // not of the cause. This route builds its own sheets and cannot resolve a name against them
  // (see the call site's own comment on why), so a non-integer op.sheet gets its own, honest
  // error instead of borrowing the range one.
  if (!Number.isInteger(sheetIdx)) {
    return { error: `operation ${i}: sheet ${JSON.stringify(op.sheet ?? 0)} is not a non-negative integer -- this route cannot resolve a sheet name (only a numeric index is accepted here; the co-editing route accepts a sheet name)` }
  }
  const rangeMsg = xlsxSheetCount !== null
    ? `this document has ${xlsxSheetCount} sheet(s) (valid indices: 0..${xlsxSheetCount - 1})`
    : 'the sheet count is not known on this route -- only a non-negative integer index can be checked here'
  return { error: `operation ${i}: sheet index ${sheetIdx} is out of range -- ${rangeMsg}` }
}

// Builds the creation script for any of the three cores. The core-specific preamble is the only
// place that knows how to reach the document/sheet/slide object.
//
// E6: `slideCount` is the number of slides the OPENED seed actually has (E7's mag-generator
// output, or the demo's own slide count for anything opened as-is) -- NOT discovered here, it
// has to come from the caller, because a DocBuilder script cannot read its own input's slide
// count before running (and by the time it runs, it is too late to refuse cleanly). Defaults
// to 1 to match this function's pre-E6 behaviour (always slide 0) for callers that pass none.
// `sheetCount` is xlsx's analogous parameter, and null (not
// defaulted) when the caller omits it, so resolveXlsxSheetLine can tell "not supplied" apart
// from "supplied as 1" -- same reasoning as checkCoeditSlideIndex's own pptxSlideCount !== null
// branch.
function buildCreateScript({ core, operations, outName, slideCount, sheetCount }) {
  if (!['docx', 'xlsx', 'pptx'].includes(core)) throw new Error(`buildCreateScript: unknown core: ${core}`)
  const ops = Array.isArray(operations) ? operations : []
  if (!ops.length) throw new Error('buildCreateScript: no operations -- the script would be a no-op')
  const pptxSlideCount = Number(slideCount ?? 1)
  const xlsxSheetCount = sheetCount === undefined || sheetCount === null ? null : Number(sheetCount)

  const preamble = {
    docx: ['var oDocument = Api.GetDocument();'],
    // No fixed `Api.GetActiveSheet()` here (same shape as
    // pptx's E6 below): every xlsx operation now gets its own `oWorksheet` binding via
    // sheetLine in the per-operation loop, so a preamble line here would always be overwritten
    // before the first operation runs -- dead code that only looks load-bearing.
    xlsx: [],
    // No `oSlide` here (pre-E6 had `GetSlideByIndex(0)` fixed here): E6 lets each operation
    // target its own slide, so the lookup moves into the per-operation loop below.
    pptx: ['var oPresentation = Api.GetPresentation();'],
  }[core]

  const lines = ['builder.OpenFile("__DOC_URL__");', ...preamble]
  const applied = []
  // Every operation gets an entry here, whatever happens to it
  // -- a bad operation is NAMED, never silently dropped, but it no longer aborts its
  // siblings either: a caller who asked for three things and got two is told about all three,
  // not left to infer the third from an exception that swallowed the two that would have worked.
  const report = []
  for (const [i, op] of ops.entries()) {
    const spec = OPERATIONS[op.type]
    if (!spec) {
      const message = `operation ${i}: unknown type ${JSON.stringify(op.type)} (known: ${Object.keys(OPERATIONS).join(', ')})`
      report.push({ index: i, type: op.type, outcome: 'nem-tamogatott', sourceRoute: CREATE_ROUTE, reason: message })
      continue
    }
    if (!spec.cores.includes(core)) {
      const message = `operation ${i}: type "${op.type}" is not available in the ${core} core (available: ${spec.cores.join(', ')})`
      report.push({ index: i, type: op.type, outcome: 'nem-tamogatott', sourceRoute: CREATE_ROUTE, reason: message })
      continue
    }
    // E6, restored after the merge with the E8 report-loop rewrite silently dropped it (no
    // conflict marker -- two non-overlapping edits to the same loop, git took master's rewrite
    // wholesale): each pptx operation targets its own slide, so the GetSlideByIndex lookup is
    // per-operation, not a single fixed preamble line. Without this, EVERY pptx operation from
    // the create route referenced an undeclared `oSlide` in the generated script.
    let slideLine = null
    if (core === 'pptx') {
      const slideIdx = Number(op.slide ?? 0)
      // An index past the actual slide count is a named error, not a silent no-op (a request
      // for slide 5 of a 1-slide deck must not quietly land on slide 0).
      if (!Number.isInteger(slideIdx) || slideIdx < 0 || slideIdx >= pptxSlideCount) {
        const message = `operation ${i}: slide index ${JSON.stringify(op.slide ?? 0)} is out of range -- this document has ${pptxSlideCount} slide(s) (valid indices: 0..${pptxSlideCount - 1})`
        report.push({ index: i, type: op.type, outcome: 'hiba', sourceRoute: CREATE_ROUTE, reason: message })
        continue
      }
      slideLine = `var oSlide = oPresentation.GetSlideByIndex(${slideIdx});`
    }
    // Same shape as slideLine above (computed
    // UNCONDITIONALLY, default sheet index 0), split into its own function purely to keep this
    // loop's own complexity down -- same reasoning as checkCoeditSlideIndex below.
    let sheetLine = null
    if (core === 'xlsx') {
      const resolved = resolveXlsxSheetLine(op, i, xlsxSheetCount)
      if (resolved.error) {
        report.push({ index: i, type: op.type, outcome: 'hiba', sourceRoute: CREATE_ROUTE, reason: resolved.error })
        continue
      }
      sheetLine = resolved.line
    }
    let opLines
    try {
      opLines = spec.emit(op, core)
    } catch (err) {
      // notSupportedError() (a genuinely missing capability, e.g. an allowlisted value the
      // builder would silently drop) vs. a plain Error (a malformed call to a capability that
      // DOES exist, e.g. a heading number out of range) -- see that function's own comment.
      report.push({ index: i, type: op.type, outcome: err.notSupported ? 'nem-tamogatott' : 'hiba', sourceRoute: CREATE_ROUTE, reason: err.message })
      continue
    }
    if (slideLine) lines.push(slideLine)
    if (sheetLine) lines.push(sheetLine)
    lines.push(...opLines)
    applied.push({ index: i, type: op.type, sourceRoute: CREATE_ROUTE })
    report.push({ index: i, type: op.type, outcome: 'alkalmazva', sourceRoute: CREATE_ROUTE, reason: null })
  }

  if (!applied.length) {
    // Nothing survived -- the script would be exactly as pointless as an empty operations list
    // (a resave of the mag, nothing more), so this stays a hard failure like that case always
    // was. The per-operation reasons travel WITH the throw (err.report), not just folded into
    // the message, so a caller that catches this can still build the same itemized answer E8
    // gives for a mixed batch -- an all-refused batch is not a different kind of failure.
    const reasons = report.map((r) => `operation ${r.index} (${r.type}): ${r.reason}`).join('; ')
    const err = new Error(`buildCreateScript: no operation could be applied -- ${reasons}`)
    err.report = report
    throw err
  }

  lines.push(`builder.SaveFile(${jsString(core)}, ${jsString(outName || 'eredmeny.' + core)});`, 'builder.CloseFile();')
  return { script: lines.join('\n') + '\n', applied, report }
}

// A SECOND translator over the SAME `operations` list and the
// SAME per-operation OPERATIONS[type].emit(op, core) -- but for the co-editing route (an
// ALREADY-OPEN document, driven via editor.callCommand(function(){ ... }), not the DocBuilder
// create route). The operation schema does not change; only the target context and the
// envelope around the emitted lines do:
//   DocBuilder route:  builder.OpenFile(...) / <op lines> / builder.SaveFile(...) + CloseFile()
//   co-editing route:  <preamble, unchanged> / <the SAME op lines> / (no open/save -- the
//                       document is already open, and callCommand's own save happens outside
//                       this body, driven by the caller)
// MEASURED: callCommand works on all three cores
// and the change survives a save -- this function does not re-measure that, it reuses it.
const COEDIT_ROUTE = 'coediting-callcommand'

// E6's named-error slide-index check (buildCreateScript),
// adapted for a route where the upper bound is not always knowable -- see buildCoeditScript's
// own header comment for why. Split out purely to keep that function's branching complexity
// down (a bare `if` inline here duplicated the same shape as buildCreateScript's own check, and
// the resulting function crossed into "high complexity" on qlty smells).
function checkCoeditSlideIndex(op, i, pptxSlideCount) {
  const slideIdx = Number(op.slide ?? 0)
  // E6's other named direction: an index past the actual slide count is a named error, not
  // a silent no-op (a request for slide 5 of a 1-slide deck must not quietly land on slide 0).
  // The upper bound only applies when the caller supplied a slideCount -- see this
  // function's own header comment for why it cannot be assumed otherwise.
  if (!Number.isInteger(slideIdx) || slideIdx < 0 || (pptxSlideCount !== null && slideIdx >= pptxSlideCount)) {
    const rangeMsg = pptxSlideCount !== null
      ? `this document has ${pptxSlideCount} slide(s) (valid indices: 0..${pptxSlideCount - 1})`
      : 'the slide count is not known on this route -- only a non-negative integer index can be checked here'
    throw new Error(`operation ${i}: slide index ${JSON.stringify(op.slide ?? 0)} is out of range -- ${rangeMsg}`)
  }
  return slideIdx
}

// xlsx write operations (table/text/
// formula/numberFormat/columnWidth/border/conditionalFormatting/chart/sheetDisplay/...) all emit
// lines against a FIXED `oWorksheet` reference, bound ONCE in buildCoeditScript's preamble
// (`Api.GetActiveSheet()`). MEASURED: a mid-batch
// `sheet.active` operation does not rebind that reference -- the JS variable still points at
// whichever sheet was active when the script started, so every subsequent write lands there
// regardless of which sheet the caller intended. This is the named mechanism behind "the 2nd/3rd
// sheet stays empty". The fix follows pptx's own `oSlide`-per-operation
// precedent exactly: `op.sheet` (optional, number index or string name) rebinds `oWorksheet` via
// `Api.GetSheet(...)` IMMEDIATELY BEFORE that operation's own emitted lines, so each operation
// can target its own sheet within one script -- same var-redeclaration mechanism `oSlide` already
// relies on (JS `var` permits repeated declaration+reassignment in the same function body).
// Same "honest reduction" as checkCoeditSlideIndex: there is no live sheet count/name list on
// this route either (see that function's own header comment), so only the TYPE is checked here
// (a non-negative integer or a non-empty string) -- an out-of-range index or an unknown name is
// NOT checkable client-side, and is refused, at runtime, INSIDE the generated script instead
// (see the `if (!oWorksheet) throw ...` line below) rather than risked as a silent write to
// whatever GetSheet's failure mode happens to return.
function checkCoeditSheetTarget(op, i) {
  const target = op.sheet
  if (typeof target === 'number') {
    if (!Number.isInteger(target) || target < 0) {
      throw new Error(`operation ${i}: sheet index ${JSON.stringify(target)} is invalid -- must be a non-negative integer (0-based)`)
    }
    return Number(target)
  }
  if (typeof target === 'string') {
    if (!target.trim()) throw new Error(`operation ${i}: sheet name must not be empty`)
    return target
  }
  throw new Error(`operation ${i}: \`sheet\` must be a non-negative integer (0-based index) or a non-empty string (sheet name), got ${JSON.stringify(target)}`)
}

// Split out of buildCoeditScript's own loop purely to keep ITS complexity down (qlty smells --
// the try/catch + ternary inline there pushed buildCoeditScript from 27 to 39; moved here it
// contributes to a small, single-purpose function instead). Returns {ok:true, line} (line is
// null when `op.sheet` was not given -- the no-op case) or {ok:false, reason}; never throws.
function resolveCoeditSheetLine(op, i) {
  // `sheet` (the operation TYPE, rename/visible/active/delete) already has its own `op.target`,
  // resolving into a separate `oTargetSheet` variable (see that operation's own emit()) -- `op.sheet`
  // (the FIELD, this function's concern) would be a confusing second meaning on that one type.
  if (op.type === 'sheet' || op.sheet === undefined || op.sheet === null) return { ok: true, line: null }
  let sheetTarget
  try {
    sheetTarget = checkCoeditSheetTarget(op, i)
  } catch (err) {
    return { ok: false, reason: err.message }
  }
  const targetArg = typeof sheetTarget === 'number' ? sheetTarget : jsString(sheetTarget)
  const notFoundMsg = jsString(`operation ${i}: sheet ${JSON.stringify(op.sheet)} not found in this workbook`)
  return { ok: true, line: `var oWorksheet = Api.GetSheet(${targetArg}); if (!oWorksheet) { throw new Error(${notFoundMsg}); }` }
}

// Unlike buildCreateScript, this route has no reliable
// source for the slide count -- it targets an ALREADY-OPEN document via callCommand, and (as of
// this fix) the only caller (coedit.cjs's writeOperationsToDocument -> the coedit_write_operations
// MCP tool) does not discover or pass one; there is no preliminary Api call reading the live
// presentation's slide count before building this script. So `slideCount` is OPTIONAL here: when
// a future caller does know it (e.g. having queried it via its own callCommand round trip first),
// passing it gets the SAME upper-bound check buildCreateScript has; when omitted (today's only
// caller), only the lower bound (a non-negative integer index) is enforced -- an honest reduction
// in what can be checked, not a silently dropped one, and NOT a fabricated default like 1 (which
// would misreport every multi-slide document's valid indices as out of range).
// A pptx lap-sor feloldasa kiemelve, UGYANOLYAN
// alakban mint a mar meglevo (xlsx-oldali) resolveCoeditSheetLine -- ok:true/line vagy
// ok:false/reason, a report-mezot a hivo (resolveCoeditOperation) allitja ossze. A ket
// hivas emiatt szimmetrikus, es a resolveCoeditOperation sajat return-szama is csokken.
function resolveCoeditSlideLine(op, i, core, pptxSlideCount) {
  if (core !== 'pptx') return { ok: true, line: null }
  let slideIdx
  try {
    slideIdx = checkCoeditSlideIndex(op, i, pptxSlideCount)
  } catch (err) {
    return { ok: false, reason: err.message }
  }
  return { ok: true, line: `var oSlide = oPresentation.GetSlideByIndex(${slideIdx});` }
}

// A per-operacio validalas+emit kiemelve
// buildCoeditScript sajat ciklustorzsebol -- TISZTA REFAKTOR, viselkedes valtozatlan.
// Az indok ugyanaz, mint a nested-function-still-counts-toward-parent-complexity leletnel
// (koteg08, resolveShapeFill/buildGradientStop): a qlty a ciklustorzs SAJAT elagazasait a
// BEZARO fuggveny szamlajara irja, tehat a kiemeles az, ami TENYLEG csokkenti a komplexitast --
// nem csak athelyezi ugyanazt a szamot egy uj nevre.
function resolveCoeditOperation(op, i, core, pptxSlideCount) {
  const spec = OPERATIONS[op.type]
  if (!spec) {
    const message = `operation ${i}: unknown type ${JSON.stringify(op.type)} (known: ${Object.keys(OPERATIONS).join(', ')})`
    return { ok: false, report: { index: i, type: op.type, outcome: 'nem-tamogatott', sourceRoute: COEDIT_ROUTE, reason: message } }
  }
  if (!spec.cores.includes(core)) {
    const message = `operation ${i}: type "${op.type}" is not available in the ${core} core (available: ${spec.cores.join(', ')})`
    return { ok: false, report: { index: i, type: op.type, outcome: 'nem-tamogatott', sourceRoute: COEDIT_ROUTE, reason: message } }
  }
  const slideResolved = resolveCoeditSlideLine(op, i, core, pptxSlideCount)
  // `sheet` is meaningless on the `sheet` operation itself
  // (it already has its own `op.target`, resolving into a SEPARATE `oTargetSheet` variable --
  // see that operation's emit()) -- only the OTHER xlsx write operations, which all reference
  // the shared `oWorksheet`, need this rebind.
  const sheetResolved = core === 'xlsx' ? resolveCoeditSheetLine(op, i) : { ok: true, line: null }
  if (!slideResolved.ok || !sheetResolved.ok) {
    const reason = slideResolved.ok ? sheetResolved.reason : slideResolved.reason
    return { ok: false, report: { index: i, type: op.type, outcome: 'hiba', sourceRoute: COEDIT_ROUTE, reason } }
  }
  let opLines
  try {
    opLines = spec.emit(op, core)
  } catch (err) {
    return { ok: false, report: { index: i, type: op.type, outcome: err.notSupported ? 'nem-tamogatott' : 'hiba', sourceRoute: COEDIT_ROUTE, reason: err.message } }
  }
  const lines = []
  if (slideResolved.line) lines.push(slideResolved.line)
  if (sheetResolved.line) lines.push(sheetResolved.line)
  lines.push(...opLines)
  // This is the
  // CLIENT-SIDE validation success path -- the operation's shape checked out, nothing has been
  // sent to the Document Server yet at this point in the call. 'elkuldve-nem-verifikalt' ("sent,
  // not verified"), not 'alkalmazva' ("applied") -- the word "applied" is exactly the false claim
  // this whole investigation started from (the autoFilter case,
  // and the cell case before it). For the two operation types with
  // post-save verification wired (autoFilter, table/formula cells), coedit.cjs's own
  // applyAutoFilterVerificationToReport/applyCellVerificationToReport overwrite this value with
  // the real, package-verified outcome (vegrehajtva/kihagyva-idempotens/megtagadva) once the write
  // actually happens and the saved file is re-read. For every OTHER operation type (no verification
  // wired yet), this value is the FINAL one a caller sees -- and it now says exactly what is known:
  // the request was well-formed and was sent, not that it succeeded.
  return {
    ok: true,
    lines,
    applied: { index: i, type: op.type, sourceRoute: COEDIT_ROUTE },
    report: { index: i, type: op.type, outcome: 'elkuldve-nem-verifikalt', sourceRoute: COEDIT_ROUTE, reason: null },
  }
}

function buildCoeditScript({ core, operations, slideCount }) {
  if (!['docx', 'xlsx', 'pptx'].includes(core)) throw new Error(`buildCoeditScript: unknown core: ${core}`)
  const ops = Array.isArray(operations) ? operations : []
  // Same reasoning as buildCreateScript's own empty-operations guard: an empty list must not
  // produce a no-op callCommand body -- the call should never HAPPEN at all, so a caller cannot
  // mistake "nothing to do" for "ran and did nothing".
  if (!ops.length) throw new Error('buildCoeditScript: no operations -- the call would be a no-op')
  const pptxSlideCount = slideCount === undefined || slideCount === null ? null : Number(slideCount)

  const preamble = {
    docx: ['var oDocument = Api.GetDocument();'],
    xlsx: ['var oWorksheet = Api.GetActiveSheet();'],
    // `oPresentation` declared here (not a fixed-slide `oSlide`):
    // the per-operation loop below looks up its OWN slide by index, same
    // division as buildCreateScript's preamble/loop split -- see that function's own comment
    // for why the lookup cannot be a single fixed preamble line once E6 let each operation
    // target its own slide.
    pptx: ['var oPresentation = Api.GetPresentation();'],
  }[core]

  const applied = []
  const report = []
  const pendingLines = []
  // VALIDATE every
  // operation first, APPLY only if every single one validates. Unlike buildCreateScript, this
  // route writes into an ALREADY-OPEN, co-edited document -- a partially-applied batch cannot be
  // rolled back safely (another editor may write between our read and our own "undo", so a
  // rewrite-the-whole-file rollback would erase THEIR change, not just ours). So this loop still
  // evaluates every operation, same as before
  // (a bad entry does not abort its siblings' validation), but nothing collected here
  // reaches `lines` unless the WHOLE batch is clean -- see the all-or-nothing gate below.
  for (const [i, op] of ops.entries()) {
    const result = resolveCoeditOperation(op, i, core, pptxSlideCount)
    report.push(result.report)
    if (!result.ok) continue
    pendingLines.push(...result.lines)
    applied.push(result.applied)
  }

  if (applied.length !== ops.length) rejectCoeditBatch(ops, applied, report)

  // The function BODY only -- wrapping it as `editor.callCommand(function(){ ... })` (or
  // `new Function(...)`) is the caller's job, same division as buildCreateScript returning the
  // DocBuilder script body without the HTTP POST that actually sends it.
  const lines = [...preamble, ...pendingLines]
  return { script: lines.join('\n') + '\n', applied, report }
}

// Split out purely
// to keep buildCoeditScript's own branching complexity down, same reasoning as
// checkCoeditSlideIndex's own split above. Covers both a mixed batch and the fully-degenerate case
// where every operation is bad -- the same gate, not two: the caller above already named every BAD
// operation's own reason; the entries still marked 'elkuldve-nem-verifikalt' validated fine on
// their own, but nothing was actually written for them either, because the batch as a whole is
// rejected -- so they are relabelled here, before the caller ever sees them, rather than shipping
// a report that names a NEVER-SENT operation the same as one that was sent (the earlier fix
// renamed the success-path value from 'alkalmazva'; this function's
// own check follows). Always throws -- never returns.
function rejectCoeditBatch(ops, applied, report) {
  const bad = report.filter((r) => r.outcome !== 'elkuldve-nem-verifikalt')
  const badList = bad.map((r) => `${r.index} (${r.type})`).join(', ')
  for (const r of report) {
    if (r.outcome === 'elkuldve-nem-verifikalt') {
      r.outcome = 'nem-alkalmazva'
      r.reason = `this operation validated on its own, but the batch was rejected because ` +
        `operation(s) ${badList} did not -- this route never writes a partial batch, so nothing ` +
        `in this call was applied`
    }
  }
  const reasons = bad.map((r) => `operation ${r.index} (${r.type}): ${r.reason}`).join('; ')
  const err = new Error(`buildCoeditScript: ${bad.length} of ${ops.length} operation(s) failed validation -- nothing was applied (all-or-nothing route): ${reasons}`)
  err.report = report
  throw err
}


// Populate the registry BEFORE requiring
// any split-out operations file (see lib-operations-registry.cjs's own comment for why the
// order matters). Every function/const the split files might reference goes in, by shorthand --
// each file destructures only the subset it actually uses.
const OPERATIONS_REGISTRY = require('./lib-operations-registry.cjs')
Object.assign(OPERATIONS_REGISTRY, {
  COEDIT_ROUTE,
  CREATE_ROUTE,
  DATA_URI_RE,
  DEFAULT_COLUMN_SPACE_TWIPS,
  IMAGE_MAGIC,
  KNOWN_ALIGNMENTS,
  KNOWN_BOOKMARK_REF_TYPES,
  KNOWN_CELL_TEXT_DIRECTIONS,
  KNOWN_CELL_VALIGNS,
  KNOWN_CHART_TYPES,
  KNOWN_HIGHLIGHT_COLORS,
  KNOWN_LEGEND_POSITIONS,
  KNOWN_PATTERN_TYPES,
  KNOWN_PLACEHOLDER_TYPES,
  KNOWN_PPTX_CELL_VALIGNS,
  KNOWN_SECTION_TYPES,
  KNOWN_THEME_NAMES,
  KNOWN_TRANSITION_EFFECTS,
  KNOWN_TRANSITION_SPEEDS,
  KNOWN_VERTICAL_TEXT_ALIGNS,
  KNOWN_VERT_ALIGNS,
  KNOWN_WORDART_TRANSFORMS,
  OUTCOME,
  PPTX_NUM_TYPES,
  SECTION2_DOCX_KEYS,
  applyAutoFilterVerificationToReport,
  applyCellVerificationToReport,
  applyChartFormatting,
  applyDocxParagraphBookmarkRef,
  applyDocxParagraphComments,
  applyDocxParagraphCustomStyle,
  applyDocxParagraphHyperlink,
  applyDocxParagraphNotes,
  applyDocxTableCaption,
  applyDocxTableCellRefinements,
  applyDocxTableMerge,
  applyDocxTableRowRefinements,
  applyDocxTableSplit,
  applyParagraphIndentSpacing,
  applyPptxCellMargin,
  applyPptxCellTextDirection,
  applyPptxCellVAlign,
  applyPptxParagraphIndent,
  applyPptxParagraphListType,
  applyPptxShapePostCreationFields,
  applyPptxTableCellSettings,
  applyPptxTableLook,
  applyPptxTableLookAndStructuralExtras,
  applyPptxTableMerge,
  applyPptxTableStructuralExtras,
  applyShapeGeometry,
  applyShapePlaceholder,
  b64url,
  bookmarksFromDocumentXml,
  buildCoeditScript,
  buildCreateScript,
  buildDocxTableLines,
  buildEditScript,
  buildFillExpression,
  buildGradientFillExpression,
  buildGradientStop,
  buildGradientStopExpr,
  buildPatternFillExpression,
  buildPptxTextShapeLines,
  buildSlideGroupLines,
  buildSlideGroupShapeLines,
  buildSlideRemoveObjectLine,
  buildSlideTransitionLines,
  cellRefsPresentInSheetXml,
  checkCoeditSheetTarget,
  checkCoeditSlideIndex,
  checkFreshness,
  classifyDsResponse,
  commentsFromCommentsXml,
  commentsFromPptxCommentXml,
  decodeXmlEntities,
  documentMetadataFromCoreAndAppXml,
  emitDocxRun,
  emitPageSetupDocxSection2,
  endnotesFromEndnotesXml,
  findMatchesInParagraphs,
  footnotesFromFootnotesXml,
  forEachCell,
  jsString,
  makeMarker,
  markerInDocumentXml,
  notSupportedError,
  ooxmlRelationshipTargetByType,
  pageSetupDocx,
  paragraphTextsFromDocumentXml,
  paragraphTextsFromSlideXml,
  parseHexColor,
  pptxCommentsPartFromSlideRelsXml,
  pptxLayoutAndThemeFromSlideLayoutXml,
  pptxObjectXfrmFields,
  pptxSlideOrderFromPresentationXml,
  rejectCoeditBatch,
  resolveCoeditOperation,
  resolveCoeditSheetLine,
  resolveCoeditSlideLine,
  resolveImageSrc,
  resolveOoxmlRelativeTarget,
  resolveShapeFill,
  resolveShapeObjectFill,
  resolveSlideBackgroundLine,
  resolveXlsxCellText,
  resolveXlsxSheetFile,
  resolveXlsxSheetLine,
  rgbArg,
  signJwt,
  slideContentSummaryFromSlideXml,
  sniffImageMime,
  tablesFromDocumentXml,
  validateDocxTableRefinements,
  validateHyperlink,
  validatePageSetupDocxSection2,
  validateRgbColor,
  xlsxAutoFilterRangeInSheetXml,
  xlsxAutoFilterVerificationReport,
  xlsxCellVerificationReport,
  xlsxRequestedCellRefs,
  xlsxRequestsAutoFilterVerification,
  xlsxTableCellRefs,
})
const OPERATIONS = {
  ...require('./lib-operations-docx.cjs'),
  ...require('./lib-operations-xlsx.cjs'),
  ...require('./lib-operations-pptx.cjs'),
}
module.exports = {
  checkFreshness,
  b64url,
  signJwt,
  makeMarker,
  jsString,
  buildEditScript,
  buildCreateScript,
  buildCoeditScript,
  OPERATIONS,
  CREATE_ROUTE,
  COEDIT_ROUTE,
  classifyDsResponse,
  markerInDocumentXml,
  paragraphTextsFromDocumentXml,
  commentsFromCommentsXml,
  footnotesFromFootnotesXml,
  endnotesFromEndnotesXml,
  bookmarksFromDocumentXml,
  tablesFromDocumentXml,
  paragraphTextsFromSlideXml,
  pptxSlideOrderFromPresentationXml,
  pptxCommentsPartFromSlideRelsXml,
  commentsFromPptxCommentXml,
  slideContentSummaryFromSlideXml,
  documentMetadataFromCoreAndAppXml,
  resolveOoxmlRelativeTarget,
  ooxmlRelationshipTargetByType,
  pptxLayoutAndThemeFromSlideLayoutXml,
  findMatchesInParagraphs,
  cellRefsPresentInSheetXml,
  resolveXlsxCellText,
  xlsxTableCellRefs,
  xlsxRequestedCellRefs,
  xlsxCellVerificationReport,
  applyCellVerificationToReport,
  xlsxAutoFilterRangeInSheetXml,
  xlsxRequestsAutoFilterVerification,
  xlsxAutoFilterVerificationReport,
  applyAutoFilterVerificationToReport,
  resolveXlsxSheetFile,
  OUTCOME,
}
