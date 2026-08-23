'use strict'
// Operations split boundary decision: this operation touches multiple cores (docx, pptx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { KNOWN_ALIGNMENTS, KNOWN_HIGHLIGHT_COLORS, applyDocxParagraphBookmarkRef, applyDocxParagraphCustomStyle, applyDocxParagraphHyperlink, applyDocxParagraphNotes, applyParagraphIndentSpacing, buildPptxTextShapeLines, jsString, notSupportedError, parseHexColor } = require('../lib-operations-registry.cjs')
module.exports = {
  cores: ['docx', 'pptx'],
  // A paragraph in the text core; a text box on a slide. Same intent, different objects.
  //
  // *** THE DOCX BRANCH GAINED A FULL SET OF PARAGRAPH/RUN FIELDS,
  // ALL OPTIONAL -- so `{type:'text', text, bold, size}` still
  // produces exactly the same lines it always has; nothing here fires unless the caller asks.
  // Two of the new fields needed an ALLOWLIST for the same reason `chartType` does: an
  // unrecognised value does NOT throw inside the builder, it silently produces no XML at all
  // (measured 2026-08-15: SetJc("nincsilyen") and SetHighlight("nincsilyenszin") both ran to
  // completion with the paragraph TEXT intact and the formatting element simply absent from
  // word/document.xml -- the same silent-no-op class as an unknown chartType). `align` is the
  // four values SetJc actually accepts on this route (left/right/center/both -- all four
  // individually run and confirmed present in the saved package). `highlight` is the full
  // 17-value ST_HighlightColor set (all 17 individually run and confirmed present). `color`
  // needs no allowlist -- it is parsed into three byte values here, so the builder never sees
  // a string SetColor could silently drop. `heading` maps to the built-in "Heading 1" ..
  // "Heading 9" paragraph styles via oDocument.GetStyle(name) + oParagraph.SetStyle(...) --
  // measured present on this instance for all nine levels, and the saved package's
  // word/styles.xml resolves the emitted numeric w:pStyle id back to the right w:name.
  //
  // `listType` (bullet/numbered): WORKS, via `oDocument.CreateNumbering(...)`, NOT
  // `Api.CreateNumbering(...)` (2026-08-17, the owner's own
  // steer after the `Api.`-level refusal below). The two are different objects: the
  // document-bound one returns a real numbering with `GetLevel(0..8)`, and paragraphs attach
  // to it via the TWO-ARGUMENT `oParagraph.SetNumPr(oNumbering, levelIndex)` -- the
  // one-argument `SetNumPr(level)` some older notes assumed is silently inert. MEASURED on
  // all 4 matrix cells (numbered/bullet x docbuilder-create/co-editing), package-verified
  // (`word/numbering.xml` gets a real `abstractNum`+`num`, every list paragraph's `<w:pPr>`
  // carries `<w:numPr><w:ilvl w:val="N"/><w:numId .../></w:numPr>`, list text is never written
  // as a literal character), AND editor-rendering-verified via a co-editing screenshot
  // (correctly nested 1./a./b./2. numbering, letter level restarting per parent item). See the
  // card comments (14499, 14705, 14714) for the full evidence.
  //
  // The `Api.CreateNumbering(...)` refusal this comment used to describe is UNCHANGED and
  // still correct for THAT call -- it still kills the docbuilder-create job outright and still
  // silently drops content on the co-editing route (both re-confirmed 2026-08-17, see the
  // card). This operation never calls it; it only ever calls the document-bound
  // `oDocument.CreateNumbering(...)`, which is a different object with no such failure mode on
  // either route.
  //
  // ONE NUMBERING PER `listType` PER SCRIPT, not per call: consecutive `text` operations with
  // the same `listType` in one `operations` array share a single `oNumbering_<listType>`
  // script-local variable (`var x = (typeof x !== 'undefined' && x) ? x : oDocument.
  // CreateNumbering(...)`, safe under `var` hoisting since every operation's lines land in one
  // concatenated script) -- this is what keeps a multi-item list as ONE list instead of a new
  // list starting over at 1. MI NEM TARTOZIK BELE: two SEPARATE numbered lists in the same
  // document (each restarting at 1) is not offered -- every `numbered` op in a script joins
  // the same running list; not measured/requested, would need a per-call `listId` to select
  // between multiple `oNumbering_*` variables.
  emit(op, core) {
    const value = jsString(op.text ?? '')
    if (core === 'docx') {
      const lines = ['var oParagraph = Api.CreateParagraph();']
      if (op.heading !== undefined && op.heading !== null) {
        const level = Number(op.heading)
        if (!Number.isInteger(level) || level < 1 || level > 9) {
          throw new Error(`text: \`heading\` must be an integer 1-9 (the built-in "Heading N" paragraph styles), got ${JSON.stringify(op.heading)}`)
        }
        lines.push(`oParagraph.SetStyle(oDocument.GetStyle(${jsString('Heading ' + level)}));`)
      }
      lines.push(`oParagraph.AddText(${value});`)
      if (op.bold) lines.push('oParagraph.SetBold(true);')
      if (op.italic) lines.push('oParagraph.SetItalic(true);')
      if (op.underline) lines.push('oParagraph.SetUnderline(true);')
      if (op.strikethrough) lines.push('oParagraph.SetStrikeout(true);')
      if (op.size) lines.push(`oParagraph.SetFontSize(${Number(op.size)});`)
      if (op.font) lines.push(`oParagraph.SetFontFamily(${jsString(op.font)});`)
      if (op.color !== undefined && op.color !== null) {
        const { r, g, b } = parseHexColor(op.color)
        lines.push(`oParagraph.SetColor(${r}, ${g}, ${b}, false);`)
      }
      if (op.highlight !== undefined && op.highlight !== null) {
        if (!KNOWN_HIGHLIGHT_COLORS.includes(op.highlight)) {
          throw notSupportedError(`text: unknown highlight ${JSON.stringify(op.highlight)} (known: ${KNOWN_HIGHLIGHT_COLORS.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently produces no highlight at all, so this tool refuses it here instead`)
        }
        lines.push(`oParagraph.SetHighlight(${jsString(op.highlight)});`)
      }
      if (op.align !== undefined && op.align !== null) {
        if (!KNOWN_ALIGNMENTS.includes(op.align)) {
          throw notSupportedError(`text: unknown align ${JSON.stringify(op.align)} (known: ${KNOWN_ALIGNMENTS.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently produces no alignment at all, so this tool refuses it here instead`)
        }
        lines.push(`oParagraph.SetJc(${jsString(op.align)});`)
      }
      // Indentation + spacing (receptek-pptx-docx.md #9).
      // All four indent/spacing-before/after fields are TWIP integers, package-confirmed
      // present with the exact called value (720 -> <w:ind w:firstLine="720"/> etc, 200 ->
      // <w:spacing w:before="200" .../>) -- no unit conversion happens on this route.
      // `spacingLine` has its own conversion note; see applyParagraphIndentSpacing() above.
      applyParagraphIndentSpacing(lines, op)
      if (op.listType !== undefined && op.listType !== null) {
        if (op.listType !== 'bullet' && op.listType !== 'numbered') {
          throw new Error(`text: \`listType\` must be "bullet" or "numbered", got ${JSON.stringify(op.listType)}`)
        }
        let listLevel = 0
        if (op.listLevel !== undefined && op.listLevel !== null) {
          listLevel = Number(op.listLevel)
          if (!Number.isInteger(listLevel) || listLevel < 0 || listLevel > 8) {
            throw new Error(`text: \`listLevel\` must be an integer 0-8, got ${JSON.stringify(op.listLevel)}`)
          }
        }
        const numVar = `oNumbering_${op.listType}`
        lines.push(`var ${numVar} = (typeof ${numVar} !== 'undefined' && ${numVar}) ? ${numVar} : oDocument.CreateNumbering(${jsString(op.listType)});`)
        lines.push(`oParagraph.SetNumPr(${numVar}, ${listLevel});`)
      }
      applyDocxParagraphCustomStyle(lines, op)
      applyDocxParagraphHyperlink(lines, op)
      lines.push('oDocument.Push(oParagraph);')
      applyDocxParagraphNotes(lines, op)
      applyDocxParagraphBookmarkRef(lines, op)
      return lines
    }
    // see buildPptxTextShapeLines() above for the `paragraphs`
    // field this adds -- kept out of this already-flagged (qlty smells) function on purpose.
    return buildPptxTextShapeLines(op)
  },
}
