'use strict'
// Operations split boundary decision: this operation touches multiple cores (docx, pptx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { KNOWN_VERT_ALIGNS, emitDocxRun, jsString, notSupportedError, parseHexColor, validateHyperlink } = require('../lib-operations-registry.cjs')

  // E14 -- MULTIPLE independently-styled runs inside ONE paragraph -- distinct from `text` (a
  // single style for the whole paragraph/box) and `shape` (no per-run text styling at all). One
  // Api.CreateRun() per entry in `runs`, each styled independently and appended with AddElement.
  // Originally pptx-only, own shape+paragraph (receptek-pptx-docx.md #3). Later
  // work added the docx branch (referencia-docx-builder.js /
  // euro-demo-docx.js lines 84-99, its own 1. chapter demo sentence) -- the pptx branch below is
  // BYTE-IDENTICAL to before that addition (differential-run proof verified at the time).
  //
  // *** THE TRAP, MEASURED (receptek-pptx-docx.md line 19-22): `SetFontSize` takes HALF-POINTS
  // on pptx too, not just docx -- SetFontSize(24) -> 12pt in the saved package. The pptx branch
  // does NOT rescale: `fontSize` here is the SAME half-point unit as `text`'s own `size` field
  // (both docx and pptx) -- the existing, established convention in this file. A caller wanting
  // 24pt text passes fontSize: 48. Silently doubling it here would create a second, conflicting
  // convention for the same concept one operation over. The docx branch below uses `size` (not
  // `fontSize`) per run, matching `text`'s own docx field name -- each branch matches its
  // sibling operation's convention for its own core, not the other branch's field name. ***
  //
  // *** DOCX RUN-LEVEL SetColor TAKES A TRAILING BOOL, PPTX RUN-LEVEL DOES NOT (both measured,
  // see euro-demo-docx.js:94 vs the pptx branch's own comment below) -- this is not an
  // inconsistency to fix, it is what each core's API actually accepts. ***
module.exports = {
  cores: ['docx', 'pptx'],
  emit(op, core) {
    const items = Array.isArray(op.runs) ? op.runs : []
    // Same discipline as `table`'s empty-rows refusal: an empty run list produces an empty
    // paragraph that otherwise looks entirely normal -- indistinguishable from a caller error
    // unless refused here, at the boundary this tool owns. Applies to both cores.
    if (!items.length) throw new Error('runs: `runs` is empty -- an empty run list would silently produce a blank paragraph')
    // Run.AddHyperlink is pptx-only IN THIS UNIT, same reasoning as `shape`'s hyperlink field
    // above -- package-verified against a pptx seed only.
    if (core === 'docx' && items.some((r) => r.hyperlink !== undefined && r.hyperlink !== null)) {
      throw notSupportedError('runs: hyperlink is pptx-only in this unit -- not verified for docx (Run.AddHyperlink was probed against a pptx seed only), refused rather than applied unverified')
    }
    if (core === 'docx') {
      const lines = ['var oParagraph = Api.CreateParagraph();']
      items.forEach((r, i) => emitDocxRun(`oRun${i}`, r, lines))
      lines.push('oDocument.Push(oParagraph);')
      return lines
    }
    const w = Number(op.width ?? 9000000)
    const h = Number(op.height ?? 800000)
    const x = Number(op.x ?? 800000)
    const y = Number(op.y ?? 800000)
    const lines = [
      `var oShape = Api.CreateShape("rect", ${w}, ${h}, Api.CreateNoFill(), Api.CreateStroke(0, Api.CreateNoFill()));`,
      `oShape.SetPosition(${x}, ${y});`,
      'var oContent = oShape.GetDocContent();',
      'var oParagraph = oContent.GetElement(0);',
      'oParagraph.RemoveAllElements();',
    ]
    items.forEach((r, i) => {
      const v = `oRun${i}`
      lines.push(`var ${v} = Api.CreateRun();`, `${v}.AddText(${jsString(String(r.text ?? ''))});`)
      if (r.bold) lines.push(`${v}.SetBold(true);`)
      if (r.italic) lines.push(`${v}.SetItalic(true);`)
      if (r.underline) lines.push(`${v}.SetUnderline(true);`)
      if (r.strikethrough) lines.push(`${v}.SetStrikeout(true);`)
      if (r.fontSize) lines.push(`${v}.SetFontSize(${Number(r.fontSize)});`)
      if (r.color !== undefined && r.color !== null) {
        const { r: rr, g, b } = parseHexColor(r.color)
        // Run-level SetColor takes three bytes, no trailing bool (recipe #3) -- unlike the
        // paragraph-level SetColor used elsewhere in this file, which takes a fourth argument.
        lines.push(`${v}.SetColor(${rr}, ${g}, ${b});`)
      }
      // Reuses KNOWN_VERT_ALIGNS and the exact call shape
      // emitDocxRun already uses on docx -- see that constant's comment for why this is safe
      // to reuse, not re-derived.
      if (r.vertAlign !== undefined && r.vertAlign !== null) {
        if (!KNOWN_VERT_ALIGNS.includes(r.vertAlign)) {
          throw notSupportedError(`runs: unknown vertAlign ${JSON.stringify(r.vertAlign)} (known: ${KNOWN_VERT_ALIGNS.join(', ')})`)
        }
        lines.push(`${v}.SetVertAlign(${jsString(r.vertAlign)});`)
      }
      lines.push(`oParagraph.AddElement(${v});`)
      // AddHyperlink's own precondition (recovered via toString()): the run must already be
      // part of a paragraph (`this.Run.GetParagraph()` non-null) -- MUST come after
      // AddElement, calling it earlier throws no error but silently returns null.
      if (r.hyperlink !== undefined && r.hyperlink !== null) {
        const { url, tooltip } = validateHyperlink(r.hyperlink, `runs[${i}]`)
        lines.push(`${v}.AddHyperlink(${jsString(url)}, ${jsString(tooltip)});`)
      }
    })
    lines.push('oSlide.AddObject(oShape);')
    return lines
  },
}
