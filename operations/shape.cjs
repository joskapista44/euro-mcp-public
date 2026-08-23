'use strict'
// Operations split boundary decision: this operation touches multiple cores (pptx, docx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { KNOWN_VERTICAL_TEXT_ALIGNS, applyPptxShapePostCreationFields, applyShapeGeometry, jsString, notSupportedError, resolveShapeFill } = require('../lib-operations-registry.cjs')

  // E6 -- SHAPE: pptx (absolute-positioned, `oSlide.AddObject`) and docx --
  // a standalone shape, distinct from `text`'s always-borderless,
  // always-unfilled text box (lib.cjs `text` core === 'pptx' branch above). This op exposes
  // the fill and the border, which `text` deliberately does not.
  //
  // *** DOCX HAS NO SLIDE-ABSOLUTE POSITION, SO `SetPosition` IS PPTX-ONLY. *** A docx shape is
  // a drawing object hung on a paragraph, the same embedding pattern as `image`'s docx branch
  // above (`AddDrawing` + `oDocument.Push`) -- measured against the reference generator
  // (an internal demo-docx script), which never calls `SetPosition` on its docx
  // shape and instead does `pShape.AddDrawing(oShape); oDocument.Push(pShape);`.
module.exports = {
  cores: ['pptx', 'docx'],
  emit(op, core) {
    // Shape.SetHyperlink is pptx-only IN THIS UNIT: package-verified on pptx (see
    // validateHyperlink's comment), but never probed against a docx seed (none available at
    // implementation time) -- refused rather than silently applied unverified or silently
    // dropped, same discipline as pptx-only fields elsewhere in this file (e.g. columnWidths).
    if (op.hyperlink !== undefined && op.hyperlink !== null && core !== 'pptx') {
      throw notSupportedError('shape: hyperlink is pptx-only in this unit -- not verified for docx (Shape.SetHyperlink was probed against a pptx seed only), refused rather than applied unverified')
    }
    // `rotation`/`verticalTextAlign` are pptx-only IN THIS UNIT:
    // both call shapes recovered via toString() and package-verified on a pptx seed only (no
    // docx seed file was available at implementation time) -- refused on docx rather than
    // applied unverified, same discipline as `shape.hyperlink` (see its own comment).
    if (core !== 'pptx' && ((op.rotation !== undefined && op.rotation !== null) || (op.verticalTextAlign !== undefined && op.verticalTextAlign !== null))) {
      throw notSupportedError('shape: rotation/verticalTextAlign are pptx-only in this unit -- not verified for docx, refused rather than applied unverified')
    }
    if (op.verticalTextAlign !== undefined && op.verticalTextAlign !== null && !KNOWN_VERTICAL_TEXT_ALIGNS.includes(op.verticalTextAlign)) {
      // Shape.SetVerticalTextAlign's own source (toString()-recovered) is a switch with NO
      // default case -- an unrecognised value does not throw, it silently does nothing.
      throw notSupportedError(`shape: unknown verticalTextAlign ${JSON.stringify(op.verticalTextAlign)} (known: ${KNOWN_VERTICAL_TEXT_ALIGNS.join(', ')}) -- an unrecognised value does NOT throw inside the builder, it silently leaves the alignment unset, so this tool refuses it here instead`)
    }
    const shapeType = String(op.shapeType ?? 'rect')
    const w = Number(op.width ?? 2000000)
    const h = Number(op.height ?? 2000000)
    const rgb = (c) => c.map(Number).join(', ')
    // `fill` now ALSO accepts {type:'gradient'|'pattern',...}
    // (pptx-only in this unit) alongside D2's existing docx-only fillGradient/fillPattern --
    // resolveShapeFill's own header comments explain why the two stay separate rather than
    // merged into one cross-core syntax. A bare [r,g,b] array still means solid, unchanged.
    const fill = resolveShapeFill(op, core, rgb)
    // CreateStroke's width is EMU, like the shape's own size -- unlike table's cell-border
    // setter (SetCellBorder*, size_mm + Fill-object), which is a different method on a
    // different object and is NOT the API this line calls.
    const borderWidth = Number(op.borderWidth ?? 0)
    const borderFill = op.borderColor ? `Api.CreateSolidFill(Api.CreateRGBColor(${rgb(op.borderColor)}))` : 'Api.CreateNoFill()'
    const lines = [
      `var oShape = Api.CreateShape(${jsString(shapeType)}, ${w}, ${h}, ${fill}, Api.CreateStroke(${borderWidth}, ${borderFill}));`,
    ]
    applyShapeGeometry(lines, op, core)
    if (core === 'pptx') {
      const x = Number(op.x ?? 500000)
      const y = Number(op.y ?? 500000)
      lines.push(`oShape.SetPosition(${x}, ${y});`)
    }
    applyPptxShapePostCreationFields(lines, op, core)
    if (op.text) {
      lines.push(
        'var oShapeContent = oShape.GetDocContent();',
        'var oShapePara = oShapeContent.GetElement(0);',
        'oShapePara.RemoveAllElements();',
        'var oShapeRun = Api.CreateRun();',
        `oShapeRun.AddText(${jsString(op.text)});`,
        'oShapePara.AddElement(oShapeRun);',
      )
    }
    if (core === 'docx') {
      lines.push('var oShapePar = Api.CreateParagraph();', 'oShapePar.AddDrawing(oShape);', 'oDocument.Push(oShapePar);')
    } else {
      lines.push('oSlide.AddObject(oShape);')
    }
    return lines
  },
}
