'use strict'
// Operations split boundary decision: this operation touches multiple cores (docx, pptx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { KNOWN_WORDART_TRANSFORMS, buildFillExpression, jsString, notSupportedError, rgbArg, validateRgbColor } = require('../lib-operations-registry.cjs')

  // `Api.CreateWordArt(textPr, text, transform, fill, stroke, rotAngle, width, height[, x, y])`
  // was measured INDEPENDENTLY on two cores by two different units, landing at the same time:
  // D2 (docx, positional `color`/AddDrawing+Push) and K8 (pptx, object
  // `fill`/AddObject, this unit). The two branches are kept
  // DELIBERATELY SEPARATE rather than unified onto one field-set: D2 never measured a
  // gradient/pattern fill through CreateWordArt on docx, and this unit never measured docx's
  // AddDrawing+Push attach path -- each branch only claims what it actually ran.
module.exports = {
  cores: ['docx', 'pptx'],
  emit(op, core) {
    if (!op.text) throw new Error('wordArt: `text` is required')
    if (core === 'docx') {
      // D2: MEASURED via toString() against a docx seed -- a first attempt with a single
      // {text, style} OBJECT argument silently produced NO text in the saved package (the
      // internal `textPr && textPr.TextPr ? ... : null` check coerced it to null,
      // "succeeding" with an empty result); the real signature is POSITIONAL. The returned
      // drawing must be added via `oParagraph.AddDrawing(...)` -- a direct
      // `oDocument.Push(oWordArt)` silently drops it (no drawing, no error).
      const fill = op.color ? (validateRgbColor('wordArt: color', op.color), `Api.CreateSolidFill(Api.CreateRGBColor(${op.color.map(Number).join(', ')}))`) : 'Api.CreateNoFill()'
      const w = Number(op.width ?? 3000000)
      const h = Number(op.height ?? 1000000)
      const rotation = Number(op.rotation ?? 0)
      const transform = op.transform ? jsString(op.transform) : 'null'
      return [
        `var oWordArt = Api.CreateWordArt(null, ${jsString(op.text)}, ${transform}, ${fill}, Api.CreateStroke(0, Api.CreateNoFill()), ${rotation}, ${w}, ${h});`,
        'var oWordArtPar = Api.CreateParagraph();',
        'oWordArtPar.AddDrawing(oWordArt);',
        'oDocument.Push(oWordArtPar);',
      ]
    }
    // K8: MEASURED live against a pptx seed (toString()-recovered signature + package-
    // verified: <a:prstTxWarp prst="textNoShape"> round-tripped). Reuses buildFillExpression
    // (shared with `shape.fill` above) so a WordArt can also take a gradient/pattern fill.
    //
    // `x`/`y` are DELIBERATELY OPTIONAL WITHOUT A HARD DEFAULT here (unlike `shape`, which
    // always defaults to 500000/500000): CreateWordArt's OWN source auto-centers on the slide
    // when these are omitted (measured: `nIndLeft>-1?nIndLeft:(presentationWidth-artWidth)/2`,
    // same for Y) -- passing a synthetic default would silently take away a real, useful
    // behaviour of the underlying call.
    const transform = op.transform !== undefined && op.transform !== null ? String(op.transform) : 'textNoShape'
    if (!KNOWN_WORDART_TRANSFORMS.includes(transform)) {
      throw notSupportedError(`wordArt: unknown transform ${JSON.stringify(transform)} (known: ${KNOWN_WORDART_TRANSFORMS.join(', ')}) -- Api.CreateWordArt falls back to "textNoShape" silently for anything not a non-empty string, so an unrecognised name here would silently produce a DIFFERENT shape than requested, refused instead`)
    }
    const fill = buildFillExpression(op.fill, 'wordArt')
    const lineWidth = Number(op.lineWidth ?? 0)
    const lineFill = op.lineColor ? `Api.CreateSolidFill(Api.CreateRGBColor(${rgbArg(op.lineColor)}))` : 'Api.CreateNoFill()'
    const rotation = Number(op.rotation ?? 0)
    const width = Number(op.width ?? 1828800)
    const height = Number(op.height ?? 1828800)
    const hasX = op.x !== undefined && op.x !== null
    const hasY = op.y !== undefined && op.y !== null
    if (hasX !== hasY) {
      throw new Error('wordArt: `x` and `y` must both be given or both omitted -- CreateWordArt auto-centers only when BOTH are absent')
    }
    const posArgs = hasX ? `, ${Number(op.x)}, ${Number(op.y)}` : ''
    return [
      `var oArt = Api.CreateWordArt(null, ${jsString(String(op.text))}, ${jsString(transform)}, ${fill}, Api.CreateStroke(${lineWidth}, ${lineFill}), ${rotation}, ${width}, ${height}${posArgs});`,
      'oSlide.AddObject(oArt);',
    ]
  },
}
