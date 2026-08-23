'use strict'
// Operations split boundary decision: this operation touches multiple cores (docx, pptx, xlsx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { jsString, resolveImageSrc } = require('../lib-operations-registry.cjs')
module.exports = {
  cores: ['docx', 'pptx', 'xlsx'],
  emit(op, core) {
    // `data:` URIs work (measured) -- no external URL is needed, so no egress question arises.
    // `path` is the local-file route -- the caller gives a path, this
    // tool reads it and encodes it itself, so a caller never has to hand-roll a data: URI.
    // Downloading from an external URL is explicitly out of scope (egress-gate question,
    // separate decision) -- only a path already on this filesystem or an already-encoded
    // `src` are accepted.
    const src = resolveImageSrc(op)
    const w = Number(op.width ?? 2000000)
    const h = Number(op.height ?? 2000000)
    if (core === 'xlsx') return [`oWorksheet.AddImage(${jsString(src)}, ${w}, ${h}, ${Number(op.col ?? 1)}, 0, ${Number(op.row ?? 1)}, 0);`]
    const lines = [`var oImage = Api.CreateImage(${jsString(src)}, ${w}, ${h});`]
    if (core === 'pptx') {
      lines.push(`oImage.SetPosition(${Number(op.x ?? 1000000)}, ${Number(op.y ?? 1000000)});`, 'oSlide.AddObject(oImage);')
    } else {
      lines.push('var oImgPara = Api.CreateParagraph();', 'oImgPara.AddDrawing(oImage);', 'oDocument.Push(oImgPara);')
    }
    return lines
  },
}
