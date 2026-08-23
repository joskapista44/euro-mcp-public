'use strict'
// Operations split boundary decision: this operation touches multiple cores (xlsx, docx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { notSupportedError, pageSetupDocx } = require('../lib-operations-registry.cjs')

  // MEASURED via full method enumeration on oWorksheet,
  // GetRange(), and Api itself (not guessed from the OOXML <pageSetup>/<pageMargins> schema) --
  // exactly seven Set* methods exist in this area:
  //   SetLeftMargin/SetRightMargin/SetTopMargin/SetBottomMargin -- WORK. Unit is MILLIMETRES,
  //     not inches: measured SetLeftMargin(2) -> saved pageMargins left="0.078740157..." which is
  //     exactly 2mm in inches (1mm = 0.03937007874in), not 2in and not 2cm.
  //   SetPrintHeadings/SetPrintGridlines -- WORK, plain booleans, produce <printOptions>.
  //   SetPageOrientation -- EXISTS (does not throw) but is a CONFIRMED NO-OP: tried a string
  //     ("landscape"), a boolean (true) and a number (1) -- saved orientation stayed "portrait"
  //     every time. Calibrated against a deliberately-unknown method name first (that DOES throw,
  //     "blocked"), so this silence is measured inertness, not a missed exception.
  // Paper size, fitToWidth/fitToHeight and print area have NO corresponding method ANYWHERE in
  // this surface (Worksheet, Range, and Api all enumerated) -- a structural gap, same shape as
  // freezePanes above, not a bug in this tool. The <sheetPr>/fitToPage two-part contract the
  // reference generator relies on (receptek-xlsx.md's own note) is therefore not reachable from
  // this API at all, not just partially: refused by name below, not silently dropped.
  // DOCX BRANCH (the docx-leltar's largest gap, `Section` 0/30):
  // see pageSetupDocx() above emitDocxRun -- pulled out to its own function (same reason as
  // emitDocxRun) with the full measurement writeup on units/argument order/orientation.
module.exports = {
  cores: ['xlsx', 'docx'],
  emit(op, core) {
    if (core === 'docx') return pageSetupDocx(op)
    const UNSUPPORTED = ['orientation', 'paperSize', 'fitToWidth', 'fitToHeight', 'printArea']
    const askedUnsupported = UNSUPPORTED.filter((k) => op[k] !== undefined)
    if (askedUnsupported.length) {
      throw notSupportedError(`pageSetup: ${askedUnsupported.join(', ')} not supported -- measured: no Set* method exists on this DocBuilder instance for paper size, fit-to-page or print area (orientation HAS a method, SetPageOrientation, but it is a confirmed no-op: string/boolean/number arguments all tried, none changed the saved orientation)`)
    }
    const SUPPORTED = ['marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'printHeadings', 'printGridlines']
    if (!SUPPORTED.some((k) => op[k] !== undefined)) {
      throw new Error(`pageSetup: at least one of ${SUPPORTED.join(', ')} is required`)
    }
    const lines = []
    if (op.marginLeft !== undefined) lines.push(`oWorksheet.SetLeftMargin(${Number(op.marginLeft)});`)
    if (op.marginRight !== undefined) lines.push(`oWorksheet.SetRightMargin(${Number(op.marginRight)});`)
    if (op.marginTop !== undefined) lines.push(`oWorksheet.SetTopMargin(${Number(op.marginTop)});`)
    if (op.marginBottom !== undefined) lines.push(`oWorksheet.SetBottomMargin(${Number(op.marginBottom)});`)
    if (op.printHeadings !== undefined) lines.push(`oWorksheet.SetPrintHeadings(${Boolean(op.printHeadings)});`)
    if (op.printGridlines !== undefined) lines.push(`oWorksheet.SetPrintGridlines(${Boolean(op.printGridlines)});`)
    return lines
  },
}
