'use strict'
// Operations split boundary decision: this operation touches multiple cores (docx, pptx, xlsx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { applyDocxTableCellRefinements, applyDocxTableMerge, applyDocxTableRowRefinements, applyPptxTableCellSettings, applyPptxTableLookAndStructuralExtras, applyPptxTableMerge, buildDocxTableLines, jsString, notSupportedError, validateDocxTableRefinements, xlsxTableCellRefs } = require('../lib-operations-registry.cjs')
module.exports = {
  cores: ['docx', 'pptx', 'xlsx'],
  // E2 -- style options, all optional and all independently switchable:
  //   header (default true)   bold + filled first row, on the KIADOTT csúcson
  //   zebra  (default false)  every second BODY row (i.e. row index 2, 4, ... counting the
  //                           header as row 0) gets a light fill -- the header itself is never
  //                           re-styled by this, `header` and `zebra` do not fight over row 0
  //   border (default true)   a full grid, all four sides + inside lines where the core has them
  //   merge  xlsx: array of "A1:C1"-style range strings, merged AFTER the values are written.
  //          docx: array of [startRow, startCol, endRow, endCol]
  //          0-indexed tuples -- see applyDocxTableMerge for the measured call shape. pptx is
  //          still out of scope for this field (E2's original "MI NEM TARTOZIK BELE" stands
  //          there; only the docx half of it was ever measured).
  // docx-K2, all table-wide unless noted:
  //   repeatHeaderRow (bool)      row 0 repeats as a page header on multi-page tables
  //   rowHeights (array, docx)   one entry per row, twips, "atLeast" only -- see the emit()
  //                              comment by GetRow(i).SetHeight for why no other rule is offered
  //   verticalAlign (docx)      "top"/"center"/"bottom", every cell
  //   noWrap (bool, docx)       every cell
  //   cellMarginTop/Bottom/Left/Right (docx, each independently optional)  twips, every cell
  // Colors default to the palette measured live against 3ea5e28 (euro-demo-docx.js / the E2
  // probe session): dark-blue header, near-white zebra, blue border. Every call below is one
  // that was run against the real Document Server and checked in the SAVED PACKAGE, not just
  // read off an `ok:true` -- this operation table has a documented graveyard of calls that
  // answer green and change nothing (SetBorders("All", ...), Range.SetAutoFilter(true), ...).
  emit(op, core) {
    const rows = Array.isArray(op.rows) ? op.rows : []
    // A zero-row table is refused rather than emitted: an empty table looks exactly like a
    // passing test, and that is the failure this whole tool set keeps running into.
    if (!rows.length) throw new Error('table: `rows` is empty -- an empty table cannot be told apart from a successful one')
    const header = op.header !== false
    const zebra = Boolean(op.zebra)
    const border = op.border !== false
    const headerFill = op.headerColor ?? [0x1f, 0x38, 0x64]
    const headerText = op.headerTextColor ?? [0xff, 0xff, 0xff]
    const zebraFill = op.zebraColor ?? [0xf7, 0xf9, 0xfc]
    const borderColor = op.borderColor ?? [0x1a, 0x5f, 0xb4]
    const rgb = (c) => c.map(Number).join(', ')

    if (core === 'xlsx') {
      const at = String(op.at ?? 'A1')
      const startCol = at.replace(/\d+/g, '')
      const startRow = Number(at.replace(/\D+/g, '') || 1)
      const cols = Math.max(...rows.map((r) => (Array.isArray(r) ? r.length : 1)))
      const colLetter = (j) => String.fromCharCode(startCol.charCodeAt(0) + j)
      const rowRange = (i) => `${colLetter(0)}${startRow + i}:${colLetter(cols - 1)}${startRow + i}`
      // The ref half of this line (colLetter(j) +
      // (startRow + i)) is now `xlsxTableCellRefs(op)`, not re-derived here -- same row-major
      // (i then j) order as the loop below it produces the values in, so zipping the two
      // arrays index-for-index reproduces exactly what the two inline expressions did. This is
      // the ONE place that computes a table's cell refs; the presence-check side calls the same
      // function, so the two can never silently disagree about which cells were asked for.
      const cellRefs = xlsxTableCellRefs(op)
      const lines = []
      let refIdx = 0
      rows.forEach((sor) => {
        const cells = Array.isArray(sor) ? sor : [sor]
        cells.forEach((cella) => {
          lines.push(`oWorksheet.GetRange(${jsString(cellRefs[refIdx])}).SetValue(${jsString(cella)});`)
          refIdx += 1
        })
      })
      rows.forEach((_sor, i) => {
        if (header && i === 0) {
          lines.push(
            `oWorksheet.GetRange(${jsString(rowRange(0))}).SetBold(true);`,
            `oWorksheet.GetRange(${jsString(rowRange(0))}).SetFillColor(Api.CreateColorFromRGB(${rgb(headerFill)}));`,
            `oWorksheet.GetRange(${jsString(rowRange(0))}).SetFontColor(Api.CreateColorFromRGB(${rgb(headerText)}));`,
          )
        } else if (zebra && i % 2 === 0) {
          lines.push(`oWorksheet.GetRange(${jsString(rowRange(i))}).SetFillColor(Api.CreateColorFromRGB(${rgb(zebraFill)}));`)
        }
      })
      if (border) {
        // *** "All"/"Outline" ARE A SILENT NO-OP ON THIS INSTANCE (measured, receptek-xlsx.md
        // #5) -- the full grid needs the six real positions named one by one. ***
        const full = `${colLetter(0)}${startRow}:${colLetter(cols - 1)}${startRow + rows.length - 1}`
        for (const pos of ['Left', 'Top', 'Right', 'Bottom', 'InsideVertical', 'InsideHorizontal']) {
          lines.push(`oWorksheet.GetRange(${jsString(full)}).SetBorders(${jsString(pos)}, "Thin", Api.CreateColorFromRGB(${rgb(borderColor)}));`)
        }
      }
      const merges = Array.isArray(op.merge) ? op.merge : []
      for (const range of merges) {
        lines.push(`oWorksheet.GetRange(${jsString(range)}).Merge(false);`)
      }
      return lines
    }

    if (core === 'docx') {
      const lines = buildDocxTableLines('oTable', op, { n: 0 })
      lines.push('oDocument.Push(oTable);')
      return lines
    }

    const cols = Math.max(...rows.map((r) => (Array.isArray(r) ? r.length : 1)))
    // E2b -- per-column width, docx only. `columnWidths`: array of percentages, one per column,
    // need not sum to 100. MEASURED 2026-08-15 against a previously shipped build: pptx has NO
    // equivalent -- the full `for...in` surface of its Table/Row/Cell objects (reflected live,
    // `.toString()`'d where a name looked promising) has no per-gridCol setter, only AddColumn/
    // RemoveColumn (structural, not a resize) and the table-wide SetSize this operation already
    // uses. That is a NAMED NO, not an unexplored gap -- a `columnWidths` request on pptx is
    // refused rather than silently ignored, same reasoning as the empty-table refusal above.
    if (core === 'pptx' && op.columnWidths) {
      // Tagged notSupportedError (message unchanged) --
      // this is one of the four canonical NEM-TAMOGATOTT examples the card names by hand, and
      // an untagged plain Error here would have been misclassified as HIBA by the report.
      throw notSupportedError('table: `columnWidths` is not available in the pptx core -- measured 2026-08-15: the Table/Row/Cell API surface on this Document Server has no per-column width setter (only AddColumn/RemoveColumn, which restructure the grid rather than resize it); pptx has table-wide SetSize only')
    }
    // *** ON docx, "dxa" (twips) IS A SILENT NO-OP FOR Cell.SetWidth -- MEASURED, not assumed:
    // three cells set to 1000/3000/5000 dxa came back with `tcW w:type="auto"` on every one,
    // unchanged. "percent" DOES land (10/30/60 -> tcW w:w=500/1500/3000, exactly *50 -- the same
    // ratio the table-level SetWidth("percent", ...) call already uses), so that is the only unit
    // this operation emits. SetTableLayout("fixed") is paired with it: "auto" layout is free to
    // recompute widths from content regardless of what tcW says. ***
    const lines = [`var oTable = Api.CreateTable(${cols}, ${rows.length});`]
    // *** THE CELL IS REACHED DIFFERENTLY IN THE TWO CORES, and getting it wrong does not throw:
    // it kills the whole job with no output and no error message (measured 2026-08-15, bisected
    // down to this single call). Slides need GetRow(i).GetCell(j); the text core takes GetCell(i, j).
    const cella = (i, j) => (core === 'pptx' ? `oTable.GetRow(${i}).GetCell(${j})` : `oTable.GetCell(${i}, ${j})`)
    if (core === 'pptx') {
      // A table on a slide is a graphic frame: without a position and a size it has nowhere to go.
      lines.push(
        `oTable.SetPosition(${Number(op.x ?? 1500000)}, ${Number(op.y ?? 1500000)});`,
        `oTable.SetSize(${Number(op.width ?? 8000000)}, ${Number(op.height ?? 2000000)});`,
      )
    } else {
      // docx: a percent width, so the table fills its column regardless of the page size.
      lines.push(`oTable.SetWidth("percent", ${Number(op.widthPercent ?? 100)});`)
    }
    if (core === 'docx' && border) {
      // Table-level, both outer and inner lines in one call each -- measured (receptek-pptx-docx.md #11).
      for (const side of ['Top', 'Bottom', 'Left', 'Right', 'InsideH', 'InsideV']) {
        lines.push(`oTable.SetTableBorder${side}("single", 4, 0, ${rgb(borderColor)});`)
      }
    }
    if (core === 'docx' && Array.isArray(op.columnWidths) && op.columnWidths.length) {
      lines.push('oTable.SetTableLayout("fixed");')
    }
    const colWidths = core === 'docx' && Array.isArray(op.columnWidths) ? op.columnWidths : null

    // docx-K2 table refinement, all four MEASURED against a real
    // DocBuilder round-trip (2026-08-17), unzipped and read back from word/document.xml:
    //   verticalAlign (top/center/bottom, table-wide) -> Cell.SetVerticalAlign, all three
    //     individually confirmed present as <w:vAlign w:val="..."/> in the saved cell tcPr.
    //   noWrap (table-wide bool) -> Cell.SetNoWrap -- confirmed present as an empty <w:noWrap/>.
    //   cellMarginTop/Bottom/Left/Right (twips, table-wide, each independently optional) ->
    //     Cell.SetCellMargin<Side> -- all four confirmed present in <w:tcMar> with the exact
    //     called values (dxa units, no conversion).
    //   textDirection is refused -- see validateDocxTableRefinements for the measurement.
    if (core === 'docx') validateDocxTableRefinements(op)
    rows.forEach((sor, i) => {
      const cells = Array.isArray(sor) ? sor : [sor]
      const isHeader = header && i === 0
      const isZebra = !isHeader && zebra && i % 2 === 0
      cells.forEach((ertek, j) => {
        const cell = cella(i, j)
        if (core === 'docx') {
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
          // E2b -- set on EVERY row's cell for the column, not just the header: a fixed-layout
          // table reads tcW per cell, and leaving later rows at their default risks an
          // inconsistent grid.
          if (colWidths && colWidths[j] != null) {
            lines.push(`${cell}.SetWidth("percent", ${Number(colWidths[j])});`)
          }
          applyDocxTableCellRefinements(lines, op, cell)
        } else {
          // pptx: bold/color live on the RUN, not the paragraph (receptek-pptx-docx.md #3+#4) --
          // and cell.SetCellBorderTop/... crashes the whole job if handed a Stroke object instead
          // of (size_mm, ApiFill); measured live 2026-08-15 (bisected: Stroke -> outcome:"blocked").
          const run = `o${core === 'pptx' ? 'Run' : 'R'}_${i}_${j}`
          lines.push(
            `var ${run} = Api.CreateRun();`,
            `${run}.AddText(${jsString(ertek)});`,
            `${cell}.GetContent().GetElement(0).RemoveAllElements();`,
            `${cell}.GetContent().GetElement(0).AddElement(${run});`,
          )
          if (isHeader) {
            lines.push(
              `${run}.SetBold(true);`,
              `${run}.SetColor(${rgb(headerText)});`,
              `${cell}.SetShd(Api.CreateSolidFill(Api.CreateRGBColor(${rgb(headerFill)})));`,
            )
          } else if (isZebra) {
            lines.push(`${cell}.SetShd(Api.CreateSolidFill(Api.CreateRGBColor(${rgb(zebraFill)})));`)
          }
          if (border) {
            // fSize is MILLIMETRES here (measured: output EMU = fSize * 36000), not EMU and not
            // an eighth-of-a-point like the docx table-border call -- a different unit per core.
            const fill = `Api.CreateSolidFill(Api.CreateRGBColor(${rgb(borderColor)}))`
            for (const side of ['Top', 'Bottom', 'Left', 'Right']) {
              lines.push(`${cell}.SetCellBorder${side}(0.35, ${fill});`)
            }
          }
        }
      })
    })
    if (core === 'docx') applyDocxTableMerge(lines, op, rows.length, cols, cella)
    if (core === 'docx') applyDocxTableRowRefinements(lines, op)
    if (core === 'pptx') applyPptxTableCellSettings(lines, op, rows.length, cols, cella)
    if (core === 'pptx') applyPptxTableMerge(lines, op, rows.length, cols, cella)
    applyPptxTableLookAndStructuralExtras(lines, op, core)
    lines.push(core === 'docx' ? 'oDocument.Push(oTable);' : 'oSlide.AddObject(oTable);')
    return lines
  },
}
