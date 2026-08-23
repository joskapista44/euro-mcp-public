'use strict'
// This file owns the XLSX-only OPERATIONS entries (extracted verbatim -- same emit()/validate()
// bodies, same comments, same order). The multi-core entries that also touch xlsx (per each
// one's own `cores` array) are pulled in from their own individual modules under operations/,
// not duplicated here. lib.cjs remains the aggregator: it populates
// lib-operations-registry.cjs with every helper function/const these files may need BEFORE
// requiring any of them, so this file (and its siblings) can destructure from a plain,
// already-populated object -- no circular require back into lib.cjs, and no function wrapper
// around the entries themselves (a wrapper would roll every nested emit()'s complexity up into
// one artificially huge factory-function score, which is exactly what the first draft of this
// split did before qlty smells caught it).
const { KNOWN_THEME_NAMES, jsString, notSupportedError, rgbArg, validateRgbColor } = require('./lib-operations-registry.cjs')
const table = require('./operations/table.cjs')
const image = require('./operations/image.cjs')
const chart = require('./operations/chart.cjs')
const pageSetup = require('./operations/pageSetup.cjs')
module.exports = {
  table,
  image,
  chart,
  pageSetup,

  // --- xlsx workbook operations. Every call shape below comes from an internal package-verified
  // recipe document -- the section numbers in the comments point at the recipe that
  // proves it, including its bad-input behaviour.
  numberFormat: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('numberFormat: `at` (a cell/range reference) is required')
      if (!op.format) throw new Error('numberFormat: `format` is required')
      return [`oWorksheet.GetRange(${jsString(op.at)}).SetNumberFormat(${jsString(op.format)});`]
    },
  },

  formula: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('formula: `at` (a cell reference) is required')
      const f = String(op.formula ?? '')
      // Measured (receptek-xlsx.md #1): SetValue stores whatever string it is given verbatim -- a
      // string not starting with "=" would silently become a plain text value, not a formula. That
      // is a different, already-covered capability (the `table` operation), so this one refuses
      // rather than quietly doing something other than what its name promises.
      if (!f.startsWith('=')) throw new Error('formula: `formula` must start with "=" -- SetValue stores any string verbatim, so a non-formula string here would silently become a plain value instead of a live formula (use `table` for plain values)')
      return [`oWorksheet.GetRange(${jsString(op.at)}).SetValue(${jsString(f)});`]
    },
  },

  columnWidth: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.column) throw new Error('columnWidth: `column` (a column letter, e.g. "B") is required')
      const width = Number(op.width)
      // Measured (receptek-xlsx.md #3): a negative width does NOT throw in the builder -- it
      // silently HIDES the column instead (hidden="1", width="0"). Refused here, at the boundary
      // this tool owns, so the caller gets a named error instead of a column that quietly vanishes.
      if (!Number.isFinite(width) || width <= 0) {
        throw new Error('columnWidth: `width` must be a positive number -- measured: a negative value does not error, it silently HIDES the column instead (hidden="1", width="0")')
      }
      return [`oWorksheet.GetRange(${jsString(op.column + '1')}).SetColumnWidth(${width});`]
    },
  },

  // SetBorders WORKS and is package-verified (receptek-xlsx.md #5) but was only reachable from
  // inside the `table` operation's own fixed six-position loop -- a caller could not ask for a
  // border on an arbitrary range directly.
  border: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('border: `at` (a range reference) is required')
      // The eight positions that are measured to actually draw a border on this DocBuilder
      // instance (receptek-xlsx.md #5, bisected against the two that do not).
      const VALID_POSITIONS = ['Left', 'Top', 'Right', 'Bottom', 'InsideVertical', 'InsideHorizontal', 'DiagonalDown', 'DiagonalUp']
      const position = String(op.position ?? '')
      // Measured (receptek-xlsx.md #5): "All" and "Outline" are a SILENT no-op on this
      // DocBuilder instance -- the call returns success and the saved package's borders count
      // stays at 1 (the package's own default "none" element), no border XML is added. A
      // previous reference script had this marked "confirmed working" from the call's own
      // ok:true answer alone; the saved package told a different story. Refused here rather
      // than repeating that mistake for every future caller.
      if (!VALID_POSITIONS.includes(position)) {
        throw notSupportedError(`border: unknown position ${JSON.stringify(op.position)} (known: ${VALID_POSITIONS.join(', ')}) -- "All" and "Outline" in particular do NOT throw inside the builder, they silently add no border to the saved package at all (borders count stays at the package's default 1, "none"); a full box needs the four edge positions named individually`)
      }
      const style = op.style ? String(op.style) : 'Thin'
      const color = op.color ?? [0, 0, 0]
      validateRgbColor('border', color)
      return [`oWorksheet.GetRange(${jsString(op.at)}).SetBorders(${jsString(position)}, ${jsString(style)}, Api.CreateColorFromRGB(${rgbArg(color)}));`]
    },
  },

  // SetFillColor/SetFontColor WORK and are package-verified (receptek-xlsx.md #6) but were only
  // reachable from inside the `table` operation's own header/zebra styling -- a caller could not
  // color an arbitrary range directly. Two operations, not one combined `cellStyle`: every other
  // entry in this table is single-concern (columnWidth, border, autoFilter, ...), and a caller
  // wanting only the font color (or only the fill) should not have to pass a no-op value for
  // the other.
  fillColor: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('fillColor: `at` (a range reference) is required')
      validateRgbColor('fillColor', op.color)
      // Pure white/black map to an indexed-color slot here too, not exact RGB -- see the
      // `fontColor` entry below for the measured detail, it applies to both properties.
      return [`oWorksheet.GetRange(${jsString(op.at)}).SetFillColor(Api.CreateColorFromRGB(${rgbArg(op.color)}));`]
    },
  },

  fontColor: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('fontColor: `at` (a range reference) is required')
      validateRgbColor('fontColor', op.color)
      // Measured (receptek-xlsx.md #6 found this for white on SetFontColor; extended and
      // live-verified 2026-08-16): pure white and pure black do
      // NOT come back as exact rgb="FFFFFFFF"/"FF000000" in the saved package on EITHER
      // SetFontColor or SetFillColor -- the builder maps both onto a built-in indexed-color slot
      // instead (<color indexed="65"/> for white, indexed="64" for black, consistently on both
      // properties). This is not a sign the call failed; a caller (or a gate) checking the saved
      // package for exact RGB equality on white/black needs to accept the indexed form too.
      return [`oWorksheet.GetRange(${jsString(op.at)}).SetFontColor(Api.CreateColorFromRGB(${rgbArg(op.color)}));`]
    },
  },

  autoFilter: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.range) throw new Error('autoFilter: `range` is required')
      // Measured (receptek-xlsx.md #7): the ONLY working call takes ZERO arguments --
      // SetAutoFilter(true) looks like the natural boolean-setter shape and is a silent no-op.
      //
      // SetAutoFilter() is a TOGGLE on this Document Server instance -- calling it a SECOND
      // time on a range that already carries a filter REMOVES it instead of being a safe no-op
      // (measured 2026-08-17, package-verified on a disposable file: two consecutive no-arg
      // calls to the same range leave zero <autoFilter> elements). This is exactly how the real
      // W-969 source file lost its filter twice -- a later script re-called this op on a range
      // an earlier one had already filtered. The document is the only place this state lives
      // (separate script runs share no state with this translator), so the guard has to ask the
      // DOCUMENT at runtime, not check anything here at build time.
      return [
        '(function() {',
        '  var af = oWorksheet.GetAutoFilter();',
        '  var already = af.GetFilterMode();',
        '  var existingRange = already ? af.GetRange().GetAddress() : null;',
        `  var wantedRange = ${jsString(op.range)};`,
        '  if (already && existingRange === wantedRange) {',
        '    // Exactly what was asked is already there. A no-op is correct AND safe: calling',
        '    // SetAutoFilter() again here is precisely the toggle that would remove it.',
        '  } else if (already) {',
        '    throw new Error("autoFilter: the sheet already has a filter on " + existingRange + ", refusing to touch " + wantedRange + " -- calling SetAutoFilter() again is a TOGGLE on this instance and could remove the existing one; remove it explicitly first if you mean to replace it")',
        '  } else {',
        '    oWorksheet.GetRange(wantedRange).SetAutoFilter();',
        '  }',
        '})();',
      ]
    },
  },

  sort: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.range) throw new Error('sort: `range` is required')
      if (!op.keyCell) throw new Error('sort: `keyCell` (a single cell inside `range` naming the sort column) is required')
      // Measured (receptek-xlsx.md #8, NO 1): the direction argument to SetSort is INERT -- every
      // call sorts ascending regardless of what is passed. There is no working descending call
      // today, so this operation refuses a `descending` request rather than silently ignoring it.
      if (op.descending) {
        throw new Error('sort: descending order is not supported -- measured: the direction argument to SetSort has no effect on this DocBuilder instance, every call sorts ascending regardless of the value passed')
      }
      return [`oWorksheet.GetRange(${jsString(op.range)}).SetSort(oWorksheet.GetRange(${jsString(op.keyCell)}), true);`]
    },
  },

  conditionalFormatting: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.range) throw new Error('conditionalFormatting: `range` is required')
      const variant = String(op.variant ?? '')
      const conditions = `oWorksheet.GetRange(${jsString(op.range)}).GetFormatConditions()`
      const WORKING = {
        colorScale: 'AddColorScale',
        dataBar: 'AddDatabar',
        iconSet: 'AddIconSetCondition',
        top10: 'AddTop10',
        aboveAverage: 'AddAboveAverage',
        uniqueValues: 'AddUniqueValues',
      }
      if (WORKING[variant]) return [`${conditions}.${WORKING[variant]}();`]
      // Measured (receptek-xlsx.md #9.7): FormatConditions.Add(...) -- the generic, cell-value/
      // expression-based rule -- is INERT on this DocBuilder instance: five different argument
      // shapes tried, every one silently returned an empty object and added nothing to the
      // package, while its six sibling methods above all work. Named here so a caller asking for
      // "cellValue" or "expression" gets that explanation, not a generic "unknown variant".
      if (variant === 'cellValue' || variant === 'expression') {
        throw new Error(`conditionalFormatting: variant "${variant}" is not supported -- measured: FormatConditions.Add() is inert on this DocBuilder instance regardless of argument shape (5 shapes tried, 0 package output; its 6 sibling methods all work). Supported variants: ${Object.keys(WORKING).join(', ')}`)
      }
      throw new Error(`conditionalFormatting: unknown variant ${JSON.stringify(variant)} (known: ${Object.keys(WORKING).join(', ')}; cellValue/expression are confirmed not-supported, see receptek-xlsx.md #9.7)`)
    },
  },

  // The earlier "not supported" conclusion above was WRONG --
  // measured (2026-08-16, live DocBuilder run, package-verified): the setter is not on Worksheet
  // (GetFreezePanes there really has no Set* counterpart, that part still holds), it is on `Api`:
  // `Api.SetFreezePanesType(mode)`. Reading its own source (Api.SetFreezePanesType.toString()):
  //   "row"    -> ALWAYS freezes exactly row 1, regardless of any selection
  //   "column" -> ALWAYS freezes exactly column A, regardless of any selection
  //   "cell"   -> freezes everything above/left of the CURRENTLY SELECTED cell -- there is no
  //               direct row/column-index parameter, so this mode needs a Range.Select() first
  // All three confirmed in the saved package's xl/worksheets/sheet1.xml <pane .../> element
  // (xSplit/ySplit/topLeftCell match the selected cell for "cell"; "row"/"column" always produce
  // ySplit="1"/xSplit="1" at topLeftCell="A2"/"B1" no matter what was selected beforehand).
  freezePanes: {
    cores: ['xlsx'],
    emit(op) {
      const VALID_MODES = ['row', 'column', 'cell']
      const mode = String(op.mode ?? '')
      if (!VALID_MODES.includes(mode)) {
        throw new Error(`freezePanes: \`mode\` must be one of ${VALID_MODES.join(', ')}, got ${JSON.stringify(op.mode)}`)
      }
      if (mode !== 'cell' && op.cell !== undefined) {
        throw new Error(`freezePanes: \`cell\` is only used with mode "cell" -- "${mode}" always freezes exactly the first row/column, measured to ignore any selection`)
      }
      const lines = []
      if (mode === 'cell') {
        if (!op.cell) throw new Error('freezePanes: `cell` (a cell reference, e.g. "B2") is required when mode is "cell"')
        lines.push(`oWorksheet.GetRange(${jsString(op.cell)}).Select();`)
      }
      lines.push(`Api.SetFreezePanesType(${jsString(mode)});`)
      return lines
    },
  },

  // MEASURED -- `ApiRange.SetRowHeight(height)` WORKS (a single
  // argument; unlike ApiWorksheet.SetRowHeight, the row is already implied by the range's own
  // reference, e.g. GetRange("A3:C3").SetRowHeight(40) -> <row r="3" ht="40" customHeight="1"/>).
  // `ApiRange.SetHidden(bool)` is CONFIRMED INERT: called on a single cell (A1) or a 3-row range
  // (B1:B3), neither the sheet XML's <row>/<cols> elements nor the cell's own style (cellXfs,
  // which would carry a `<protection hidden="1"/>` if this meant cell-protection-hiding) show any
  // change at all -- cellXfs count stays at 1 either way. Refused client-side rather than emitted
  // as a silent no-op.
  rangeSize: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeSize: `at` is required')
      if (op.hidden !== undefined) {
        throw notSupportedError('rangeSize: `hidden` not supported -- measured: ApiRange.SetHidden() is inert on this DocBuilder instance (tried on a single cell and a 3-row range, neither the sheet XML nor the cell style changed at all)')
      }
      if (op.rowHeight === undefined) throw new Error('rangeSize: `rowHeight` is required (the only supported field)')
      return [`oWorksheet.GetRange(${jsString(op.at)}).SetRowHeight(${Number(op.rowHeight)});`]
    },
  },

  // MEASURED, all three working -- `ApiRange.SetAlignVertical/
  // SetAlignHorizontal/SetReadingOrder`, all string-valued. Package-verified in xl/styles.xml
  // cellXfs: SetAlignVertical("center")+SetAlignHorizontal("right") ->
  // <alignment horizontal="right" vertical="center"/>; SetReadingOrder("rtl") ->
  // <alignment readingOrder="2"/> (OOXML's own 0=context/1=ltr/2=rtl encoding -- the string "rtl"
  // maps to the numeric 2, not passed through literally).
  rangeAlign: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeAlign: `at` is required')
      const SUPPORTED = ['vertical', 'horizontal', 'readingOrder']
      if (!SUPPORTED.some((k) => op[k] !== undefined)) {
        throw new Error(`rangeAlign: at least one of ${SUPPORTED.join(', ')} is required`)
      }
      const range = `oWorksheet.GetRange(${jsString(op.at)})`
      const lines = []
      if (op.vertical !== undefined) lines.push(`${range}.SetAlignVertical(${jsString(op.vertical)});`)
      if (op.horizontal !== undefined) lines.push(`${range}.SetAlignHorizontal(${jsString(op.horizontal)});`)
      if (op.readingOrder !== undefined) lines.push(`${range}.SetReadingOrder(${jsString(op.readingOrder)});`)
      return lines
    },
  },

  // MEASURED -- `ApiRange.SetItalic/SetStrikeout/SetWrap(bool)`
  // all WORK (package-verified: italic+strikeout add a new xl/styles.xml <font> entry with
  // <i/><strike/>; wrap sets `wrapText="1"` in the cellXfs <alignment>). `ApiRange.SetUnderline
  // (bool)` is CONFIRMED INERT, tested alone (not just alongside the other two, to rule out one
  // masking another): the saved package's <fonts> count stays at 1, no <u/> element anywhere.
  // Refused client-side, same reasoning as `rangeSize`'s `hidden` field above.
  rangeFontStyle: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeFontStyle: `at` is required')
      if (op.underline !== undefined) {
        throw notSupportedError('rangeFontStyle: `underline` not supported -- measured: ApiRange.SetUnderline() is inert on this DocBuilder instance (tested alone: the saved package\'s font count does not change, no <u/> element anywhere)')
      }
      const SUPPORTED = ['italic', 'strikeout', 'wrap']
      if (!SUPPORTED.some((k) => op[k] !== undefined)) {
        throw new Error(`rangeFontStyle: at least one of ${SUPPORTED.join(', ')} is required`)
      }
      const range = `oWorksheet.GetRange(${jsString(op.at)})`
      const lines = []
      if (op.italic !== undefined) lines.push(`${range}.SetItalic(${Boolean(op.italic)});`)
      if (op.strikeout !== undefined) lines.push(`${range}.SetStrikeout(${Boolean(op.strikeout)});`)
      if (op.wrap !== undefined) lines.push(`${range}.SetWrap(${Boolean(op.wrap)});`)
      return lines
    },
  },

  // MEASURED, working -- `ApiRange.UnMerge()`. Package-verified:
  // a prior Merge(false) produces <mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells> in
  // xl/worksheets/sheet1.xml; calling UnMerge() on the same range afterward removes the element
  // entirely (not merely empties it -- the whole <mergeCells> tag is gone, matching a fresh
  // never-merged sheet's own output byte-for-byte on this field).
  unmergeRange: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('unmergeRange: `at` is required')
      return [`oWorksheet.GetRange(${jsString(op.at)}).UnMerge();`]
    },
  },

  // MEASURED, working -- `ApiWorksheet.SetName(name)` and
  // `ApiWorksheet.SetVisible(bool)` on the CURRENT sheet (`oWorksheet`, `Api.GetActiveSheet()`
  // from the preamble), or on ANY other sheet via `op.target` (`Api.GetSheet(nameOrIndex)`).
  // Both call forms of GetSheet measured live: a string name AND a 0-based numeric index each
  // return the expected sheet (verified via each one's own SetName landing on the right
  // xl/workbook.xml <sheet> element, byte-checked against the untouched siblings as NEG control).
  // `SetActive()` measured live too: on a non-active sheet fetched by name, the saved
  // xl/workbook.xml bookViews/workbookView `activeTab` attribute moves to that sheet's own
  // 0-based index -- confirmed against the pre-call default (whichever sheet the seed already
  // had active).
  // `Delete()` measured live on a NON-active sheet (fetched by name): the job completes normally
  // (does NOT kill it), the sheet's <sheet> element is gone from xl/workbook.xml, the remaining
  // sheets keep their own names/order, only sheetId renumbers -- and activeTab is untouched
  // (still points at the sheet that was already active, unaffected by a sibling's removal).
  // `Move()` WORKS -- the refusal that previously stood here was measured against the WRONG
  // call shape: `oWorksheet.Move(n)`, a single
  // number. Reflected off the live instance (`String(oWorksheet.Move)`), the real signature is
  // `Move(before, after)` -- TWO arguments, and its own guard requires EXACTLY ONE of them to be
  // an `ApiWorksheet` instance (`if (bb&&ba || !bb&&!ba) throwException(...)`) -- a bare number
  // satisfies neither, so the function's OWN guard throws "Incorrect parametrs.", which reads
  // exactly like a job-kill from the outside (uncaught exception, no saved output). With the
  // correct shape it is package-verified: `Move(undefined, oTarget)` moved a sheet after
  // another, `xl/workbook.xml`'s `<sheet>` order changed to match, sheet names/content untouched.
  sheet: {
    cores: ['xlsx'],
    emit(op) {
      const SUPPORTED = ['name', 'visible', 'active', 'delete', 'move']
      const asked = SUPPORTED.filter((k) => op[k] !== undefined)
      if (!asked.length) {
        throw new Error(`sheet: at least one of ${SUPPORTED.join(', ')} is required`)
      }
      if (op.delete && asked.length > 1) {
        throw new Error('sheet: `delete` cannot be combined with `name`/`visible`/`active`/`move` -- the sheet would no longer exist for them to apply to')
      }
      const lines = []
      let ref = 'oWorksheet'
      if (op.target !== undefined) {
        const targetArg = typeof op.target === 'number' ? Number(op.target) : jsString(op.target)
        lines.push(`var oTargetSheet = Api.GetSheet(${targetArg});`)
        ref = 'oTargetSheet'
      }
      if (op.name !== undefined) lines.push(`${ref}.SetName(${jsString(op.name)});`)
      if (op.visible !== undefined) lines.push(`${ref}.SetVisible(${Boolean(op.visible)});`)
      if (op.active) lines.push(`${ref}.SetActive();`)
      if (op.move !== undefined) {
        if (typeof op.move !== 'object' || op.move === null) {
          throw new Error('sheet: `move` must be an object with exactly one of `before`/`after` (a sheet name or 0-based index) -- e.g. { after: "Munka1" }')
        }
        const hasBefore = op.move.before !== undefined
        const hasAfter = op.move.after !== undefined
        if (hasBefore === hasAfter) {
          throw new Error('sheet: `move` needs EXACTLY ONE of `before`/`after` -- measured: ApiWorksheet.Move(before, after) requires exactly one argument to be a worksheet instance, the other stays undefined')
        }
        const moveArg = hasBefore ? op.move.before : op.move.after
        const moveRef = typeof moveArg === 'number' ? Number(moveArg) : jsString(moveArg)
        lines.push(`var oMoveTarget = Api.GetSheet(${moveRef});`)
        lines.push(`${ref}.Move(${hasBefore ? 'oMoveTarget, undefined' : 'undefined, oMoveTarget'});`)
      }
      if (op.delete) lines.push(`${ref}.Delete();`)
      return lines
    },
  },

  // MEASURED, working -- `ApiWorksheet.SetRowHeight(rowIndex,
  // height)`, `ApiWorksheet.SetDisplayGridlines(bool)`, `ApiWorksheet.SetDisplayHeadings(bool)`.
  // Live DocBuilder run, package-verified in xl/worksheets/sheet1.xml:
  //   SetDisplayGridlines(false) -> <sheetView showGridLines="0" .../> (absent = shown, the default)
  //   SetDisplayHeadings(false)  -> <sheetView showRowColHeaders="0" .../> (same absent-is-default shape)
  //   SetRowHeight(3, 40)        -> the row at index 3 gets <row r="4" ht="40" customHeight="1"/>
  // MEASURED TRAP: SetRowHeight's row index is 0-BASED -- calling it with 3 set the FOURTH row
  // (r="4"), not the third. Every other row/column reference in this file is 1-based (matching
  // cell-reference convention, e.g. columnWidth's "B1" anchor), so this op accepts a 1-based
  // `row` from the caller and does the -1 conversion here, keeping the API's 0-based quirk out
  // of this tool's own interface. NEG control: a run with none of these three calls leaves
  // sheetView/row1 with no ht/customHeight/showGridLines/showRowColHeaders attributes at all.
  sheetDisplay: {
    cores: ['xlsx'],
    emit(op) {
      const lines = []
      const hasRow = op.row !== undefined || op.rowHeight !== undefined
      if (hasRow) {
        if (op.row === undefined) throw new Error('sheetDisplay: `rowHeight` requires `row`')
        if (op.rowHeight === undefined) throw new Error('sheetDisplay: `row` requires `rowHeight`')
        const row = Number(op.row)
        if (!Number.isInteger(row) || row < 1) {
          throw new Error(`sheetDisplay: \`row\` must be a positive integer (1-based), got ${JSON.stringify(op.row)}`)
        }
        lines.push(`oWorksheet.SetRowHeight(${row - 1}, ${Number(op.rowHeight)});`)
      }
      if (op.gridlines !== undefined) lines.push(`oWorksheet.SetDisplayGridlines(${Boolean(op.gridlines)});`)
      if (op.headings !== undefined) lines.push(`oWorksheet.SetDisplayHeadings(${Boolean(op.headings)});`)
      if (!lines.length) {
        throw new Error('sheetDisplay: at least one of `row`+`rowHeight`, `gridlines`, `headings` is required')
      }
      return lines
    },
  },

  // MEASURED, working -- `ApiRange.AddComment(text, author)`.
  // Package-verified: produces a new `xl/comments1.xml` part with
  // <comments><authors><author>...</author></authors><commentList><comment ref="A1" .../>...
  rangeComment: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeComment: `at` is required')
      if (!op.text) throw new Error('rangeComment: `text` is required')
      return [`oWorksheet.GetRange(${jsString(op.at)}).AddComment(${jsString(op.text)}, ${jsString(op.author ?? '')});`]
    },
  },

  // MEASURED, both working -- `ApiRange.Select()` (no args) and
  // `ApiRange.SetOrientation(degrees)` (a NUMBER, degrees of text rotation -- not a keyword like
  // xlsx's own `sort`/`conditionalFormatting` string enums). Package-verified: Select() on B2 ->
  // xl/worksheets/sheet1.xml <selection activeCell="B2" .../> (the seed's own default is A1);
  // SetOrientation(45) -> xl/styles.xml cellXfs <alignment textRotation="45"/>.
  rangeSelect: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeSelect: `at` is required')
      const SUPPORTED = ['select', 'rotation']
      if (!SUPPORTED.some((k) => op[k] !== undefined)) {
        throw new Error(`rangeSelect: at least one of ${SUPPORTED.join(', ')} is required`)
      }
      const range = `oWorksheet.GetRange(${jsString(op.at)})`
      const lines = []
      if (op.select) lines.push(`${range}.Select();`)
      if (op.rotation !== undefined) lines.push(`${range}.SetOrientation(${Number(op.rotation)});`)
      return lines
    },
  },

  // MEASURED, all three working -- `ApiRange.Delete(direction)`/
  // `Insert(direction)` (string direction, e.g. "up"/"down" -- both confirmed to shift the sheet's
  // remaining rows correctly: Delete("up") on A2 pulled A3's value up into row 2; Insert("down")
  // on A2 pushed the old A2 down into row 3, leaving a blank styled A2) and `AutoFit(rows, cols)`
  // (two booleans; AutoFit(false, true) on A1 produced xl/worksheets/sheet1.xml's own <cols><col
  // bestFit="1" min="1" max="1" width="34.28125"/></cols> -- a real computed best-fit width, not
  // a placeholder).
  rangeEdit: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeEdit: `at` is required')
      const SUPPORTED = ['delete', 'insert', 'autoFitRows', 'autoFitColumns']
      if (!SUPPORTED.some((k) => op[k] !== undefined)) {
        throw new Error(`rangeEdit: at least one of ${SUPPORTED.join(', ')} is required`)
      }
      const range = `oWorksheet.GetRange(${jsString(op.at)})`
      const lines = []
      if (op.delete !== undefined) lines.push(`${range}.Delete(${jsString(op.delete)});`)
      if (op.insert !== undefined) lines.push(`${range}.Insert(${jsString(op.insert)});`)
      if (op.autoFitRows !== undefined || op.autoFitColumns !== undefined) {
        lines.push(`${range}.AutoFit(${Boolean(op.autoFitRows)}, ${Boolean(op.autoFitColumns)});`)
      }
      return lines
    },
  },

  // MEASURED, working -- `ApiWorksheet.SetHyperlink(range, address,
  // subAddress, tooltip)`. Live DocBuilder run, package-verified: `SetHyperlink("B1", url, "",
  // "some text")` produces BOTH a `xl/worksheets/_rels/sheet1.xml.rels` Relationship
  // (Type=".../hyperlink", Target=url, TargetMode="External") AND the sheet's own
  // `<hyperlinks><hyperlink r:id="rId1" ref="B1" tooltip="some text"/></hyperlinks>`.
  // MEASURED TRAP: the call takes a STRING range reference ("B1"), NOT a Range object --
  // `SetHyperlink(oWorksheet.GetRange("B1"), url)` measured to kill the whole job (outcome
  // "blocked"), unlike almost every other xlsx op in this file which calls GetRange() first.
  // MEASURED TRAP #2: the 3rd positional argument has NO observed effect on an external URL (a
  // distinctive marker string in that position appears nowhere in the saved package -- sheet XML
  // or rels); only the 4th argument lands, as the `tooltip` attribute. AND it only lands when
  // FOUR arguments are passed -- calling with three (range, address, tooltip) drops the tooltip
  // silently (measured: no `tooltip` attribute at all). This op therefore always passes an empty
  // 3rd argument when a tooltip is requested, and omits both trailing arguments otherwise (a
  // 2-argument call is confirmed to produce a plain hyperlink, no tooltip attribute).
  hyperlink: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.range) throw new Error('hyperlink: `range` is required')
      if (!op.address) throw new Error('hyperlink: `address` is required')
      if (op.tooltip !== undefined) {
        return [`oWorksheet.SetHyperlink(${jsString(op.range)}, ${jsString(op.address)}, "", ${jsString(op.tooltip)});`]
      }
      return [`oWorksheet.SetHyperlink(${jsString(op.range)}, ${jsString(op.address)});`]
    },
  },

  // The ARGUMENT ORDER here was previously BACKWARDS. Reflected off the live instance
  // (`String(oWorksheet.AddProtectedRange)`), the
  // real signature is `AddProtectedRange(sTitle, sDataRange)` -- title FIRST, the range as a
  // STRING second. This entry used to call it `(range, title)`, and the OLD comment's own
  // "evidence" confirms the swap rather than refuting it: it describes the saved
  // `<userProtectedRange name="A1:B2">` as proof the call worked -- but `name` should hold the
  // human TITLE, and "A1:B2" is the range, landing in the title slot precisely because the
  // arguments were swapped. Package-verified with the CORRECT order (title, range):
  // `AddProtectedRange("Cim", "A1:B2")` -> `<userProtectedRange name="Cim" sqref="A1:B2">`.
  // `title` stays required -- that part of the old finding holds, just not the reason: the
  // function's OWN guard (`isValidTitle && isValidRef`, both must be non-empty strings) throws
  // "The title or dataRange is invalid" on a missing/empty one, an uncaught exception that reads
  // like a job-kill from the outside, not a distinct "kills the job" failure mode.
  protectedRange: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.range) throw new Error('protectedRange: `range` is required')
      if (!op.title) {
        throw new Error('protectedRange: `title` is required -- measured: AddProtectedRange(sTitle, sDataRange) throws "The title or dataRange is invalid" when either argument is missing/empty (its own validation, not a distinct job-kill)')
      }
      return [`oWorksheet.AddProtectedRange(${jsString(op.title)}, ${jsString(op.range)});`]
    },
  },

  // MEASURED -- `ApiRange.Copy(destRange)` and
  // `ApiRange.Cut(destRange)` both WORK, and both take the destination as a Range OBJECT argument
  // (unlike `hyperlink`/`protectedRange` in koteg02, which take a STRING). Copy leaves the source
  // in place and duplicates the value into the destination; Cut moves it (the source cell is gone
  // from the saved sheetData entirely, not merely cleared to empty).
  //
  // `ApiRange.Paste(rangeFrom)` WORKS -- it was previously refused here on the theory that it
  // silently no-ops; that measurement
  // called it with ZERO arguments (`oRange.Paste()`), which is not its real shape. Reflected off
  // the live instance, the signature is ONE argument, a source Range object, called ON THE
  // DESTINATION: `oDest.Paste(oSource)` -- the reverse orientation of `Copy`/`Cut` above (those
  // are called on the source with the destination as the argument). Package-verified: the
  // destination cell genuinely receives the source's value (same shared-string index as the
  // source in the saved package). `paste` below is therefore a STRING -- the SOURCE range to
  // paste FROM, with `at` staying the destination, same "at is where this op targets" contract
  // every other rangeClipboard field already uses.
  //
  // `ApiRange.PasteSpecial(...)` REMAINS refused, and this part of the old finding holds:
  // measured with a real Copy() staged first (not the zero-argument shape above), it still
  // completes the job but writes nothing to the destination -- a silent no-op, same class as
  // `SetPageOrientation`/`chartType` elsewhere in this file. It goes through `Asc.editor`'s own
  // clipboard/paste helper (`AscCommon.g_specialPasteHelper` + `oApi.asc_Paste()`), which has
  // nothing to act on in this headless DocBuilder context -- unlike `Paste`, which operates
  // directly on the two Range objects with no editor/clipboard dependency.
  rangeClipboard: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeClipboard: `at` is required')
      if (op.pasteSpecial !== undefined) {
        throw notSupportedError('rangeClipboard: `pasteSpecial` not supported -- measured: ApiRange.PasteSpecial() is a silent no-op on this DocBuilder instance (the job completes, but the destination cell is not written; re-measured 2026-08-17 with a real staged Copy()). Use `paste` (ApiRange.Paste, no editor/clipboard dependency) or `copyTo`/`cutTo` instead')
      }
      const SUPPORTED = ['copyTo', 'cutTo', 'paste']
      if (!SUPPORTED.some((k) => op[k] !== undefined)) {
        throw new Error(`rangeClipboard: at least one of ${SUPPORTED.join(', ')} is required`)
      }
      const range = `oWorksheet.GetRange(${jsString(op.at)})`
      const lines = []
      if (op.copyTo !== undefined) lines.push(`${range}.Copy(oWorksheet.GetRange(${jsString(op.copyTo)}));`)
      if (op.cutTo !== undefined) lines.push(`${range}.Cut(oWorksheet.GetRange(${jsString(op.cutTo)}));`)
      if (op.paste !== undefined) {
        if (typeof op.paste !== 'string' || !op.paste) {
          throw new Error('rangeClipboard: `paste` must be a non-empty cell/range reference string -- the SOURCE to paste FROM (`at` is the destination, the reverse of `copyTo`/`cutTo`)')
        }
        lines.push(`${range}.Paste(oWorksheet.GetRange(${jsString(op.paste)}));`)
      }
      return lines
    },
  },

  // CONFIRMED INERT, both -- `ApiRange.Find(what)` and
  // `ApiRange.Replace(what, replacement)`. Find measured to return `null` even on a SINGLE cell
  // holding the EXACT search text (no ambiguity, no range confusion) -- tried with just the
  // search string, and with a full 9-argument call (after/lookIn/lookAt/searchOrder/
  // searchDirection/matchCase/matchByte/searchFormat). Replace measured to return `undefined`
  // and leave the cell's own text completely unchanged (the saved package's sharedStrings.xml
  // keeps the original word, tried both with and without trailing boolean arguments). Refused
  // client-side -- always, not input-dependent, same shape as `freezePanes` above.
  rangeSearch: {
    cores: ['xlsx'],
    emit() {
      throw notSupportedError('rangeSearch: not supported -- measured: ApiRange.Find() always returns null (even on a single cell holding the exact search text) and ApiRange.Replace() always returns undefined and leaves the cell text unchanged, in every call shape tried, on this DocBuilder instance')
    },
  },

  // MEASURED, working -- `ApiRange.SetFormulaArray(formula)`.
  // Package-verified: SetFormulaArray("=B1:B3*2") on C1:C3 -> the top-left cell (C1) carries
  // <f t="array" ref="C1:C3">B1:B3*2</f>, the other cells in the range get an empty <f/> sharing
  // the same array -- the standard OOXML array-formula shape (one spilled definition, not one
  // formula per cell).
  rangeFormulaArray: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeFormulaArray: `at` is required')
      if (!op.formula) throw new Error('rangeFormulaArray: `formula` is required')
      return [`oWorksheet.GetRange(${jsString(op.at)}).SetFormulaArray(${jsString(op.formula)});`]
    },
  },

  // MEASURED, both working -- `ApiRange.Offset(rows, cols)` and
  // `ApiRange.Resize(rows, cols)` both RETURN A RANGE OBJECT rather than mutating anything
  // themselves (worth checking explicitly): the call itself writes nothing to the
  // package, so this op WRITES A VALUE through the returned range and the measurement is what
  // landed at the OFFSET/RESIZED cell, not whether the call itself succeeded. Package-verified:
  // GetRange("A1").Offset(2, 1).SetValue(...) landed the value at B3 (row+2, col+1 from A1, 0-
  // indexed like ApiWorksheet.SetRowHeight -- NOT the 1-based convention this file uses at its
  // own interface boundary elsewhere); GetRange("A1").Resize(2, 2).SetValue(...) landed the SAME
  // value at all four cells of A1:B2 (Resize grows the range in place, it does not move it).
  // `value` is therefore a required field, not incidental -- an op with no value would prove
  // nothing about whether the offset/resize itself worked.
  rangeOffsetWrite: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('rangeOffsetWrite: `at` is required')
      if (op.value === undefined) {
        throw new Error('rangeOffsetWrite: `value` is required -- Offset/Resize return a Range but write nothing themselves, so this op needs something to write to prove where it landed')
      }
      let range = `oWorksheet.GetRange(${jsString(op.at)})`
      if (op.offsetRows !== undefined || op.offsetCols !== undefined) {
        range = `${range}.Offset(${Number(op.offsetRows ?? 0)}, ${Number(op.offsetCols ?? 0)})`
      }
      if (op.resizeRows !== undefined || op.resizeCols !== undefined) {
        range = `${range}.Resize(${Number(op.resizeRows ?? 1)}, ${Number(op.resizeCols ?? 1)})`
      }
      return [`${range}.SetValue(${jsString(op.value)});`]
    },
  },

  // MEASURED, working -- `Api.AddDefName(name, ref, isHidden)`
  // creates a real workbook-scope defined name, confirmed in the saved package's own
  // xl/workbook.xml (<definedNames><definedName name="..." hidden="0|1">ref</definedName>...).
  // A previous positive control had assumed this was already proven by the
  // reference generator (an internal demo-xlsx script) -- it was not: that script writes
  // raw OOXML with Python's zipfile, never calling the DocBuilder API at all, so it only shows
  // the FORMAT accepts a defined name, not that this API does. Re-measured from zero, live.
  //
  // BAD-INPUT BEHAVIOUR, MEASURED, all three tried: a duplicate `name`, an invalid `name`
  // (leading digit, embedded space), and a `ref` naming a sheet that does not exist -- ALL THREE
  // hard-fail the WHOLE builder script (server outcome "blocked", no output file at all), never a
  // silent no-op or a silently-mangled name/ref. This is the opposite failure shape from
  // `columnWidth`'s negative-width or `border`'s "All"/"Outline" above (both silently do
  // something other than what was asked) -- here a bad call is refused loudly by the server
  // itself, so this operation does not need to pre-validate name syntax or range shape itself.
  definedName: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.name) throw new Error('definedName: `name` is required')
      if (!op.ref) throw new Error('definedName: `ref` (a sheet-qualified range reference, e.g. "Munka1!$A$1:$B$2") is required')
      const hidden = Boolean(op.hidden)
      return [`Api.AddDefName(${jsString(op.name)}, ${jsString(op.ref)}, ${hidden});`]
    },
  },

  // lap-szintu rajz-objektumok, ApiWorksheet-en.
  //   shape: `ApiWorksheet.AddShape(sType, x, y, cx, cy)` -- MEASURED, working, CSOMAG-SZINTEN
  //     igazolva a buildCreateScript() SAJAT kimenetevel futtatva (nem csak izolalt probaval):
  //     a mentett .xlsx-ben `xl/drawings/drawing1.xml` + `xl/worksheets/_rels/sheet1.xml.rels`
  //     valodi tartalommal jon letre.
  //   wordArt: MEASURED, CONFIRMED INERT -- `ApiWorksheet.AddWordArt(...)` MINDIG
  //     "Asc.editor.CreateNoFill is not a function" hibat dob, FUGGETLENUL az argumentumok
  //     szamatol/sorrendjetol (harom alak probalva: (text,style), (style,text), csak text) --
  //     tehat nem egy konkret hivas-alak hibaja, hanem a metodus maga toresponkius ezen a
  //     DocBuilder peldanyon. Refused, mint a `freezePanes` regi allapota volt.
  //   ole: MEASURED, CONFIRMED NEM VALODI OLE-OBJEKTUM -- `ApiWorksheet.AddOleObject(sImageSrc,
  //     sName, sProgId)` NEM DOB HIBAT es a job lefut, DE a csomag-szintu ellenorzes
  //     (buildCreateScript() sajat kimenetevel, nem csak izolalt hivassal) megmutatta, hogy a
  //     mentett fajlban CSAK egy sima `<xdr:pic>` kepelem jon letre `xl/drawings/`-ben -- SEM
  //     `xl/embeddings/*` resz, SEM valodi OLE-struktura NEM keletkezik, a `name` es `progId`
  //     ERTEKEK CSENDBEN ELVESZNEK, es a horgony-koordinatak (`colOff`/`rowOff`) garbage
  //     ertekre allnak (-9223372036854775808, INT64_MIN szentinel). Egy 7-argumentumos
  //     probalkozas (pozicio/meret parameterekkel, az AddImage mintajara) A TELJES DOCBUILDER
  //     JOBOT MEGOLTE (a szerver valasz nelkul maradt) -- ugyanaz a sulyossagi osztaly, mint az
  //     ApiRange.Paste(). Sem a 3-, sem a 7-argumentumos alak nem ad hasznalhato eredmenyt --
  //     refused, a `paste`/`pasteSpecial` mintajara.
  //   replaceImage: MEASURED, CONFIRMED CSENDES NO-OP -- `ApiWorksheet.ReplaceCurrentImage(src)`
  //     NEM DOB HIBAT, es az IZOLALT probaban a `GetAllImages().length` valtozatlan maradt (ami
  //     eleinte "mukodik"-nek nezett ki), DE a csomag-szintu ellenorzes (buildCreateScript()
  //     sajat kimenetevel, before/after byte-osszehasonlitassal a mentett `xl/media/image1.png`
  //     tartalman) megmutatta, hogy a kep BAJTRA AZONOS marad az EREDETIVEL -- a csere TENYLEGESEN
  //     NEM TORTENIK MEG. Ugyanaz a csapda-osztaly, mint amit ez a fajl mar dokumental:
  //     "no error thrown" nem bizonyitek, csak a tenyleges ertek before/after osszevetese az.
  //     Refused.
  sheetDrawing: {
    cores: ['xlsx'],
    emit(op) {
      if (op.wordArt !== undefined) {
        throw notSupportedError('sheetDrawing: `wordArt` not supported -- measured: ApiWorksheet.AddWordArt() always throws "Asc.editor.CreateNoFill is not a function" on this DocBuilder instance, in every call shape tried ((text,style), (style,text), text-only) -- not argument-shape dependent, the method itself is broken here')
      }
      if (op.ole !== undefined) {
        throw notSupportedError('sheetDrawing: `ole` not supported -- measured: ApiWorksheet.AddOleObject() does not create a real OLE object in the saved package in the working (3-arg) call shape (only a plain picture element appears, `name`/`progId` are silently discarded, anchor coordinates come out as garbage) -- and a fuller (7-arg) call shape crashes the whole job outright, same severity class as ApiRange.Paste()')
      }
      if (op.replaceImage !== undefined) {
        throw notSupportedError('sheetDrawing: `replaceImage` not supported -- measured: ApiWorksheet.ReplaceCurrentImage() is a silent no-op in the saved package (the media part stays byte-identical to the original image) even though no error is thrown and even though ApiWorksheet.GetAllImages().length looks unchanged in-script -- package-level verification (not just an in-script check) is required to see this')
      }
      const SUPPORTED = ['shape']
      if (!SUPPORTED.some((k) => op[k] !== undefined)) {
        throw new Error(`sheetDrawing: at least one of ${SUPPORTED.join(', ')} is required`)
      }
      const s = op.shape
      if (!s.type) throw new Error('sheetDrawing.shape: `type` is required (e.g. "rect")')
      const x = Number(s.x ?? 1)
      const y = Number(s.y ?? 1)
      const w = Number(s.width ?? 2000000)
      const h = Number(s.height ?? 1000000)
      return [`oWorksheet.AddShape(${jsString(s.type)}, ${x}, ${y}, ${w}, ${h});`]
    },
  },

  // lap-szintu vagolap + FormatAsTable, ApiWorksheet-en.
  //   `ApiWorksheet.FormatAsTable(sRange, sStyleName)` -- MEASURED, working, DE A SORREND SZAMIT:
  //     (style, range) alakban "Cannot read property 'c1' of null"-t dob (catchable), (range, style)
  //     alakban hibatlanul lefut. A hiba-uzenet a fordított sorrendet sugallja, nem a metodus
  //     hianyat -- ezert catchable Error, nem notSupportedError.
  //   `ApiWorksheet.Paste(...)` -- MEASURED, CONFIRMED INERT (csendes no-op, NEM omlik ossze,
  //     ELLENTETBEN az ApiRange.Paste()-vel, ami a teljes jobot megoli): harom alak probalva
  //     (string cel, Range-objektum cel, cel nelkul egyetlen forras-argumentummal) -- egyik sem
  //     valtoztatta meg a cel cella erteket. Refused itt, kulon nevesitve, hogy ne keveredjen
  //     az ApiRange szintu, MAR landolt `rangeClipboard` copyTo/cutTo megoldassal.
  sheetFormatTable: {
    cores: ['xlsx'],
    emit(op) {
      if (op.paste !== undefined) {
        throw notSupportedError('sheetFormatTable: `paste` not supported -- measured: ApiWorksheet.Paste() is a silent no-op on this DocBuilder instance in every call shape tried (string dest, Range dest, no dest) -- use the already-landed `rangeClipboard` (copyTo/cutTo) instead')
      }
      if (!op.range) throw new Error('sheetFormatTable: `range` is required')
      if (!op.style) throw new Error('sheetFormatTable: `style` is required (e.g. "TableStyleMedium2")')
      return [`oWorksheet.FormatAsTable(${jsString(op.range)}, ${jsString(op.style)});`]
    },
  },

  // pivot-tabla, Api szinten (NEM ApiWorksheet-en).
  //   `Api.InsertPivotNewWorksheet(oSourceRange)` -- MEASURED, working, uj munkalapot hoz letre
  //     a pivotnak (Api.GetSheets().length 1-rol 2-re nott a hivas utan).
  //   `Api.InsertPivotExistingWorksheet(oSourceRange, oDestRange)` -- MEASURED, working, DE KET
  //     Range-objektum kell, NEM egy Range + egy Worksheet: egy (range, worksheet, "cel-cim")
  //     3-argumentumos probalkozas "pivotRef.GetWorksheet is not a function"-t dobott (catchable) --
  //     a helyes alak (range, destRange) 2 argumentum, mindketto GetRange()-bol szarmazo objektum.
  //   `Api.RefreshAllPivots()` es `Api.GetAllPivotTables()` -- MEASURED, mindketto mukodik,
  //     package-verified a pivots.length ertekkel.
  pivotTable: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.range) throw new Error('pivotTable: `range` (the source data range, e.g. "A1:B4") is required')
      const source = `oWorksheet.GetRange(${jsString(op.range)})`
      const lines = []
      if (op.at !== undefined) {
        lines.push(`Api.InsertPivotExistingWorksheet(${source}, oWorksheet.GetRange(${jsString(op.at)}));`)
      } else {
        lines.push(`Api.InsertPivotNewWorksheet(${source});`)
      }
      if (op.refresh) lines.push('Api.RefreshAllPivots();')
      return lines
    },
  },

  // MEASURED via full method enumeration on Api AND
  // ApiWorksheet (not guessed), then a write-known-value-then-query round trip, since a
  // DocBuilder script has no return channel to the caller other than the saved package -- the
  // query result is written into the caller-chosen `at` cell as text, same "observe through the
  // saved file" discipline as every other entry in this table.
  //   `Api.GetComments()` -- EXISTS (does not throw) but is a CONFIRMED WRONG-SCOPE method on
  //     this DocBuilder instance: it returns an empty array EVEN AFTER a comment was just added
  //     via ApiRange.AddComment() in the same script (measured against a matching zero-comment
  //     NEG. KONTROLL -- both give count=0, so the presence of a real comment changes nothing).
  //   `ApiWorksheet.GetComments()` and `Api.GetAllComments()` -- BOTH measured working: after
  //     AddComment("text","author"), GetComments() returns count=1 with GetText()/GetAuthorName()
  //     round-tripping the exact values written; with two comments added in sequence, the
  //     returned array preserves insertion order (index 0 = first added). Sheet-scoped
  //     `oWorksheet.GetComments()` used here to match every other entry in this table (all bound
  //     to oWorksheet, not Api).
  sheetComments: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.at) throw new Error('sheetComments: `at` (the cell to write the query result into) is required -- this is a query operation, its answer travels through the saved package, not a direct return value')
      return [
        '(function () {',
        '  var __c = oWorksheet.GetComments();',
        '  var __parts = [];',
        '  for (var __i = 0; __i < __c.length; __i++) { __parts.push(__c[__i].GetAuthorName() + ":" + __c[__i].GetText()); }',
        `  oWorksheet.GetRange(${jsString(op.at)}).SetValue("count=" + __c.length + (__parts.length ? ";" + __parts.join("|") : ""));`,
        '})();',
      ]
    },
  },

  // `Api.GetThemesColors()` does NOT return colors despite the
  // name -- MEASURED: it returns the array of 24 built-in THEME NAMES ("Aspect", "Blue Green",
  // ..., "Yellow") available on this DocBuilder instance, unchanged before/after a SetThemeColors
  // call. `Api.SetThemeColors(colors)` with an array of Api.CreateColorFromRGB(...) objects (the
  // shape every other color-setting entry in this table uses, e.g. `border`/`fillColor`) is a
  // CONFIRMED SILENT NO-OP -- does not throw, but the saved xl/theme/theme1.xml stayed
  // byte-identical to the untouched default. `Api.SetThemeColors("Aspect")` (a NAME STRING, one
  // of the values GetThemesColors() itself returns) is the measured working call shape: the
  // saved theme1.xml's active colour scheme (the first 12 <a:srgbClr>) changed to Aspect's
  // documented palette (323232/E3DED1/F07F09/... vs. the untouched default's
  // 1F497D/EEECE1/4F81BD/...), with the previous scheme preserved alongside it in the same part
  // (OOXML's extraClrSchemeLst) -- a real, package-verified write, not a claim from the call's
  // own return value. An UNRECOGNIZED name is ALSO a confirmed silent no-op (measured:
  // SetThemeColors("NemLetezoTemaNev123") does not throw, saved theme1.xml stays at the default)
  // -- refused client-side here, same discipline as `border`'s "All"/"Outline" no-op names,
  // rather than repeating that silent-no-op mistake for a caller.
  sheetTheme: {
    cores: ['xlsx'],
    emit(op) {
      if (!op.name) throw new Error(`sheetTheme: \`name\` is required -- one of the built-in theme names: ${KNOWN_THEME_NAMES.join(', ')}`)
      if (!KNOWN_THEME_NAMES.includes(op.name)) {
        throw notSupportedError(`sheetTheme: unknown theme name ${JSON.stringify(op.name)} (known: ${KNOWN_THEME_NAMES.join(', ')}) -- measured: an unrecognized name does not throw inside the builder, it silently changes nothing (the saved theme1.xml stays byte-identical to the pre-call default)`)
      }
      return [`Api.SetThemeColors(${jsString(op.name)});`]
    },
  },

  // `Api.RecalculateAllFormulas()` EXISTS and does not throw,
  // but is a CONFIRMED REDUNDANT no-op on this DocBuilder instance in every scenario tried: a
  // formula's cached <v> already updates the moment its precedent cell is written via SetValue,
  // with or without this call, and that holds BOTH within a single script AND across a
  // reopen-the-saved-file / edit-the-precedent / resave cycle (measured: A1=5, B1="=A1*2" (cached
  // <v>=10), saved and reopened as a fresh job, A1 rewritten to 99 -- B1's cached <v> is already
  // 198 with NO recalculate call, and identically 198 with one). Refused here so a caller does
  // not believe this call is doing work it is not; this is a "confirmed redundant" finding, not
  // "confirmed broken" -- if a case is later found where a cached value genuinely goes stale
  // (e.g. through the co-editing path, which this batch did not touch), that is new evidence, not
  // something this refusal rules out.
  recalculateFormulas: {
    cores: ['xlsx'],
    emit() {
      throw notSupportedError('recalculateFormulas: not supported -- measured: Api.RecalculateAllFormulas() exists and does not throw, but is a confirmed redundant no-op here -- a formula\'s cached value already updates automatically the instant its precedent cell is written, both in-script and across a reopen-edit-resave cycle, identically with or without this call')
    },
  },
}
