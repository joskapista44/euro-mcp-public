'use strict'
// Operations split boundary decision: this operation touches multiple cores (docx, pptx, xlsx), and the 4 largest
// OPERATIONS entries are exactly the multi-core ones -- a single shared bucket for all 8 would
// have recreated the collision surface this whole split exists to remove. Each multi-core key
// gets its own module instead; every per-core aggregator that needs it (per its own `cores`
// array) requires this file directly.
const { KNOWN_CHART_TYPES, applyChartFormatting, jsString, notSupportedError } = require('../lib-operations-registry.cjs')

  // Api.CreateChart's real signature (recovered via toString() on
  // this DocBuilder instance, receptek-xlsx.md #10): CreateChart(sType, aSeries, aSeriesNames,
  // aCatNames, nWidth, nHeight, ...) -- this is a PRESENTATION-INTERNAL API, so it always throws
  // on xlsx (setParent(private_GetCurrentSlide()) -> null deref). The xlsx entry point is the
  // Worksheet's OWN AddChart(sDataRange, bInRows, sType, ...), which pulls data from CELLS
  // instead of taking literal arrays -- a structurally different call shape, not just a
  // different embedding step, which is why xlsx has its own branch below rather than sharing
  // the docx/pptx code path.
  //
  // KNOWN_CHART_TYPES is an ALLOWLIST, not an exhaustive list of everything the engine accepts --
  // each entry was individually run and its package checked for an actual chart part. This
  // matters because an unrecognised type string does NOT throw here: it silently produces NO
  // chart at all (measured: "line" -> job completes, valid docx written, zero chart parts,
  // zero error). Without this allowlist a typo'd type would look exactly like success.
module.exports = {
  cores: ['docx', 'pptx', 'xlsx'],
  emit(op, core) {
    const chartType = String(op.chartType ?? '')
    if (!KNOWN_CHART_TYPES.includes(chartType)) {
      throw notSupportedError(`chart: unknown chartType ${JSON.stringify(chartType)} (known: ${KNOWN_CHART_TYPES.join(', ')}) -- an unrecognised type does NOT throw inside the builder, it silently produces no chart at all, so this tool refuses it here instead`)
    }
    const w = Number(op.width ?? 120 * 36000)
    const h = Number(op.height ?? 70 * 36000)

    if (core === 'xlsx') {
      if (!op.dataRange) throw new Error('chart: `dataRange` (e.g. "Munka1!$A$1:$B$4") is required for xlsx -- the xlsx chart reads cell data, it does not take literal series')
      const inRows = op.inRows ? 'true' : 'false'
      const style = Number(op.styleIndex ?? 2)
      const fromCol = Number(op.fromCol ?? 3)
      const colOffset = Number(op.colOffset ?? 0)
      const fromRow = Number(op.fromRow ?? 1)
      const rowOffset = Number(op.rowOffset ?? 0)
      return [`oWorksheet.AddChart(${jsString(op.dataRange)}, ${inRows}, ${jsString(chartType)}, ${style}, ${w}, ${h}, ${fromCol}, ${colOffset}, ${fromRow}, ${rowOffset});`]
    }

    // docx/pptx: the chart carries its OWN literal data -- there is no underlying spreadsheet.
    if (!Array.isArray(op.series) || !op.series.length) throw new Error('chart: `series` (array of numeric arrays, one per data series) is required')
    if (!op.series.every((s) => Array.isArray(s) && s.length)) throw new Error('chart: every entry in `series` must be a non-empty array of numbers')
    if (!Array.isArray(op.categories) || !op.categories.length) throw new Error('chart: `categories` (array of axis labels) is required')
    const seriesNames = Array.isArray(op.seriesNames) && op.seriesNames.length === op.series.length
      ? op.seriesNames
      : op.series.map((_s, i) => `Sorozat ${i + 1}`)
    const createCall = `Api.CreateChart(${jsString(chartType)}, ${JSON.stringify(op.series)}, ${JSON.stringify(seriesNames)}, ${JSON.stringify(op.categories)}, ${w}, ${h})`
    if (core === 'pptx') {
      const lines = [
        `var oChart = ${createCall};`,
        `oChart.SetPosition(${Number(op.x ?? 800000)}, ${Number(op.y ?? 800000)});`,
      ]
      applyChartFormatting(lines, op)
      lines.push('oSlide.AddObject(oChart);')
      return lines
    }
    return [
      `var oChart = ${createCall};`,
      'var oChartPara = Api.CreateParagraph();',
      'oChartPara.AddDrawing(oChart);',
      'oDocument.Push(oChartPara);',
    ]
  },
}
