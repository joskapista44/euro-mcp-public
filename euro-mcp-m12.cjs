'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m11.cjs')
const coedit = require('./coedit.cjs')
const rangeReader = require('./range-reader.cjs')

function textResult(payload) {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] }
}

server.tool(
  'office_read_range',
  'Reads a rectangular range from the CURRENT in-memory state of an ALREADY EXISTING xlsx ' +
    'workbook through the live ONLYOFFICE spreadsheet co-edit session. One request reads the whole ' +
    'range and returns a 2D cell matrix with address, row/column, raw/unformatted value, calculated ' +
    'value, displayed text, formula, inferred data type and number format where the live editor API ' +
    'supports them. It never downloads/parses OOXML and never falls back to DocBuilder. Unsupported ' +
    'fields are returned explicitly instead of guessed.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the existing xlsx workbook'),
    sheet: z.string().min(1).describe('Worksheet name, e.g. Sales'),
    range: z.string().regex(/^\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$/).describe('Rectangular A1 range, e.g. A1:H50'),
  },
  async ({ file_id, sheet, range }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })

    const r = await rangeReader.readRangeLive({
      url: cred.url,
      user: cred.user,
      pass: cred.pass,
      fileId: file_id,
      sheet,
      range,
      loadPlaywright: coedit.loadPlaywright,
    })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      identitasForras: hivo.forras || null,
      hitelesitoForras: cred.forras || null,
      ncUser: cred.user,
      source: r.source || null,
      editor: r.editor || null,
      apiHely: r.apiHely || null,
      sheet: r.sheet ?? sheet,
      range: r.range ?? range,
      rows: r.rows ?? null,
      columns: r.columns ?? null,
      cellCount: r.cellCount ?? null,
      cells: r.cells ?? null,
      unsupported: r.unsupported ?? null,
      maxCells: r.maxCells ?? null,
      error: r.error || r.indok || null,
    })
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
