'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m14.cjs')
const coedit = require('./coedit.cjs')
const inspector = require('./workbook-inspector.cjs')
const rangeReader = require('./range-reader.cjs')
const bulkWriter = require('./bulk-writer.cjs')
const verification = require('./verification-contract.cjs')

function textResult(payload) { return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] } }

const scalar = z.union([z.string(), z.number(), z.boolean(), z.null()])
const valueMatrix = z.array(z.array(scalar).min(1)).min(1)
const formulaMatrix = z.array(z.array(z.union([z.string(), z.null()])).min(1)).min(1)

server.tool(
  'office_write_range_verified',
  'Runs the M1 verification contract against the CURRENT live ONLYOFFICE co-edit state: inspect workbook, read target, edit target, then read back and verify. Every phase must prove source=live-coedit-editor and outcome=ok; any unsupported, ambiguous or mismatching phase fails closed. No DocBuilder, OOXML or downloaded-XLSX fallback is used.',
  {
    file_id: z.string(),
    sheet: z.string().min(1),
    range: z.string().min(1),
    values: valueMatrix,
    formulas: formulaMatrix.optional(),
  },
  async ({ file_id, sheet, range, values, formulas }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })

    const common = { url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, loadPlaywright: coedit.loadPlaywright }
    const result = await verification.runRangeWriteContract({
      inspect: () => inspector.inspectWorkbookLive(common),
      read: () => rangeReader.readRangeLive({ ...common, sheet, range }),
      edit: () => bulkWriter.writeBulkLive({ ...common, sheet, range, values, formulas: formulas || null }),
      values,
      formulas: formulas || null,
    })
    return textResult({
      ...result,
      callerId: hivo.id,
      identitasForras: hivo.forras || null,
      hitelesitoForras: cred.forras || null,
      ncUser: cred.user,
      sheet,
      range,
    })
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
