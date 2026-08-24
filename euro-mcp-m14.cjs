'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m13.cjs')
const coedit = require('./coedit.cjs')
const bulkWriter = require('./bulk-writer.cjs')

function textResult(payload) { return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] } }

const scalar = z.union([z.string(), z.number(), z.boolean(), z.null()])
const valueMatrix = z.array(z.array(scalar).min(1)).min(1)
const formulaMatrix = z.array(z.array(z.union([z.string(), z.null()])).min(1)).min(1)

server.tool(
  'office_write_range',
  'Writes a rectangular 2D value matrix, optionally overlaid by a same-sized formula matrix, into the CURRENT live ONLYOFFICE spreadsheet co-edit session in one callCommand round trip. Null value cells are cleared. Formula entries must begin with =. Dimensions must exactly match the target A1 range. No saved-file, OOXML, or DocBuilder edit fallback is used.',
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
    const r = await bulkWriter.writeBulkLive({
      url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id,
      sheet, range, values, formulas: formulas || null,
      loadPlaywright: coedit.loadPlaywright,
    })
    return textResult({ ...r, callerId: hivo.id, identitasForras: hivo.forras || null, hitelesitoForras: cred.forras || null, ncUser: cred.user })
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
