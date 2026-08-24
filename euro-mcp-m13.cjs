'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m12.cjs')
const coedit = require('./coedit.cjs')
const workbookOps = require('./workbook-ops.cjs')

function textResult(payload) { return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] } }

const operationSchema = z.object({
  type: z.enum([
    'sheet.create', 'sheet.copy', 'sheet.delete', 'sheet.rename', 'sheet.move',
    'range.clear', 'range.copy', 'range.move',
    'rows.insert', 'rows.delete', 'columns.insert', 'columns.delete',
  ]),
  sheet: z.string().min(1).optional(),
  name: z.string().min(1).optional(),
  position: z.enum(['before', 'after']).optional(),
  referenceSheet: z.string().min(1).optional(),
  range: z.string().min(1).optional(),
  targetSheet: z.string().min(1).optional(),
  targetRange: z.string().min(1).optional(),
})

server.tool(
  'office_workbook_operation',
  'Performs one core workbook edit against the CURRENT live ONLYOFFICE spreadsheet co-edit session. ' +
    'Supported operation families: worksheet create/copy/delete/rename/move, range clear/copy/move, ' +
    'and row/column insert/delete. Unsupported editor methods fail closed; there is no saved-file rewrite or builder fallback.',
  { file_id: z.string(), operation: operationSchema },
  async ({ file_id, operation }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await workbookOps.runOperationLive({
      url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id,
      operation, loadPlaywright: coedit.loadPlaywright,
    })
    return textResult({ ...r, callerId: hivo.id, identitasForras: hivo.forras || null, hitelesitoForras: cred.forras || null, ncUser: cred.user })
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
