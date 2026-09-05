'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m42.cjs')
const coedit = require('./coedit.cjs')
const layout = require('./live-layout.cjs')
const workbookOps = require('./workbook-ops.cjs')

function textResult(payload) { return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] } }
async function context(file_id) {
  const h = coedit.detectCallerId()
  if (!h.ok) return { error: textResult({ ok: false, outcome: 'azonossag-hiany', error: h.indok }) }
  const c = await coedit.credentialsFor(h.id)
  if (!c.ok) return { error: textResult({ ok: false, outcome: 'konfig-hiany', callerId: h.id, error: c.indok }) }
  return {
    common: { url: c.url, user: c.user, pass: c.pass, fileId: file_id, loadPlaywright: coedit.loadPlaywright },
    meta: { callerId: h.id, ncUser: c.user, identitasForras: h.forras || null, hitelesitoForras: c.forras || null },
  }
}

const spec = z.object({
  type: z.enum([
    'column.width', 'row.height', 'columns.hidden', 'rows.hidden',
    'autofit.columns', 'autofit.rows',
    'rows.insert', 'rows.delete', 'columns.insert', 'columns.delete',
  ]),
  sheet: z.string().min(1),
  range: z.string().min(1),
  width: z.number().positive().optional(),
  height: z.number().positive().optional(),
  hidden: z.boolean().optional(),
}).strict()

server.tool(
  'office_layout',
  'M4.3 live row/column layout operations in the CURRENT ONLYOFFICE spreadsheet editor. ' +
    'Supports verified column width, row height, row/column hide-show, and AutoFit. ' +
    'Existing M1.3 row/column insert-delete is reused rather than reimplemented; because the live API exposes no generic structural postcondition getter, those operations return explicit UNKNOWN verification after a successful mutation. ' +
    'No saved-file, WebDAV, OOXML or DocBuilder fallback is used as live success evidence.',
  { file_id: z.string(), layout: spec },
  async ({ file_id, layout: requested }) => {
    const c = await context(file_id)
    if (c.error) return c.error

    if (/^(rows|columns)\.(insert|delete)$/.test(requested.type)) {
      const result = await workbookOps.runOperationLive({
        ...c.common,
        operation: { type: requested.type, sheet: requested.sheet, range: requested.range },
      })
      if (!result.ok) return textResult({ ...result, ...c.meta })
      return textResult({
        ...result,
        ...c.meta,
        reusedComponent: 'workbook-ops.cjs',
        verification: {
          status: 'UNKNOWN',
          reason: 'the existing live structural operation succeeded, but there is no generic ONLYOFFICE row/column insert-delete getter that proves the requested structural postcondition without relying on workbook contents',
        },
      })
    }

    const result = await layout.runLayoutLive({ ...c.common, spec: requested })
    return textResult({ ...result, ...c.meta })
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
