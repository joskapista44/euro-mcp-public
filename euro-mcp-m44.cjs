'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m43.cjs')
const coedit = require('./coedit.cjs')
const structure = require('./live-structure.cjs')

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
  type: z.enum(['range.merge', 'range.unmerge', 'sheet.inspect', 'sheet.move']),
  sheet: z.string().min(1).optional(),
  range: z.string().min(1).optional(),
  referenceSheet: z.string().min(1).optional(),
  position: z.enum(['before', 'after']).optional(),
}).strict()

server.tool(
  'office_structure',
  'M4.4 live worksheet structure operations in the CURRENT ONLYOFFICE spreadsheet editor. ' +
    'Supports verified merge/unmerge, ordered sheet inspection and verified sheet moves. ' +
    'Verification is fail-closed and uses only the live editor session; no saved-file, WebDAV, OOXML or DocBuilder fallback is used.',
  { file_id: z.string(), structure: spec },
  async ({ file_id, structure: requested }) => {
    const c = await context(file_id)
    if (c.error) return c.error
    const result = await structure.runStructureLive({ ...c.common, spec: requested })
    return textResult({ ...result, ...c.meta })
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
