'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp.cjs')
const coedit = require('./coedit.cjs')
const inspector = require('./workbook-inspector.cjs')

function textResult(payload) {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] }
}

server.tool(
  'office_inspect_workbook',
  'Reads the CURRENT in-editor structure of an ALREADY EXISTING xlsx workbook through the live ' +
    'ONLYOFFICE spreadsheet co-edit session. Returns sheet order/names/visibility/active state, ' +
    'used ranges and dimensions, AutoFilter state, freeze-pane type, defined names, and table ' +
    'metadata where the deployed editor exposes a live table-collection API. This tool never ' +
    'downloads/parses OOXML and never falls back to DocBuilder; unsupported live API fields are ' +
    'returned explicitly in `unsupported` instead of being guessed.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the existing xlsx workbook to inspect'),
  },
  async ({ file_id }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })

    const r = await inspector.inspectWorkbookLive({
      url: cred.url,
      user: cred.user,
      pass: cred.pass,
      fileId: file_id,
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
      activeSheet: r.activeSheet ?? null,
      sheetCount: r.sheetCount ?? null,
      sheets: r.sheets ?? null,
      freezePanes: r.freezePanes ?? null,
      definedNames: r.definedNames ?? null,
      unsupported: r.unsupported ?? null,
      error: r.error || r.indok || null,
    })
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
