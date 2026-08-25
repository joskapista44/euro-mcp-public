'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m2.cjs')
const coedit = require('./coedit.cjs')
const formulaInspector = require('./formula-inspector.cjs')

function textResult(payload) { return { content: [{ type:'text', text:JSON.stringify(payload,null,2) }] } }

server.tool(
  'office_inspect_formulas',
  'Inspects formulas in a target range from the CURRENT live ONLYOFFICE spreadsheet editor in one callCommand round trip. Distinguishes formula/constant/blank where GetFormula proves it, returns formula text plus current calculated value and display text, and returns unknown/null rather than guessing when a field is unsupported. No DocBuilder, OOXML, ZIP/XML or downloaded-XLSX fallback.',
  { file_id:z.string(), sheet:z.string().min(1), range:z.string().min(1) },
  async ({file_id,sheet,range}) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})
    const result = await formulaInspector.inspectFormulaLive({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,sheet,range,loadPlaywright:coedit.loadPlaywright})
    return textResult({...result,callerId:hivo.id,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null,ncUser:cred.user})
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
