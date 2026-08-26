'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m35.cjs')
const coedit = require('./coedit.cjs')
const formulaVerification = require('./formula-verification.cjs')

function textResult(payload){return {content:[{type:'text',text:JSON.stringify(payload,null,2)}]}}

server.tool(
  'office_verify_formulas',
  'Verifies live spreadsheet formulas after a formula operation through ONLYOFFICE spreadsheeteditor/callCommand. Compares exact formula read-back against an optional expected 2D formula matrix and returns calculated/display results plus proven error status. Mismatch and unsupported/ambiguous evidence are explicit; no DocBuilder, OOXML, XLSX, ZIP, or downloaded-package fallback.',
  {
    file_id:z.string(),
    sheet:z.string().min(1),
    range:z.string().min(1),
    expected_formulas:z.array(z.array(z.string().nullable())).nullable().optional()
  },
  async ({file_id,sheet,range,expected_formulas})=>{
    const hivo=coedit.detectCallerId(); if(!hivo.ok)return textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})
    const cred=await coedit.credentialsFor(hivo.id); if(!cred.ok)return textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})
    const result=await formulaVerification.verifyFormulaLive({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,sheet,range,expectedFormulas:expected_formulas||null,loadPlaywright:coedit.loadPlaywright})
    return textResult({...result,callerId:hivo.id,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null,ncUser:cred.user})
  }
)

if(require.main===module){const transport=new StdioServerTransport();server.connect(transport)}
module.exports={server}
