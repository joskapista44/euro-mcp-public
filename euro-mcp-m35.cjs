'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m34.cjs')
const coedit = require('./coedit.cjs')
const formulaErrors = require('./formula-error-intelligence.cjs')

function textResult(payload){return {content:[{type:'text',text:JSON.stringify(payload,null,2)}]}}

server.tool(
  'office_diagnose_formula_errors',
  'Diagnoses formula errors in a live ONLYOFFICE spreadsheeteditor/callCommand session. Returns formula text, calculated value, display text, proven error status/type, per-cell evidence, and a range summary. Error classification is fail-closed: constants never become formula errors and unavailable evidence becomes unknown. No DocBuilder, OOXML, XLSX, ZIP, or downloaded-package fallback.',
  {file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1)},
  async ({file_id,sheet,range})=>{
    const hivo=coedit.detectCallerId(); if(!hivo.ok)return textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})
    const cred=await coedit.credentialsFor(hivo.id); if(!cred.ok)return textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})
    const result=await formulaErrors.diagnoseFormulaErrorsLive({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,sheet,range,loadPlaywright:coedit.loadPlaywright})
    return textResult({...result,callerId:hivo.id,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null,ncUser:cred.user})
  }
)

if(require.main===module){const transport=new StdioServerTransport();server.connect(transport)}
module.exports={server}
