'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m35.cjs')
const coedit = require('./coedit.cjs')
const formulaVerification = require('./formula-verification-intelligence.cjs')

function textResult(payload){return {content:[{type:'text',text:JSON.stringify(payload,null,2)}]}}

server.tool(
  'office_verify_formula_range',
  'Verifies formula operation results in a live ONLYOFFICE spreadsheeteditor/callCommand session. Compares formula read-back with an optional expected formula matrix, includes calculated/display results and proven formula-error diagnosis, and fails closed to unknown when live evidence is incomplete or ambiguous. Suitable for direct writes, fills, bulk ranges and cross-sheet formulas. No DocBuilder, OOXML, XLSX, ZIP, downloaded-package, or offline fallback.',
  {file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1),expected_formulas:z.array(z.array(z.string())).optional()},
  async ({file_id,sheet,range,expected_formulas})=>{
    const hivo=coedit.detectCallerId(); if(!hivo.ok)return textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})
    const cred=await coedit.credentialsFor(hivo.id); if(!cred.ok)return textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})
    const result=await formulaVerification.verifyFormulaRangeLive({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,sheet,range,expectedFormulas:expected_formulas||null,loadPlaywright:coedit.loadPlaywright})
    return textResult({...result,callerId:hivo.id,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null,ncUser:cred.user})
  }
)

if(require.main===module){const transport=new StdioServerTransport();server.connect(transport)}
module.exports={server}
