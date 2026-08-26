'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m33.cjs')
const coedit = require('./coedit.cjs')
const crossSheet = require('./cross-sheet-formulas.cjs')

function textResult(payload){return {content:[{type:'text',text:JSON.stringify(payload,null,2)}]}}

server.tool(
  'office_write_cross_sheet_formulas',
  'Writes sheet-qualified formulas through the live ONLYOFFICE spreadsheeteditor/callCommand path, validates every referenced worksheet, and verifies formula text by same-session live read-back. Supports quoted sheet names, cross-sheet ranges, and relative/absolute/mixed A1 references. No DocBuilder, OOXML, ZIP/XML or downloaded-XLSX fallback.',
  {file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1),formulas:z.array(z.array(z.string().min(1))).min(1)},
  async ({file_id,sheet,range,formulas})=>{
    const hivo=coedit.detectCallerId(); if(!hivo.ok)return textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})
    const cred=await coedit.credentialsFor(hivo.id); if(!cred.ok)return textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})
    const result=await crossSheet.writeCrossSheetLive({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,sheet,range,formulas,loadPlaywright:coedit.loadPlaywright})
    return textResult({...result,callerId:hivo.id,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null,ncUser:cred.user})
  }
)

server.tool(
  'office_qualify_cross_sheet_reference',
  'Builds a safely quoted sheet-qualified A1 cell/range reference for use in a formula. This helper does not edit the workbook.',
  {sheet:z.string().min(1),reference:z.string().min(1)},
  async ({sheet,reference})=>textResult(crossSheet.qualifyA1(sheet,reference))
)

if(require.main===module){const transport=new StdioServerTransport();server.connect(transport)}
module.exports={server}
