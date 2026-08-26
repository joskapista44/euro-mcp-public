'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m32.cjs')
const coedit = require('./coedit.cjs')
const formulaFill = require('./formula-fill.cjs')

function textResult(payload){return {content:[{type:'text',text:JSON.stringify(payload,null,2)}]}}

server.tool(
  'office_fill_formula',
  'Fills a one-dimensional target range from one live source formula cell, preserving relative, absolute and mixed A1 references and verifying formulas by same-session live read-back. No DocBuilder, OOXML, ZIP/XML or downloaded-XLSX fallback.',
  {file_id:z.string(),sheet:z.string().min(1),source_cell:z.string().min(1),target_range:z.string().min(1)},
  async ({file_id,sheet,source_cell,target_range})=>{
    const hivo=coedit.detectCallerId(); if(!hivo.ok)return textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})
    const cred=await coedit.credentialsFor(hivo.id); if(!cred.ok)return textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})
    const result=await formulaFill.fillFormulaLive({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,sheet,sourceCell:source_cell,targetRange:target_range,loadPlaywright:coedit.loadPlaywright})
    return textResult({...result,callerId:hivo.id,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null,ncUser:cred.user})
  }
)

if(require.main===module){const transport=new StdioServerTransport();server.connect(transport)}
module.exports={server}
