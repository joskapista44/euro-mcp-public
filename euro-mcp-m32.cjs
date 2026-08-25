'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m31.cjs')
const coedit = require('./coedit.cjs')
const formulaWriter = require('./formula-writer.cjs')

function textResult(payload){ return {content:[{type:'text',text:JSON.stringify(payload,null,2)}]} }

server.tool(
  'office_write_formulas',
  'Writes a 2D formula matrix into the CURRENT live ONLYOFFICE spreadsheet editor and verifies formula text by live read-back in the same editor session. No DocBuilder, OOXML, ZIP/XML or downloaded-XLSX fallback.',
  {file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1),formulas:z.array(z.array(z.string().min(1))).min(1)},
  async ({file_id,sheet,range,formulas})=>{
    const hivo=coedit.detectCallerId(); if(!hivo.ok)return textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})
    const cred=await coedit.credentialsFor(hivo.id); if(!cred.ok)return textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})
    const result=await formulaWriter.writeFormulaLive({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,sheet,range,formulas,loadPlaywright:coedit.loadPlaywright})
    return textResult({...result,callerId:hivo.id,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null,ncUser:cred.user})
  }
)

if(require.main===module){const transport=new StdioServerTransport();server.connect(transport)}
module.exports={server}
