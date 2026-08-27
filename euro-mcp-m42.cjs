'use strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m2.cjs')
const coedit = require('./coedit.cjs')
const nf = require('./number-format-intelligence.cjs')

function textResult(payload){ return {content:[{type:'text',text:JSON.stringify(payload,null,2)}]} }
async function context(file_id){ const h=coedit.detectCallerId(); if(!h.ok)return {error:textResult({ok:false,outcome:'azonossag-hiany',error:h.indok})}; const c=await coedit.credentialsFor(h.id); if(!c.ok)return {error:textResult({ok:false,outcome:'konfig-hiany',callerId:h.id,error:c.indok})}; return {common:{url:c.url,user:c.user,pass:c.pass,fileId:file_id,loadPlaywright:coedit.loadPlaywright},meta:{callerId:h.id,ncUser:c.user}} }

const preset=z.object({kind:z.enum(['currency','percent','date','time','datetime','decimal']),decimals:z.number().int().min(0).max(12).optional(),grouped:z.boolean().optional(),symbol:z.string().min(1).optional(),pattern:z.string().min(1).optional()}).strict()
const custom=z.object({mask:z.string().min(1)}).strict()

server.tool('office_number_format','Applies and verifies an Excel/ONLYOFFICE number format in the CURRENT live spreadsheet editor. Presets: currency, percent, date, time, datetime, decimal; or provide a custom mask. Verification reads the same live range before and after formatting and proves that the number-format mask changed as requested while the underlying raw value stayed unchanged. Returns rawValue, value and displayText separately per cell. No WebDAV, downloaded XLSX, OOXML or DocBuilder fallback.',{file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1),number_format:z.union([preset,custom])},async(args)=>{
  const c=await context(args.file_id); if(c.error)return c.error
  const result=await nf.formatNumberRangeLive({...c.common,sheet:args.sheet,range:args.range,numberFormat:args.number_format})
  return textResult({...result,...c.meta})
})

if(require.main===module){ const transport=new StdioServerTransport(); server.connect(transport) }
module.exports={server}
