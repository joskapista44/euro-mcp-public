' strict'

const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')
const { server } = require('./euro-mcp-m15.cjs')
const coedit = require('./coedit.cjs')
const formatting = require('./live-formatting.cjs')

function textResult(payload){ return { content:[{type:'text',text:JSON.stringify(payload,null,2)}] } }
async function context(file_id){
  const hivo=coedit.detectCallerId(); if(!hivo.ok)return {error:textResult({ok:false,outcome:'azonossag-hiany',error:hivo.indok})}
  const cred=await coedit.credentialsFor(hivo.id); if(!cred.ok)return {error:textResult({ok:false,outcome:'konfig-hiany',callerId:hivo.id,error:cred.indok})}
  return { common:{url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,loadPlaywright:coedit.loadPlaywright}, meta:{callerId:hivo.id,ncUser:cred.user,identitasForras:hivo.forras||null,hitelesitoForras:cred.forras||null} }
}

const rgb=z.tuple([z.number().int().min(0).max(255),z.number().int().min(0).max(255),z.number().int().min(0).max(255)])
const formatSchema=z.object({
  bold:z.boolean().optional(), italic:z.boolean().optional(), fontName:z.string().min(1).optional(), fontSize:z.number().positive().max(409).optional(),
  fontColor:rgb.optional(), fillColor:rgb.optional(),
  border:z.object({index:z.enum(['Top','Bottom','Left','Right','InsideHorizontal','InsideVertical','DiagonalDown','DiagonalUp']),style:z.string().min(1),color:rgb}).optional(),
  alignHorizontal:z.enum(['left','right','center','justify']).optional(), alignVertical:z.enum(['center','bottom','top','distributed','justify']).optional(),
  wrap:z.boolean().optional(), numberFormat:z.string().min(1).optional(),
}).strict()

server.tool('office_format_range','Formats a range in the CURRENT live ONLYOFFICE spreadsheet editor. Supports bold/italic/font name/font size/font color/fill/borders/alignment/wrap and arbitrary number-format masks including currency, percent and date. Fails closed if the live Office API does not expose a requested operation. No DocBuilder, OOXML, or downloaded-XLSX fallback.',{file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1),format:formatSchema},async(args)=>{
  const c=await context(args.file_id); if(c.error)return c.error
  const result=await formatting.formatRangeLive({...c.common,sheet:args.sheet,range:args.range,format:args.format})
  return textResult({...result,...c.meta})
})

server.tool('office_resize_range','Resizes rows/columns for a range in the CURRENT live ONLYOFFICE spreadsheet editor. Supports explicit column width, explicit row height, and row/column autofit. No saved-file fallback.',{file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1),column_width:z.number().positive().optional(),row_height:z.number().positive().optional(),autofit_rows:z.boolean().optional(),autofit_columns:z.boolean().optional()},async(args)=>{
  const c=await context(args.file_id); if(c.error)return c.error
  const result=await formatting.layoutRangeLive({...c.common,sheet:args.sheet,range:args.range,columnWidth:args.column_width,rowHeight:args.row_height,autofitRows:!!args.autofit_rows,autofitColumns:!!args.autofit_columns})
  return textResult({...result,...c.meta})
})

server.tool('office_merge_range','Merges or unmerges a range in the CURRENT live ONLYOFFICE spreadsheet editor. Merge supports across=true for row-wise merges. No saved-file fallback.',{file_id:z.string(),sheet:z.string().min(1),range:z.string().min(1),operation:z.enum(['merge','unmerge']),across:z.boolean().optional()},async(args)=>{
  const c=await context(args.file_id); if(c.error)return c.error
  const result=await formatting.mergeRangeLive({...c.common,sheet:args.sheet,range:args.range,operation:args.operation,across:!!args.across})
  return textResult({...result,...c.meta})
})

server.tool('office_freeze_panes','Controls freeze panes in the CURRENT live ONLYOFFICE spreadsheet editor. Actions: rows, columns, at, unfreeze. For rows/columns value is the count; for at value is an A1 range/cell. No saved-file fallback.',{file_id:z.string(),sheet:z.string().min(1),action:z.enum(['rows','columns','at','unfreeze']),value:z.union([z.number().int().min(0),z.string().min(1)]).optional()},async(args)=>{
  const c=await context(args.file_id); if(c.error)return c.error
  if((args.action==='rows'||args.action==='columns') && !Number.isInteger(args.value))return textResult({ok:false,outcome:'invalid-freeze-value',error:'rows/columns require integer value'})
  if(args.action==='at' && typeof args.value!=='string')return textResult({ok:false,outcome:'invalid-freeze-value',error:'at requires an A1 range/cell string'})
  const result=await formatting.freezePanesLive({...c.common,sheet:args.sheet,action:args.action,value:args.value})
  return textResult({...result,...c.meta})
})

if(require.main===module){ const transport=new StdioServerTransport(); server.connect(transport) }
module.exports={server}
