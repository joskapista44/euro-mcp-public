'use strict'
const assert=require('assert/strict'),mcp=require('./xlsx-batch-mcp.cjs')
const ok=mcp.schema.safeParse({file_id:'123',operations:[{intent:'set_first_page_number',sheet:'Report',value:5}]});assert.equal(ok.success,true)
const clear=mcp.schema.safeParse({file_id:'123',operations:[{intent:'set_first_page_number',sheet:'Report',value:null}]});assert.equal(clear.success,true)
const bad=mcp.schema.safeParse({file_id:'123',operations:[{intent:'set_first_page_number',sheet:'Report',value:'5'}]});assert.equal(bad.success,false)
const order=mcp.schema.safeParse({file_id:'123',operations:[{intent:'set_first_page_number',sheet:'Report',value:5,pageOrder:'overThenDown'}]});assert.equal(order.success,false)
console.log('XLSX FIRST PAGE NUMBER MCP SCHEMA STATIC: PASS')
