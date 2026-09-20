'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
const ops=[{intent:'set_header_footer',sheet:'Report',slot:'oddHeader',value:'&CReport'},{intent:'set_header_footer',sheet:'Report',slot:'oddFooter',value:'&C&P / &N'}]
const p=batch.planTask({operations:ops});assert.equal(p.ok,true);assert.deepEqual(p.steps.map(x=>x.family),['header-footer','header-footer'])
assert.equal(mcp.schema.safeParse({file_id:'1',operations:ops}).success,true)
assert.equal(batch.planTask({operations:[ops[0],ops[0]]}).ok,false)
console.log('XLSX HEADER FOOTER BATCH STATIC: PASS')
