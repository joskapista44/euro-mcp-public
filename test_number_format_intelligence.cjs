'use strict'
const assert=require('assert')
const fs=require('fs')
const nf=require('./number-format-intelligence.cjs')

assert.deepEqual(nf.buildNumberFormat({kind:'currency'}),{ok:true,kind:'currency',mask:'€ #,##0.00'})
assert.deepEqual(nf.buildNumberFormat({kind:'percent',decimals:1}),{ok:true,kind:'percent',mask:'0.0%'})
assert.deepEqual(nf.buildNumberFormat({kind:'decimal',decimals:3,grouped:false}),{ok:true,kind:'decimal',mask:'0.000'})
assert.deepEqual(nf.buildNumberFormat({kind:'date'}),{ok:true,kind:'date',mask:'yyyy-mm-dd'})
assert.deepEqual(nf.buildNumberFormat({kind:'time'}),{ok:true,kind:'time',mask:'hh:mm:ss'})
assert.deepEqual(nf.buildNumberFormat({kind:'datetime'}),{ok:true,kind:'datetime',mask:'yyyy-mm-dd hh:mm:ss'})
assert.deepEqual(nf.buildNumberFormat({mask:'[$€-hu-HU] #,##0.00'}),{ok:true,kind:'custom',mask:'[$€-hu-HU] #,##0.00'})
assert.equal(nf.buildNumberFormat({kind:'decimal',decimals:13}).ok,false)

function read(raw,value,text,format){return {ok:true,cells:[[{address:'A1',rawValue:raw,value,displayText:text,numberFormat:format}]]}}
const write={ok:true,outcome:'ok'}
let r=nf.verifyNumberFormatTransition(read(1234.5,1234.5,'1234.5','General'),write,read(1234.5,1234.5,'€ 1,234.50','€ #,##0.00'),'€ #,##0.00')
assert.equal(r.ok,true); assert.equal(r.verification.outcome,'pass'); assert.equal(r.verification.cells[0].displayText,'€ 1,234.50'); assert.equal(r.verification.cells[0].checks.underlyingValuePreserved.status,'pass')

r=nf.verifyNumberFormatTransition(read(0.15,0.15,'0.15','General'),write,read(0.15,0.15,'15.00%','0.00%'),'0.00%')
assert.equal(r.ok,true); assert.equal(r.verification.cells[0].rawValue,0.15); assert.equal(r.verification.cells[0].displayText,'15.00%')

r=nf.verifyNumberFormatTransition(read(45500,45500,'45500','General'),write,read(45500,45500,'2024-07-27','yyyy-mm-dd'),'yyyy-mm-dd')
assert.equal(r.ok,true); assert.equal(r.verification.cells[0].value,45500)

r=nf.verifyNumberFormatTransition(read(1,1,'1','General'),write,read(1,1,'1.00','0.00'),'#,##0.00')
assert.equal(r.ok,false); assert.equal(r.outcome,'verification-mismatch'); assert.deepEqual(r.verification.mismatches,['A1'])

r=nf.verifyNumberFormatTransition(read(1,1,'1','General'),write,read(2,2,'2.00','0.00'),'0.00')
assert.equal(r.ok,false); assert.equal(r.verification.cells[0].checks.underlyingValuePreserved.status,'fail')

r=nf.verifyNumberFormatTransition(read(1,1,'1','General'),write,{ok:false,outcome:'callback-timeout'},'0.00')
assert.equal(r.ok,false); assert.equal(r.outcome,'verification-unknown'); assert.equal(r.verification.outcome,'unknown')

const src=fs.readFileSync('number-format-intelligence.cjs','utf8')
assert.match(src,/readRangeInFrame/)
assert.match(src,/formattingCommand/)
assert.doesNotMatch(src,/WebDAV|PROPFIND|OOXML|writeFile|unzip/i)
const entry=fs.readFileSync('euro-mcp-m42.cjs','utf8')
assert.match(entry,/office_number_format/)
assert.match(entry,/require\('\.\/euro-mcp-m36\.cjs'\)/)
assert.doesNotMatch(entry,/require\('\.\/euro-mcp-m2\.cjs'\)/)
const pkg=JSON.parse(fs.readFileSync('package.json','utf8'))
assert.equal(pkg.main,'euro-mcp-m42.cjs')
assert.match(pkg.scripts['test:js'],/test_number_format_intelligence\.cjs/)
console.log('test_number_format_intelligence: OK')
