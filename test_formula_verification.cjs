'use strict'
const assert=require('assert')
const fv=require('./formula-verification.cjs')

function inspected(cells, unsupported=[]){return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:'Calc',range:'B2:C3',rows:2,columns:2,cellCount:4,cells,unsupported}}
function f(address,formula,value,text){return {address,row:Number(address.match(/\d+/)[0]),column:1,formulaStatus:'formula',formula,calculatedValue:value,displayText:text,cellType:'formula'}}
function c(address,value,text=value){return {address,row:Number(address.match(/\d+/)[0]),column:1,formulaStatus:'constant',formula:null,calculatedValue:value,displayText:text,cellType:value===''?'blank':typeof value}}

let r=fv.verifyInspection(inspected([[f('B2','=A2*2',4,'4'),f('C2','=Sheet2!A1',9,'9')],[f('B3','=A3*2',6,'6'),f('C3','=SUM(A1:A3)',12,'12')]]),[['=A2*2','=Sheet2!A1'],['=A3*2','=SUM(A1:A3)']])
assert.equal(r.ok,true);assert.equal(r.verificationStatus,'verified');assert.equal(r.summary.matched,4);assert.equal(r.summary.formulaErrors,0)

r=fv.verifyInspection(inspected([[f('B2','=A2*2',4,'4'),f('C2','=Wrong!A1','#REF!','#REF!')],[f('B3','=A3*3',9,'9'),c('C3','')]]),[['=A2*2','=Sheet2!A1'],['=A3*2',null]])
assert.equal(r.verificationStatus,'mismatch');assert.equal(r.summary.mismatched,2);assert.equal(r.summary.formulaErrors,1);assert.equal(r.errors[0].errorType,'REF')

r=fv.verifyInspection(inspected([[f('B2','=A2*2',4,'4'),f('C2','=Sheet2!A1',9,'9')],[f('B3','=A3*2',6,'6'),f('C3','=SUM(A1:A3)',12,'12')]], [{field:'GetText',reason:'unavailable'}]),[['=A2*2','=Sheet2!A1'],['=A3*2','=SUM(A1:A3)']])
assert.equal(r.verificationStatus,'unknown');assert.ok(r.summary.unknown>0)

r=fv.verifyInspection(inspected([[f('B2','=A2*2','#DIV/0!','#DIV/0!'),c('C2','x')],[f('B3','=A3*2',6,'6'),c('C3','')]]),null)
assert.equal(r.verificationStatus,'observed');assert.equal(r.summary.unchecked,4);assert.equal(r.summary.formulaErrors,1)

r=fv.verifyInspection({...inspected([[f('B2','=A2*2',4,'4')]]),rows:1,columns:1,cellCount:1,range:'B2'},[['=A2*2','=extra']])
assert.equal(r.ok,false);assert.equal(r.outcome,'invalid-expectation')

const source=require('fs').readFileSync('./formula-verification.cjs','utf8')+require('fs').readFileSync('./euro-mcp-m36.cjs','utf8')
for(const forbidden of ['DocBuilder','readFileSync(".xlsx','adm-zip','unzipper','sheetjs']) assert.equal(source.includes(forbidden),false)
assert.ok(source.includes('spreadsheeteditor'));assert.ok(source.includes('callCommand'))
console.log('formula verification: ok')
