'use strict'

const assert=require('assert')
const {conditionalFormattingCommand}=require('./live-conditional-formatting.cjs')

function makeApi(){
  const rules=[]
  const fc={
    GetCount(){return rules.length},
    GetItem(i){return rules[i]||null},
    Add(type,operator,formula1,formula2){
      const r={type,operator,formula1,formula2,priority:rules.length+1,fill:null,
        GetType(){return this.type},GetOperator(){return this.operator},GetFormula1(){return this.formula1},GetFormula2(){return this.formula2},GetPriority(){return this.priority},
        SetPriority(v){this.priority=v},SetFillColor(v){this.fill=v},GetFillColor(){return this.fill},GetAppliesTo(){return {GetAddress(){return '$A$1:$A$3'}}},
        Delete(){const p=rules.indexOf(this);if(p>=0)rules.splice(p,1)} }
      rules.push(r);return r
    }
  }
  const range={GetFormatConditions(){return fc}}
  const sheet={GetRange(){return range}}
  return {api:{GetSheet(){return sheet},CreateColorFromRGB(r,g,b){return {GetRGB(){return (r<<16)|(g<<8)|b}}}},rules}
}

const env=makeApi();global.Api=env.api
let r=conditionalFormattingCommand({type:'cf.add',sheet:'Sheet1',range:'A1:A3',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'200',fillColor:[255,0,0],priority:1}})
assert.equal(r.ok,true);assert.equal(r.verification.status,'PASS');assert.equal(r.afterCount,1);assert.equal(r.actual.type,'xlCellValue');assert.equal(r.actual.operator,'xlGreater');assert.equal(r.actual.formula1,'200');assert.equal(r.actual.priority,1)
r=conditionalFormattingCommand({type:'cf.inspect',sheet:'Sheet1',range:'A1:A3'});assert.equal(r.count,1);assert.equal(r.rules[0].type,'xlCellValue')
r=conditionalFormattingCommand({type:'cf.delete',sheet:'Sheet1',range:'A1:A3',index:0});assert.equal(r.ok,true);assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actualCount,0)

// Serialized editor-context regression: command must not depend on module closure state.
const serialized=new Function('Api','spec',`return (${conditionalFormattingCommand.toString()})(spec)`)
const env2=makeApi();r=serialized(env2.api,{type:'cf.add',sheet:'Sheet1',range:'A1:A3',rule:{type:'xlCellValue',operator:'xlBetween',formula1:'100',formula2:'200'}})
assert.equal(r.ok,true);assert.equal(r.actual.formula2,'200')

console.log('test_live_conditional_formatting.cjs: OK')
