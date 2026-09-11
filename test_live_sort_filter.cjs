'use strict'
const assert=require('assert')
const {sortFilterCommand}=require('./live-sort-filter.cjs')
let failed=0
function check(n,fn){try{fn();console.log('OK    '+n)}catch(e){failed++;console.log('FAIL  '+n+'\n      '+(e.stack||e))}}
function fixture(opts={}){
 const state={values:opts.values||[['Name','Score'],['B',2],['A',1]],filter:false,mode:false,filters:[]}
 const range={
  GetAddress(){return '$A$1:$B$3'},GetValue(){return state.values},
  SetSort(keyRange,order){if(!opts.ignoreSort){const h=state.values[0],rows=state.values.slice(1).sort((a,b)=>String(a[0]).localeCompare(String(b[0])));state.values=[h,...(order==='xlDescending'?rows.reverse():rows)]}},
  SetAutoFilter(field,criteria1,operator,criteria2){state.filter=true;if(field!=null){state.mode=true;state.filters=[{GetOperator(){return operator},GetCriteria1(){return criteria1},GetCriteria2(){return criteria2},GetOn(){return true}}]}}
 }
 const af={GetRange(){return state.filter?range:null},GetFilterMode(){return state.mode},GetFilters(){return state.filters},ApplyFilter(){},ShowAllData(){state.mode=false;state.filters=[]}}
 return {state,api:{GetSheet(n){return n==='Sheet1'?{GetRange(){return range},GetAutoFilter(){return af}}:null}}}
}
function withApi(f,fn){const old=global.Api;global.Api=f.api;try{return fn()}finally{if(old===undefined)delete global.Api;else global.Api=old}}
check('sort requires keyRange and verifies before/after readback',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'sort.apply',sheet:'Sheet1',range:'A1:B3',keyRange:'A1:A3',order:'asc',hasHeaders:true}));assert.equal(r.verification.status,'PASS');assert.deepStrictEqual(r.after,[['Name','Score'],['A',1],['B',2]])})
check('sort without keyRange fails closed',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'sort.apply',sheet:'Sheet1',range:'A1:B3',order:'asc'}));assert.equal(r.outcome,'invalid-operation')})
check('filter inspect reports live state',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'filter.inspect',sheet:'Sheet1'}));assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.present,false)})
check('filter enable verifies live range exists',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'filter.enable',sheet:'Sheet1',range:'A1:B3'}));assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.present,true)})
check('filter set verifies readable criteria',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'filter.set',sheet:'Sheet1',range:'A1:B3',field:1,criteria1:'A',operator:'xlOr'}));assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.filterMode,true);assert.equal(r.verification.actual.filters[0].criteria1,'A')})
check('filter reapply returns public state readback',()=>{const f=fixture();f.state.filter=true;f.state.mode=true;const r=withApi(f,()=>sortFilterCommand({type:'filter.reapply',sheet:'Sheet1'}));assert.equal(r.verification.status,'PASS')})
check('filter clear verifies filters empty',()=>{const f=fixture();f.state.filter=true;f.state.mode=true;f.state.filters=[{}];const r=withApi(f,()=>sortFilterCommand({type:'filter.clear',sheet:'Sheet1'}));assert.equal(r.verification.status,'PASS');assert.deepStrictEqual(r.verification.actual.filters,[])})
check('invalid sort order fails closed',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'sort.apply',sheet:'Sheet1',range:'A1:B3',keyRange:'A1:A3',order:'sideways'}));assert.equal(r.outcome,'invalid-operation')})
check('unknown operation fails closed',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'wat',sheet:'Sheet1'}));assert.equal(r.outcome,'invalid-operation')})
console.log('\n'+(failed===0?'MIND OK':failed+' FAILED'));process.exit(failed?1:0)
