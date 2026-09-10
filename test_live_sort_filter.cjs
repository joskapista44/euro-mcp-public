'use strict'
const assert=require('assert')
const {sortFilterCommand}=require('./live-sort-filter.cjs')
let failed=0
function check(n,fn){try{fn();console.log('OK    '+n)}catch(e){failed++;console.log('FAIL  '+n+'\n      '+(e.stack||e))}}
function fixture(opts={}){
 const state={values:opts.values||[['Name','Score'],['B',2],['A',1]],filter:false,mode:false,filters:[]}
 const range={GetAddress(){return '$A$1:$B$3'},GetValue(){return state.values},SetSort(){if(!opts.ignoreSort)state.values=[state.values[0],state.values[2],state.values[1]]},SetAutoFilter(){state.filter=true}}
 const af={GetRange(){return state.filter?range:null},GetFilterMode(){return state.mode},GetFilters(){return state.filters},ApplyFilter(c,x){state.mode=true;state.filters=[{column:c,criteria:x}]},ShowAllData(){state.mode=false;state.filters=[]}}
 return {state,api:{GetSheet(n){return n==='Sheet1'?{GetRange(){return range},GetAutoFilter(){return af}}:null}}}
}
function withApi(f,fn){const old=global.Api;global.Api=f.api;try{return fn()}finally{if(old===undefined)delete global.Api;else global.Api=old}}
check('sort returns PASS only with before/after readback',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'sort.apply',sheet:'Sheet1',range:'A1:B3',key:1,order:'asc',hasHeaders:true}));assert.equal(r.verification.status,'PASS');assert.deepStrictEqual(r.after,[['Name','Score'],['A',1],['B',2]])})
check('filter inspect reports live state',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'filter.inspect',sheet:'Sheet1'}));assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.present,false)})
check('filter enable verifies live range exists',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'filter.enable',sheet:'Sheet1',range:'A1:B3'}));assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.present,true)})
check('filter apply verifies readable filter state',()=>{const f=fixture();f.state.filter=true;const r=withApi(f,()=>sortFilterCommand({type:'filter.apply',sheet:'Sheet1',column:1,criteria:'A'}));assert.equal(r.verification.status,'PASS');assert.equal(r.filterMode,true)})
check('filter clear verifies filter mode false',()=>{const f=fixture();f.state.filter=true;f.state.mode=true;const r=withApi(f,()=>sortFilterCommand({type:'filter.clear',sheet:'Sheet1'}));assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual,false)})
check('invalid sort order fails closed',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'sort.apply',sheet:'Sheet1',range:'A1:B3',order:'sideways'}));assert.equal(r.outcome,'invalid-operation')})
check('unknown operation fails closed',()=>{const f=fixture();const r=withApi(f,()=>sortFilterCommand({type:'wat',sheet:'Sheet1'}));assert.equal(r.outcome,'invalid-operation')})
console.log('\n'+(failed===0?'MIND OK':failed+' FAILED'));process.exit(failed?1:0)
