'use strict'
const assert=require('assert')
const {filterObserveCommand}=require('./xlsx-filter-observer.cjs')
const agent=require('./xlsx-agent-filter-task.cjs')
let failed=0
async function check(name,fn){try{await fn();console.log('OK '+name)}catch(e){failed++;console.error('FAIL '+name+'\n'+(e?.stack||e))}}
function fixture(){
 const state={range:null,filters:[],writes:0}
 const range={GetAddress(){return '$A$1:$C$5'},SetAutoFilter(field,c1,op,c2){state.writes++;state.range=range;state.filters=[{filter:{ColId:field-1},GetOn(){return true},GetOperator(){return op},GetCriteria1(){return '='+c1},GetCriteria2(){return c2??null}}]}}
 const af={GetRange(){return state.range},GetFilterMode(){return state.filters.length>0},GetFilters(){return state.filters}}
 return {state,api:{GetSheet:n=>n==='S'?{GetRange:()=>range,GetAutoFilter:()=>af}:null}}
}
function withApi(f,fn){const old=global.Api;global.Api=f.api;try{return fn()}finally{if(old===undefined)delete global.Api;else global.Api=old}}
;(async()=>{
 await check('observer applies exact field and criteria',async()=>{const f=fixture(),r=withApi(f,()=>filterObserveCommand({sheet:'S',range:'A1:C5',field:3,criteria1:'A',criteria2:null,operator:'xlOr',apply:true}));assert.equal(r.ok,true);assert.equal(r.state.filters[0].field,3);assert.equal(f.state.writes,1)})
 await check('observer recognizes exact no-op',async()=>{const f=fixture();withApi(f,()=>filterObserveCommand({sheet:'S',range:'A1:C5',field:3,criteria1:'A',criteria2:null,operator:'xlOr',apply:true}));const r=withApi(f,()=>filterObserveCommand({sheet:'S',range:'A1:C5',field:3,criteria1:'A',criteria2:null,operator:'xlOr',apply:false}));assert.equal(r.noOp,true)})
 await check('observer fails closed without field identity',async()=>{const f=fixture();f.state.range={GetAddress(){return '$A$1:$C$5'}};f.state.filters=[{filter:{},GetOn(){return true}}];const r=withApi(f,()=>filterObserveCommand({sheet:'S',range:'A1:C5',field:3,criteria1:'A',operator:'xlOr',apply:false}));assert.equal(r.outcome,'filter-state-unverifiable')})
 await check('planner rejects field outside range',async()=>{assert.equal(agent.planTask({operations:[{intent:'filter_range',sheet:'S',range:'A1:C5',field:4,criteria1:'A'}]}).ok,false)})
 await check('executor applies then whole-verifies',async()=>{let phase=0,writes=0;const empty={ok:true,source:'live-coedit-editor',noOp:false,state:{measurable:true,present:false,range:null,filterMode:false,filters:[]},verification:{measurable:true,match:false}},set={ok:true,source:'live-coedit-editor',noOp:true,state:{measurable:true,present:true,range:'A1:C5',filterMode:true,filters:[{field:3,operator:'xlOr',criteria1:'=A',criteria2:null,on:true}]},verification:{measurable:true,match:true}};const api={inspect:async()=>({ok:true,authority:'LIVE_READ',sheets:[{name:'S'}]}),filterObserved:async(_op,apply)=>{if(apply){phase=1;writes++;return {...set,noOp:false,applied:true}}return phase?set:empty}};const r=await agent.executeFilterTask({task:{operations:[{intent:'filter_range',sheet:'S',range:'A1:C5',field:3,criteria1:'A',operator:'xlOr'}]},api});assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(writes,1)})
 console.log(failed?'XLSX FILTER STATIC: FAIL':'XLSX FILTER STATIC: PASS');process.exitCode=failed?1:0
})()
