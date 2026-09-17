'use strict'
const assert=require('assert')
const {sortObserveCommand}=require('./xlsx-sort-observer.cjs')
const agent=require('./xlsx-agent-sort-task.cjs')
let failed=0
async function check(name,fn){try{await fn();console.log('OK '+name)}catch(e){failed++;console.error('FAIL '+name+'\n'+(e?.stack||e))}}
function fixture(values){const state={values:JSON.parse(JSON.stringify(values)),writes:0};const range={GetValue(){return state.values},SetSort(key,order){state.writes++;const h=state.values[0],rows=state.values.slice(1).sort((a,b)=>String(a[0]).localeCompare(String(b[0])));state.values=[h,...(order==='xlDescending'?rows.reverse():rows)]}};return {state,api:{GetSheet:n=>n==='S'?{GetRange:()=>range}:null}}}
function withApi(f,fn){const old=global.Api;global.Api=f.api;try{return fn()}finally{if(old===undefined)delete global.Api;else global.Api=old}}
;(async()=>{
 await check('observer detects sorted no-op',async()=>{const f=fixture([['Name'],['A'],['B']]),r=withApi(f,()=>sortObserveCommand({sheet:'S',range:'A1:A3',keyRange:'A1:A3',order:'asc',hasHeaders:true,apply:false}));assert.equal(r.noOp,true);assert.equal(f.state.writes,0)})
 await check('observer applies and verifies sort',async()=>{const f=fixture([['Name'],['B'],['A']]),r=withApi(f,()=>sortObserveCommand({sheet:'S',range:'A1:A3',keyRange:'A1:A3',order:'asc',hasHeaders:true,apply:true}));assert.equal(r.ok,true);assert.equal(r.applied,true);assert.equal(f.state.writes,1)})
 await check('planner rejects key outside range',async()=>{assert.equal(agent.planTask({operations:[{intent:'sort_range',sheet:'S',range:'A1:B3',keyRange:'C1:C3'}]}).ok,false)})
 await check('executor applies then whole-verifies',async()=>{let phase=0,writes=0;const unsorted={ok:true,source:'live-coedit-editor',noOp:false,matrix:[['Name'],['B'],['A']],verification:{measurable:true,match:false,keys:['B','A']}},sorted={ok:true,source:'live-coedit-editor',noOp:true,matrix:[['Name'],['A'],['B']],verification:{measurable:true,match:true,keys:['A','B']}};const api={inspect:async()=>({ok:true,authority:'LIVE_READ',sheets:[{name:'S'}]}),sortObserved:async(_op,apply)=>{if(apply){writes++;phase=1;return {...sorted,noOp:false,applied:true}}return phase?sorted:unsorted}};const r=await agent.executeSortTask({task:{operations:[{intent:'sort_range',sheet:'S',range:'A1:A3',keyRange:'A1:A3',order:'asc'}]},api});assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(writes,1)})
 console.log(failed?'XLSX SORT STATIC: FAIL':'XLSX SORT STATIC: PASS');process.exitCode=failed?1:0
})()
