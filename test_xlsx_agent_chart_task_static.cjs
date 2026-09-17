'use strict'
const assert=require('assert')
const agent=require('./xlsx-agent-chart-task.cjs')
const persistent=require('./xlsx-persistent-chart.cjs')
const sheetInventory={ok:true,authority:'LIVE_READ',sheets:[{name:'Sheet1'}]}
function task(extra={}){return {operations:[{intent:'set_chart',sheet:'Sheet1',name:'Revenue Chart',range:'A1:B3',chartType:'bar',title:'Revenue',width:3600000,height:2160000,expectedSeriesCount:1,...extra}]}}
function response(state,match,noOp=match){return {ok:true,source:'live-coedit-editor',noOp,state,applied:noOp?undefined:true,verification:{measurable:true,match}}}
async function run(name,fn){await fn();console.log('OK',name)}
;(async()=>{
 await run('planner canonicalizes exact measurable chart state',async()=>{const p=agent.planTask(task());assert.equal(p.ok,true);assert.equal(p.operation.title,'Revenue');assert.equal(p.operation.expectedSeriesCount,1)})
 await run('planner rejects missing semantic series count',async()=>{assert.equal(agent.planTask(task({expectedSeriesCount:undefined})).ok,false)})
 await run('state matcher rejects duplicate names',async()=>{const s=persistent.stateOf({charts:[{name:'C'},{name:'C'}]},{name:'C'});assert.equal(s.count,2);assert.equal(persistent.match(s,{intent:'delete_chart'}),false)})
 await run('executor applies then whole-verifies',async()=>{let n=0;const absent={measurable:true,present:false,count:0,name:'Revenue Chart'},present={measurable:true,present:true,count:1,name:'Revenue Chart',chartType:'bar',title:'Revenue',width:3600000,height:2160000,seriesCount:1};const r=await agent.executeChartTask({task:task(),api:{inspect:async()=>sheetInventory,chartObserved:async(_op,apply)=>{n++;return apply?response(present,true,false):n<3?response(absent,false,false):response(present,true,true)}}});assert.equal(r.ok,true);assert.equal(r.noOp,false)})
 await run('executor recognizes exact no-op',async()=>{const present={measurable:true,present:true,count:1,name:'Revenue Chart',chartType:'bar',title:'Revenue',width:3600000,height:2160000,seriesCount:1};const r=await agent.executeChartTask({task:task(),api:{inspect:async()=>sheetInventory,chartObserved:async()=>response(present,true,true)}});assert.equal(r.ok,true);assert.equal(r.noOp,true)})
 await run('executor fails closed on changed mutation boundary',async()=>{let n=0;const a={measurable:true,present:false,count:0,name:'Revenue Chart'},b={...a,count:2,present:true};const r=await agent.executeChartTask({task:task(),api:{inspect:async()=>sheetInventory,chartObserved:async()=>response(++n===1?a:b,false,false)}});assert.equal(r.ok,false);assert.equal(r.outcome,'xlsx-chart-task-state-changed-before-mutation')})
 console.log('XLSX CHART STATIC: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
