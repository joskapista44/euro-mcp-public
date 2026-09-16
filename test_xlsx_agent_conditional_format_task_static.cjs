'use strict'
const assert=require('assert'),m=require('./xlsx-agent-conditional-format-task.cjs')
const task=rule=>({operations:[{intent:'add_conditional_format',sheet:'S',range:'A1:A3',rule}]})
const rule={type:'xlCellValue',operator:'xlGreater',formula1:'10',fillColor:[255,0,0],priority:1}
assert.equal(m.planTask(task(rule)).ok,true)
assert.equal(m.planTask(task({...rule,fillColor:[999,0,0]})).outcome,'xlsx-cf-task-invalid-rule')
assert.equal(m.matchRule({type:'xlCellValue',operator:'xlGreater',formula1:'10',fillColor:0xff0000,priority:1},rule),true)
assert.equal(m.matchRule({type:'xlCellValue',operator:'xlGreater',formula1:'11',fillColor:0xff0000,priority:1},rule),false)
function inv(){return {ok:true,authority:'LIVE_READ',sheets:[{name:'S'}]}}
;(async()=>{
 let mutations=0
 const already=await m.executeConditionalFormatTask({task:task(rule),api:{inspect:async()=>inv(),cfObserved:async spec=>spec.type==='cf.inspect'?{ok:true,authority:'LIVE_READ',count:1,rules:[{index:0,type:'xlCellValue',operator:'xlGreater',formula1:'10',fillColor:0xff0000,priority:1}]}:(mutations++,{ok:true,authority:'LIVE_VERIFY'})}})
 assert.equal(already.ok,true);assert.equal(already.noOp,true);assert.equal(mutations,0)
 const absentDelete=await m.executeConditionalFormatTask({task:{operations:[{intent:'delete_conditional_format',sheet:'S',range:'A1:A3',rule}]},api:{inspect:async()=>inv(),cfObserved:async()=>({ok:true,authority:'LIVE_READ',count:0,rules:[]})}})
 assert.equal(absentDelete.ok,true);assert.equal(absentDelete.noOp,true)
 const ambiguous=await m.executeConditionalFormatTask({task:task(rule),api:{inspect:async()=>inv(),cfObserved:async()=>({ok:true,authority:'LIVE_READ',count:2,rules:[{index:0,...rule,fillColor:0xff0000},{index:1,...rule,fillColor:0xff0000}]})}})
 assert.equal(ambiguous.ok,false);assert.equal(ambiguous.outcome,'xlsx-cf-task-rule-identity-ambiguous')
 console.log('XLSX AGENT CONDITIONAL FORMAT TASK STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
