'use strict'
const assert=require('assert')
const m=require('./xlsx-agent-freeze-task.cjs')
function inv(names){return {ok:true,authority:'LIVE_READ',sheets:names.map((name,index)=>({name,index}))}}
function observed({match=true,noOp=true,applied=false,measurable=true}={}){return {ok:measurable&&match,source:'live-coedit-editor',noOp,applied,verification:{measurable,match,actual:match?{r1:0,c1:0,r2:1,c2:16383}:null}}}
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'rows',count:2}]}).ok,true)
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'columns',count:2}]}).ok,true)
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'at',range:'C4'}]}).operation.range,'C4')
assert.equal(m.planTask({operations:[{intent:'unfreeze_panes',sheet:'S'}]}).ok,true)
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'rows',count:0}]}).outcome,'xlsx-freeze-task-invalid-count')
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'at',range:'bad'}]}).outcome,'xlsx-freeze-task-invalid-range')
;(async()=>{
  let calls=0
  const noOp=await m.executeFreezeTask({task:{operations:[{intent:'freeze_panes',sheet:'S',mode:'rows',count:2}]},api:{inspect:async()=>inv(['S']),freezeObserved:async(op,apply)=>{calls++;return observed({noOp:!apply,applied:apply})}}})
  assert.equal(noOp.ok,true);assert.equal(noOp.noOp,true);assert.equal(calls,1)
  const seq=[observed({match:false,noOp:false}),observed({match:true,noOp:false,applied:true}),observed({match:true,noOp:true})]
  const applied=await m.executeFreezeTask({task:{operations:[{intent:'freeze_panes',sheet:'S',mode:'at',range:'C4'}]},api:{inspect:async()=>inv(['S']),freezeObserved:async()=>seq.shift()}})
  assert.equal(applied.ok,true);assert.equal(applied.authority,'LIVE_VERIFY');assert.equal(applied.noOp,false)
  const unknown=await m.executeFreezeTask({task:{operations:[{intent:'unfreeze_panes',sheet:'S'}]},api:{inspect:async()=>inv(['S']),freezeObserved:async()=>observed({measurable:false})}})
  assert.equal(unknown.outcome,'xlsx-freeze-task-precheck-unverifiable')
  console.log('XLSX AGENT FREEZE TASK STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
