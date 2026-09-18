'use strict'
const assert=require('assert')
const m=require('./xlsx-agent-freeze-task.cjs')
const {freezeObserveCommand}=require('./xlsx-freeze-observer.cjs')
function inv(names){return {ok:true,authority:'LIVE_READ',sheets:names.map((name,index)=>({name,index}))}}
function observed({match=true,noOp=true,applied=false,measurable=true}={}){return {ok:measurable&&match,source:'live-coedit-editor',noOp,applied,verification:{measurable,match,actual:match?{r1:0,c1:0,r2:1,c2:16383}:null}}}
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'rows',count:2}]}).ok,true)
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'columns',count:2}]}).ok,true)
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'at',range:'C4'}]}).operation.range,'C4')
assert.equal(m.planTask({operations:[{intent:'unfreeze_panes',sheet:'S'}]}).ok,true)
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'rows',count:0}]}).outcome,'xlsx-freeze-task-invalid-count')
assert.equal(m.planTask({operations:[{intent:'freeze_panes',sheet:'S',mode:'at',range:'bad'}]}).outcome,'xlsx-freeze-task-invalid-range')
{
  let active='Data'
  const location=box=>({range:{bbox:box},GetAddress:()=>box.r2===3?'A1:A4':'A1:A2'})
  const panes={Plan:{GetLocation:()=>location({r1:0,c1:0,r2:3,c2:0})},Data:{GetLocation:()=>location({r1:0,c1:0,r2:1,c2:0})}}
  const sheets={}
  for(const name of ['Plan','Data'])sheets[name]={GetName:()=>name,SetActive:()=>{active=name},GetRange:address=>({range:{bbox:{r1:Number(address.slice(1))-1,c1:0,r2:Number(address.slice(1))-1,c2:0}}}),GetFreezePanes:()=>panes[active]}
  const oldApi=global.Api,oldAsc=global.AscCommon;global.Api={GetSheet:name=>sheets[name],GetActiveSheet:()=>sheets[active]};global.AscCommon={gc_nMaxCol0:16383,gc_nMaxRow0:1048575}
  try{const r=freezeObserveCommand('Plan',{intent:'freeze_panes',mode:'at',range:'A4'},false);assert.equal(r.ok,true);assert.equal(r.noOp,true);assert.equal(active,'Plan')}finally{if(oldApi===undefined)delete global.Api;else global.Api=oldApi;if(oldAsc===undefined)delete global.AscCommon;else global.AscCommon=oldAsc}
}
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
