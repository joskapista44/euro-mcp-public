'use strict'

function clone(v){return v===undefined?undefined:JSON.parse(JSON.stringify(v))}
function validOp(op){return op&&op.intent==='write_range'&&typeof op.sheet==='string'&&typeof op.range==='string'&&Array.isArray(op.values)}

function planTask(task){
  if(!task||!Array.isArray(task.operations)||!task.operations.length)return {ok:false,outcome:'xlsx-task-operations-required',authority:'PLAN_ONLY',writeAllowed:false}
  const operations=[]
  for(let i=0;i<task.operations.length;i++){
    const op=task.operations[i]
    if(!validOp(op))return {ok:false,outcome:'xlsx-task-intent-not-supported',authority:'PLAN_ONLY',writeAllowed:false,index:i,intent:op?.intent||null}
    operations.push({index:i,intent:'write_range',sheet:op.sheet,range:op.range,values:clone(op.values),formulas:clone(op.formulas||null)})
  }
  return {ok:true,outcome:'xlsx-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operations,constraints:clone(task.constraints||{})}
}

function sheetFromInventory(inventory,name){
  const hits=(inventory?.sheets||[]).filter(s=>s?.name===name)
  if(hits.length!==1)return {ok:false,outcome:hits.length?'xlsx-target-sheet-ambiguous':'xlsx-target-sheet-not-found',count:hits.length}
  return {ok:true,sheet:hits[0]}
}

function sameExpectedCell(cell,value,formula){
  if(formula!==null&&formula!==undefined)return cell?.formula===formula
  if(value===null||value===undefined)return cell?.dataType==='blank'||((cell?.value===null||cell?.value===undefined||cell?.value==='')&&!cell?.formula)
  return Object.is(cell?.rawValue,value)||Object.is(cell?.value,value)||String(cell?.displayText)===String(value)
}

function rangeAlreadySatisfied(read,op){
  if(!read?.ok||!Array.isArray(read.cells)||read.cells.length!==op.values.length)return false
  for(let r=0;r<op.values.length;r++){
    if(!Array.isArray(read.cells[r])||read.cells[r].length!==op.values[r].length)return false
    for(let c=0;c<op.values[r].length;c++)if(!sameExpectedCell(read.cells[r][c],op.values[r][c],op.formulas?.[r]?.[c]??null))return false
  }
  return true
}

async function executeTask({task,api}){
  const plan=planTask(task)
  if(!plan.ok)return plan
  const initial=await api.inspect()
  if(!initial?.ok||initial.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-task-initial-live-inventory-failed',authority:'PLAN_ONLY',writeAllowed:false,plan,initial}
  const receipt=[]
  for(const op of plan.operations){
    const fresh=await api.inspect()
    if(!fresh?.ok||fresh.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,index:op.index,receipt}
    const identity=sheetFromInventory(fresh,op.sheet)
    if(!identity.ok)return {...identity,ok:false,authority:'PLAN_ONLY',writeAllowed:false,index:op.index,receipt}
    const observed=await api.readRange({sheet:op.sheet,range:op.range})
    if(!observed?.ok||observed.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-task-target-read-failed',authority:'PLAN_ONLY',writeAllowed:false,index:op.index,receipt}
    if(rangeAlreadySatisfied(observed,op)){
      receipt.push({index:op.index,intent:op.intent,status:'NO_OP_SATISFIED',resolvedTarget:{sheet:op.sheet,range:op.range},verification:'LIVE_VERIFY'})
      continue
    }
    const applied=await api.writeRangeVerified({sheet:op.sheet,range:op.range,values:op.values,formulas:op.formulas})
    if(!applied?.ok||applied.authority!=='LIVE_VERIFY')return {ok:false,outcome:'xlsx-task-primitive-failed',authority:applied?.authority||'DISPATCH_ONLY',writeAllowed:false,index:op.index,receipt,result:applied||null}
    receipt.push({index:op.index,intent:op.intent,status:'APPLIED',resolvedTarget:{sheet:op.sheet,range:op.range},verification:'LIVE_VERIFY'})
  }
  const finalInventory=await api.inspect()
  if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-task-final-inventory-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,receipt}
  const checks=[]
  for(const op of plan.operations){
    const identity=sheetFromInventory(finalInventory,op.sheet)
    if(!identity.ok)return {ok:false,outcome:'xlsx-task-final-target-identity-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,receipt,index:op.index}
    const read=await api.readRange({sheet:op.sheet,range:op.range})
    const pass=read?.authority==='LIVE_READ'&&rangeAlreadySatisfied(read,op)
    checks.push({index:op.index,intent:op.intent,sheet:op.sheet,range:op.range,status:pass?'PASS':'FAIL'})
    if(!pass)return {ok:false,outcome:'xlsx-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,receipt,wholeTaskVerification:{ok:false,checks}}
  }
  const noOp=receipt.every(x=>x.status==='NO_OP_SATISFIED')
  return {ok:true,outcome:noOp?'xlsx-task-already-satisfied':'xlsx-task-applied-and-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp,plan,receipt,wholeTaskVerification:{ok:true,outcome:'verified',authority:'LIVE_VERIFY',checks},freshBoundaryCount:plan.operations.length+2}
}

module.exports={planTask,sheetFromInventory,rangeAlreadySatisfied,executeTask}
