'use strict'
const transport=require('./xlsx-persistent-page-number-order.cjs')
function planTask(task){const op=task?.operations?.[0],ok=task?.operations?.length===1&&op?.intent==='set_first_page_number'&&typeof op.sheet==='string'&&op.sheet.trim()&&(op.value===null||Number.isInteger(op.value)&&op.value>0);return ok?{ok:true,outcome:'xlsx-first-page-number-task-planned',authority:'PLAN_ONLY',operation:{index:0,intent:op.intent,sheet:op.sheet,firstPageNumber:op.value}}:{ok:false,outcome:'xlsx-first-page-number-task-invalid',authority:'PLAN_ONLY'}}
async function firstPageNumberObserved(api,op){return transport.immediate(api.session,op,false)}
async function executeFirstPageNumberTask(api,task){const p=planTask(task);if(!p.ok)return p;const pre=await firstPageNumberObserved(api,p.operation);if(pre.ok&&pre.noOp)return {...pre,outcome:'xlsx-first-page-number-task-already-satisfied',authority:'LIVE_VERIFY'};const r=await transport.immediate(api.session,p.operation,true);return {...r,outcome:r.ok?'xlsx-first-page-number-task-live-verified':r.outcome,authority:r.ok?'LIVE_VERIFY':'LIVE_READ'}}
module.exports={planTask,firstPageNumberObserved,executeFirstPageNumberTask}
