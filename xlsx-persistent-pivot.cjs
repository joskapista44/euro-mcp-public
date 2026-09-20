'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-pivot-task.cjs')
const {pivotObserveCommand}=require('./xlsx-pivot-observer.cjs')
function optionsOf({fileId,credentials,timeoutMs=30000,pollMs=50}={}){if(!credentials?.url||!credentials?.user||!credentials?.pass)return null;return {url:credentials.url,user:credentials.user,pass:credentials.pass,fileId,timeoutMs,pollMs}}
async function callObserved(session,spec,apply){const body=`return (${pivotObserveCommand.toString()})(${JSON.stringify({...spec,apply:!!apply})});`;return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};const timer=setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),8000);try{e.callCommand(new Function(body),false,v=>{clearTimeout(timer);finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v)})}catch(err){clearTimeout(timer);finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message||err)})}}),{where:session.apiWhere,body})}
async function runCommand(session,spec,apply){
 const first=await callObserved(session,spec,apply)
 if(!apply||first?.ok||first?.outcome!=='pivot-semantic-mismatch'||first?.applied!==true)return first
 // Pivot insertion can become observable only after the mutation callCommand
 // returns. Cross one fresh editor-command boundary in the SAME session and
 // perform read-only semantic observation; never sleep or redispatch mutation.
 const observed=await callObserved(session,spec,false)
 if(observed?.ok&&observed?.verification?.measurable===true&&observed?.verification?.match===true&&observed?.noOp===true)return {...observed,outcome:'pivot-live-verified-after-command-boundary',noOp:false,applied:true,mutation:first,postMutationObservation:observed}
 // In a large co-edit batch the first read boundary can expose a transient
 // public pivot object whose getters are not hydrated yet (measured as an
 // observer error such as null.map). Cross ONE additional read-only command
 // boundary. Never sleep and never redispatch the mutation.
 const transientError=observed?.outcome==='pivot-operation-error'||observed?.outcome==='pivot-state-unverifiable'
 const stillAbsent=observed?.ok===true&&observed?.outcome==='pivot-observed'&&observed?.verification?.measurable===true&&observed?.verification?.match===false&&observed?.state?.present===false
 if(transientError||stillAbsent){
  const observed2=await callObserved(session,spec,false)
  if(observed2?.ok&&observed2?.verification?.measurable===true&&observed2?.verification?.match===true&&observed2?.noOp===true)return {...observed2,outcome:'pivot-live-verified-after-second-command-boundary',noOp:false,applied:true,mutation:first,postMutationObservation:observed,secondPostMutationObservation:observed2}
  return {...first,postMutationObservation:observed,secondPostMutationObservation:observed2}
 }
 return {...first,postMutationObservation:observed}
}
async function executePivotTaskInPersistentSession(options={}){const sessionOptions=optionsOf(options);if(!sessionOptions)return {ok:false,outcome:'xlsx-persistent-credentials-required',authority:'PLAN_ONLY',writeAllowed:false};return persistent.withPersistentXlsxSession(sessionOptions,api=>agent.executePivotTask({task:options.task,api:{...api,pivotObserved:async(spec,apply)=>{const r=await runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}))}
module.exports={callObserved,runCommand,executePivotTaskInPersistentSession}
