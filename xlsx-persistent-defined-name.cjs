'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-defined-name-task.cjs')
const {definedNameObserveCommand}=require('./xlsx-defined-name-observer.cjs')
function optionsOf({fileId,credentials,timeoutMs=30000,pollMs=50}={}){if(!credentials?.url||!credentials?.user||!credentials?.pass)return null;return {url:credentials.url,user:credentials.user,pass:credentials.pass,fileId,timeoutMs,pollMs}}
async function runImmediateCommand(session,spec,apply){const body=`return (${definedNameObserveCommand.toString()})(${JSON.stringify({...spec,apply:!!apply})});`;return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};const timer=setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),5000);try{e.callCommand(new Function(body),false,v=>{clearTimeout(timer);finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v)})}catch(err){clearTimeout(timer);finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message||err)})}}),{where:session.apiWhere,body})}
async function runCommand(session,spec,apply,{timeoutMs=15000,pollMs=25}={}){
 const immediate=await runImmediateCommand(session,spec,apply)
 // In a co-editing workbook AddDefName acquires a collaboration lock
 // asynchronously. callCommand can therefore return after dispatch but before
 // the new name is visible to GetDefName. Dispatch exactly once, then poll only
 // the public semantic getter in the same editor session.
 const pendingAdd=apply&&spec?.intent==='set_defined_name'&&immediate?.applied===true&&immediate?.outcome==='defined-name-semantic-mismatch'&&immediate?.state?.measurable===true&&immediate.state.present===false
 if(!pendingAdd)return immediate
 const started=Date.now();let observed=null,reads=0
 while(Date.now()-started<=timeoutMs){
  observed=await runImmediateCommand(session,spec,false);reads++
  if(observed?.ok&&observed?.noOp===true&&observed?.verification?.match===true)return {...immediate,ok:true,outcome:'defined-name-live-verified',noOp:false,applied:true,state:observed.state,verification:{measurable:true,match:true},collaborationReadback:{outcome:'visible-after-collaboration-lock',reads,waitMs:Date.now()-started}}
  if(observed?.state?.measurable===true&&observed.state.present===true)return {...immediate,collaborationReadback:{outcome:'visible-with-semantic-mismatch',reads,waitMs:Date.now()-started,state:observed.state}}
  if(Date.now()-started>=timeoutMs)break
  await new Promise(resolve=>setTimeout(resolve,pollMs))
 }
 return {...immediate,collaborationReadback:{outcome:'collaboration-lock-timeout',reads,waitMs:Date.now()-started,state:observed?.state||null}}
}
async function executeDefinedNameTaskInPersistentSession(options={}){const sessionOptions=optionsOf(options);if(!sessionOptions)return {ok:false,outcome:'xlsx-persistent-credentials-required',authority:'PLAN_ONLY',writeAllowed:false};return persistent.withPersistentXlsxSession(sessionOptions,api=>agent.executeDefinedNameTask({task:options.task,api:{...api,definedNameObserved:async(spec,apply)=>{const r=await runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}))}
module.exports={runImmediateCommand,runCommand,executeDefinedNameTaskInPersistentSession}
