'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-freeze-task.cjs')
const {freezeObserveCommand}=require('./xlsx-freeze-observer.cjs')
function optionsOf({fileId,credentials,timeoutMs=30000,pollMs=50}={}){if(!credentials?.url||!credentials?.user||!credentials?.pass)return null;return {url:credentials.url,user:credentials.user,pass:credentials.pass,fileId,timeoutMs,pollMs}}
async function runCommand(session,operation,apply){
  const body=`return (${freezeObserveCommand.toString()}).apply(null, ${JSON.stringify([operation.sheet,operation,!!apply])});`
  return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};const timer=setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),5000);try{e.callCommand(new Function(body),false,v=>{clearTimeout(timer);finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v)})}catch(err){clearTimeout(timer);finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message||err)})}}),{where:session.apiWhere,body})
}
async function executeFreezeTaskInPersistentSession(options={}){
  const sessionOptions=optionsOf(options)
  if(!sessionOptions)return {ok:false,outcome:'xlsx-persistent-credentials-required',authority:'PLAN_ONLY',writeAllowed:false}
  return persistent.withPersistentXlsxSession(sessionOptions,api=>agent.executeFreezeTask({task:options.task,api:{...api,freezeObserved:async(operation,apply)=>{const result=await runCommand(api.session,operation,apply);if(apply&&result?.ok&&result.applied===true)api.session.markWrite();return result}}}))
}
module.exports={runCommand,executeFreezeTaskInPersistentSession}
