'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const agentTask=require('./xlsx-agent-task.cjs')
const classifier=require('./xlsx-copy-chain-retry-classifier.cjs')
function optionsOf({fileId,credentials,timeoutMs=30000,pollMs=50}={}){if(!credentials?.url||!credentials?.user||!credentials?.pass)return null;return {url:credentials.url,user:credentials.user,pass:credentials.pass,fileId,timeoutMs,pollMs}}
async function executeCopyChainTaskInPersistentSession(options={}){
 const sessionOptions=optionsOf(options);if(!sessionOptions)return {ok:false,outcome:'xlsx-persistent-credentials-required',authority:'PLAN_ONLY',writeAllowed:false}
 const plan=agentTask.planTask(options.task);if(!plan.ok)return plan
 return persistent.withPersistentXlsxSession(sessionOptions,async api=>{
   const initial=await api.inspect();if(!initial?.ok||initial.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-copy-chain-retry-initial-live-inventory-failed',authority:'PLAN_ONLY',writeAllowed:false}
   const retry=await classifier.classifyCopyChainRetry({plan,api,inventory:initial})
   if(retry.matched)return {...retry,plan,freshBoundaryCount:1}
   return agentTask.executeTask({task:options.task,api})
 })
}
module.exports={executeCopyChainTaskInPersistentSession}
