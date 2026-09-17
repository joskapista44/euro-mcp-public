'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-chart-task.cjs')
const {runChartInFrame}=require('./live-charts.cjs')
function stateOf(observed,op){
 const matches=(observed?.charts||[]).filter(c=>String(c?.name)===op.name)
 if(matches.length!==1)return {measurable:true,present:matches.length>0,count:matches.length,name:op.name}
 const c=matches[0]
 return {measurable:true,present:true,count:1,name:c.name,chartType:c.chartType,title:c.title==null?null:String(c.title).replace(/[\r\n]+$/g,''),width:c.width,height:c.height,seriesCount:c.seriesCount}
}
function match(state,op){
 if(op.intent==='delete_chart')return state.measurable&&state.present===false&&state.count===0
 return state.measurable&&state.present&&state.count===1&&String(state.chartType)===op.chartType&&state.title===op.title&&Number(state.width)===op.width&&Number(state.height)===op.height&&Number(state.seriesCount)===op.expectedSeriesCount
}
async function chartObserved(session,op,apply){
 const inspect=()=>runChartInFrame(session.frame,session.apiWhere,{type:'chart.inspect',sheet:op.sheet},8000)
 const beforeRead=await inspect()
 if(!beforeRead?.ok||beforeRead?.verification?.status!=='PASS')return {ok:false,outcome:'chart-state-unverifiable',source:'live-coedit-editor',readback:beforeRead}
 const before=stateOf(beforeRead,op),satisfied=match(before,op)
 if(!apply)return {ok:true,outcome:satisfied?'chart-already-satisfied':'chart-observed',source:'live-coedit-editor',noOp:satisfied,state:before,verification:{measurable:true,match:satisfied}}
 if(satisfied)return {ok:true,outcome:'chart-already-satisfied',source:'live-coedit-editor',noOp:true,state:before,verification:{measurable:true,match:true}}
 let command
 if(op.intent==='delete_chart'){
  if(before.count!==1)return {ok:false,outcome:'chart-delete-identity-conflict',source:'live-coedit-editor',state:before,verification:{measurable:true,match:false}}
  command={type:'chart.delete',sheet:op.sheet,name:op.name}
 }else if(before.count===0){
  command={type:'chart.create',sheet:op.sheet,range:op.range,chartType:op.chartType,name:op.name,title:op.title,width:op.width,height:op.height,inRows:op.inRows}
 }else if(before.count===1&&String(before.chartType)===op.chartType&&Number(before.seriesCount)===op.expectedSeriesCount){
  command={type:'chart.modify',sheet:op.sheet,name:op.name,title:op.title,width:op.width,height:op.height}
 }else return {ok:false,outcome:'chart-set-identity-conflict',source:'live-coedit-editor',state:before,verification:{measurable:true,match:false}}
 const changed=await runChartInFrame(session.frame,session.apiWhere,command,8000)
 if(!changed?.ok||changed?.verification?.status!=='PASS')return {ok:false,outcome:'chart-mutation-unverified',source:'live-coedit-editor',state:before,mutation:changed,verification:{measurable:true,match:false}}
 const afterRead=await inspect()
 if(!afterRead?.ok||afterRead?.verification?.status!=='PASS')return {ok:false,outcome:'chart-post-state-unverifiable',source:'live-coedit-editor',mutation:changed,readback:afterRead}
 const after=stateOf(afterRead,op),pass=match(after,op)
 return {ok:pass,outcome:pass?'chart-live-verified':'chart-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before,state:after,mutation:changed,verification:{measurable:true,match:pass}}
}
async function executeChartTaskInPersistentSession(options={}){
 if(!options.url||!options.user||!options.pass)return {ok:false,outcome:'xlsx-persistent-credentials-required',authority:'PLAN_ONLY',writeAllowed:false}
 return persistent.withPersistentXlsxSession(options,api=>agent.executeChartTask({task:options.task,api:{...api,chartObserved:async(op,apply)=>{const r=await chartObserved(api.session,op,apply);if(apply&&(r?.applied||r?.mutation?.ok))api.session.markWrite();return r}}}))
}
const runCommand=chartObserved
module.exports={stateOf,match,chartObserved,runCommand,executeChartTaskInPersistentSession}
