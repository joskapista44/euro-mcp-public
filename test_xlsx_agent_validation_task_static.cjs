'use strict'
const assert=require('assert')
const {validationObserveCommand}=require('./xlsx-validation-observer.cjs')
const agent=require('./xlsx-agent-validation-task.cjs')
let failed=0
async function check(name,fn){try{await fn();console.log('OK '+name)}catch(e){failed++;console.error('FAIL '+name+'\n'+(e?.stack||e))}}
function fixture(){
 const s={type:'xlValidateInputOnly',alertStyle:null,operator:null,formula1:'',formula2:'',ignoreBlank:true,inCellDropdown:true,inputMessage:'',inputTitle:'',showError:true,showInput:true,errorMessage:'',errorTitle:'',writes:0}
 const v={Add(t,a,o,f1,f2){s.writes++;Object.assign(s,{type:t,alertStyle:a,operator:o,formula1:String(f1),formula2:f2==null?'':String(f2)})},Modify(t,a,o,f1,f2){s.writes++;Object.assign(s,{type:t,alertStyle:a,operator:o,formula1:String(f1),formula2:f2==null?'':String(f2)})},Delete(){s.writes++;Object.assign(s,{type:'xlValidateInputOnly',alertStyle:null,operator:null,formula1:'',formula2:''})}}
 for(const k of ['Type','AlertStyle','Operator','Formula1','Formula2','IgnoreBlank','InCellDropdown','InputMessage','InputTitle','ShowError','ShowInput','ErrorMessage','ErrorTitle'])v['Get'+k]=()=>s[k.charAt(0).toLowerCase()+k.slice(1)]
 for(const k of ['IgnoreBlank','InCellDropdown','InputMessage','InputTitle','ShowError','ShowInput','ErrorMessage','ErrorTitle'])v['Set'+k]=x=>{s[k.charAt(0).toLowerCase()+k.slice(1)]=x}
 const range={GetAddress(){return '$B$2:$B$5'},GetValidation(){return v}}
 return {s,api:{GetSheet:n=>n==='S'?{GetRange:()=>range}:null}}
}
function setSpec(extra={}){return {intent:'set_validation',sheet:'S',range:'B2:B5',validationType:'xlValidateWholeNumber',alertStyle:'xlValidAlertStop',operator:'xlBetween',formula1:'1',formula2:'10',...extra}}
function withApi(f,fn){const old=global.Api;global.Api=f.api;try{return fn()}finally{if(old===undefined)delete global.Api;else global.Api=old}}
;(async()=>{
 await check('observer adds and exactly verifies durable validation core',async()=>{const f=fixture(),r=withApi(f,()=>validationObserveCommand({...setSpec(),apply:true}));assert.equal(r.ok,true);assert.equal(r.applied,true);assert.equal(r.state.type,'xlValidateWholeNumber');assert.equal(r.state.formula2,'10');assert.equal(f.s.writes,1)})
 await check('observer recognizes exact set no-op',async()=>{const f=fixture();withApi(f,()=>validationObserveCommand({...setSpec(),apply:true}));const r=withApi(f,()=>validationObserveCommand({...setSpec(),apply:false}));assert.equal(r.noOp,true);assert.equal(r.verification.match,true)})
 await check('observer modifies mismatching validation',async()=>{const f=fixture();withApi(f,()=>validationObserveCommand({...setSpec(),apply:true}));const r=withApi(f,()=>validationObserveCommand({...setSpec({formula1:'2',formula2:'20',alertStyle:'xlValidAlertWarning'}),apply:true}));assert.equal(r.ok,true);assert.equal(r.state.formula1,'2');assert.equal(r.state.alertStyle,'xlValidAlertWarning');assert.equal(f.s.writes,2)})
 await check('observer clears and verifies absent state',async()=>{const f=fixture();withApi(f,()=>validationObserveCommand({...setSpec(),apply:true}));const r=withApi(f,()=>validationObserveCommand({intent:'clear_validation',sheet:'S',range:'B2:B5',apply:true}));assert.equal(r.ok,true);assert.equal(r.state.absent,true);assert.equal(f.s.writes,2)})
 await check('observer recognizes clear no-op',async()=>{const f=fixture(),r=withApi(f,()=>validationObserveCommand({intent:'clear_validation',sheet:'S',range:'B2:B5',apply:false}));assert.equal(r.noOp,true)})
 await check('observer fails closed when a semantic getter is missing',async()=>{const f=fixture();const range=f.api.GetSheet('S').GetRange();delete range.GetValidation().GetErrorTitle;const r=withApi(f,()=>validationObserveCommand({...setSpec(),apply:false}));assert.equal(r.outcome,'validation-state-unverifiable')})
 await check('planner rejects invalid type and malformed range',async()=>{assert.equal(agent.planTask({operations:[{...setSpec(),validationType:'wat'}]}).ok,false);assert.equal(agent.planTask({operations:[{...setSpec(),range:'B5:B2'}]}).ok,false)})
 await check('planner rejects non-persistable validation options',async()=>{const p=agent.planTask({operations:[{...setSpec(),inputMessage:'Enter value'}]});assert.equal(p.ok,false);assert.equal(p.outcome,'xlsx-validation-task-option-not-persistable')})
 await check('planner canonicalizes durable deterministic state',async()=>{const p=agent.planTask({operations:[{intent:'set_validation',sheet:'S',range:'B2:B5',validationType:'xlValidateWholeNumber',formula1:1,formula2:10}]});assert.equal(p.ok,true);assert.equal(p.operation.formula1,'1');assert.equal(Object.hasOwn(p.operation,'inputMessage'),false)})
 await check('executor applies then whole-verifies',async()=>{let phase=0,writes=0;const absent={ok:true,source:'live-coedit-editor',noOp:false,state:{measurable:true,address:'B2:B5',absent:true,present:false},verification:{measurable:true,match:false}},set={ok:true,source:'live-coedit-editor',noOp:true,state:{measurable:true,address:'B2:B5',absent:false,present:true},verification:{measurable:true,match:true}};const api={inspect:async()=>({ok:true,authority:'LIVE_READ',sheets:[{name:'S'}]}),validationObserved:async(_op,apply)=>{if(apply){phase=1;writes++;return {...set,noOp:false,applied:true}}return phase?set:absent}};const r=await agent.executeValidationTask({task:{operations:[{intent:'set_validation',sheet:'S',range:'B2:B5',validationType:'xlValidateWholeNumber',formula1:'1',formula2:'10'}]},api});assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(writes,1)})
 console.log(failed?'XLSX VALIDATION STATIC: FAIL':'XLSX VALIDATION STATIC: PASS');process.exitCode=failed?1:0
})()
