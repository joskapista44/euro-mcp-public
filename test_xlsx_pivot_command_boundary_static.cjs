'use strict'
const assert=require('assert/strict')
const Module=require('module')
const originalLoad=Module._load
let calls=[],bodies=[]
Module._load=function(request,parent,isMain){
 if(request==='./xlsx-persistent-session.cjs')return {}
 if(request==='./xlsx-agent-pivot-task.cjs')return {}
 if(request==='./xlsx-pivot-observer.cjs')return {pivotObserveCommand:function pivotObserveCommand(spec){return spec}}
 return originalLoad.call(this,request,parent,isMain)
}
const pivot=require('./xlsx-persistent-pivot.cjs')
Module._load=originalLoad
function session(results){
 let i=0
 return {apiWhere:'window.editor',frame:{evaluate:async arg=>{calls.push(i);bodies.push(arg.body);return results[i++]}}}
}
;(async()=>{
 const spec={intent:'create_pivot',name:'P'}
 calls=[]
 let s=session([{ok:true,outcome:'pivot-live-verified',applied:true,verification:{measurable:true,match:true}}])
 let r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,true);assert.equal(calls.length,1)

 calls=[]
 s=session([
  {ok:false,outcome:'pivot-semantic-mismatch',applied:true,verification:{measurable:true,match:false},state:{present:false}},
  {ok:true,outcome:'pivot-already-satisfied',noOp:true,verification:{measurable:true,match:true},state:{present:true}}
 ])
 r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,true);assert.equal(r.applied,true);assert.equal(r.outcome,'pivot-live-verified-after-command-boundary');assert.equal(calls.length,2);assert.match(bodies[0],/\\"apply\\":true/);assert.match(bodies[1],/\\"apply\\":false/)
 assert.equal(r.postMutationObservation.noOp,true)

 calls=[]
 s=session([
  {ok:false,outcome:'pivot-semantic-mismatch',applied:true,verification:{measurable:true,match:false},state:{present:false}},
  {ok:true,outcome:'pivot-observed',noOp:false,verification:{measurable:true,match:false},state:{present:false}}
 ])
 r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,false);assert.equal(r.outcome,'pivot-semantic-mismatch');assert.equal(calls.length,2)
 assert.equal(r.postMutationObservation.state.present,false)

 calls=[]
 s=session([{ok:false,outcome:'pivot-operation-error',applied:false}])
 r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,false);assert.equal(calls.length,1)

 calls=[]
 s=session([{ok:false,outcome:'pivot-semantic-mismatch',applied:true},{ok:true,noOp:true,verification:{measurable:true,match:true}}])
 r=await pivot.runCommand(s,spec,false)
 assert.equal(r.ok,false);assert.equal(calls.length,1)
 console.log('XLSX PIVOT COMMAND BOUNDARY STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
