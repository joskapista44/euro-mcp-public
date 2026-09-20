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
 return {apiWhere:'window.editor',frame:{evaluate:async(_fn,arg)=>{calls.push(i);bodies.push(arg.body);return results[i++]}}}
}
;(async()=>{
 const spec={intent:'create_pivot',name:'P'}
 calls=[];bodies=[]
 let s=session([{ok:true,outcome:'pivot-live-verified',applied:true,verification:{measurable:true,match:true}}])
 let r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,true);assert.equal(calls.length,1)

 calls=[];bodies=[]
 s=session([
  {ok:false,outcome:'pivot-semantic-mismatch',applied:true,verification:{measurable:true,match:false},state:{present:false}},
  {ok:true,outcome:'pivot-already-satisfied',noOp:true,verification:{measurable:true,match:true},state:{present:true}}
 ])
 r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,true);assert.equal(r.applied,true);assert.equal(r.outcome,'pivot-live-verified-after-command-boundary');assert.equal(calls.length,2);assert.match(bodies[0],/"apply":true/);assert.match(bodies[1],/"apply":false/)
 assert.equal(r.postMutationObservation.noOp,true)

 calls=[];bodies=[]
 s=session([
  {ok:false,outcome:'pivot-semantic-mismatch',applied:true,verification:{measurable:true,match:false},state:{present:false}},
  {ok:true,outcome:'pivot-observed',noOp:false,verification:{measurable:true,match:false},state:{present:false}}
 ])
 r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,false);assert.equal(r.outcome,'pivot-semantic-mismatch');assert.equal(calls.length,3)
 assert.equal(r.postMutationObservation.state.present,false);assert.equal(r.secondPostMutationObservation,undefined)

 calls=[];bodies=[]
 s=session([
  {ok:false,outcome:'pivot-semantic-mismatch',applied:true,verification:{measurable:true,match:false},state:{present:false}},
  {ok:false,outcome:'pivot-operation-error',stage:'init',error:"Cannot read properties of null (reading 'map')"},
  {ok:true,outcome:'pivot-already-satisfied',noOp:true,verification:{measurable:true,match:true},state:{present:true}}
 ])
 r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,true);assert.equal(r.applied,true);assert.equal(r.outcome,'pivot-live-verified-after-second-command-boundary');assert.equal(calls.length,3)
 assert.match(bodies[1],/"apply":false/);assert.match(bodies[2],/"apply":false/)
 assert.equal(r.secondPostMutationObservation.state.present,true)

  calls=[];bodies=[]
 s=session([{ok:false,outcome:'pivot-operation-error',applied:false}])
 r=await pivot.runCommand(s,spec,true)
 assert.equal(r.ok,false);assert.equal(calls.length,1)

 calls=[];bodies=[]
 s=session([{ok:false,outcome:'pivot-semantic-mismatch',applied:true},{ok:true,noOp:true,verification:{measurable:true,match:true}}])
 r=await pivot.runCommand(s,spec,false)
 assert.equal(r.ok,false);assert.equal(calls.length,1)
 // Source-level guard for the full-batch context fix: existing-sheet insert
 // must activate and verify the exact destination sheet before insertion.
 const fs=require('fs'),observerSource=fs.readFileSync(require.resolve('./xlsx-pivot-observer.cjs'),'utf8')
 assert.match(observerSource,/destinationSheet,'SetActive'/)
 assert.match(observerSource,/pivot-destination-activation-failed/)
  console.log('XLSX PIVOT COMMAND BOUNDARY STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
