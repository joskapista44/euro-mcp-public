'use strict'
const assert=require('assert/strict')
const Module=require('module'),load=Module._load
let calls=[],bodies=[]
Module._load=function(r,p,m){if(r==='./xlsx-persistent-session.cjs'||r==='./xlsx-agent-freeze-task.cjs')return {};if(r==='./xlsx-freeze-observer.cjs')return {freezeObserveCommand:function freezeObserveCommand(){}};return load.call(this,r,p,m)}
const freeze=require('./xlsx-persistent-freeze.cjs');Module._load=load
function session(results){let i=0;return {apiWhere:'window.editor',frame:{evaluate:async(_f,arg)=>{calls.push(i);bodies.push(arg.body);return results[i++]}}}}
;(async()=>{
 const op={intent:'freeze_panes',sheet:'Sales_Data',mode:'at',range:'A3'}
 let s=session([{ok:false,outcome:'freeze-verification-mismatch-or-unavailable',applied:true,noOp:false,verification:{measurable:true,actual:null,match:false}},{ok:true,outcome:'ok',applied:false,noOp:true,verification:{measurable:true,actual:{r1:0,c1:0,r2:2,c2:0},match:true}}])
 let r=await freeze.runCommand(s,op,true)
 assert.equal(r.ok,true);assert.equal(r.outcome,'freeze-live-verified-after-command-boundary');assert.equal(r.applied,true);assert.equal(r.noOp,false);assert.equal(calls.length,2);assert.match(bodies[0],/true/);assert.match(bodies[1],/false/)
 calls=[];bodies=[]
 s=session([{ok:false,outcome:'freeze-verification-mismatch-or-unavailable',applied:true,noOp:false,verification:{measurable:true,actual:null,match:false}},{ok:true,outcome:'freeze-observed',applied:false,noOp:false,verification:{measurable:true,actual:null,match:false}},{ok:true,outcome:'ok',applied:false,noOp:true,verification:{measurable:true,actual:{r1:0,c1:0,r2:3,c2:0},match:true}}])
 r=await freeze.runCommand(s,op,true)
 assert.equal(r.ok,true);assert.equal(r.outcome,'freeze-live-verified-after-second-command-boundary');assert.equal(r.noOp,false);assert.equal(r.applied,true);assert.equal(calls.length,3);assert.match(bodies[2],/false/)
 calls=[];bodies=[]
 s=session([{ok:false,outcome:'freeze-verification-mismatch-or-unavailable',applied:true,noOp:false,verification:{measurable:true,actual:null,match:false}},{ok:true,outcome:'freeze-observed',applied:false,noOp:false,verification:{measurable:true,actual:null,match:false}},{ok:true,outcome:'freeze-observed',applied:false,noOp:false,verification:{measurable:true,actual:null,match:false}}])
 r=await freeze.runCommand(s,op,true)
 assert.equal(r.ok,false);assert.equal(r.outcome,'freeze-verification-mismatch-or-unavailable');assert.equal(r.secondPostMutationObservation.verification.match,false);assert.equal(calls.length,3)
 calls=[];bodies=[]
 s=session([{ok:true,outcome:'ok',applied:false,noOp:true,verification:{measurable:true,match:true}}])
 r=await freeze.runCommand(s,op,false);assert.equal(r.ok,true);assert.equal(calls.length,1)
 console.log('XLSX FREEZE COMMAND BOUNDARY STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
