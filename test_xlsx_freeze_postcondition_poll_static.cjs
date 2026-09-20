'use strict'
const assert=require('assert/strict'),Module=require('module'),load=Module._load
let calls=[],bodies=[],results=[]
Module._load=function(r,p,m){if(r==='./xlsx-persistent-session.cjs'||r==='./xlsx-agent-freeze-task.cjs')return {};if(r==='./xlsx-freeze-observer.cjs')return {freezeObserveCommand:function(){}};return load.call(this,r,p,m)}
const freeze=require('./xlsx-persistent-freeze.cjs');Module._load=load
function session(xs){let i=0;return {apiWhere:'window.editor',frame:{evaluate:async(_f,arg)=>{calls.push(i);bodies.push(arg.body);return xs[i++]}}}}
;(async()=>{
 const op={intent:'freeze_panes',sheet:'Sales_Data',mode:'at',range:'A3'}
 const miss={ok:false,outcome:'freeze-verification-mismatch-or-unavailable',applied:false,noOp:false,verification:{measurable:true,actual:null,match:false}}
 const first={...miss,applied:true}
 const hit={ok:true,outcome:'ok',applied:false,noOp:true,verification:{measurable:true,actual:{r1:0,c1:0,r2:2,c2:0},match:true}}
 calls=[];bodies=[];let r=await freeze.runCommand(session([first,miss,miss,hit]),op,true)
 assert.equal(r.ok,true);assert.equal(r.outcome,'freeze-live-verified-after-postcondition-poll');assert.equal(r.noOp,false);assert.equal(r.applied,true);assert.equal(r.postMutationObservations.length,3);assert.equal(calls.length,4)
 for(let i=1;i<bodies.length;i++)assert.match(bodies[i],/false/)
 calls=[];bodies=[];r=await freeze.runCommand(session([first,{...miss,verification:{measurable:true,actual:{r1:0,c1:0,r2:1,c2:0},match:false}}]),op,true)
 assert.equal(r.ok,false);assert.equal(r.postMutationObservations.length,1);assert.equal(calls.length,2)
 calls=[];bodies=[];r=await freeze.runCommand(session([hit]),op,false);assert.equal(r.ok,true);assert.equal(calls.length,1)
 console.log('XLSX FREEZE POSTCONDITION POLL STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
