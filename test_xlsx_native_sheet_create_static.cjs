'use strict'

const assert=require('assert')
const {runNativeSheetCreateInFrame,runNativeSheetMoveInFrame}=require('./workbook-ops.cjs')

function fixture({blocked=false,duplicate=false}={}){
 const callbacks=new Map(),names=['Sheet1']
 if(duplicate)names.push('Sales')
 const editor={
  canEdit:()=>!blocked,
  collaborativeEditing:{getGlobalLock:()=>false},
  asc_isProtectedWorkbook:()=>false,
  asc_getWorksheetsCount:()=>names.length,
  asc_getWorksheetName:i=>names[i],
  asc_registerCallback:(name,cb)=>callbacks.set(name,cb),
  asc_unregisterCallback:(name,cb)=>{if(callbacks.get(name)===cb)callbacks.delete(name)},
  asc_addWorksheet:name=>{setTimeout(()=>names.push(name),5)}
 }
 const frame={evaluate:async(fn,args)=>{const previous=global.window;global.window={Asc:{editor}};try{return await fn(args)}finally{global.window=previous}}}
 return {frame,names}
}

;(async()=>{
 const asyncLock=fixture()
 const created=await runNativeSheetCreateInFrame(asyncLock.frame,'window.Asc.editor',{type:'sheet.create',name:'Sales'},200)
 assert.equal(created.ok,true);assert.equal(created.dispatchApi,'asc_addWorksheet');assert.deepEqual(asyncLock.names,['Sheet1','Sales'])
 const duplicate=await runNativeSheetCreateInFrame(fixture({duplicate:true}).frame,'window.Asc.editor',{type:'sheet.create',name:'Sales'},50)
 assert.equal(duplicate.ok,false);assert.equal(duplicate.outcome,'already-exists')
 const blocked=await runNativeSheetCreateInFrame(fixture({blocked:true}).frame,'window.Asc.editor',{type:'sheet.create',name:'Sales'},50)
 assert.equal(blocked.ok,false);assert.equal(blocked.outcome,'create-precondition-blocked')
 const order=['Sheet1','Anchor','Dashboard'],callbacks=new Map()
 const moveEditor={
  canEdit:()=>true,collaborativeEditing:{getGlobalLock:()=>false},asc_isProtectedWorkbook:()=>false,
  asc_getWorksheetsCount:()=>order.length,asc_getWorksheetName:i=>order[i],
  asc_registerCallback:(name,cb)=>callbacks.set(name,cb),asc_unregisterCallback:(name,cb)=>{if(callbacks.get(name)===cb)callbacks.delete(name)},
  asc_moveWorksheet:(where,[source])=>setTimeout(()=>{const [item]=order.splice(source,1);order.splice(source<where?where-1:where,0,item)},5)
 }
 const moveFrame={evaluate:async(fn,args)=>{const previous=global.window;global.window={Asc:{editor:moveEditor}};try{return await fn(args)}finally{global.window=previous}}}
 const moved=await runNativeSheetMoveInFrame(moveFrame,'window.Asc.editor',{type:'sheet.move',sheet:'Dashboard',referenceSheet:'Sheet1',position:'before'},200)
 assert.equal(moved.ok,true);assert.equal(moved.dispatchApi,'asc_moveWorksheet');assert.deepEqual(order,['Dashboard','Sheet1','Anchor'])
 console.log('XLSX NATIVE SHEET STRUCTURE STATIC: PASS (async collaborative locks)')
})().catch(error=>{console.error(error);process.exitCode=1})
