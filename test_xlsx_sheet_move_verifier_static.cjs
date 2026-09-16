'use strict'
const assert=require('assert'),{verifySheetMoveSemantic}=require('./xlsx-sheet-move-verifier.cjs')
const inv=arr=>({ok:true,authority:'LIVE_READ',sheets:arr.map((name,index)=>({name,index}))})
let r=verifySheetMoveSemantic({before:inv(['A','B','C']),after:inv(['B','A','C']),sheet:'B',referenceSheet:'A',position:'before'});assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY')
r=verifySheetMoveSemantic({before:inv(['A','B','C']),after:inv(['A','C','B']),sheet:'B',referenceSheet:'C',position:'after'});assert.equal(r.ok,true)
r=verifySheetMoveSemantic({before:inv(['A','B','C']),after:inv(['A','B','C']),sheet:'B',referenceSheet:'A',position:'before'});assert.equal(r.ok,false);assert.equal(r.outcome,'sheet-move-position-mismatch')
r=verifySheetMoveSemantic({before:inv(['A','B','C']),after:inv(['B','A','D']),sheet:'B',referenceSheet:'A',position:'before'});assert.equal(r.ok,false);assert.equal(r.outcome,'sheet-move-membership-changed')
console.log('XLSX SHEET MOVE VERIFIER STATIC: PASS')
