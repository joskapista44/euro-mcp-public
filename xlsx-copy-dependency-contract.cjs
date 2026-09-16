'use strict'

function validName(v){return typeof v==='string'&&v.trim()!==''}
function validateCopyDependencies(operations){
  const copies=(operations||[]).filter(op=>op?.intent==='copy_sheet')
  const targets=new Map()
  for(const cp of copies){
    if(!validName(cp.sheet)||!validName(cp.name)||cp.sheet===cp.name)return {ok:false,outcome:'xlsx-task-invalid-copy',index:cp.index??null}
    if(targets.has(cp.name))return {ok:false,outcome:'xlsx-task-copy-target-conflict',sheet:cp.name,firstIndex:targets.get(cp.name).index,secondIndex:cp.index}
    targets.set(cp.name,cp)
    const earlier=(operations||[]).filter(op=>op.index<cp.index)
    if(earlier.some(op=>op.intent==='delete_sheet'&&op.sheet===cp.sheet))return {ok:false,outcome:'xlsx-task-copy-source-deleted-before-copy',copyIndex:cp.index,sheet:cp.sheet}
    const later=(operations||[]).filter(op=>op.index>cp.index)
    const targetDelete=later.find(op=>op.intent==='delete_sheet'&&op.sheet===cp.name)
    if(targetDelete)return {ok:false,outcome:'xlsx-task-copy-then-delete-not-supported',copyIndex:cp.index,dependencyIndex:targetDelete.index,sheet:cp.name}
    const targetCopy=later.find(op=>op.intent==='copy_sheet'&&op.sheet===cp.name)
    if(targetCopy)return {ok:false,outcome:'xlsx-task-copy-from-copy-not-supported',copyIndex:cp.index,dependencyIndex:targetCopy.index,sheet:cp.name}
  }
  return {ok:true,outcome:'xlsx-task-copy-dependencies-valid'}
}

function copyProducedSheet(name,opIndex,operations){
  const producers=(operations||[]).filter(op=>op.intent==='copy_sheet'&&op.index<opIndex&&op.name===name)
  if(producers.length!==1)return null
  return producers[0]
}

function resolveCopyDependency(name,opIndex,operations,inventory){
  const direct=(inventory?.sheets||[]).filter(s=>s?.name===name)
  if(direct.length===1)return {ok:true,name,viaCopy:false}
  if(direct.length>1)return {ok:false,outcome:'xlsx-target-sheet-ambiguous',count:direct.length}
  const producer=copyProducedSheet(name,opIndex,operations)
  if(!producer)return {ok:false,outcome:'xlsx-target-sheet-not-found',count:0}
  return {ok:false,outcome:'xlsx-copy-dependent-target-missing',producerIndex:producer.index,sheet:name}
}

module.exports={validateCopyDependencies,copyProducedSheet,resolveCopyDependency}
