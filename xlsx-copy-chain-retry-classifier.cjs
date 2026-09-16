'use strict'

const {sameCell,canonicalRange}=require('./xlsx-sheet-copy-verifier.cjs')
const {normalizeFormula}=require('./verification-contract.cjs')

function sheetMeta(inventory,name){const hits=(inventory?.sheets||[]).filter(s=>s?.name===name);return hits.length===1?{ok:true,sheet:hits[0]}:{ok:false,count:hits.length}}
function colNumber(s){let n=0;for(const ch of s)n=n*26+(ch.charCodeAt(0)-64);return n}
function parseCell(s){const m=String(s||'').replace(/\$/g,'').toUpperCase().match(/^([A-Z]+)([1-9][0-9]*)$/);return m?{column:colNumber(m[1]),row:Number(m[2])}:null}
function parseRange(s){const p=String(s||'').split(':');const a=parseCell(p[0]),b=parseCell(p[1]||p[0]);return a&&b?{start:{row:Math.min(a.row,b.row),column:Math.min(a.column,b.column)},end:{row:Math.max(a.row,b.row),column:Math.max(a.column,b.column)}}:null}
function finalName(name,ops,afterIndex=-1){let current=name,guard=0;while(guard++<=ops.length){const next=ops.find(x=>x.intent==='rename_sheet'&&x.index>afterIndex&&x.sheet===current);if(!next)return current;current=next.name}return null}
function contains(range,row,column){return range&&row>=range.start.row&&row<=range.end.row&&column>=range.start.column&&column<=range.end.column}
function sameExpectedCell(cell,value,formula){if(formula!==null&&formula!==undefined)return normalizeFormula(cell?.formula)===normalizeFormula(formula);if(value===null||value===undefined)return cell?.dataType==='blank'||((cell?.value===null||cell?.value===undefined||cell?.value==='')&&!cell?.formula);return Object.is(cell?.rawValue,value)||Object.is(cell?.value,value)||String(cell?.displayText)===String(value)}
function writesForCopy(copy,ops){const terminal=finalName(copy.name,ops,copy.index);return ops.filter(x=>x.intent==='write_range'&&x.index>copy.index&&finalName(x.sheet,ops,x.index)===terminal)}
async function classifyCopyChainRetry({plan,api,inventory}){
 if(!plan?.ok||!Array.isArray(plan.operations)||!inventory?.ok||inventory.authority!=='LIVE_READ')return {matched:false}
 const copies=plan.operations.filter(x=>x.intent==='copy_sheet')
 if(copies.length!==1)return {matched:false}
 const copy=copies[0],terminal=finalName(copy.name,plan.operations,copy.index)
 if(!terminal||terminal===copy.name)return {matched:false}
 const source=sheetMeta(inventory,copy.sheet),intermediate=sheetMeta(inventory,copy.name),target=sheetMeta(inventory,terminal)
 if(!source.ok||intermediate.count!==0||!target.ok)return {matched:false}
 const writes=writesForCopy(copy,plan.operations)
 if(!writes.length)return {matched:false}
 const sr=source.sheet.usedRange,tr=target.sheet.usedRange
 if(!sr||!tr||canonicalRange(sr)!==canonicalRange(tr))return {matched:true,ok:false,outcome:'xlsx-copy-chain-retry-used-range-mismatch',authority:'LIVE_READ',writeAllowed:false,sourceRange:sr||null,targetRange:tr||null}
 const sourceRead=await api.readRange({sheet:copy.sheet,range:sr}),targetRead=await api.readRange({sheet:terminal,range:tr})
 if(!sourceRead?.ok||sourceRead.authority!=='LIVE_READ'||!targetRead?.ok||targetRead.authority!=='LIVE_READ')return {matched:true,ok:false,outcome:'xlsx-copy-chain-retry-live-read-failed',authority:'LIVE_READ',writeAllowed:false}
 const used=parseRange(sr);if(!used)return {matched:true,ok:false,outcome:'xlsx-copy-chain-retry-used-range-unparseable',authority:'LIVE_READ',writeAllowed:false}
 const parsedWrites=writes.map(w=>({...w,parsed:parseRange(w.range)}));if(parsedWrites.some(w=>!w.parsed))return {matched:true,ok:false,outcome:'xlsx-copy-chain-retry-write-range-unparseable',authority:'LIVE_READ',writeAllowed:false}
 for(let r=0;r<sourceRead.rows;r++)for(let c=0;c<sourceRead.columns;c++){
   const row=used.start.row+r,column=used.start.column+c,covering=parsedWrites.filter(w=>contains(w.parsed,row,column))
   const targetCell=targetRead.cells?.[r]?.[c]
   if(!covering.length){if(!sameCell(sourceRead.cells?.[r]?.[c],targetCell))return {matched:true,ok:false,outcome:'xlsx-copy-chain-retry-copy-baseline-mismatch',authority:'LIVE_READ',writeAllowed:false,row,column}}
   else {const w=covering[covering.length-1],wr=row-w.parsed.start.row,wc=column-w.parsed.start.column;if(!sameExpectedCell(targetCell,w.values?.[wr]?.[wc],w.formulas?.[wr]?.[wc]??null))return {matched:true,ok:false,outcome:'xlsx-copy-chain-retry-downstream-write-mismatch',authority:'LIVE_READ',writeAllowed:false,index:w.index,row,column}}
 }
 const receipts=plan.operations.map(op=>({index:op.index,intent:op.intent,status:'NO_OP_SATISFIED',resolvedTarget:op.intent==='copy_sheet'?{source:op.sheet,target:terminal,originalTarget:op.name}:op.intent==='write_range'?{sheet:terminal,originalSheet:op.sheet,range:op.range}:op.intent==='rename_sheet'?{from:op.sheet,to:terminal}:{},identityProof:'fresh-live-inventory+copy-chain-final-state-proof',verification:'LIVE_VERIFY'}))
 return {matched:true,ok:true,outcome:'xlsx-copy-chain-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,receipt:receipts,wholeTaskVerification:{ok:true,outcome:'verified-copy-chain-final-state',authority:'LIVE_VERIFY',checks:receipts.map(x=>({index:x.index,intent:x.intent,status:'PASS'}))},retryClassification:{source:copy.sheet,originalTarget:copy.name,terminalTarget:terminal,downstreamWrites:writes.map(x=>x.index)}}
}
module.exports={parseCell,parseRange,finalName,writesForCopy,classifyCopyChainRetry}
