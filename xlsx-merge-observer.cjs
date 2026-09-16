'use strict'
function mergeObserveCommand(spec,apply){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function bad(outcome,error,extra){return Object.assign({ok:false,outcome,source:'live-coedit-editor',authority:'LIVE_READ',error},extra||{})}
  function norm(v){var s=String(v||'').replace(/\$/g,'').toUpperCase(),bang=s.lastIndexOf('!');return bang>=0?s.slice(bang+1):s}
  function first(v){return norm(v).split(':')[0]}
  function state(range,target){var probe=range;try{if(has(range,'GetCells'))probe=range.GetCells(1,1)||range}catch(_){}var area=null;try{area=probe&&probe.MergeArea}catch(_){}if(!area||!has(area,'GetAddress'))return {ok:false,outcome:'merge-readback-unavailable'};var actual=null;try{actual=norm(area.GetAddress())}catch(_){return {ok:false,outcome:'merge-readback-unavailable'}};var expected=norm(target),single=first(target);return {ok:true,actual,expected,single,state:actual===expected?'merged-exact':actual===single?'unmerged':'merged-other'}}
  try{
    if(!spec||!spec.sheet||!spec.range||!['merge_range','unmerge_range'].includes(spec.intent))return bad('invalid-merge-spec','sheet, range and merge intent are required')
    var sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null;if(!sh||!has(sh,'GetRange'))return bad('sheet-not-found','sheet/range API unavailable')
    var r=sh.GetRange(spec.range);if(!r)return bad('range-not-found','range not found')
    var before=state(r,spec.range);if(!before.ok)return bad(before.outcome,'merge area readback unavailable')
    var desired=spec.intent==='merge_range'?'merged-exact':'unmerged'
    if(!apply)return {ok:true,outcome:before.state===desired?'merge-already-satisfied':before.state==='merged-other'?'merge-state-conflict':'merge-ready',source:'live-coedit-editor',authority:'LIVE_READ',noOp:before.state===desired,state:before.state,actual:before.actual,expected:before.expected}
    if(before.state===desired)return {ok:true,outcome:'merge-already-satisfied',source:'live-coedit-editor',authority:'LIVE_VERIFY',noOp:true,state:before.state,actual:before.actual,expected:before.expected,applied:[]}
    if(before.state==='merged-other')return bad('merge-state-conflict','target belongs to a different merged area',{state:before.state,actual:before.actual,expected:before.expected})
    var method=spec.intent==='merge_range'?'Merge':'UnMerge';if(!has(r,method))return bad('merge-api-unavailable','ApiRange.'+method+' unavailable')
    var ret=spec.intent==='merge_range'?r.Merge(false):r.UnMerge();if(ret===false)return bad('merge-operation-failed','ApiRange.'+method+' returned false')
    var after=state(r,spec.range);if(!after.ok)return bad(after.outcome,'post-mutation merge readback unavailable')
    var pass=after.state===desired
    return {ok:pass,outcome:pass?'merge-live-verified':'merge-live-verify-failed',source:'live-coedit-editor',authority:pass?'LIVE_VERIFY':'LIVE_READ',noOp:false,state:after.state,actual:after.actual,expected:after.expected,applied:pass?[spec.intent]:[],verification:{status:pass?'PASS':'FAIL',expected:desired,actual:after.state}}
  }catch(err){return bad('merge-observer-error',String(err&&err.message||err))}
}
module.exports={mergeObserveCommand}
