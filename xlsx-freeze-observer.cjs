'use strict'
function freezeObserveCommand(sheetName,operation,apply){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
  function boxOf(range){var b=range&&range.range&&range.range.bbox;return b?{r1:b.r1,c1:b.c1,r2:b.r2,c2:b.c2}:null}
  function same(a,b){if(a===null||b===null)return a===b;return a.r1===b.r1&&a.c1===b.c1&&a.r2===b.r2&&a.c2===b.c2}
  try{
    if(!has(Api,'GetSheet'))return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(sheetName)
    if(!sheet||!has(sheet,'GetFreezePanes'))return fail('sheet-or-freeze-api-unavailable','worksheet/freeze panes API unavailable')
    var fp=sheet.GetFreezePanes()
    if(!fp||!has(fp,'GetLocation'))return fail('freeze-readback-unavailable','ApiFreezePanes.GetLocation is unavailable')
    var expected=null,targetRange=null
    if(operation.intent==='freeze_panes'){
      if(operation.mode==='rows')expected={r1:0,c1:0,r2:operation.count-1,c2:AscCommon.gc_nMaxCol0}
      else if(operation.mode==='columns')expected={r1:0,c1:0,r2:AscCommon.gc_nMaxRow0,c2:operation.count-1}
      else if(operation.mode==='at'){
        if(!has(sheet,'GetRange'))return fail('freeze-target-unavailable','ApiWorksheet.GetRange is unavailable')
        targetRange=sheet.GetRange(operation.range)
        var targetBox=boxOf(targetRange)
        if(!targetBox)return fail('freeze-target-unavailable','freeze target range bbox is unavailable')
        expected={r1:0,c1:0,r2:targetBox.r2,c2:targetBox.c2}
      }else return fail('invalid-freeze-mode','mode must be rows, columns or at')
    }else if(operation.intent!=='unfreeze_panes')return fail('invalid-freeze-intent','unsupported freeze intent')
    function observe(){
      var location=fp.GetLocation(),actual=boxOf(location),address=location&&has(location,'GetAddress')?String(location.GetAddress()):null
      return {measurable:location===null||actual!==null,actual:actual,address:address,expected:expected,match:same(actual,expected)}
    }
    var before=observe(),noOp=before.measurable&&before.match
    if(apply&&!noOp){
      if(operation.intent==='unfreeze_panes'){
        if(!has(fp,'Unfreeze'))return fail('unsupported','ApiFreezePanes.Unfreeze is unavailable')
        fp.Unfreeze()
      }else if(operation.mode==='rows'){
        if(!has(fp,'FreezeRows'))return fail('unsupported','ApiFreezePanes.FreezeRows is unavailable')
        fp.FreezeRows(operation.count)
      }else if(operation.mode==='columns'){
        if(!has(fp,'FreezeColumns'))return fail('unsupported','ApiFreezePanes.FreezeColumns is unavailable')
        fp.FreezeColumns(operation.count)
      }else{
        if(!has(fp,'FreezeAt'))return fail('unsupported','ApiFreezePanes.FreezeAt is unavailable')
        fp.FreezeAt(targetRange)
      }
    }
    var after=apply?observe():before
    return {ok:after.measurable&&after.match,outcome:after.measurable&&after.match?'ok':'freeze-verification-mismatch-or-unavailable',source:'live-coedit-editor',sheet:sheetName,operation:operation,applied:apply&&!noOp,noOp:noOp,before:before,verification:after}
  }catch(err){return fail('freeze-error',String(err&&err.message||err))}
}
module.exports={freezeObserveCommand}
