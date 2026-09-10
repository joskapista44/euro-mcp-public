'use strict'

function dataValidationCommand(spec) {
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function fail(outcome,error,extra){return Object.assign({ok:false,outcome,source:'live-coedit-editor',error},extra||{})}
  function unknown(operation,reason,extra){return Object.assign({ok:false,outcome:'unknown',source:'live-coedit-editor',operation,verification:{status:'UNKNOWN',reason}},extra||{})}
  function pass(operation,expected,actual,extra){return Object.assign({ok:true,outcome:'ok',source:'live-coedit-editor',operation,verification:{status:'PASS',expected,actual}},extra||{})}
  function mismatch(operation,expected,actual,extra){return Object.assign(fail('verification-failed','live validation read-back does not match requested state',{operation,verification:{status:'FAIL',expected,actual}}),extra||{})}
  function snapshot(v){
    if(!v)return null
    var names=['Type','AlertStyle','Operator','Formula1','Formula2','IgnoreBlank','InCellDropdown','InputMessage','InputTitle','ShowError','ShowInput','ErrorMessage','ErrorTitle']
    var out={}
    for(var i=0;i<names.length;i++){var n=names[i],g='Get'+n;if(has(v,g)){try{out[n.charAt(0).toLowerCase()+n.slice(1)]=v[g]()}catch(_){}}}
    return out
  }
  function coreEqual(a,b){return !!a&&!!b&&a.type===b.type&&a.alertStyle===b.alertStyle&&a.operator===b.operator&&String(a.formula1)===String(b.formula1)&&String(a.formula2||'')===String(b.formula2||'')}
  try{
    if(!spec||typeof spec.type!=='string')return fail('invalid-operation','type is required')
    if(!spec.sheet||!spec.range)return fail('invalid-operation','sheet and range are required',{operation:spec.type})
    if(!has(Api,'GetSheet'))return unknown(spec.type,'Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(spec.sheet);if(!sheet)return fail('sheet-not-found','worksheet not found',{operation:spec.type})
    if(!has(sheet,'GetRange'))return unknown(spec.type,'ApiWorksheet.GetRange is unavailable')
    var range=sheet.GetRange(spec.range);if(!range||!has(range,'GetValidation'))return unknown(spec.type,'ApiRange.GetValidation is unavailable')
    var validation=range.GetValidation();if(!validation)return unknown(spec.type,'ApiRange.GetValidation returned no validation object')

    if(spec.type==='validation.inspect')return pass(spec.type,'live-validation-state',snapshot(validation),{sheet:spec.sheet,range:spec.range})

    if(spec.type==='validation.add'||spec.type==='validation.modify'){
      var method=spec.type==='validation.add'?'Add':'Modify'
      if(!has(validation,method))return unknown(spec.type,'ApiValidation.'+method+' is unavailable')
      if(!spec.validationType)return fail('invalid-operation','validationType is required',{operation:spec.type})
      var alertStyle=spec.alertStyle||'xlValidAlertStop',operator=spec.operator||'xlBetween'
      validation[method](spec.validationType,alertStyle,operator,spec.formula1,spec.formula2)
      var current=range.GetValidation(),actual=snapshot(current)
      var expected={type:spec.validationType,alertStyle:alertStyle,operator:operator,formula1:spec.formula1,formula2:spec.formula2==null?'':spec.formula2}
      if(!coreEqual(expected,actual))return mismatch(spec.type,expected,actual,{sheet:spec.sheet,range:spec.range})
      return pass(spec.type,expected,actual,{sheet:spec.sheet,range:spec.range})
    }

    if(spec.type==='validation.options'){
      var setters={ignoreBlank:'SetIgnoreBlank',inCellDropdown:'SetInCellDropdown',inputMessage:'SetInputMessage',inputTitle:'SetInputTitle',showError:'SetShowError',showInput:'SetShowInput',errorMessage:'SetErrorMessage',errorTitle:'SetErrorTitle'}
      var requested={}
      for(var k in setters){if(Object.prototype.hasOwnProperty.call(spec,k)){var s=setters[k];if(!has(validation,s))return unknown(spec.type,'ApiValidation.'+s+' is unavailable');validation[s](spec[k]);requested[k]=spec[k]}}
      var actual2=snapshot(range.GetValidation())
      for(var q in requested){if(actual2[q]!==requested[q])return mismatch(spec.type,requested,actual2,{sheet:spec.sheet,range:spec.range})}
      return pass(spec.type,requested,actual2,{sheet:spec.sheet,range:spec.range})
    }

    if(spec.type==='validation.delete'){
      if(!has(validation,'Delete'))return unknown(spec.type,'ApiValidation.Delete is unavailable')
      validation.Delete()
      var after=snapshot(range.GetValidation())
      var absent=!after||after.type==='xlValidateInputOnly'||after.type===null||after.type===undefined
      if(!absent)return mismatch(spec.type,'validation-absent',after,{sheet:spec.sheet,range:spec.range})
      return pass(spec.type,'validation-absent',after,{sheet:spec.sheet,range:spec.range})
    }
    return fail('invalid-operation','unknown validation operation: '+spec.type,{operation:spec.type})
  }catch(e){return fail('operation-error',String(e&&e.message?e.message:e),{operation:spec&&spec.type?spec.type:null})}
}

async function runDataValidationInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body=`return (${dataValidationCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({u,timeout,commandBody})=>new Promise(resolve=>{
    const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor
    if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'})
    let settled=false;const finish=v=>{if(!settled){settled=true;resolve(v)}}
    try{editor.callCommand(new Function(commandBody),false,v=>finish(v===undefined?{ok:false,outcome:'ures-callback',source:'live-coedit-editor',error:'callCommand returned undefined'}:v))}
    catch(e){finish({ok:false,outcome:'callcommand-dobott',source:'live-coedit-editor',error:String(e&&e.message?e.message:e)})}
    setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'data validation callback timed out'}),timeout)
  }),{u:apiHely,timeout:timeoutMs,commandBody:body})
}
module.exports={dataValidationCommand,runDataValidationInFrame}
