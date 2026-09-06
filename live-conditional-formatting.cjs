'use strict'

const LIVE_SOURCE = 'live-coedit-editor'

function conditionalFormattingCommand(spec) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error,extra){ var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error}; if(extra)for(var k in extra)x[k]=extra[k]; return x }
  function value(o, getter, prop) {
    try { if (getter && has(o,getter)) return o[getter](); if (o && o[prop] !== undefined) return o[prop] } catch (_) {}
    return null
  }
  function colorValue(c) {
    if (c == null || c === 'No Fill') return c
    try { if(has(c,'GetRGB')) return c.GetRGB(); if(has(c,'GetHex')) return c.GetHex() } catch (_) {}
    return null
  }
  function describe(rule,index) {
    if(!rule) return null
    var applies=null
    try { var ar=has(rule,'GetAppliesTo')?rule.GetAppliesTo():rule.AppliesTo; if(ar&&has(ar,'GetAddress')) applies=String(ar.GetAddress()) } catch(_){}
    return {
      index:index,
      type:value(rule,'GetType','Type'),
      operator:value(rule,'GetOperator','Operator'),
      formula1:value(rule,'GetFormula1','Formula1'),
      formula2:value(rule,'GetFormula2','Formula2'),
      priority:value(rule,'GetPriority','Priority'),
      fillColor:colorValue(value(rule,'GetFillColor','FillColor')),
      appliesTo:applies
    }
  }
  function collection(spec) {
    if(!has(Api,'GetSheet')) return {error:'Api.GetSheet is unavailable'}
    var sheet=null; try{sheet=Api.GetSheet(spec.sheet)}catch(_){}
    if(!sheet||!has(sheet,'GetRange')) return {error:'worksheet/range API unavailable'}
    var range=null; try{range=sheet.GetRange(spec.range)}catch(_){}
    if(!range) return {error:'target range could not be resolved'}
    if(!has(range,'GetFormatConditions')) return {error:'ApiRange.GetFormatConditions is unavailable'}
    var fc=null; try{fc=range.GetFormatConditions()}catch(err){return {error:String(err&&err.message?err.message:err)}}
    if(!fc||!has(fc,'GetCount')||!has(fc,'GetItem')) return {error:'format condition collection introspection is unavailable'}
    return {sheet:sheet,range:range,fc:fc}
  }
  function inventory(fc) {
    var n=fc.GetCount(), out=[]
    for(var i=0;i<n;i++){ var r=null; try{r=fc.GetItem(i)}catch(_){}; if(!r){try{r=fc.GetItem(i+1)}catch(_){}}; out.push(describe(r,i)) }
    return out
  }
  function same(a,b){ return String(a==null?'':a)===String(b==null?'':b) }
  try {
    if(!spec||typeof spec!=='object') return fail('invalid-operation','spec is required')
    if(!spec.sheet||!spec.range) return fail('invalid-operation','sheet and range are required')
    var c=collection(spec); if(c.error)return fail('unsupported',c.error)
    var fc=c.fc
    if(spec.type==='cf.inspect') {
      var rules=inventory(fc)
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,range:spec.range,count:rules.length,rules:rules,verification:{status:'PASS',actual:rules}}
    }
    if(spec.type==='cf.add') {
      if(!has(fc,'Add'))return fail('unsupported','ApiFormatConditions.Add is unavailable')
      if(!spec.rule||!spec.rule.type)return fail('invalid-operation','rule.type is required')
      var before=fc.GetCount(), rule=null
      try{rule=fc.Add(spec.rule.type,spec.rule.operator||null,spec.rule.formula1==null?null:String(spec.rule.formula1),spec.rule.formula2==null?null:String(spec.rule.formula2))}catch(err){return fail('operation-error',String(err&&err.message?err.message:err))}
      if(!rule)return fail('operation-error','ApiFormatConditions.Add returned null')
      if(spec.rule.fillColor){ if(!has(rule,'SetFillColor')||!has(Api,'CreateColorFromRGB'))return fail('unsupported','conditional fill-color API unavailable'); rule.SetFillColor(Api.CreateColorFromRGB(spec.rule.fillColor[0],spec.rule.fillColor[1],spec.rule.fillColor[2])) }
      if(spec.rule.priority!=null){ if(!has(rule,'SetPriority'))return fail('unsupported','ApiFormatCondition.SetPriority is unavailable'); rule.SetPriority(spec.rule.priority) }
      var after=fc.GetCount(), actual=describe(rule,after-1), mismatch=[]
      if(after!==before+1)mismatch.push('count')
      if(!same(actual.type,spec.rule.type))mismatch.push('type')
      if(spec.rule.operator!=null&&!same(actual.operator,spec.rule.operator))mismatch.push('operator')
      if(spec.rule.formula1!=null&&!same(actual.formula1,spec.rule.formula1))mismatch.push('formula1')
      if(spec.rule.formula2!=null&&!same(actual.formula2,spec.rule.formula2))mismatch.push('formula2')
      if(spec.rule.priority!=null&&!same(actual.priority,spec.rule.priority))mismatch.push('priority')
      return mismatch.length?fail('verification-failed','live conditional-formatting readback mismatch',{beforeCount:before,afterCount:after,actual:actual,mismatches:mismatch}):{ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,range:spec.range,beforeCount:before,afterCount:after,actual:actual,verification:{status:'PASS',expected:spec.rule,actual:actual}}
    }
    if(spec.type==='cf.delete') {
      var beforeRules=inventory(fc), index=spec.index==null?0:Number(spec.index)
      if(!Number.isInteger(index)||index<0||index>=beforeRules.length)return fail('invalid-operation','index is outside live rule inventory',{count:beforeRules.length})
      var target=null; try{target=fc.GetItem(index)}catch(_){}; if(!target){try{target=fc.GetItem(index+1)}catch(_){}}
      if(!target||!has(target,'Delete'))return fail('unsupported','ApiFormatCondition.Delete is unavailable')
      target.Delete(); var afterRules=inventory(fc)
      if(afterRules.length!==beforeRules.length-1)return fail('verification-failed','rule count did not decrease after delete',{before:beforeRules,after:afterRules})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,range:spec.range,deleted:beforeRules[index],rules:afterRules,verification:{status:'PASS',expectedCount:beforeRules.length-1,actualCount:afterRules.length}}
    }
    return fail('invalid-operation','unknown conditional-formatting operation: '+spec.type)
  } catch(err){ return fail('conditional-formatting-error',String(err&&err.message?err.message:err)) }
}

async function runConditionalFormattingInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body=`return (${conditionalFormattingCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({u,body,timeout})=>new Promise((resolve)=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};if(!editor||typeof editor.callCommand!=='function')return finish({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'});try{editor.callCommand(new Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor',error:'callCommand callback returned undefined'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'conditional-formatting callback timed out'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})
}

module.exports={conditionalFormattingCommand,runConditionalFormattingInFrame}
