'use strict'

function chartDataObjectCommand(spec){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function safe(o,n){try{return has(o,n)?o[n]():null}catch(_){return null}}
  function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
  function unknown(op,reason){return {ok:true,outcome:'unknown',source:'live-coedit-editor',operation:op,verification:{status:'UNKNOWN',reason:reason}}}
  try{
    if(!spec||!spec.sheet)return fail('invalid-operation','spec and sheet are required')
    var sh=Api.GetSheet(spec.sheet);if(!sh)return fail('sheet-not-found','worksheet not found')
    var charts=sh.GetAllCharts()||[],chart=null,index=-1
    for(var i=0;i<charts.length;i++)if(spec.name!=null&&String(safe(charts[i],'GetName'))===String(spec.name)){chart=charts[i];index=i;break}
    if(!chart&&spec.index!=null){index=Number(spec.index);if(Number.isInteger(index)&&index>=0&&index<charts.length)chart=charts[index]}
    if(!chart)return fail('invalid-operation','chart target not found',{count:charts.length})
    if(spec.type==='chart.object.inspect')return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,actual:{name:safe(chart,'GetName'),width:safe(chart,'GetWidth'),height:safe(chart,'GetHeight')},capabilities:{rename:has(chart,'SetName')&&has(chart,'GetName'),resize:has(chart,'SetSize')&&has(chart,'GetWidth')&&has(chart,'GetHeight'),positionSetter:has(chart,'SetPosition'),positionReadback:has(chart,'GetPosition')||has(chart,'GetPosX')||has(chart,'GetPosY'),copy:has(chart,'Copy')&&has(sh,'AddDrawing'),seriesNameSetter:has(chart,'SetSeriaName'),seriesValuesSetter:has(chart,'SetSeriaValues'),seriesXValuesSetter:has(chart,'SetSeriaXValues'),categoryFormulaSetter:has(chart,'SetCatFormula')},verification:{status:'PASS'}}
    if(spec.type==='chart.object.rename'){
      if(!spec.newName)return fail('invalid-operation','newName is required');if(!has(chart,'SetName'))return fail('unsupported','SetName unavailable');if(!has(chart,'GetName'))return unknown(spec.type,'GetName unavailable')
      var before=safe(chart,'GetName');chart.SetName(String(spec.newName));var after=safe(chart,'GetName');if(String(after)!==String(spec.newName))return fail('verification-failed','name readback mismatch',{expectedName:String(spec.newName),actualName:after})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,beforeName:before,actualName:after,verification:{status:'PASS',expectedName:String(spec.newName),actualName:String(after)}}
    }
    if(spec.type==='chart.object.resize'){
      var w=Number(spec.width),h=Number(spec.height);if(!Number.isFinite(w)||!Number.isFinite(h)||w<=0||h<=0)return fail('invalid-operation','positive width and height are required');if(!has(chart,'SetSize'))return fail('unsupported','SetSize unavailable');if(!has(chart,'GetWidth')||!has(chart,'GetHeight'))return unknown(spec.type,'size readback unavailable')
      var beforeSize={width:safe(chart,'GetWidth'),height:safe(chart,'GetHeight')};chart.SetSize(w,h);var actual={width:safe(chart,'GetWidth'),height:safe(chart,'GetHeight')};if(Number(actual.width)!==w||Number(actual.height)!==h)return fail('verification-failed','size readback mismatch',{expected:{width:w,height:h},actual:actual})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,before:beforeSize,actual:actual,verification:{status:'PASS',expected:{width:w,height:h},actual:actual}}
    }
    if(spec.type==='chart.object.position')return has(chart,'SetPosition')?unknown(spec.type,'SetPosition exists but no public semantic position getter; mutation not attempted'):fail('unsupported','SetPosition unavailable')
    if(spec.type==='chart.object.copy')return fail('unsupported','deployed runtime lacks public Copy/AddDrawing path')
    if(spec.type==='chart.data.seriesName'||spec.type==='chart.data.seriesValues'||spec.type==='chart.data.seriesXValues'||spec.type==='chart.data.categoryFormula')return unknown(spec.type,'setter exists but deployed runtime has no public semantic series-data getter; mutation not attempted')
    return fail('invalid-operation','unknown chart data/object operation: '+spec.type)
  }catch(err){return fail('chart-data-object-error',String(err&&err.message?err.message:err))}
}

async function runChartDataObjectInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body='return ('+chartDataObjectCommand.toString()+')('+JSON.stringify(spec)+');'
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};if(!editor||typeof editor.callCommand!=='function')return finish({ok:false,outcome:'nincs-api',source:'live-coedit-editor'});try{editor.callCommand(Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})
}

module.exports={chartDataObjectCommand,runChartDataObjectInFrame}
