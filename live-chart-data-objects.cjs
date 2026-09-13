'use strict'

function chartDataObjectCommand(spec){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function safe(o,n){try{return has(o,n)?o[n]():null}catch(_){return null}}
  function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
  function unknown(op,reason){return {ok:true,outcome:'unknown',source:'live-coedit-editor',operation:op,verification:{status:'UNKNOWN',reason:reason}}}
  function normalizeFormula(value){return typeof value==='string'&&value.charAt(0)==='='?value.slice(1):value}
  function decodeSeriesName(value){if(typeof value!=='string')return value;if(value.slice(0,2)==='="'&&value.slice(-1)==='"')return value.slice(2,-1).replace(/""/g,'"');return value}
  function same(a,b){return JSON.stringify(a)===JSON.stringify(b)}
  try{
    if(!spec||!spec.sheet)return fail('invalid-operation','spec and sheet are required')
    var sh=Api.GetSheet(spec.sheet);if(!sh)return fail('sheet-not-found','worksheet not found')
    var charts=sh.GetAllCharts()||[],chart=null,index=-1
    for(var i=0;i<charts.length;i++)if(spec.name!=null&&String(safe(charts[i],'GetName'))===String(spec.name)){chart=charts[i];index=i;break}
    if(!chart&&spec.index!=null){index=Number(spec.index);if(Number.isInteger(index)&&index>=0&&index<charts.length)chart=charts[index]}
    if(!chart)return fail('invalid-operation','chart target not found',{count:charts.length})
    if(spec.type==='chart.object.inspect'){var inspectSeries=has(chart,'GetAllSeries')?(chart.GetAllSeries()||[]):[],inspectS0=inspectSeries.length?inspectSeries[0]:null;return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,actual:{name:safe(chart,'GetName'),width:safe(chart,'GetWidth'),height:safe(chart,'GetHeight')},capabilities:{rename:has(chart,'SetName')&&has(chart,'GetName'),resize:has(chart,'SetSize')&&has(chart,'GetWidth')&&has(chart,'GetHeight'),positionSetter:has(chart,'SetPosition'),positionReadback:has(chart,'GetPosition'),copy:has(chart,'Copy')&&has(sh,'AddDrawing'),seriesNameSetter:has(chart,'SetSeriaName'),seriesNameReadback:has(inspectS0,'GetName'),seriesValuesSetter:has(chart,'SetSeriaValues'),seriesValuesReadback:has(inspectS0,'GetValues'),seriesXValuesSetter:has(chart,'SetSeriaXValues'),seriesXValuesReadback:has(inspectS0,'GetXValues'),categoryFormulaSetter:has(chart,'SetCatFormula'),categoryFormulaReadback:has(inspectS0,'GetCatFormula')},verification:{status:'PASS'}}}
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
    if(spec.type==='chart.object.position'){
      if(!has(chart,'SetPosition'))return fail('unsupported','SetPosition unavailable')
      if(!has(chart,'GetPosition'))return unknown(spec.type,'GetPosition unavailable; mutation not attempted')
      var expectedPosition={fromCol:Number(spec.fromCol),colOffset:Number(spec.colOffset),fromRow:Number(spec.fromRow),rowOffset:Number(spec.rowOffset)}
      if(!Number.isFinite(expectedPosition.fromCol)||!Number.isFinite(expectedPosition.colOffset)||!Number.isFinite(expectedPosition.fromRow)||!Number.isFinite(expectedPosition.rowOffset))return fail('invalid-operation','fromCol, colOffset, fromRow and rowOffset are required')
      chart.SetPosition(expectedPosition.fromCol,expectedPosition.colOffset,expectedPosition.fromRow,expectedPosition.rowOffset)
      var actualPosition=safe(chart,'GetPosition')
      if(!same(actualPosition,expectedPosition))return fail('verification-failed','position readback mismatch',{expected:expectedPosition,actual:actualPosition})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,actual:actualPosition,verification:{status:'PASS',expected:expectedPosition,actual:actualPosition}}
    }
    if(spec.type==='chart.object.copy')return fail('unsupported','live spreadsheet ApiChart exposes no public Copy() and worksheet exposes no public AddDrawing attach path')
    if(spec.type==='chart.data.seriesName'||spec.type==='chart.data.seriesValues'||spec.type==='chart.data.seriesXValues'||spec.type==='chart.data.categoryFormula'){
      var series=has(chart,'GetAllSeries')?(chart.GetAllSeries()||[]):[],seriesIndex=spec.seriesIndex==null?0:Number(spec.seriesIndex),ser=Number.isInteger(seriesIndex)&&seriesIndex>=0&&seriesIndex<series.length?series[seriesIndex]:null
      if(!ser)return fail('invalid-operation','chart series target not found')
      if(spec.type==='chart.data.seriesName'){
        if(!has(chart,'SetSeriaName'))return fail('unsupported','SetSeriaName unavailable')
        if(!has(ser,'GetName'))return unknown(spec.type,'GetName unavailable; mutation not attempted')
        if(spec.value==null)return fail('invalid-operation','value is required')
        chart.SetSeriaName(String(spec.value),seriesIndex);var actualName=safe(ser,'GetName'),normalizedName=decodeSeriesName(actualName)
        if(normalizedName!==String(spec.value))return fail('verification-failed','series name readback mismatch',{expected:String(spec.value),actual:actualName,normalizedActual:normalizedName})
        return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,actual:actualName,normalizedActual:normalizedName,verification:{status:'PASS',expected:String(spec.value),actual:normalizedName}}
      }
      var setter=spec.type==='chart.data.seriesValues'?'SetSeriaValues':spec.type==='chart.data.seriesXValues'?'SetSeriaXValues':'SetCatFormula'
      var getter=spec.type==='chart.data.seriesValues'?'GetValues':spec.type==='chart.data.seriesXValues'?'GetXValues':'GetCatFormula'
      if(!has(chart,setter))return fail('unsupported',setter+' unavailable')
      if(!has(ser,getter))return unknown(spec.type,getter+' unavailable; mutation not attempted')
      if(spec.range==null)return fail('invalid-operation','range is required')
      if(setter==='SetCatFormula')chart[setter](String(spec.range));else chart[setter](String(spec.range),seriesIndex)
      var actualFormula=safe(ser,getter),normalizedFormula=normalizeFormula(actualFormula),expectedFormula=normalizeFormula(String(spec.range))
      if(normalizedFormula!==expectedFormula)return fail('verification-failed','series formula readback mismatch',{expected:expectedFormula,actual:actualFormula,normalizedActual:normalizedFormula})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,actual:actualFormula,normalizedActual:normalizedFormula,verification:{status:'PASS',expected:expectedFormula,actual:normalizedFormula}}
    }
    return fail('invalid-operation','unknown chart data/object operation: '+spec.type)
  }catch(err){return fail('chart-data-object-error',String(err&&err.message?err.message:err))}
}

async function runChartDataObjectInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body='return ('+chartDataObjectCommand.toString()+')('+JSON.stringify(spec)+');'
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};if(!editor||typeof editor.callCommand!=='function')return finish({ok:false,outcome:'nincs-api',source:'live-coedit-editor'});try{editor.callCommand(Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})
}

module.exports={chartDataObjectCommand,runChartDataObjectInFrame}
