'use strict'

function chartSeriesCommand(spec){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
  function safe(o,n){try{return has(o,n)?o[n]():null}catch(_){return null}}
  function sheetOf(name){try{return has(Api,'GetSheet')?Api.GetSheet(name):null}catch(_){return null}}
  function chartsOf(sheet){if(!has(sheet,'GetAllCharts'))return null;try{return sheet.GetAllCharts()||[]}catch(_){return null}}
  function targetOf(charts,s){var index=s.index==null?0:Number(s.index),target=null;if(s.name!=null){for(var i=0;i<charts.length;i++)if(String(safe(charts[i],'GetName'))===String(s.name)){target=charts[i];index=i;break}}else if(Number.isInteger(index)&&index>=0&&index<charts.length)target=charts[index];return {target:target,index:index}}
  function seriesOf(chart){if(!has(chart,'GetAllSeries'))return null;try{return chart.GetAllSeries()||[]}catch(_){return null}}
  function describeSeries(s,index){return {index:index,chartType:safe(s,'GetChartType'),classType:safe(s,'GetClassType')}}
  try{
    if(!spec||typeof spec!=='object')return fail('invalid-operation','spec is required')
    if(!spec.sheet)return fail('invalid-operation','sheet is required')
    var sheet=sheetOf(spec.sheet);if(!sheet)return fail('sheet-not-found','worksheet not found',{sheet:spec.sheet})
    var charts=chartsOf(sheet);if(charts===null)return fail('unsupported','ApiWorksheet.GetAllCharts is unavailable')
    var found=targetOf(charts,spec),chart=found.target
    if(!chart)return fail('invalid-operation','chart target not found',{count:charts.length})
    var series=seriesOf(chart);if(series===null)return fail('unsupported','ApiChart.GetAllSeries is unavailable')

    if(spec.type==='chart.series.inspect'){
      var inventory=[];for(var i=0;i<series.length;i++)inventory.push(describeSeries(series[i],i))
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,count:inventory.length,series:inventory,verification:{status:'PASS',actual:inventory}}
    }

    if(spec.type==='chart.series.add'){
      if(!has(chart,'AddSeria'))return fail('unsupported','ApiChart.AddSeria is unavailable')
      if(!spec.nameRange||!spec.valuesRange)return fail('invalid-operation','nameRange and valuesRange are required')
      var before=series.length
      try{chart.AddSeria(String(spec.nameRange),String(spec.valuesRange),spec.xValuesRange==null?undefined:String(spec.xValuesRange))}catch(err){return fail('operation-error',String(err&&err.message?err.message:err))}
      var after=seriesOf(chart);if(after===null)return fail('unsupported','series readback unavailable after add')
      if(after.length!==before+1)return fail('verification-failed','series count did not increase after AddSeria',{beforeCount:before,afterCount:after.length})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,beforeCount:before,afterCount:after.length,actual:describeSeries(after[after.length-1],after.length-1),verification:{status:'PASS',expectedCount:before+1,actualCount:after.length}}
    }

    if(spec.type==='chart.series.remove'){
      var ri=Number(spec.seriesIndex)
      if(!Number.isInteger(ri)||ri<0||ri>=series.length)return fail('invalid-operation','valid seriesIndex is required',{seriesCount:series.length})
      if(!has(chart,'RemoveSeria'))return fail('unsupported','ApiChart.RemoveSeria is unavailable')
      var removed=describeSeries(series[ri],ri),beforeRemove=series.length
      try{chart.RemoveSeria(ri)}catch(err){return fail('operation-error',String(err&&err.message?err.message:err))}
      var remaining=seriesOf(chart);if(remaining===null)return fail('unsupported','series readback unavailable after remove')
      if(remaining.length!==beforeRemove-1)return fail('verification-failed','series count did not decrease after RemoveSeria',{beforeCount:beforeRemove,afterCount:remaining.length})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,removed:removed,beforeCount:beforeRemove,afterCount:remaining.length,verification:{status:'PASS',expectedCount:beforeRemove-1,actualCount:remaining.length}}
    }

    if(spec.type==='chart.series.changeType'){
      var si=Number(spec.seriesIndex)
      if(!Number.isInteger(si)||si<0||si>=series.length||!spec.chartType)return fail('invalid-operation','valid seriesIndex and chartType are required',{seriesCount:series.length})
      var target=series[si]
      if(!has(target,'ChangeChartType'))return fail('unsupported','ApiChartSeries.ChangeChartType is unavailable')
      if(!has(target,'GetChartType'))return {ok:true,outcome:'unknown',source:'live-coedit-editor',operation:spec.type,verification:{status:'UNKNOWN',reason:'ApiChartSeries.GetChartType is unavailable'}}
      var beforeType=safe(target,'GetChartType')
      try{target.ChangeChartType(String(spec.chartType))}catch(err){return fail('operation-error',String(err&&err.message?err.message:err))}
      var reread=seriesOf(chart);if(reread===null||!reread[si])return fail('verification-failed','series disappeared during readback')
      var afterType=safe(reread[si],'GetChartType')
      if(afterType==null)return {ok:true,outcome:'unknown',source:'live-coedit-editor',operation:spec.type,beforeType:beforeType,verification:{status:'UNKNOWN',reason:'series type readback unavailable after change'}}
      if(String(afterType)!==String(spec.chartType))return fail('verification-failed','series chart type readback mismatch',{beforeType:beforeType,actualType:afterType,expectedType:spec.chartType})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,seriesIndex:si,beforeType:beforeType,actualType:afterType,verification:{status:'PASS',expectedType:String(spec.chartType),actualType:String(afterType)}}
    }

    return fail('invalid-operation','unknown chart series operation: '+spec.type)
  }catch(err){return fail('chart-series-error',String(err&&err.message?err.message:err))}
}

async function runChartSeriesInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body=`return (${chartSeriesCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};if(!editor||typeof editor.callCommand!=='function')return finish({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'});try{editor.callCommand(new Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})
}

module.exports={chartSeriesCommand,runChartSeriesInFrame}
