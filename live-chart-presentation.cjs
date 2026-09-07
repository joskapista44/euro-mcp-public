'use strict'

function chartPresentationCommand(spec){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
  function safe(o,n,args){try{return has(o,n)?o[n].apply(o,args||[]):null}catch(_){return null}}
  function sheetOf(name){try{return has(Api,'GetSheet')?Api.GetSheet(name):null}catch(_){return null}}
  function chartsOf(sheet){if(!has(sheet,'GetAllCharts'))return null;try{return sheet.GetAllCharts()||[]}catch(_){return null}}
  function targetOf(charts,s){var index=s.index==null?0:Number(s.index),target=null;if(s.name!=null){for(var i=0;i<charts.length;i++)if(String(safe(charts[i],'GetName'))===String(s.name)){target=charts[i];index=i;break}}else if(Number.isInteger(index)&&index>=0&&index<charts.length)target=charts[index];return {target:target,index:index}}
  function support(chart){
    var names=['SetLegendPos','SetLegendFontSize','SetHorAxisTitle','SetVerAxisTitle','SetHorAxisLabelsFontSize','SetVertAxisLabelsFontSize','SetHorAxisOrientation','SetVerAxisOrientation','SetAxieNumFormat','SetShowDataLabels','SetShowPointDataLabel','ApplyChartStyle','SetPlotAreaFill','SetPlotAreaOutLine','SetSeriesFill','SetSeriesOutLine','SetMarkerFill','SetMarkerOutLine']
    var out={};for(var i=0;i<names.length;i++)out[names[i]]=has(chart,names[i])
    out.GetLegendPos=has(chart,'GetLegendPos');out.GetHorAxisTitle=has(chart,'GetHorAxisTitle');out.GetVerAxisTitle=has(chart,'GetVerAxisTitle');out.GetShowDataLabels=has(chart,'GetShowDataLabels');out.GetChartStyle=has(chart,'GetChartStyle')
    return out
  }
  function unknown(op,chartIndex,changed,reason,extra){var x={ok:true,outcome:'unknown',source:'live-coedit-editor',operation:op,sheet:spec.sheet,chartIndex:chartIndex,changed:changed,verification:{status:'UNKNOWN',reason:reason}};if(extra)for(var k in extra)x[k]=extra[k];return x}
  try{
    if(!spec||typeof spec!=='object')return fail('invalid-operation','spec is required')
    if(!spec.sheet)return fail('invalid-operation','sheet is required')
    var sheet=sheetOf(spec.sheet);if(!sheet)return fail('sheet-not-found','worksheet not found',{sheet:spec.sheet})
    var charts=chartsOf(sheet);if(charts===null)return fail('unsupported','ApiWorksheet.GetAllCharts is unavailable')
    var found=targetOf(charts,spec),chart=found.target;if(!chart)return fail('invalid-operation','chart target not found',{count:charts.length})
    if(spec.type==='chart.presentation.inspect')return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,support:support(chart),verification:{status:'PASS',actual:support(chart)}}
    if(spec.type==='chart.presentation.legend'){
      if(spec.position==null&&spec.fontSize==null)return fail('invalid-operation','position or fontSize is required')
      var changed=[]
      if(spec.position!=null){if(!has(chart,'SetLegendPos'))return fail('unsupported','ApiChart.SetLegendPos is unavailable');chart.SetLegendPos(String(spec.position));changed.push('position')}
      if(spec.fontSize!=null){if(!has(chart,'SetLegendFontSize'))return fail('unsupported','ApiChart.SetLegendFontSize is unavailable');chart.SetLegendFontSize(Number(spec.fontSize));changed.push('fontSize')}
      if(spec.position!=null&&has(chart,'GetLegendPos')){var actual=safe(chart,'GetLegendPos');if(actual==null)return unknown(spec.type,found.index,changed,'legend position getter returned no value');if(String(actual)!==String(spec.position))return fail('verification-failed','legend position readback mismatch',{expected:String(spec.position),actual:String(actual)});if(spec.fontSize==null)return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,changed:changed,verification:{status:'PASS',expected:{position:String(spec.position)},actual:{position:String(actual)}}}}
      return unknown(spec.type,found.index,changed,'deployed runtime exposes legend setters but no complete public getter for requested presentation state')
    }
    if(spec.type==='chart.presentation.axisTitles'){
      if(spec.horizontal==null&&spec.vertical==null)return fail('invalid-operation','horizontal or vertical axis title is required')
      var c=[]
      if(spec.horizontal!=null){if(!has(chart,'SetHorAxisTitle'))return fail('unsupported','ApiChart.SetHorAxisTitle is unavailable');chart.SetHorAxisTitle(String(spec.horizontal),spec.fontSize==null?11:Number(spec.fontSize),spec.bold===true);c.push('horizontal')}
      if(spec.vertical!=null){if(!has(chart,'SetVerAxisTitle'))return fail('unsupported','ApiChart.SetVerAxisTitle is unavailable');chart.SetVerAxisTitle(String(spec.vertical),spec.fontSize==null?11:Number(spec.fontSize),spec.bold===true);c.push('vertical')}
      return unknown(spec.type,found.index,c,'deployed runtime exposes axis-title setters but no public axis-title getter')
    }
    if(spec.type==='chart.presentation.dataLabels'){
      if(!has(chart,'SetShowDataLabels'))return fail('unsupported','ApiChart.SetShowDataLabels is unavailable')
      var showSer=!!spec.showSeriesName,showCat=!!spec.showCategoryName,showVal=spec.showValue!==false,showPct=!!spec.showPercent
      chart.SetShowDataLabels(showSer,showCat,showVal,showPct)
      return unknown(spec.type,found.index,['dataLabels'],'deployed runtime exposes SetShowDataLabels but no public data-label state getter',{requested:{showSeriesName:showSer,showCategoryName:showCat,showValue:showVal,showPercent:showPct}})
    }
    if(spec.type==='chart.presentation.style'){
      if(spec.style==null)return fail('invalid-operation','style is required')
      if(!has(chart,'ApplyChartStyle'))return fail('unsupported','ApiChart.ApplyChartStyle is unavailable')
      chart.ApplyChartStyle(Number(spec.style))
      return unknown(spec.type,found.index,['style'],'deployed runtime exposes ApplyChartStyle but no public chart-style getter',{requested:{style:Number(spec.style)}})
    }
    return fail('invalid-operation','unknown chart presentation operation: '+spec.type)
  }catch(err){return fail('chart-presentation-error',String(err&&err.message?err.message:err))}
}

async function runChartPresentationInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body=`return (${chartPresentationCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};if(!editor||typeof editor.callCommand!=='function')return finish({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'});try{editor.callCommand(new Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})
}

module.exports={chartPresentationCommand,runChartPresentationInFrame}
