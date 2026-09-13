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
    out.GetLegendPos=has(chart,'GetLegendPos');out.GetHorAxisTitle=has(chart,'GetHorAxisTitle');out.GetVerAxisTitle=has(chart,'GetVerAxisTitle');out.GetDataLabels=has(chart,'GetDataLabels');out.GetChartStyle=has(chart,'GetChartStyle')
    return out
  }
  function unknown(op,chartIndex,changed,reason,extra){var x={ok:true,outcome:'unknown',source:'live-coedit-editor',operation:op,sheet:spec.sheet,chartIndex:chartIndex,changed:changed,verification:{status:'UNKNOWN',reason:reason}};if(extra)for(var k in extra)x[k]=extra[k];return x}
  function trimParagraphEnd(value){return typeof value==='string'?value.replace(/[\r\n]+$/,''):value}
  function same(a,b){return JSON.stringify(a)===JSON.stringify(b)}
  try{
    if(!spec||typeof spec!=='object')return fail('invalid-operation','spec is required')
    if(!spec.sheet)return fail('invalid-operation','sheet is required')
    var sheet=sheetOf(spec.sheet);if(!sheet)return fail('sheet-not-found','worksheet not found',{sheet:spec.sheet})
    var charts=chartsOf(sheet);if(charts===null)return fail('unsupported','ApiWorksheet.GetAllCharts is unavailable')
    var found=targetOf(charts,spec),chart=found.target;if(!chart)return fail('invalid-operation','chart target not found',{count:charts.length})
    if(spec.type==='chart.presentation.inspect')return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,support:support(chart),verification:{status:'PASS',actual:support(chart)}}
    if(spec.type==='chart.presentation.legend'){
      if(spec.position==null&&spec.fontSize==null)return fail('invalid-operation','position or fontSize is required')
      if(spec.fontSize!=null)return unknown(spec.type,found.index,[],'legend font-size readback is unavailable; mutation not attempted')
      if(!has(chart,'SetLegendPos'))return fail('unsupported','ApiChart.SetLegendPos is unavailable')
      if(!has(chart,'GetLegendPos'))return unknown(spec.type,found.index,[],'ApiChart.GetLegendPos is unavailable; mutation not attempted')
      var requestedPosition=String(spec.position)
      chart.SetLegendPos(requestedPosition)
      var actual=safe(chart,'GetLegendPos')
      if(actual==null)return fail('verification-failed','legend position getter returned no value',{expected:requestedPosition,actual:actual})
      if(String(actual)!==requestedPosition)return fail('verification-failed','legend position readback mismatch',{expected:requestedPosition,actual:String(actual)})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,changed:['position'],verification:{status:'PASS',expected:{position:requestedPosition},actual:{position:String(actual)}}}
    }
    if(spec.type==='chart.presentation.axisTitles'){
      if(spec.horizontal==null&&spec.vertical==null)return fail('invalid-operation','horizontal or vertical axis title is required')
      if(spec.horizontal!=null&&(!has(chart,'SetHorAxisTitle')||!has(chart,'GetHorAxisTitle')))return unknown(spec.type,found.index,[],'horizontal axis setter/getter pair unavailable; mutation not attempted')
      if(spec.vertical!=null&&(!has(chart,'SetVerAxisTitle')||!has(chart,'GetVerAxisTitle')))return unknown(spec.type,found.index,[],'vertical axis setter/getter pair unavailable; mutation not attempted')
      var changed=[],expectedTitles={},actualTitles={}
      if(spec.horizontal!=null){
        var expectedH=String(spec.horizontal)
        chart.SetHorAxisTitle(expectedH,spec.fontSize==null?11:Number(spec.fontSize),spec.bold===true)
        var rawH=safe(chart,'GetHorAxisTitle'),actualH=trimParagraphEnd(rawH)
        if(actualH!==expectedH)return fail('verification-failed','horizontal axis title readback mismatch',{expected:expectedH,actual:rawH,normalizedActual:actualH})
        changed.push('horizontal');expectedTitles.horizontal=expectedH;actualTitles.horizontal=actualH
      }
      if(spec.vertical!=null){
        var expectedV=String(spec.vertical)
        chart.SetVerAxisTitle(expectedV,spec.fontSize==null?11:Number(spec.fontSize),spec.bold===true)
        var rawV=safe(chart,'GetVerAxisTitle'),actualV=trimParagraphEnd(rawV)
        if(actualV!==expectedV)return fail('verification-failed','vertical axis title readback mismatch',{expected:expectedV,actual:rawV,normalizedActual:actualV})
        changed.push('vertical');expectedTitles.vertical=expectedV;actualTitles.vertical=actualV
      }
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,changed:changed,verification:{status:'PASS',expected:expectedTitles,actual:actualTitles}}
    }
    if(spec.type==='chart.presentation.dataLabels'){
      if(!has(chart,'SetShowDataLabels'))return fail('unsupported','ApiChart.SetShowDataLabels is unavailable')
      if(!has(chart,'GetDataLabels'))return unknown(spec.type,found.index,[],'ApiChart.GetDataLabels is unavailable; mutation not attempted')
      var showSer=!!spec.showSeriesName,showCat=!!spec.showCategoryName,showVal=spec.showValue!==false,showPct=!!spec.showPercent
      var expectedLabels={showSerName:showSer,showCatName:showCat,showVal:showVal,showPercent:showPct}
      chart.SetShowDataLabels(showSer,showCat,showVal,showPct)
      var actualLabels=safe(chart,'GetDataLabels')
      if(!same(actualLabels,expectedLabels))return fail('verification-failed','data-label state readback mismatch',{expected:expectedLabels,actual:actualLabels})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,changed:['dataLabels'],verification:{status:'PASS',expected:expectedLabels,actual:actualLabels}}
    }
    if(spec.type==='chart.presentation.style'){
      if(spec.style==null)return fail('invalid-operation','style is required')
      if(!has(chart,'ApplyChartStyle'))return fail('unsupported','ApiChart.ApplyChartStyle is unavailable')
      if(!has(chart,'GetChartStyle'))return unknown(spec.type,found.index,[],'ApiChart.GetChartStyle is unavailable; mutation not attempted')
      var expectedStyle=Number(spec.style)
      if(!Number.isFinite(expectedStyle))return fail('invalid-operation','style must be numeric')
      var styleSet=chart.ApplyChartStyle(expectedStyle)
      var actualStyle=safe(chart,'GetChartStyle')
      if(Number(actualStyle)!==expectedStyle)return fail('verification-failed','chart-style readback mismatch',{expected:expectedStyle,actual:actualStyle,setResult:styleSet})
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,chartIndex:found.index,changed:['style'],setResult:styleSet,verification:{status:'PASS',expected:{style:expectedStyle},actual:{style:Number(actualStyle)}}}
    }
    return fail('invalid-operation','unknown chart presentation operation: '+spec.type)
  }catch(err){return fail('chart-presentation-error',String(err&&err.message?err.message:err))}
}

async function runChartPresentationInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body=`return (${chartPresentationCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};if(!editor||typeof editor.callCommand!=='function')return finish({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'});try{editor.callCommand(new Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})
}

module.exports={chartPresentationCommand,runChartPresentationInFrame}
