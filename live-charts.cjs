'use strict'

function chartCommand(spec) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error,extra){ var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error}; if(extra)for(var k in extra)x[k]=extra[k]; return x }
  function safe(o,n){ try{return has(o,n)?o[n]():null}catch(_){return null} }
  function normText(v){ return v==null?v:String(v).replace(/[\r\n]+$/g,'') }
  function describe(c,index){
    if(!c)return null
    var series=null
    try{series=has(c,'GetAllSeries')?(c.GetAllSeries()||[]).length:null}catch(_){}
    return {index:index,name:safe(c,'GetName'),chartType:safe(c,'GetChartType'),title:safe(c,'GetTitle'),width:safe(c,'GetWidth'),height:safe(c,'GetHeight'),seriesCount:series,classType:safe(c,'GetClassType')}
  }
  function sheetOf(name){ try{return has(Api,'GetSheet')?Api.GetSheet(name):null}catch(_){return null} }
  function chartsOf(sheet){ if(!has(sheet,'GetAllCharts'))return null; try{return sheet.GetAllCharts()||[]}catch(_){return null} }
  function drawingsOf(sheet){ if(!has(sheet,'GetAllDrawings'))return null; try{return sheet.GetAllDrawings()||[]}catch(_){return null} }
  function targetOf(charts,spec){
    var index=spec.index==null?0:Number(spec.index),target=null
    if(spec.name!=null){for(var i=0;i<charts.length;i++)if(String(safe(charts[i],'GetName'))===String(spec.name)){target=charts[i];index=i;break}}
    else if(Number.isInteger(index)&&index>=0&&index<charts.length)target=charts[index]
    return {target:target,index:index}
  }
  try{
    if(!spec||typeof spec!=='object')return fail('invalid-operation','spec is required')
    if(!spec.sheet)return fail('invalid-operation','sheet is required')
    var sheet=sheetOf(spec.sheet); if(!sheet)return fail('sheet-not-found','worksheet not found',{sheet:spec.sheet})
    var charts=chartsOf(sheet); if(charts===null)return fail('unsupported','ApiWorksheet.GetAllCharts is unavailable')
    if(spec.type==='chart.inspect'){
      var inventory=[]; for(var i=0;i<charts.length;i++)inventory.push(describe(charts[i],i))
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,count:inventory.length,charts:inventory,verification:{status:'PASS',actual:inventory}}
    }
    if(spec.type==='chart.create'){
      if(!has(sheet,'AddChart'))return fail('unsupported','ApiWorksheet.AddChart is unavailable')
      if(!spec.range||!spec.chartType)return fail('invalid-operation','range and chartType are required')
      var before=charts.length, range=String(spec.range), q=String(spec.sheet).replace(/'/g,"''"), dataRange="'"+q+"'!"+range
      var style=spec.style==null?1:Number(spec.style), width=spec.width==null?3600000:Number(spec.width), height=spec.height==null?2160000:Number(spec.height)
      var fromCol=spec.fromCol==null?5:Number(spec.fromCol), colOffset=spec.colOffset==null?0:Number(spec.colOffset), fromRow=spec.fromRow==null?0:Number(spec.fromRow), rowOffset=spec.rowOffset==null?0:Number(spec.rowOffset)
      var chart=null
      try{chart=sheet.AddChart(dataRange,!!spec.inRows,String(spec.chartType),style,width,height,fromCol,colOffset,fromRow,rowOffset)}catch(err){return fail('operation-error',String(err&&err.message?err.message:err))}
      if(!chart)return fail('operation-error','ApiWorksheet.AddChart returned null')
      if(spec.name!=null){if(!has(chart,'SetName'))return fail('unsupported','ApiChart.SetName is unavailable');chart.SetName(String(spec.name))}
      if(spec.title!=null){if(!has(chart,'SetTitle'))return fail('unsupported','ApiChart.SetTitle is unavailable');chart.SetTitle(String(spec.title),spec.titleFontSize==null?12:Number(spec.titleFontSize),spec.titleBold!==false)}
      if(spec.legendPos!=null){if(!has(chart,'SetLegendPos'))return fail('unsupported','ApiChart.SetLegendPos is unavailable');chart.SetLegendPos(String(spec.legendPos))}
      var after=chartsOf(sheet); if(after===null)return fail('unsupported','chart readback unavailable after create')
      var actual=null, idx=-1
      if(spec.name!=null){for(var j=0;j<after.length;j++){if(String(safe(after[j],'GetName'))===String(spec.name)){actual=describe(after[j],j);idx=j;break}}}
      if(!actual&&after.length===before+1){idx=after.length-1;actual=describe(after[idx],idx)}
      var mismatches=[]
      if(after.length!==before+1)mismatches.push('count')
      if(!actual)mismatches.push('created-chart')
      if(actual&&actual.chartType!=null&&String(actual.chartType)!==String(spec.chartType))mismatches.push('chartType')
      if(spec.name!=null&&actual&&actual.name!=null&&String(actual.name)!==String(spec.name))mismatches.push('name')
      if(spec.title!=null&&actual&&actual.title!=null&&normText(actual.title)!==normText(spec.title))mismatches.push('title')
      if(actual&&actual.width!=null&&Number(actual.width)!==width)mismatches.push('width')
      if(actual&&actual.height!=null&&Number(actual.height)!==height)mismatches.push('height')
      if(mismatches.length)return fail('verification-failed','live chart readback mismatch',{beforeCount:before,afterCount:after.length,actual:actual,mismatches:mismatches})
      var unknown=[]
      if(actual.chartType==null)unknown.push('chartType'); if(spec.title!=null&&actual.title==null)unknown.push('title'); if(actual.width==null)unknown.push('width'); if(actual.height==null)unknown.push('height')
      if(unknown.length)return {ok:true,outcome:'unknown',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,beforeCount:before,afterCount:after.length,actual:actual,verification:{status:'UNKNOWN',reason:'required live chart getters unavailable',unknown:unknown}}
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,beforeCount:before,afterCount:after.length,actual:actual,verification:{status:'PASS',expected:{chartType:spec.chartType,name:spec.name||null,title:spec.title||null,width:width,height:height},actual:actual}}
    }
    if(spec.type==='chart.modify'){
      var found=targetOf(charts,spec), target=found.target, index=found.index
      if(!target)return fail('invalid-operation','chart target not found',{count:charts.length})
      var beforeDesc=describe(target,index), changed=[]
      if(spec.title!=null){if(!has(target,'SetTitle'))return fail('unsupported','ApiChart.SetTitle is unavailable');target.SetTitle(String(spec.title),spec.titleFontSize==null?12:Number(spec.titleFontSize),spec.titleBold!==false);changed.push('title')}
      if(spec.width!=null||spec.height!=null){
        if(!has(target,'SetSize'))return fail('unsupported','ApiChart.SetSize is unavailable')
        var newWidth=spec.width==null?beforeDesc.width:Number(spec.width),newHeight=spec.height==null?beforeDesc.height:Number(spec.height)
        if(newWidth==null||newHeight==null)return fail('unsupported','chart size readback is required for partial SetSize')
        target.SetSize(Number(newWidth),Number(newHeight));changed.push('size')
      }
      if(spec.legendPos!=null){if(!has(target,'SetLegendPos'))return fail('unsupported','ApiChart.SetLegendPos is unavailable');target.SetLegendPos(String(spec.legendPos));changed.push('legendPos')}
      if(!changed.length)return fail('invalid-operation','no supported chart modification requested')
      var afterCharts=chartsOf(sheet); if(afterCharts===null)return fail('unsupported','chart readback unavailable after modify')
      var reread=targetOf(afterCharts,spec), actual=reread.target?describe(reread.target,reread.index):null
      if(!actual)return fail('verification-failed','modified chart disappeared during live readback')
      var mm=[],unk=[]
      if(spec.title!=null){if(actual.title==null)unk.push('title');else if(normText(actual.title)!==normText(spec.title))mm.push('title')}
      if(spec.width!=null){if(actual.width==null)unk.push('width');else if(Number(actual.width)!==Number(spec.width))mm.push('width')}
      if(spec.height!=null){if(actual.height==null)unk.push('height');else if(Number(actual.height)!==Number(spec.height))mm.push('height')}
      if(mm.length)return fail('verification-failed','live chart modification readback mismatch',{before:beforeDesc,actual:actual,mismatches:mm,changed:changed})
      if(unk.length)return {ok:true,outcome:'unknown',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,before:beforeDesc,actual:actual,changed:changed,verification:{status:'UNKNOWN',reason:'required live chart getters unavailable',unknown:unk}}
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,before:beforeDesc,actual:actual,changed:changed,verification:{status:'PASS',actual:actual}}
    }
    if(spec.type==='chart.delete'){
      var f=targetOf(charts,spec), target=f.target, index=f.index
      if(!target)return fail('invalid-operation','chart target not found',{count:charts.length})
      var deleted=describe(target,index), deleteTarget=target, deleteVia='chart'
      if(!has(deleteTarget,'Delete')){
        var drawings=drawingsOf(sheet), drawing=null
        if(drawings!==null){
          for(var d=0;d<drawings.length;d++){
            var dn=safe(drawings[d],'GetName'), dc=safe(drawings[d],'GetClassType')
            if(spec.name!=null&&String(dn)===String(spec.name)){drawing=drawings[d];break}
            if(spec.name==null&&String(dc)==='chart'&&d===index){drawing=drawings[d];break}
          }
        }
        if(!drawing||!has(drawing,'Delete'))return fail('unsupported','chart deletion is unavailable on ApiChart and matching ApiDrawing')
        deleteTarget=drawing; deleteVia='drawing'
      }
      deleteTarget.Delete(); var remaining=chartsOf(sheet)
      if(remaining===null)return fail('unsupported','chart readback unavailable after delete')
      if(remaining.length!==charts.length-1)return fail('verification-failed','chart count did not decrease after delete',{beforeCount:charts.length,afterCount:remaining.length,deleteVia:deleteVia})
      if(spec.name!=null){for(var m=0;m<remaining.length;m++)if(String(safe(remaining[m],'GetName'))===String(spec.name))return fail('verification-failed','named chart still exists after delete')}
      return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:spec.type,sheet:spec.sheet,deleted:deleted,deleteVia:deleteVia,beforeCount:charts.length,afterCount:remaining.length,verification:{status:'PASS',expectedCount:charts.length-1,actualCount:remaining.length}}
    }
    return fail('invalid-operation','unknown chart operation: '+spec.type)
  }catch(err){return fail('chart-error',String(err&&err.message?err.message:err))}
}

async function runChartInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body=`return (${chartCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};if(!editor||typeof editor.callCommand!=='function')return finish({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'});try{editor.callCommand(new Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})
}

module.exports={chartCommand,runChartInFrame}
