'use strict'

function chartCapabilityCommand(sheetName) {
  function yes(o, name) { return !!o && typeof o[name] === 'function' }
  var sheet = Api.GetSheet(sheetName)
  if (!sheet) return { ok:false, outcome:'sheet-not-found', source:'live-coedit-editor', sheet:sheetName }
  var result = {
    ok:true,
    outcome:'capability-probe',
    source:'live-coedit-editor',
    sheet:sheetName,
    worksheetMethods:{
      AddChart:yes(sheet,'AddChart'),
      GetAllCharts:yes(sheet,'GetAllCharts'),
      GetAllDrawings:yes(sheet,'GetAllDrawings')
    },
    existingChartCount:null,
    chartMethods:null,
    error:null
  }
  if (result.worksheetMethods.GetAllCharts) {
    try {
      var charts=sheet.GetAllCharts() || []
      result.existingChartCount=charts.length
      if (charts.length) {
        var c=charts[0]
        result.chartMethods={
          GetChartType:yes(c,'GetChartType'),
          GetAllSeries:yes(c,'GetAllSeries'),
          GetSeries:yes(c,'GetSeries'),
          GetTitle:yes(c,'GetTitle'),
          SetTitle:yes(c,'SetTitle'),
          SetLegendPos:yes(c,'SetLegendPos'),
          SetHorAxisTitle:yes(c,'SetHorAxisTitle'),
          SetVerAxisTitle:yes(c,'SetVerAxisTitle'),
          GetWidth:yes(c,'GetWidth'),
          GetHeight:yes(c,'GetHeight'),
          Delete:yes(c,'Delete')
        }
      }
    } catch(e) { result.error=String(e&&e.message?e.message:e) }
  }
  return result
}

async function runChartCapabilityInFrame(frame, apiHely, sheetName, timeout=15000) {
  const body=`return (${chartCapabilityCommand.toString()})(${JSON.stringify(sheetName)});`
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{
    const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor
    let done=false
    const finish=v=>{if(!done){done=true;resolve(v)}}
    try { editor.callCommand(new Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v)) }
    catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}
    setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)
  }),{u:apiHely,body,timeout})
}

module.exports={chartCapabilityCommand,runChartCapabilityInFrame}
