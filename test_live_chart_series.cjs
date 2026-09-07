'use strict'
const assert=require('assert')
const {chartSeriesCommand}=require('./live-chart-series.cjs')

function makeApi(){
  const mkSeries=type=>{let t=type;return {GetClassType:()=> 'chartSeries',GetChartType:()=>t,ChangeChartType:v=>(t=v,true)}}
  const charts=[]
  const chart={
    GetName:()=> 'M52_CHART',
    GetAllSeries:()=>series.slice(),
    AddSeria:()=>{series.push(mkSeries('bar'))},
    RemoveSeria:i=>{series.splice(i,1);return true}
  }
  const series=[mkSeries('bar'),mkSeries('line')]
  charts.push(chart)
  const sheet={GetAllCharts:()=>charts.slice()}
  return {GetSheet:n=>n==='Sheet1'?sheet:null}
}

function main(){
  global.Api=makeApi()
  let r=chartSeriesCommand({type:'chart.series.inspect',sheet:'Sheet1',name:'M52_CHART'})
  assert.equal(r.verification.status,'PASS');assert.equal(r.count,2);assert.equal(r.series[0].chartType,'bar');assert.equal(r.series[1].chartType,'line')

  r=chartSeriesCommand({type:'chart.series.changeType',sheet:'Sheet1',name:'M52_CHART',seriesIndex:0,chartType:'area'})
  assert.equal(r.verification.status,'PASS');assert.equal(r.beforeType,'bar');assert.equal(r.actualType,'area')

  r=chartSeriesCommand({type:'chart.series.add',sheet:'Sheet1',name:'M52_CHART',nameRange:"'Sheet1'!$D$1",valuesRange:"'Sheet1'!$D$2:$D$4"})
  assert.equal(r.verification.status,'PASS');assert.equal(r.beforeCount,2);assert.equal(r.afterCount,3)

  r=chartSeriesCommand({type:'chart.series.remove',sheet:'Sheet1',name:'M52_CHART',seriesIndex:2})
  assert.equal(r.verification.status,'PASS');assert.equal(r.beforeCount,3);assert.equal(r.afterCount,2)

  r=chartSeriesCommand({type:'chart.series.inspect',sheet:'Sheet1',name:'M52_CHART'})
  assert.equal(r.count,2);assert.equal(r.series[0].chartType,'area')

  const serialized=new Function('Api',`return (${chartSeriesCommand.toString()})({type:'chart.series.inspect',sheet:'Sheet1',name:'M52_CHART'});`)(global.Api)
  assert.equal(serialized.verification.status,'PASS');assert.equal(serialized.count,2)
  console.log('test_live_chart_series.cjs: OK')
}

try{main()}catch(err){console.error(err);process.exitCode=1}
