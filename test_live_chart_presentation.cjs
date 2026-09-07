'use strict'
const assert=require('assert')
const {chartPresentationCommand}=require('./live-chart-presentation.cjs')

function makeApi(withLegendGetter=false){
  const chart={
    GetName:()=> 'M53_CHART',
    SetLegendPos:v=>{chart.legend=v},
    SetLegendFontSize:v=>{chart.legendFont=v},
    SetHorAxisTitle:v=>{chart.h=v},
    SetVerAxisTitle:v=>{chart.v=v},
    SetShowDataLabels:(a,b,c,d)=>{chart.labels=[a,b,c,d]},
    ApplyChartStyle:v=>{chart.style=v}
  }
  if(withLegendGetter)chart.GetLegendPos=()=>chart.legend
  const sheet={GetAllCharts:()=>[chart]}
  return {api:{GetSheet:n=>n==='Sheet1'?sheet:null},chart}
}

function main(){
  let env=makeApi(false);global.Api=env.api
  let r=chartPresentationCommand({type:'chart.presentation.inspect',sheet:'Sheet1',name:'M53_CHART'})
  assert.equal(r.verification.status,'PASS');assert.equal(r.support.SetLegendPos,true);assert.equal(r.support.GetLegendPos,false)

  r=chartPresentationCommand({type:'chart.presentation.legend',sheet:'Sheet1',name:'M53_CHART',position:'bottom'})
  assert.equal(env.chart.legend,'bottom');assert.equal(r.verification.status,'UNKNOWN');assert.notEqual(r.outcome,'ok')

  r=chartPresentationCommand({type:'chart.presentation.axisTitles',sheet:'Sheet1',name:'M53_CHART',horizontal:'Quarter',vertical:'EUR'})
  assert.equal(env.chart.h,'Quarter');assert.equal(env.chart.v,'EUR');assert.equal(r.verification.status,'UNKNOWN')

  r=chartPresentationCommand({type:'chart.presentation.dataLabels',sheet:'Sheet1',name:'M53_CHART',showValue:true})
  assert.deepEqual(env.chart.labels,[false,false,true,false]);assert.equal(r.verification.status,'UNKNOWN')

  r=chartPresentationCommand({type:'chart.presentation.style',sheet:'Sheet1',name:'M53_CHART',style:5})
  assert.equal(env.chart.style,5);assert.equal(r.verification.status,'UNKNOWN')

  env=makeApi(true);global.Api=env.api
  r=chartPresentationCommand({type:'chart.presentation.legend',sheet:'Sheet1',name:'M53_CHART',position:'right'})
  assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.position,'right')

  const serialized=new Function('Api',`return (${chartPresentationCommand.toString()})({type:'chart.presentation.inspect',sheet:'Sheet1',name:'M53_CHART'});`)(env.api)
  assert.equal(serialized.verification.status,'PASS')
  console.log('test_live_chart_presentation.cjs: OK')
}
try{main()}catch(err){console.error(err);process.exitCode=1}
