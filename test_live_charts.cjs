'use strict'
const assert=require('assert')
const {chartCommand}=require('./live-charts.cjs')

function makeApi(){
  const charts=[]
  function mk(type,w,h){let name='Chart '+(charts.length+1),title=null;return {GetClassType:()=> 'chart',GetName:()=>name,SetName:v=>(name=v,true),GetChartType:()=>type,GetTitle:()=>title,SetTitle:v=>(title=v,true),SetLegendPos:()=>true,GetWidth:()=>w,GetHeight:()=>h,GetAllSeries:()=>[{}],Delete(){const i=charts.indexOf(this);if(i>=0)charts.splice(i,1);return true}}}
  const sheet={GetAllCharts:()=>charts.slice(),AddChart:(range,inRows,type,style,w,h)=>{const c=mk(type,w,h);charts.push(c);return c}}
  return {api:{GetSheet:n=>n==='Sheet1'?sheet:null},charts}
}

{
  const {api}=makeApi(); global.Api=api
  let r=chartCommand({type:'chart.inspect',sheet:'Sheet1'}); assert.equal(r.verification.status,'PASS'); assert.equal(r.count,0)
  r=chartCommand({type:'chart.create',sheet:'Sheet1',range:'A1:B3',chartType:'bar',name:'M51_CHART',title:'Revenue',width:3600000,height:2160000})
  assert.equal(r.verification.status,'PASS'); assert.equal(r.afterCount,1); assert.equal(r.actual.chartType,'bar'); assert.equal(r.actual.title,'Revenue'); assert.equal(r.actual.name,'M51_CHART')
  r=chartCommand({type:'chart.inspect',sheet:'Sheet1'}); assert.equal(r.count,1); assert.equal(r.charts[0].name,'M51_CHART')
  r=chartCommand({type:'chart.delete',sheet:'Sheet1',name:'M51_CHART'}); assert.equal(r.verification.status,'PASS'); assert.equal(r.afterCount,0)
}

{
  const {api}=makeApi(); global.Api=api
  const serialized=new Function('Api',`return (${chartCommand.toString()})({type:'chart.inspect',sheet:'Sheet1'});`)(api)
  assert.equal(serialized.verification.status,'PASS')
}

console.log('test_live_charts.cjs: OK')
