'use strict'
const assert=require('assert')
const {chartCommand}=require('./live-charts.cjs')

function makeApi(opts={}){
  const charts=[]
  function mk(type,w,h){
    let name='Chart '+(charts.length+1),title=null,selected=false
    const c={
      GetClassType:()=> 'chart',GetName:()=>name,SetName:v=>(name=v,true),GetChartType:()=>type,
      GetTitle:()=>title,SetTitle:v=>(title=v,true),SetLegendPos:()=>true,GetWidth:()=>w,GetHeight:()=>h,
      GetAllSeries:()=>[{}],SetSize:(nw,nh)=>(w=nw,h=nh,true),Select:()=>{selected=true;return true},
      _selected:()=>selected
    }
    if(opts.directDelete!==false)c.Delete=function(){const i=charts.indexOf(this);if(i>=0)charts.splice(i,1);return true}
    return c
  }
  const sheet={GetAllCharts:()=>charts.slice(),GetAllDrawings:()=>charts.slice(),AddChart:(range,inRows,type,style,w,h)=>{const c=mk(type,w,h);charts.push(c);return c}}
  return {api:{GetSheet:n=>n==='Sheet1'?sheet:null},charts}
}

{
  const {api}=makeApi(); global.Api=api
  let r=chartCommand({type:'chart.inspect',sheet:'Sheet1'}); assert.equal(r.verification.status,'PASS'); assert.equal(r.count,0)
  r=chartCommand({type:'chart.create',sheet:'Sheet1',range:'A1:B3',chartType:'bar',name:'M51_CHART',title:'Revenue',width:3600000,height:2160000})
  assert.equal(r.verification.status,'PASS'); assert.equal(r.afterCount,1); assert.equal(r.actual.chartType,'bar'); assert.equal(r.actual.title,'Revenue'); assert.equal(r.actual.name,'M51_CHART')
  r=chartCommand({type:'chart.modify',sheet:'Sheet1',name:'M51_CHART',title:'Revenue verified',width:4000000,height:2400000})
  assert.equal(r.verification.status,'PASS'); assert.equal(r.actual.title,'Revenue verified'); assert.equal(r.actual.width,4000000); assert.equal(r.actual.height,2400000)
  r=chartCommand({type:'chart.inspect',sheet:'Sheet1'}); assert.equal(r.count,1); assert.equal(r.charts[0].name,'M51_CHART')
  r=chartCommand({type:'chart.delete',sheet:'Sheet1',name:'M51_CHART'}); assert.equal(r.verification.status,'PASS'); assert.equal(r.afterCount,0); assert.equal(r.deleteVia,'chart')
}

{
  const {api,charts}=makeApi({directDelete:false}); global.Api=api
  let r=chartCommand({type:'chart.create',sheet:'Sheet1',range:'A1:B3',chartType:'bar',name:'M51_SELECT_DELETE',title:'Revenue',width:3600000,height:2160000})
  assert.equal(r.verification.status,'PASS')
  r=chartCommand({type:'chart.delete',sheet:'Sheet1',name:'M51_SELECT_DELETE'})
  assert.equal(r.ok,true)
  assert.equal(r.outcome,'editor-delete-key-required')
  assert.equal(r.verification.status,'PENDING')
  assert.equal(r.deleteVia,'public-select+editor-delete-key')
  assert.equal(r.beforeCount,1)
  assert.equal(charts[0]._selected(),true)
  assert.equal(charts.length,1,'selection phase must never claim or simulate deletion')
}

{
  const {api}=makeApi(); global.Api=api
  const serialized=new Function('Api',`return (${chartCommand.toString()})({type:'chart.inspect',sheet:'Sheet1'});`)(api)
  assert.equal(serialized.verification.status,'PASS')
}

console.log('test_live_charts.cjs: OK')
