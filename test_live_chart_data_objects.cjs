'use strict'
const assert=require('assert')
const {chartDataObjectCommand}=require('./live-chart-data-objects.cjs')

function makeApi(opts={}){
  const chart={
    name:'M54_CHART',width:4000000,height:2400000,
    GetName(){return opts.nameReadbackMismatch?'WRONG':this.name},
    SetName(v){this.name=v},
    GetWidth(){return opts.sizeReadbackMismatch?123:this.width},
    GetHeight(){return opts.sizeReadbackMismatch?456:this.height},
    SetSize(w,h){this.width=w;this.height=h},
    SetPosition(){this.positionMutated=true},
    SetSeriaName(){this.seriesNameMutated=true},
    SetSeriaValues(){this.seriesValuesMutated=true},
    SetSeriaXValues(){this.seriesXValuesMutated=true},
    SetCatFormula(){this.categoryMutated=true}
  }
  const sheet={GetAllCharts:()=>[chart]}
  return {api:{GetSheet:n=>n==='Sheet1'?sheet:null},chart}
}

function main(){
  let env=makeApi();global.Api=env.api
  let r=chartDataObjectCommand({type:'chart.object.inspect',sheet:'Sheet1',name:'M54_CHART'})
  assert.equal(r.verification.status,'PASS');assert.equal(r.capabilities.rename,true);assert.equal(r.capabilities.resize,true);assert.equal(r.capabilities.positionSetter,true);assert.equal(r.capabilities.positionReadback,false);assert.equal(r.capabilities.copy,false)

  r=chartDataObjectCommand({type:'chart.object.rename',sheet:'Sheet1',name:'M54_CHART',newName:'M54_RENAMED'})
  assert.equal(r.verification.status,'PASS');assert.equal(r.actualName,'M54_RENAMED');assert.equal(env.chart.name,'M54_RENAMED')

  r=chartDataObjectCommand({type:'chart.object.resize',sheet:'Sheet1',name:'M54_RENAMED',width:5000000,height:3000000})
  assert.equal(r.verification.status,'PASS');assert.deepEqual(r.actual,{width:5000000,height:3000000})

  r=chartDataObjectCommand({type:'chart.object.position',sheet:'Sheet1',name:'M54_RENAMED'})
  assert.equal(r.verification.status,'UNKNOWN');assert.equal(env.chart.positionMutated,undefined)

  for(const type of ['chart.data.seriesName','chart.data.seriesValues','chart.data.seriesXValues','chart.data.categoryFormula']){
    r=chartDataObjectCommand({type,sheet:'Sheet1',name:'M54_RENAMED'});assert.equal(r.verification.status,'UNKNOWN')
  }
  assert.equal(env.chart.seriesNameMutated,undefined);assert.equal(env.chart.seriesValuesMutated,undefined);assert.equal(env.chart.seriesXValuesMutated,undefined);assert.equal(env.chart.categoryMutated,undefined)

  r=chartDataObjectCommand({type:'chart.object.copy',sheet:'Sheet1',name:'M54_RENAMED'});assert.equal(r.outcome,'unsupported');assert.equal(r.ok,false)

  env=makeApi({nameReadbackMismatch:true});global.Api=env.api
  r=chartDataObjectCommand({type:'chart.object.rename',sheet:'Sheet1',index:0,newName:'EXPECTED'});assert.equal(r.outcome,'verification-failed');assert.equal(r.ok,false)

  env=makeApi({sizeReadbackMismatch:true});global.Api=env.api
  r=chartDataObjectCommand({type:'chart.object.resize',sheet:'Sheet1',index:0,width:5000000,height:3000000});assert.equal(r.outcome,'verification-failed');assert.equal(r.ok,false)

  env=makeApi();global.Api=env.api
  const serialized=new Function('Api','return ('+chartDataObjectCommand.toString()+')({type:"chart.object.inspect",sheet:"Sheet1",name:"M54_CHART"});')(env.api)
  assert.equal(serialized.verification.status,'PASS')
  console.log('test_live_chart_data_objects.cjs: OK')
}
try{main()}catch(err){console.error(err);process.exitCode=1}
