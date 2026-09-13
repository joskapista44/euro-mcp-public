'use strict'
const assert=require('assert')
const {chartDataObjectCommand}=require('./live-chart-data-objects.cjs')

function makeApi(opts={}){
  const series={
    name:'OLD SERIES',
    values:'Sheet1!$A$1:$A$3',
    xValues:'Sheet1!$B$1:$B$3',
    category:'Sheet1!$C$1:$C$3',

    GetName(){
      if(opts.seriesNameReadbackMismatch)return '="WRONG"'
      return '="'+String(this.name).replace(/"/g,'""')+'"'
    },
    GetValues(){
      return opts.seriesValuesReadbackMismatch
        ? '=Sheet1!$Z$1:$Z$3'
        : '='+this.values
    },
    GetXValues(){
      return opts.seriesXReadbackMismatch
        ? '=Sheet1!$Z$1:$Z$3'
        : '='+this.xValues
    },
    GetCatFormula(){
      return opts.categoryReadbackMismatch
        ? '=Sheet1!$Z$1:$Z$3'
        : '='+this.category
    }
  }

  if(opts.noSeriesNameGetter)delete series.GetName
  if(opts.noSeriesValuesGetter)delete series.GetValues
  if(opts.noSeriesXGetter)delete series.GetXValues
  if(opts.noCategoryGetter)delete series.GetCatFormula

  const chart={
    name:'M54_CHART',
    width:4000000,
    height:2400000,
    position:{fromCol:0,colOffset:0,fromRow:0,rowOffset:0},

    GetName(){return opts.nameReadbackMismatch?'WRONG':this.name},
    SetName(v){this.name=v},

    GetWidth(){return opts.sizeReadbackMismatch?123:this.width},
    GetHeight(){return opts.sizeReadbackMismatch?456:this.height},
    SetSize(w,h){this.width=w;this.height=h},

    SetPosition(fromCol,colOffset,fromRow,rowOffset){
      this.position={fromCol,colOffset,fromRow,rowOffset}
      this.positionMutated=true
    },
    GetPosition(){
      return opts.positionReadbackMismatch
        ? {fromCol:99,colOffset:0,fromRow:99,rowOffset:0}
        : this.position
    },

    GetAllSeries(){return [series]},

    SetSeriaName(v,i){
      assert.equal(i,0)
      series.name=v
      this.seriesNameMutated=true
    },
    SetSeriaValues(v,i){
      assert.equal(i,0)
      series.values=v
      this.seriesValuesMutated=true
    },
    SetSeriaXValues(v,i){
      assert.equal(i,0)
      series.xValues=v
      this.seriesXValuesMutated=true
    },
    SetCatFormula(v){
      series.category=v
      this.categoryMutated=true
    }
  }

  if(opts.noPositionGetter)delete chart.GetPosition

  const sheet={GetAllCharts:()=>[chart]}
  return {
    api:{GetSheet:n=>n==='Sheet1'?sheet:null},
    chart,
    series
  }
}

function main(){
  let env=makeApi()
  global.Api=env.api

  let r=chartDataObjectCommand({
    type:'chart.object.inspect',
    sheet:'Sheet1',
    name:'M54_CHART'
  })

  assert.equal(r.verification.status,'PASS')
  assert.equal(r.capabilities.rename,true)
  assert.equal(r.capabilities.resize,true)
  assert.equal(r.capabilities.positionSetter,true)
  assert.equal(r.capabilities.positionReadback,true)
  assert.equal(r.capabilities.copy,false)
  assert.equal(r.capabilities.seriesNameReadback,true)
  assert.equal(r.capabilities.seriesValuesReadback,true)
  assert.equal(r.capabilities.seriesXValuesReadback,true)
  assert.equal(r.capabilities.categoryFormulaReadback,true)

  r=chartDataObjectCommand({
    type:'chart.object.rename',
    sheet:'Sheet1',
    name:'M54_CHART',
    newName:'M54_RENAMED'
  })
  assert.equal(r.verification.status,'PASS')
  assert.equal(r.actualName,'M54_RENAMED')

  r=chartDataObjectCommand({
    type:'chart.object.resize',
    sheet:'Sheet1',
    name:'M54_RENAMED',
    width:5000000,
    height:3000000
  })
  assert.equal(r.verification.status,'PASS')
  assert.deepEqual(r.actual,{width:5000000,height:3000000})

  r=chartDataObjectCommand({
    type:'chart.object.position',
    sheet:'Sheet1',
    name:'M54_RENAMED',
    fromCol:2,
    colOffset:360000,
    fromRow:3,
    rowOffset:720000
  })
  assert.equal(r.verification.status,'PASS')
  assert.deepEqual(r.actual,{
    fromCol:2,
    colOffset:360000,
    fromRow:3,
    rowOffset:720000
  })

  r=chartDataObjectCommand({
    type:'chart.data.seriesName',
    sheet:'Sheet1',
    name:'M54_RENAMED',
    seriesIndex:0,
    value:'EURO "Series"'
  })
  assert.equal(r.verification.status,'PASS')
  assert.equal(r.normalizedActual,'EURO "Series"')

  r=chartDataObjectCommand({
    type:'chart.data.seriesValues',
    sheet:'Sheet1',
    name:'M54_RENAMED',
    seriesIndex:0,
    range:'Sheet1!$D$1:$D$3'
  })
  assert.equal(r.verification.status,'PASS')
  assert.equal(r.normalizedActual,'Sheet1!$D$1:$D$3')

  r=chartDataObjectCommand({
    type:'chart.data.seriesXValues',
    sheet:'Sheet1',
    name:'M54_RENAMED',
    seriesIndex:0,
    range:'Sheet1!$E$1:$E$3'
  })
  assert.equal(r.verification.status,'PASS')
  assert.equal(r.normalizedActual,'Sheet1!$E$1:$E$3')

  r=chartDataObjectCommand({
    type:'chart.data.categoryFormula',
    sheet:'Sheet1',
    name:'M54_RENAMED',
    seriesIndex:0,
    range:'Sheet1!$F$1:$F$3'
  })
  assert.equal(r.verification.status,'PASS')
  assert.equal(r.normalizedActual,'Sheet1!$F$1:$F$3')

  r=chartDataObjectCommand({
    type:'chart.object.copy',
    sheet:'Sheet1',
    name:'M54_RENAMED'
  })
  assert.equal(r.outcome,'unsupported')
  assert.equal(r.ok,false)

  /* Getter missing => fail closed, NO mutation. */
  env=makeApi({noPositionGetter:true})
  global.Api=env.api
  r=chartDataObjectCommand({
    type:'chart.object.position',
    sheet:'Sheet1',
    name:'M54_CHART',
    fromCol:2,
    colOffset:1,
    fromRow:3,
    rowOffset:2
  })
  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.positionMutated,undefined)

  env=makeApi({noSeriesNameGetter:true})
  global.Api=env.api
  r=chartDataObjectCommand({
    type:'chart.data.seriesName',
    sheet:'Sheet1',
    name:'M54_CHART',
    value:'SHOULD NOT MUTATE'
  })
  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.seriesNameMutated,undefined)

  env=makeApi({noSeriesValuesGetter:true})
  global.Api=env.api
  r=chartDataObjectCommand({
    type:'chart.data.seriesValues',
    sheet:'Sheet1',
    name:'M54_CHART',
    range:'Sheet1!$D$1:$D$3'
  })
  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.seriesValuesMutated,undefined)

  /* Wrong semantic readback => FAIL. */
  env=makeApi({positionReadbackMismatch:true})
  global.Api=env.api
  r=chartDataObjectCommand({
    type:'chart.object.position',
    sheet:'Sheet1',
    name:'M54_CHART',
    fromCol:2,
    colOffset:360000,
    fromRow:3,
    rowOffset:720000
  })
  assert.equal(r.outcome,'verification-failed')
  assert.equal(r.ok,false)

  env=makeApi({seriesValuesReadbackMismatch:true})
  global.Api=env.api
  r=chartDataObjectCommand({
    type:'chart.data.seriesValues',
    sheet:'Sheet1',
    name:'M54_CHART',
    range:'Sheet1!$D$1:$D$3'
  })
  assert.equal(r.outcome,'verification-failed')
  assert.equal(r.ok,false)

  env=makeApi({nameReadbackMismatch:true})
  global.Api=env.api
  r=chartDataObjectCommand({
    type:'chart.object.rename',
    sheet:'Sheet1',
    index:0,
    newName:'EXPECTED'
  })
  assert.equal(r.outcome,'verification-failed')
  assert.equal(r.ok,false)

  env=makeApi({sizeReadbackMismatch:true})
  global.Api=env.api
  r=chartDataObjectCommand({
    type:'chart.object.resize',
    sheet:'Sheet1',
    index:0,
    width:5000000,
    height:3000000
  })
  assert.equal(r.outcome,'verification-failed')
  assert.equal(r.ok,false)

  /* Serialization/callCommand compatibility. */
  env=makeApi()
  global.Api=env.api
  const serialized=new Function(
    'Api',
    'return ('+chartDataObjectCommand.toString()+')({type:"chart.object.inspect",sheet:"Sheet1",name:"M54_CHART"});'
  )(env.api)
  assert.equal(serialized.verification.status,'PASS')
  assert.equal(serialized.capabilities.positionReadback,true)
  assert.equal(serialized.capabilities.seriesNameReadback,true)

  console.log('test_live_chart_data_objects.cjs: OK')
}

try{
  main()
}catch(err){
  console.error(err)
  process.exitCode=1
}
