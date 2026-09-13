'use strict'

const assert=require('assert')
const {chartPresentationCommand}=require('./live-chart-presentation.cjs')

function makeApi(opts={}){
  const chart={
    name:'M53_CHART',
    legend:'bottom',
    h:'OLD H',
    v:'OLD V',
    labels:{
      showSerName:false,
      showCatName:false,
      showVal:false,
      showPercent:false
    },
    style:0,

    GetName(){
      return this.name
    },

    SetLegendPos(v){
      this.legend=v
      this.legendMutated=true
    },
    GetLegendPos(){
      return opts.legendMismatch?'left':this.legend
    },

    SetLegendFontSize(v){
      this.legendFont=v
      this.legendFontMutated=true
    },

    SetHorAxisTitle(v){
      this.h=v
      this.horizontalMutated=true
    },
    GetHorAxisTitle(){
      return opts.horizontalMismatch
        ? 'WRONG\r\n'
        : this.h+'\r\n'
    },

    SetVerAxisTitle(v){
      this.v=v
      this.verticalMutated=true
    },
    GetVerAxisTitle(){
      return opts.verticalMismatch
        ? 'WRONG\r\n'
        : this.v+'\r\n'
    },

    SetShowDataLabels(a,b,c,d){
      this.labels={
        showSerName:a,
        showCatName:b,
        showVal:c,
        showPercent:d
      }
      this.labelsMutated=true
    },
    GetDataLabels(){
      if(opts.labelsMismatch){
        return {
          showSerName:false,
          showCatName:false,
          showVal:false,
          showPercent:false
        }
      }
      return this.labels
    },

    ApplyChartStyle(v){
      this.style=v
      this.styleMutated=true
      return true
    },
    GetChartStyle(){
      return opts.styleMismatch?99:this.style
    }
  }

  if(opts.noLegendGetter)delete chart.GetLegendPos
  if(opts.noHorGetter)delete chart.GetHorAxisTitle
  if(opts.noVerGetter)delete chart.GetVerAxisTitle
  if(opts.noLabelsGetter)delete chart.GetDataLabels
  if(opts.noStyleGetter)delete chart.GetChartStyle

  const sheet={
    GetAllCharts:()=>[chart]
  }

  return {
    api:{
      GetSheet:n=>n==='Sheet1'?sheet:null
    },
    chart
  }
}

function main(){
  let env=makeApi()
  global.Api=env.api

  /* Capability inspection must see the new public getters. */
  let r=chartPresentationCommand({
    type:'chart.presentation.inspect',
    sheet:'Sheet1',
    name:'M53_CHART'
  })

  assert.equal(r.verification.status,'PASS')
  assert.equal(r.support.SetLegendPos,true)
  assert.equal(r.support.GetLegendPos,true)
  assert.equal(r.support.GetHorAxisTitle,true)
  assert.equal(r.support.GetVerAxisTitle,true)
  assert.equal(r.support.GetDataLabels,true)
  assert.equal(r.support.GetChartStyle,true)

  /* Legend position: exact public readback. */
  r=chartPresentationCommand({
    type:'chart.presentation.legend',
    sheet:'Sheet1',
    name:'M53_CHART',
    position:'right'
  })

  assert.equal(r.verification.status,'PASS')
  assert.equal(r.verification.expected.position,'right')
  assert.equal(r.verification.actual.position,'right')
  assert.equal(env.chart.legend,'right')

  /*
   * Legend font size remains fail-closed because no matching
   * public semantic getter has been established.
   */
  const legendFontBefore=env.chart.legendFont

  r=chartPresentationCommand({
    type:'chart.presentation.legend',
    sheet:'Sheet1',
    name:'M53_CHART',
    fontSize:14
  })

  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.legendFont,legendFontBefore)
  assert.equal(env.chart.legendFontMutated,undefined)

  /* Axis titles: public getters include paragraph terminators. */
  r=chartPresentationCommand({
    type:'chart.presentation.axisTitles',
    sheet:'Sheet1',
    name:'M53_CHART',
    horizontal:'Quarter',
    vertical:'EUR'
  })

  assert.equal(r.verification.status,'PASS')
  assert.deepEqual(
    r.verification.expected,
    {horizontal:'Quarter',vertical:'EUR'}
  )
  assert.deepEqual(
    r.verification.actual,
    {horizontal:'Quarter',vertical:'EUR'}
  )
  assert.equal(env.chart.h,'Quarter')
  assert.equal(env.chart.v,'EUR')

  /* Data-label semantic object readback. */
  r=chartPresentationCommand({
    type:'chart.presentation.dataLabels',
    sheet:'Sheet1',
    name:'M53_CHART',
    showSeriesName:true,
    showCategoryName:true,
    showValue:true,
    showPercent:false
  })

  assert.equal(r.verification.status,'PASS')
  assert.deepEqual(
    r.verification.actual,
    {
      showSerName:true,
      showCatName:true,
      showVal:true,
      showPercent:false
    }
  )

  /* Chart style: public input/output uses the same 0-based index. */
  r=chartPresentationCommand({
    type:'chart.presentation.style',
    sheet:'Sheet1',
    name:'M53_CHART',
    style:2
  })

  assert.equal(r.verification.status,'PASS')
  assert.equal(r.verification.expected.style,2)
  assert.equal(r.verification.actual.style,2)
  assert.equal(env.chart.style,2)

  /*
   * Getter missing => UNKNOWN before mutation.
   */

  env=makeApi({noLegendGetter:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.legend',
    sheet:'Sheet1',
    name:'M53_CHART',
    position:'left'
  })

  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.legendMutated,undefined)
  assert.equal(env.chart.legend,'bottom')

  env=makeApi({noHorGetter:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.axisTitles',
    sheet:'Sheet1',
    name:'M53_CHART',
    horizontal:'SHOULD NOT MUTATE'
  })

  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.horizontalMutated,undefined)
  assert.equal(env.chart.h,'OLD H')

  env=makeApi({noVerGetter:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.axisTitles',
    sheet:'Sheet1',
    name:'M53_CHART',
    vertical:'SHOULD NOT MUTATE'
  })

  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.verticalMutated,undefined)
  assert.equal(env.chart.v,'OLD V')

  env=makeApi({noLabelsGetter:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.dataLabels',
    sheet:'Sheet1',
    name:'M53_CHART',
    showValue:true
  })

  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.labelsMutated,undefined)

  env=makeApi({noStyleGetter:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.style',
    sheet:'Sheet1',
    name:'M53_CHART',
    style:2
  })

  assert.equal(r.verification.status,'UNKNOWN')
  assert.equal(env.chart.styleMutated,undefined)
  assert.equal(env.chart.style,0)

  /*
   * Wrong public semantic readback => verification-failed.
   */

  env=makeApi({legendMismatch:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.legend',
    sheet:'Sheet1',
    name:'M53_CHART',
    position:'right'
  })

  assert.equal(r.ok,false)
  assert.equal(r.outcome,'verification-failed')

  env=makeApi({horizontalMismatch:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.axisTitles',
    sheet:'Sheet1',
    name:'M53_CHART',
    horizontal:'Quarter'
  })

  assert.equal(r.ok,false)
  assert.equal(r.outcome,'verification-failed')

  env=makeApi({verticalMismatch:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.axisTitles',
    sheet:'Sheet1',
    name:'M53_CHART',
    vertical:'EUR'
  })

  assert.equal(r.ok,false)
  assert.equal(r.outcome,'verification-failed')

  env=makeApi({labelsMismatch:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.dataLabels',
    sheet:'Sheet1',
    name:'M53_CHART',
    showSeriesName:true,
    showCategoryName:true,
    showValue:true
  })

  assert.equal(r.ok,false)
  assert.equal(r.outcome,'verification-failed')

  env=makeApi({styleMismatch:true})
  global.Api=env.api

  r=chartPresentationCommand({
    type:'chart.presentation.style',
    sheet:'Sheet1',
    name:'M53_CHART',
    style:2
  })

  assert.equal(r.ok,false)
  assert.equal(r.outcome,'verification-failed')

  /* callCommand serialization compatibility. */
  env=makeApi()
  global.Api=env.api

  const serialized=new Function(
    'Api',
    `return (${chartPresentationCommand.toString()})({type:'chart.presentation.inspect',sheet:'Sheet1',name:'M53_CHART'});`
  )(env.api)

  assert.equal(serialized.verification.status,'PASS')
  assert.equal(serialized.support.GetLegendPos,true)
  assert.equal(serialized.support.GetHorAxisTitle,true)
  assert.equal(serialized.support.GetVerAxisTitle,true)
  assert.equal(serialized.support.GetDataLabels,true)
  assert.equal(serialized.support.GetChartStyle,true)

  console.log('test_live_chart_presentation.cjs: OK')
}

try{
  main()
}catch(err){
  console.error(err)
  process.exitCode=1
}
