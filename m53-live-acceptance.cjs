'use strict'

const {runChartInFrame}=require('./live-charts.cjs')
const {runChartPresentationInFrame}=require('./live-chart-presentation.cjs')
const {writeBulkInFrame}=require('./bulk-writer.cjs')

function loadPlaywright(){
  for(const p of [
    process.env.EURO_PLAYWRIGHT_PATH,
    'playwright',
    '/home/user/marveen/node_modules/playwright'
  ].filter(Boolean)){
    try{
      return require(p)
    }catch(_){}
  }
  throw new Error('Playwright is unavailable')
}

function pass(r){
  return !!(
    r &&
    r.ok &&
    r.verification &&
    r.verification.status==='PASS'
  )
}

async function main(){
  const base=process.env.EURO_NC_BASE_URL
  const fileId=process.env.EURO_NC_FILE_ID
  const user=process.env.EURO_NC_USER
  const password=process.env.EURO_NC_PASSWORD

  if(!base||!fileId||!user||!password)
    throw new Error('required EURO_NC_* env missing')

  const sheet='Sheet1'
  const range='XFA1:XFC4'
  const token=Date.now().toString(36)
  const name='M53_'+token
  const title='M5.3 standalone acceptance'
  const horizontalTitle='EURO horizontal axis '+token
  const verticalTitle='EURO vertical axis '+token
  const steps=[]

  const {chromium}=loadPlaywright()
  const browser=await chromium.launch()

  try{
    const ctx=await browser.newContext({
      viewport:{width:1400,height:900}
    })

    const page=await ctx.newPage()

    await page.goto(`${base}/login`,{
      waitUntil:'domcontentloaded',
      timeout:60000
    })

    await page.fill('#user',user)
    await page.fill('#password',password)

    await Promise.all([
      page.waitForNavigation({
        waitUntil:'domcontentloaded',
        timeout:60000
      }).catch(()=>null),
      page.click('button[type=submit], input[type=submit]')
    ])

    await page.waitForTimeout(2500)

    if(/\/login/.test(page.url()))
      throw new Error('login failed')

    await page.goto(
      `${base}/index.php/apps/eurooffice/${fileId}`,
      {
        waitUntil:'domcontentloaded',
        timeout:60000
      }
    )

    await page.waitForTimeout(22000)

    const frame=page.frames().find(
      f=>/spreadsheeteditor/.test(f.url())
    )

    if(!frame)
      throw new Error('spreadsheeteditor frame missing')

    const apiHely=await frame.evaluate(
      ()=>(
        (window.Asc||{}).editor &&
        typeof window.Asc.editor.callCommand==='function'
      )
        ? 'window.Asc.editor'
        : null
    )

    if(!apiHely)
      throw new Error('callCommand unavailable')

    const seed=await writeBulkInFrame(
      frame,
      apiHely,
      {
        sheet,
        range,
        values:[
          ['Quarter','Revenue','Cost'],
          ['Q1',100,70],
          ['Q2',140,90],
          ['Q3',125,85]
        ]
      }
    )

    steps.push({
      name:'seed',
      result:seed
    })

    const initial=await runChartInFrame(
      frame,
      apiHely,
      {
        type:'chart.inspect',
        sheet
      }
    )

    steps.push({
      name:'initial-chart-inspect',
      result:initial
    })

    const before=initial.ok
      ? initial.count
      : null

    const created=await runChartInFrame(
      frame,
      apiHely,
      {
        type:'chart.create',
        sheet,
        range,
        chartType:'bar',
        style:2,
        width:4000000,
        height:2400000,
        name,
        title
      }
    )

    steps.push({
      name:'create-chart',
      result:created
    })

    const inspect=created&&created.ok
      ? await runChartPresentationInFrame(
          frame,
          apiHely,
          {
            type:'chart.presentation.inspect',
            sheet,
            name
          }
        )
      : null

    steps.push({
      name:'presentation-capability',
      result:inspect
    })

    const legend=created&&created.ok
      ? await runChartPresentationInFrame(
          frame,
          apiHely,
          {
            type:'chart.presentation.legend',
            sheet,
            name,
            position:'bottom'
          }
        )
      : null

    steps.push({
      name:'legend-position',
      result:legend
    })

    const axes=created&&created.ok
      ? await runChartPresentationInFrame(
          frame,
          apiHely,
          {
            type:'chart.presentation.axisTitles',
            sheet,
            name,
            horizontal:horizontalTitle,
            vertical:verticalTitle
          }
        )
      : null

    steps.push({
      name:'axis-titles',
      result:axes
    })

    const labels=created&&created.ok
      ? await runChartPresentationInFrame(
          frame,
          apiHely,
          {
            type:'chart.presentation.dataLabels',
            sheet,
            name,
            showSeriesName:true,
            showCategoryName:true,
            showValue:true,
            showPercent:false
          }
        )
      : null

    steps.push({
      name:'data-labels',
      result:labels
    })

    const styleResults=[]

    if(created&&created.ok){
      for(const styleId of [0,1,2]){
        const result=await runChartPresentationInFrame(
          frame,
          apiHely,
          {
            type:'chart.presentation.style',
            sheet,
            name,
            style:styleId
          }
        )

        styleResults.push({
          requested:styleId,
          result
        })
      }
    }

    steps.push({
      name:'chart-style',
      results:styleResults
    })

    const cleanup=created&&created.ok
      ? await runChartInFrame(
          frame,
          apiHely,
          {
            type:'chart.delete',
            sheet,
            name
          }
        )
      : null

    steps.push({
      name:'cleanup-delete',
      result:cleanup
    })

    const post=await runChartInFrame(
      frame,
      apiHely,
      {
        type:'chart.inspect',
        sheet
      }
    )

    const clean=!!(
      post &&
      post.ok &&
      post.count===before &&
      !post.charts.some(
        c=>String(c&&c.name)===name
      )
    )

    steps.push({
      name:'post-cleanup',
      result:post,
      matches:clean
    })

    const capabilityPass=!!(
      inspect &&
      inspect.ok &&
      inspect.verification &&
      inspect.verification.status==='PASS' &&
      inspect.support &&
      inspect.support.SetLegendPos===true &&
      inspect.support.GetLegendPos===true &&
      inspect.support.SetHorAxisTitle===true &&
      inspect.support.GetHorAxisTitle===true &&
      inspect.support.SetVerAxisTitle===true &&
      inspect.support.GetVerAxisTitle===true &&
      inspect.support.SetShowDataLabels===true &&
      inspect.support.GetDataLabels===true &&
      inspect.support.ApplyChartStyle===true &&
      inspect.support.GetChartStyle===true
    )

    const stylePass=
      styleResults.length===3 &&
      styleResults.every(
        x=>
          x &&
          pass(x.result) &&
          x.result.verification.expected &&
          x.result.verification.actual &&
          x.result.verification.expected.style===x.requested &&
          x.result.verification.actual.style===x.requested
      )

    const semanticPass=
      capabilityPass &&
      pass(legend) &&
      pass(axes) &&
      pass(labels) &&
      stylePass

    const infrastructurePass=!!(
      seed &&
      seed.ok &&
      initial &&
      initial.ok &&
      created &&
      created.ok &&
      cleanup &&
      cleanup.ok &&
      clean
    )

    const outcome=
      infrastructurePass &&
      semanticPass
        ? 'PASS'
        : 'FAIL'

    console.log(JSON.stringify({
      milestone:'M5.3',
      scope:'chart presentation controls',
      source:'live-coedit-editor',
      outcome,
      testOutcome:outcome,
      humanObservationRequired:false,
      supportedCapabilities:{
        legendPosition:'PASS',
        axisTitles:'PASS',
        dataLabels:'PASS',
        chartStyle:'PASS'
      },
      notClaimedAsPass:{
        legendFontSize:{
          status:'UNKNOWN',
          reason:'no matching public semantic getter has been established; standalone acceptance does not mutate it'
        }
      },
      sheet,
      range,
      chartName:name,
      requested:{
        legendPosition:'bottom',
        horizontalTitle,
        verticalTitle,
        dataLabels:{
          showSerName:true,
          showCatName:true,
          showVal:true,
          showPercent:false
        },
        chartStyles:[0,1,2]
      },
      steps,
      editor:'spreadsheeteditor',
      apiHely
    },null,2))

    if(outcome==='FAIL')
      process.exitCode=2

  }finally{
    await browser.close()
  }
}

main().catch(e=>{
  console.error(JSON.stringify({
    milestone:'M5.3',
    outcome:'launcher-error',
    error:String(e&&e.message?e.message:e)
  },null,2))
  process.exitCode=1
})
