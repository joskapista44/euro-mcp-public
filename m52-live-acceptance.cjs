'use strict'

const {runChartInFrame}=require('./live-charts.cjs')
const {runChartSeriesInFrame}=require('./live-chart-series.cjs')
const {writeBulkInFrame}=require('./bulk-writer.cjs')
const {readRangeInFrame}=require('./range-reader.cjs')

function loadPlaywright(){for(const p of [process.env.EURO_PLAYWRIGHT_PATH,'playwright','/home/user/marveen/node_modules/playwright'].filter(Boolean)){try{return require(p)}catch(_){}}throw new Error('Playwright is unavailable')}
function normMatrix(v){return Array.isArray(v)?v.map(r=>Array.isArray(r)?r.map(x=>String(x==null?'':(x&&typeof x==='object'&&Object.prototype.hasOwnProperty.call(x,'value')?x.value:x))):r):v}
function sameMatrix(a,b){return JSON.stringify(normMatrix(a))===JSON.stringify(normMatrix(b))}
async function main(){
 const base=process.env.EURO_NC_BASE_URL,fileId=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD
 if(!base||!fileId||!user||!password)throw new Error('EURO_NC_BASE_URL, EURO_NC_FILE_ID, EURO_NC_USER and EURO_NC_PASSWORD are required')
 const sheet=process.env.EURO_M52_SHEET||'Sheet1',range=process.env.EURO_M52_RANGE||'XFA1:XFD4',chartRange='XFA1:XFC4',chartName='M52_'+Date.now().toString(36)
 const q=String(sheet).replace(/'/g,"''"),fq=r=>`'${q}'!${r}`
 const {chromium}=loadPlaywright(),browser=await chromium.launch()
 try{
  const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage()
  await page.goto(`${base}/login`,{waitUntil:'domcontentloaded',timeout:60000});await page.fill('#user',user);await page.fill('#password',password)
  await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:60000}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]);await page.waitForTimeout(2500)
  if(/\/login/.test(page.url()))throw new Error('Nextcloud login did not succeed')
  await page.goto(`${base}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000)
  const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)throw new Error('spreadsheeteditor frame did not open')
  const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':(window.editor&&typeof window.editor.callCommand==='function')?'window.editor':null);if(!apiHely)throw new Error('callCommand is unavailable')
  const steps=[];let outcome='PASS'
  const data=[['Quarter','Revenue','Cost','Forecast'],['Q1',100,70,120],['Q2',140,90,150],['Q3',125,85,145]]
  let w=await writeBulkInFrame(frame,apiHely,{sheet,range,values:data});steps.push({name:'seed-data-write',result:w});if(!w.ok)outcome='FAIL'
  let rr=await readRangeInFrame(frame,apiHely,{sheet,range});let seedMatches=!!rr.ok&&sameMatrix(rr.cells,data);steps.push({name:'seed-data-readback',result:rr,expected:data,matches:seedMatches});if(!seedMatches)outcome='FAIL'
  let initial=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'initial-chart-inspect',result:initial});if(!initial.ok)outcome='UNKNOWN'
  const before=initial.ok?initial.count:null
  if(initial.ok){
    const created=await runChartInFrame(frame,apiHely,{type:'chart.create',sheet,range:chartRange,inRows:false,chartType:'comboBarLine',style:2,width:4000000,height:2400000,name:chartName,title:'M5.2 series acceptance'});steps.push({name:'create-combo-chart',result:created})
    if(!created.ok||!created.verification||created.verification.status!=='PASS'||!created.actual||created.actual.seriesCount!==2)outcome=created&&created.verification&&created.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'

    let inspected=await runChartSeriesInFrame(frame,apiHely,{type:'chart.series.inspect',sheet,name:chartName});steps.push({name:'series-inspect-initial',result:inspected})
    if(!inspected.ok||inspected.count!==2)outcome=inspected&&inspected.verification&&inspected.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'

    let changed=await runChartSeriesInFrame(frame,apiHely,{type:'chart.series.changeType',sheet,name:chartName,seriesIndex:0,chartType:'area'});steps.push({name:'series-change-type',result:changed})
    if(!changed.ok||!changed.verification||changed.verification.status!=='PASS'||changed.actualType!=='area')outcome=changed&&changed.verification&&changed.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'

    let afterType=await runChartSeriesInFrame(frame,apiHely,{type:'chart.series.inspect',sheet,name:chartName});steps.push({name:'series-post-type-inspect',result:afterType})
    if(!afterType.ok||!afterType.series||!afterType.series[0]||afterType.series[0].chartType!=='area')outcome='FAIL'

    let added=await runChartSeriesInFrame(frame,apiHely,{type:'chart.series.add',sheet,name:chartName,nameRange:fq('$XFD$1'),valuesRange:fq('$XFD$2:$XFD$4')});steps.push({name:'series-add',result:added})
    if(!added.ok||!added.verification||added.verification.status!=='PASS'||added.beforeCount!==2||added.afterCount!==3)outcome=added&&added.verification&&added.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'

    let afterAdd=await runChartSeriesInFrame(frame,apiHely,{type:'chart.series.inspect',sheet,name:chartName});steps.push({name:'series-post-add-inspect',result:afterAdd});if(!afterAdd.ok||afterAdd.count!==3)outcome='FAIL'

    let removed=await runChartSeriesInFrame(frame,apiHely,{type:'chart.series.remove',sheet,name:chartName,seriesIndex:2});steps.push({name:'series-remove',result:removed})
    if(!removed.ok||!removed.verification||removed.verification.status!=='PASS'||removed.beforeCount!==3||removed.afterCount!==2)outcome=removed&&removed.verification&&removed.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'

    let finalSeries=await runChartSeriesInFrame(frame,apiHely,{type:'chart.series.inspect',sheet,name:chartName});steps.push({name:'series-final-inspect',result:finalSeries});if(!finalSeries.ok||finalSeries.count!==2||!finalSeries.series[0]||finalSeries.series[0].chartType!=='area')outcome='FAIL'

    let deleted=await runChartInFrame(frame,apiHely,{type:'chart.delete',sheet,name:chartName});steps.push({name:'chart-cleanup-delete',result:deleted});if(!deleted.ok||!deleted.verification||deleted.verification.status!=='PASS')outcome='FAIL'
    let postDelete=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'post-cleanup-inspect',result:postDelete});const clean=!!postDelete.ok&&postDelete.count===before&&!postDelete.charts.some(c=>String(c&&c.name)===chartName);steps.push({name:'cleanup-acceptance',matches:clean,expectedCount:before,actualCount:postDelete&&postDelete.count});if(!clean)outcome='FAIL'
  }
  console.log(JSON.stringify({milestone:'M5.2',scope:'chart series management',source:'live-coedit-editor',outcome,testOutcome:outcome,humanObservationRequired:false,sheet,range,chartName,steps,editor:'spreadsheeteditor',apiHely},null,2))
  if(outcome!=='PASS')process.exitCode=2
 }finally{await browser.close()}
}
if(require.main===module)main().catch(err=>{console.error(JSON.stringify({milestone:'M5.2',outcome:'launcher-error',error:String(err&&err.message?err.message:err)},null,2));process.exitCode=1})
