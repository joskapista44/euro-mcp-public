'use strict'

const {runChartInFrame}=require('./live-charts.cjs')
const {writeBulkInFrame}=require('./bulk-writer.cjs')
const {readRangeInFrame}=require('./range-reader.cjs')

function loadPlaywright(){for(const p of [process.env.EURO_PLAYWRIGHT_PATH,'playwright','/home/user/marveen/node_modules/playwright'].filter(Boolean)){try{return require(p)}catch(_){}}throw new Error('Playwright is unavailable')}
async function main(){
 const base=process.env.EURO_NC_BASE_URL,fileId=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD
 if(!base||!fileId||!user||!password)throw new Error('EURO_NC_BASE_URL, EURO_NC_FILE_ID, EURO_NC_USER and EURO_NC_PASSWORD are required')
 const sheet=process.env.EURO_M51_SHEET||'Sheet1',range=process.env.EURO_M51_RANGE||'XFA1:XFB4',chartName='M51_'+Date.now().toString(36)
 const {chromium}=loadPlaywright(),browser=await chromium.launch()
 try{
  const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage()
  await page.goto(`${base}/login`,{waitUntil:'domcontentloaded',timeout:60000});await page.fill('#user',user);await page.fill('#password',password)
  await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:60000}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]);await page.waitForTimeout(2500)
  if(/\/login/.test(page.url()))throw new Error('Nextcloud login did not succeed')
  await page.goto(`${base}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000)
  const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)throw new Error('spreadsheeteditor frame did not open')
  const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':(window.editor&&typeof window.editor.callCommand==='function')?'window.editor':null);if(!apiHely)throw new Error('callCommand is unavailable')
  const steps=[];let outcome='PASS',cleanupOutcome='UNKNOWN'
  const data=[['Quarter','Revenue'],['Q1',100],['Q2',140],['Q3',125]]
  let w=await writeBulkInFrame(frame,apiHely,{sheet,range,values:data});steps.push({name:'seed-data-write',result:w});if(!w.ok)outcome='FAIL'
  let rr=await readRangeInFrame(frame,apiHely,{sheet,range});steps.push({name:'seed-data-readback',result:rr});if(!rr.ok)outcome='FAIL'
  let initial=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'initial-chart-inspect',result:initial})
  if(!initial.ok){outcome='UNKNOWN'}
  const before=initial.ok?initial.count:null
  let created=null
  if(initial.ok){
   created=await runChartInFrame(frame,apiHely,{type:'chart.create',sheet,range,inRows:false,chartType:'bar',style:2,width:3600000,height:2200000,name:chartName,title:'Quarterly revenue'});steps.push({name:'create-chart',result:created})
   if(!created.ok||!created.verification||created.verification.status!=='PASS')outcome=created&&created.verification&&created.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'
   let post=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'post-create-inspect',result:post});if(!post.ok||post.count!==before+1)outcome='FAIL'
   let del=await runChartInFrame(frame,apiHely,{type:'chart.delete',sheet,name:chartName});steps.push({name:'delete-chart',result:del});if(!del.ok||!del.verification||del.verification.status!=='PASS')outcome='FAIL'
   let final=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'final-chart-inspect',result:final});cleanupOutcome=final.ok&&final.count===before?'PASS':'FAIL';if(cleanupOutcome!=='PASS')outcome='FAIL'
  }
  console.log(JSON.stringify({milestone:'M5.1',source:'live-coedit-editor',outcome,testOutcome:outcome,cleanupOutcome,humanObservationRequired:false,sheet,range,chartName,steps,editor:'spreadsheeteditor',apiHely},null,2))
 }finally{await browser.close()}
}
if(require.main===module)main().catch(err=>{console.error(JSON.stringify({milestone:'M5.1',outcome:'launcher-error',error:String(err&&err.message?err.message:err)},null,2));process.exitCode=1})
