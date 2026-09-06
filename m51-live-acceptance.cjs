'use strict'

const {runChartInFrame}=require('./live-charts.cjs')
const {writeBulkInFrame}=require('./bulk-writer.cjs')
const {readRangeInFrame}=require('./range-reader.cjs')

function loadPlaywright(){for(const p of [process.env.EURO_PLAYWRIGHT_PATH,'playwright','/home/user/marveen/node_modules/playwright'].filter(Boolean)){try{return require(p)}catch(_){}}throw new Error('Playwright is unavailable')}
function normMatrix(v){return Array.isArray(v)?v.map(r=>Array.isArray(r)?r.map(x=>String(x==null?'':(x&&typeof x==='object'&&Object.prototype.hasOwnProperty.call(x,'value')?x.value:x))):r):v}
function sameMatrix(a,b){return JSON.stringify(normMatrix(a))===JSON.stringify(normMatrix(b))}
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
  const steps=[];let outcome='PASS',cleanupOutcome='DEFERRED_RUNTIME_CAPABILITY'
  const data=[['Quarter','Revenue'],['Q1',100],['Q2',140],['Q3',125]]
  let w=await writeBulkInFrame(frame,apiHely,{sheet,range,values:data});steps.push({name:'seed-data-write',result:w});if(!w.ok)outcome='FAIL'
  let rr=await readRangeInFrame(frame,apiHely,{sheet,range});let seedMatches=!!rr.ok&&sameMatrix(rr.cells,data);steps.push({name:'seed-data-readback',result:rr,expected:data,matches:seedMatches});if(!seedMatches)outcome='FAIL'
  let initial=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'initial-chart-inspect',result:initial})
  if(!initial.ok)outcome='UNKNOWN'
  const before=initial.ok?initial.count:null
  if(initial.ok){
   let created=await runChartInFrame(frame,apiHely,{type:'chart.create',sheet,range,inRows:false,chartType:'bar',style:2,width:3600000,height:2200000,name:chartName,title:'Quarterly revenue'});steps.push({name:'create-chart',result:created})
   if(!created.ok||!created.verification||created.verification.status!=='PASS'||!created.actual||created.actual.seriesCount!==1)outcome=created&&created.verification&&created.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'
   let post=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'post-create-inspect',result:post});if(!post.ok||post.count!==before+1)outcome='FAIL'
   let modified=await runChartInFrame(frame,apiHely,{type:'chart.modify',sheet,name:chartName,title:'Quarterly revenue — verified',width:4000000,height:2400000});steps.push({name:'modify-chart',result:modified});if(!modified.ok||!modified.verification||modified.verification.status!=='PASS')outcome=modified&&modified.verification&&modified.verification.status==='UNKNOWN'?'UNKNOWN':'FAIL'
   let postModify=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'post-modify-inspect',result:postModify});let modifiedInventory=postModify.ok?postModify.charts.find(c=>String(c.name)===chartName):null;if(!modifiedInventory||String(modifiedInventory.title||'').replace(/[\r\n]+$/g,'')!=='Quarterly revenue — verified'||Number(modifiedInventory.width)!==4000000||Number(modifiedInventory.height)!==2400000||modifiedInventory.seriesCount!==1)outcome='FAIL'
   steps.push({name:'chart-delete',result:{ok:false,outcome:'deferred-runtime-capability',source:'live-coedit-editor',reason:'EuroOffice 9.3.4 public spreadsheet ApiChart/ApiDrawing objects do not expose Delete(); independently reproduced in live co-edit and DocBuilder',recheck:'EuroOffice based on ONLYOFFICE 9.4+'}})
  }
  console.log(JSON.stringify({milestone:'M5.1',source:'live-coedit-editor',outcome,testOutcome:outcome,cleanupOutcome,humanObservationRequired:false,deferred:[{operation:'chart.delete',reason:'runtime capability; recheck on EuroOffice based on ONLYOFFICE 9.4+'}],sheet,range,chartName,steps,editor:'spreadsheeteditor',apiHely},null,2))
  if(outcome!=='PASS')process.exitCode=2
 }finally{await browser.close()}
}
if(require.main===module)main().catch(err=>{console.error(JSON.stringify({milestone:'M5.1',outcome:'launcher-error',error:String(err&&err.message?err.message:err)},null,2));process.exitCode=1})
