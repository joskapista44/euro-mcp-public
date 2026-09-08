'use strict'
const {runChartInFrame}=require('./live-charts.cjs')
const {runChartDataObjectInFrame}=require('./live-chart-data-objects.cjs')
const {writeBulkInFrame}=require('./bulk-writer.cjs')
function loadPlaywright(){for(const p of [process.env.EURO_PLAYWRIGHT_PATH,'playwright','/home/user/marveen/node_modules/playwright'].filter(Boolean)){try{return require(p)}catch(_){}}throw new Error('Playwright is unavailable')}
async function main(){
 const base=process.env.EURO_NC_BASE_URL,fileId=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD
 if(!base||!fileId||!user||!password)throw new Error('required EURO_NC_* env missing')
 const sheet='Sheet1',range='XFA1:XFC4',name='M54_'+Date.now().toString(36),renamed=name+'_RENAMED',steps=[]
 const {chromium}=loadPlaywright(),browser=await chromium.launch()
 try{
  const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage();await page.goto(`${base}/login`,{waitUntil:'domcontentloaded',timeout:60000});await page.fill('#user',user);await page.fill('#password',password);await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:60000}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]);await page.waitForTimeout(2500);if(/\/login/.test(page.url()))throw new Error('login failed');await page.goto(`${base}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000)
  const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)throw new Error('spreadsheeteditor frame missing');const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':null);if(!apiHely)throw new Error('callCommand unavailable')
  const seed=await writeBulkInFrame(frame,apiHely,{sheet,range,values:[['Quarter','Revenue','Cost'],['Q1',100,70],['Q2',140,90],['Q3',125,85]]});steps.push({name:'seed',result:seed})
  const initial=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});steps.push({name:'initial-chart-inspect',result:initial});const before=initial.ok?initial.count:null
  const created=await runChartInFrame(frame,apiHely,{type:'chart.create',sheet,range,chartType:'bar',style:2,width:4000000,height:2400000,name,title:'M5.4 acceptance'});steps.push({name:'create-chart',result:created})
  const inspect=created.ok?await runChartDataObjectInFrame(frame,apiHely,{type:'chart.object.inspect',sheet,name}):null;steps.push({name:'object-capability',result:inspect})
  const rename=created.ok?await runChartDataObjectInFrame(frame,apiHely,{type:'chart.object.rename',sheet,name,newName:renamed}):null;steps.push({name:'rename',result:rename})
  const resize=rename&&rename.ok?await runChartDataObjectInFrame(frame,apiHely,{type:'chart.object.resize',sheet,name:renamed,width:5000000,height:3000000}):null;steps.push({name:'resize',result:resize})
  const position=rename&&rename.ok?await runChartDataObjectInFrame(frame,apiHely,{type:'chart.object.position',sheet,name:renamed}):null;steps.push({name:'position-capability',result:position})
  const dataTypes=['chart.data.seriesName','chart.data.seriesValues','chart.data.seriesXValues','chart.data.categoryFormula'],dataResults=[]
  for(const type of dataTypes){const r=rename&&rename.ok?await runChartDataObjectInFrame(frame,apiHely,{type,sheet,name:renamed}):null;dataResults.push(r);steps.push({name:type,result:r})}
  const copy=rename&&rename.ok?await runChartDataObjectInFrame(frame,apiHely,{type:'chart.object.copy',sheet,name:renamed}):null;steps.push({name:'copy-capability',result:copy})
  const cleanup=created.ok?await runChartInFrame(frame,apiHely,{type:'chart.delete',sheet,name:rename&&rename.ok?renamed:name}):null;steps.push({name:'cleanup-delete',result:cleanup})
  const post=await runChartInFrame(frame,apiHely,{type:'chart.inspect',sheet});const clean=!!post.ok&&post.count===before&&!post.charts.some(c=>String(c&&c.name)===name||String(c&&c.name)===renamed);steps.push({name:'post-cleanup',result:post,matches:clean})
  const pass=r=>!!r&&r.ok&&r.verification&&r.verification.status==='PASS',unknown=r=>!!r&&r.ok&&r.verification&&r.verification.status==='UNKNOWN'
  const infrastructurePass=!!seed.ok&&!!initial.ok&&!!created.ok&&!!cleanup&&cleanup.ok&&clean
  const verifiedObjectPass=pass(inspect)&&pass(rename)&&pass(resize)
  const deferredPass=unknown(position)&&dataResults.every(unknown)&&!!copy&&!copy.ok&&copy.outcome==='unsupported'
  const outcome=infrastructurePass&&verifiedObjectPass&&deferredPass?'DEFERRED':'FAIL'
  console.log(JSON.stringify({milestone:'M5.4',scope:'chart data and object management',source:'live-coedit-editor',outcome,testOutcome:outcome,humanObservationRequired:false,reason:outcome==='DEFERRED'?'rename and resize have exact public readback PASS; position and series-data setters lack public semantic readback, while copy/AddDrawing is unavailable on deployed runtime':undefined,sheet,range,chartName:name,renamedChartName:renamed,steps,editor:'spreadsheeteditor',apiHely},null,2));if(outcome==='FAIL')process.exitCode=2
 }finally{await browser.close()}
}
main().catch(e=>{console.error(JSON.stringify({milestone:'M5.4',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1})
