'use strict'

const {runConditionalFormattingInFrame}=require('./live-conditional-formatting.cjs')

function loadPlaywright(){for(const p of [process.env.EURO_PLAYWRIGHT_PATH,'playwright','/home/user/marveen/node_modules/playwright'].filter(Boolean)){try{return require(p)}catch(_){}}throw new Error('Playwright is unavailable')}
async function main(){
 const base=process.env.EURO_NC_BASE_URL,fileId=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD
 if(!base||!fileId||!user||!password)throw new Error('EURO_NC_BASE_URL, EURO_NC_FILE_ID, EURO_NC_USER and EURO_NC_PASSWORD are required')
 const sheet=process.env.EURO_M46_SHEET||'Sheet1',range=process.env.EURO_M46_RANGE||'XFD1:XFD3'
 const {chromium}=loadPlaywright(),browser=await chromium.launch()
 try{
  const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage()
  await page.goto(`${base}/login`,{waitUntil:'domcontentloaded',timeout:60000});await page.fill('#user',user);await page.fill('#password',password)
  await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:60000}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]);await page.waitForTimeout(2500)
  if(/\/login/.test(page.url()))throw new Error('Nextcloud login did not succeed')
  await page.goto(`${base}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000)
  const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)throw new Error('spreadsheeteditor frame did not open')
  const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':(window.editor&&typeof window.editor.callCommand==='function')?'window.editor':null);if(!apiHely)throw new Error('callCommand is unavailable')
  const steps=[]
  let r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.inspect',sheet,range});steps.push({name:'initial-inspect',result:r})
  // Acceptance is isolated: refuse to delete or overwrite pre-existing rules.
  if(!r.ok||r.count!==0){console.log(JSON.stringify({milestone:'M4.6',source:'live-coedit-editor',outcome:'UNKNOWN',reason:'acceptance range is not empty or not measurable',sheet,range,steps,editor:'spreadsheeteditor',apiHely},null,2));return}
  r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.add',sheet,range,rule:{type:'xlCellValue',operator:'xlGreater',formula1:'200',fillColor:[255,255,0],priority:1}});steps.push({name:'add-cell-value-rule',result:r})
  let testOutcome=r.ok&&r.verification&&r.verification.status==='PASS'?'PASS':'FAIL'
  r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.inspect',sheet,range});steps.push({name:'post-add-inspect',result:r});if(!r.ok||r.count!==1)testOutcome='FAIL'
  r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.delete',sheet,range,index:0});steps.push({name:'delete-rule',result:r});if(!r.ok||!r.verification||r.verification.status!=='PASS')testOutcome='FAIL'
  const final=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.inspect',sheet,range});steps.push({name:'final-inspect',result:final})
  const cleanupOutcome=final.ok&&final.count===0?'PASS':'FAIL';if(cleanupOutcome!=='PASS')testOutcome='FAIL'
  console.log(JSON.stringify({milestone:'M4.6',source:'live-coedit-editor',outcome:testOutcome,testOutcome,cleanupOutcome,humanObservationRequired:false,sheet,range,steps,editor:'spreadsheeteditor',apiHely},null,2))
 }finally{await browser.close()}
}
if(require.main===module)main().catch(err=>{console.error(JSON.stringify({milestone:'M4.6',outcome:'launcher-error',error:String(err&&err.message?err.message:err)},null,2));process.exitCode=1})
