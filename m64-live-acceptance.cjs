'use strict';
const { chromium } = require(process.env.EURO_PLAYWRIGHT_PATH || '/home/user/marveen/node_modules/playwright');
const { runPivotInFrame } = require('./live-pivot-tables.cjs');

async function main(){
 const base=process.env.EURO_NC_BASE_URL,id=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,pw=process.env.EURO_NC_PASSWORD;
 if(!base||!id||!user||!pw) throw new Error('required EURO_NC_* env missing');
 const browser=await chromium.launch(); const steps=[];
 try{
  const page=await browser.newPage();
  await page.goto(base+'/login',{waitUntil:'domcontentloaded',timeout:60000});
  await page.fill('#user',user); await page.fill('#password',pw); await page.click('button[type=submit], input[type=submit]'); await page.waitForTimeout(2500);
  if(/\/login/.test(page.url())) throw new Error('login failed');
  await page.goto(base+'/index.php/apps/eurooffice/'+id,{waitUntil:'domcontentloaded',timeout:60000}); await page.waitForTimeout(22000);
  const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url())); if(!frame) throw new Error('spreadsheeteditor frame missing');
  const seed=await frame.evaluate(()=>new Promise(resolve=>window.Asc.editor.callCommand(function(){try{var ws=Api.GetActiveSheet();var r=ws.GetRange('XFA60:XFC64');r.SetValues([['Region','Style','Price'],['East','A',10],['West','B',20],['East','B',30],['West','A',40]]);return {ok:true,values:r.GetValues()}}catch(e){return {ok:false,error:String(e)}}},false,resolve)));
  steps.push({step:'seed',result:{ok:seed.ok,source:'live-coedit-editor',verification:{status:seed.ok?'PASS':'FAIL',expected:'scratch pivot source written',actual:seed}}});
  async function run(step,spec){const result=await runPivotInFrame(frame,'window.Asc.editor',spec);steps.push({step,result});return result}
  const name='EURO_M64_PIVOT', renamed='EURO_M64_RENAMED';
  await run('create',{operation:'pivot.createNewWorksheet',source:'XFA60:XFC64',name});
  await run('inspect-after-create',{operation:'pivot.inspect',name});
  await run('add-fields',{operation:'pivot.addFields',name,rowFields:['Region'],columnFields:['Style']});
  await run('add-data-field',{operation:'pivot.addDataField',name,field:'Price'});
  await run('inspect-after-fields',{operation:'pivot.inspect',name});
  await run('style',{operation:'pivot.style',name,styleName:'PivotStyleMedium2'});
  await run('refresh',{operation:'pivot.refresh',name});
  await run('rename',{operation:'pivot.rename',name,newName:renamed});
  await run('inspect-after-rename',{operation:'pivot.inspect',name:renamed});
  const statuses=steps.map(x=>x.result&&x.result.verification&&x.result.verification.status);
  const pass=statuses.length===10&&statuses.every(x=>x==='PASS');
  console.log(JSON.stringify({milestone:'M6.4',source:'live-coedit-editor',humanObservationRequired:false,steps,outcome:pass?'PASS':'FAIL',verification:{status:pass?'PASS':'FAIL',operationStatuses:statuses}},null,2));
  if(!pass) process.exitCode=1;
 }finally{await browser.close()}
}
main().catch(e=>{console.error(JSON.stringify({milestone:'M6.4',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1});
