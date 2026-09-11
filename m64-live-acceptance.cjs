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

  const seed=await frame.evaluate(()=>new Promise(resolve=>window.Asc.editor.callCommand(function(){
    try{
      var r=Api.GetActiveSheet().GetRange('XFA60:XFC64');
      r.SetValue([
        ['Region','Style','Price'],
        ['East','A',10],
        ['West','B',20],
        ['East','B',30],
        ['West','A',40]
      ]);
      return {ok:true};
    }catch(e){return {ok:false,error:String(e&&e.message?e.message:e)}}
  },false,resolve)));
  steps.push({step:'seed',result:{ok:seed.ok,source:'live-coedit-editor',verification:{status:seed.ok?'PASS':'FAIL',expected:'scratch pivot source write completes',actual:seed}}});

  async function run(step,spec){const result=await runPivotInFrame(frame,'window.Asc.editor',spec);steps.push({step,result});return result}
  const name='EURO_M64_PIVOT', renamed='EURO_M64_RENAMED';
  await run('create',{operation:'pivot.createNewWorksheet',source:'XFA60:XFC64',name});
  await run('inspect-after-create',{operation:'pivot.inspect',name,expectedPresent:true});
  await run('add-fields',{operation:'pivot.addFields',name,rowFields:['Region'],columnFields:['Style']});

  // Deployed 9.3.4 can apply AddDataField successfully while its callCommand callback
  // is not delivered. Treat the mutation call as transport-only, then verify the
  // semantic effect in a fresh public-API callCommand readback.
  await runPivotInFrame(frame,'window.Asc.editor',{operation:'pivot.addDataField',name,field:'Price'});
  const dataReadback=await runPivotInFrame(frame,'window.Asc.editor',{operation:'pivot.inspect',name,expectedPresent:true});
  const dataActual=dataReadback&&dataReadback.verification&&dataReadback.verification.actual;
  const dataPass=!!dataActual&&dataActual.dataFields===1;
  steps.push({step:'add-data-field',result:{
    ok:dataPass,outcome:dataPass?'ok':'verification-failed',source:'live-coedit-editor',operation:'pivot.addDataField',
    verification:{status:dataPass?'PASS':'FAIL',expected:{dataFields:1},actual:dataActual,readback:'separate-public-inspect'}
  }});

  await run('inspect-after-fields',{operation:'pivot.inspect',name,expectedPresent:true});
  await run('style',{operation:'pivot.style',name,styleName:'PivotStyleMedium2'});
  await run('rename',{operation:'pivot.rename',name,newName:renamed});
  await run('inspect-after-rename',{operation:'pivot.inspect',name:renamed,expectedPresent:true});
  const refresh=await run('refresh-deferred',{operation:'pivot.refresh',name:renamed});

  const required=steps.filter(x=>x.step!=='refresh-deferred');
  const statuses=required.map(x=>x.result&&x.result.verification&&x.result.verification.status);
  const refreshStatus=refresh&&refresh.verification&&refresh.verification.status;
  const pass=statuses.length===8&&statuses.every(x=>x==='PASS')&&refreshStatus==='UNKNOWN';
  console.log(JSON.stringify({
    milestone:'M6.4',source:'live-coedit-editor',humanObservationRequired:false,steps,
    outcome:pass?'PASS_WITH_REFRESH_DEFERRED':'FAIL',
    verification:{status:pass?'PASS':'FAIL',operationStatuses:statuses,refreshStatus}
  },null,2));
  if(!pass) process.exitCode=1;
 }finally{await browser.close()}
}
main().catch(e=>{console.error(JSON.stringify({milestone:'M6.4',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1});
