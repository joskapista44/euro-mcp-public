'use strict';
const { chromium } = require(process.env.EURO_PLAYWRIGHT_PATH || '/home/user/marveen/node_modules/playwright');
function methods(o){if(!o)return [];const out=new Set();let x=o;for(let depth=0;x&&depth<8;depth++,x=Object.getPrototypeOf(x)){for(const n of Object.getOwnPropertyNames(x)){if(typeof o[n]==='function')out.add(n)}}return [...out].sort()}
async function main(){
 const base=process.env.EURO_NC_BASE_URL,id=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,pw=process.env.EURO_NC_PASSWORD;
 if(!base||!id||!user||!pw)throw new Error('required EURO_NC_* env missing');
 const browser=await chromium.launch();
 try{
  const page=await browser.newPage();await page.goto(base+'/login',{waitUntil:'domcontentloaded',timeout:60000});
  await page.fill('#user',user);await page.fill('#password',pw);await page.click('button[type=submit], input[type=submit]');await page.waitForTimeout(2500);if(/\/login/.test(page.url()))throw new Error('login failed');
  await page.goto(base+'/index.php/apps/eurooffice/'+id,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000);
  const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)throw new Error('spreadsheeteditor frame missing');
  const inventory=await frame.evaluate(()=>new Promise(resolve=>window.Asc.editor.callCommand(function(){
   function ownMethods(o){if(!o)return [];var out={},x=o,depth=0;while(x&&depth++<8){try{Object.getOwnPropertyNames(x).forEach(function(n){try{if(typeof o[n]==='function')out[n]=true}catch(_){}})}catch(_){};try{x=Object.getPrototypeOf(x)}catch(_){x=null}}return Object.keys(out).sort()}
   function pick(a){return a.filter(function(n){return /protect|lock|allow|permission|password/i.test(n)})}
   try{var sh=Api.GetActiveSheet(),r=sh.GetRange('XFD40');var apiM=ownMethods(Api),shM=ownMethods(sh),rM=ownMethods(r);return {ok:true,apiProtectionMethods:pick(apiM),worksheetProtectionMethods:pick(shM),rangeProtectionMethods:pick(rM),worksheetAllMethods:shM,rangeRelevantMethods:pick(rM.concat(['SetLocked','GetLocked']))};}catch(e){return {ok:false,error:String(e&&e.message?e.message:e)}}
  },false,resolve)));
  console.log(JSON.stringify({milestone:'M7',probe:'protection-runtime-capability',source:'live-coedit-editor',humanObservationRequired:false,inventory,verification:{status:inventory&&inventory.ok?'PASS':'UNKNOWN',expected:'public protection-related runtime method inventory',actual:inventory}},null,2));
 }finally{await browser.close()}
}
main().catch(e=>{console.error(JSON.stringify({milestone:'M7',probe:'protection-runtime-capability',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1});
