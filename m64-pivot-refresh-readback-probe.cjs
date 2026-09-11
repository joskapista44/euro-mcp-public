'use strict';
const { chromium } = require(process.env.EURO_PLAYWRIGHT_PATH || '/home/user/marveen/node_modules/playwright');
function probeCommand(){
 function snap(r){if(!r)return null;try{return {address:r.GetAddress(),value:r.GetValue(),text:r.GetText()};}catch(e){return {error:String(e.message||e)};}}
 function safe(fn){try{return fn();}catch(e){return {error:String(e.message||e)};}}
 try{
  var ws=Api.GetActiveSheet(),wsName=ws.GetName();
  var src=ws.GetRange('XFA80:XFC84');
  src.SetValue([['Region','Style','Price'],['East','A',10],['West','B',20],['East','B',30],['West','A',40]]);
  var p=Api.InsertPivotNewWorksheet(src);p.SetName('EURO_M64_REFRESH_PROBE');p.AddFields({rows:'Region',columns:'Style'});p.AddDataField('Price');
  function reads(){return {getDataEastA:safe(function(){return p.GetData(['East','A']);}),getDataEastB:safe(function(){return p.GetData(['East','B']);}),pivotDataEast:safe(function(){return snap(p.GetPivotData('Price',['Region','East']));}),dataBody:safe(function(){return snap(p.GetDataBodyRange());}),tableRange1:safe(function(){return snap(p.GetTableRange1());}),tableRange2:safe(function(){return snap(p.GetTableRange2());})};}
  var before=reads();
  Api.GetSheet(wsName).GetRange('XFC81').SetValue(110);
  var sourceAfter=Api.GetSheet(wsName).GetRange('XFA80:XFC84').GetValue();
  var refreshResult=safe(function(){p.RefreshTable();return true;});
  var after=reads();
  var ps=p.GetParent(),psName=ps&&ps.GetName?ps.GetName():null;
  return {ok:true,source:'live-coedit-editor',kind:'public-pivot-refresh-readback-probe',sourceSheetName:wsName,pivotSheetName:psName,sourceAfter:sourceAfter,refreshResult:refreshResult,before:before,after:after,verification:{status:'UNKNOWN',expected:'stable public semantic readback changes after RefreshTable',actual:'measurement only'}};
 }catch(e){return {ok:false,source:'live-coedit-editor',kind:'public-pivot-refresh-readback-probe',error:String(e.message||e),verification:{status:'UNKNOWN'}};}
}
async function main(){
 const base=process.env.EURO_NC_BASE_URL,id=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD;if(!base||!id||!user||!password)throw new Error('required EURO_NC_* env missing');
 const browser=await chromium.launch();try{const page=await browser.newPage();await page.goto(base+'/login',{waitUntil:'domcontentloaded',timeout:60000});await page.fill('#user',user);await page.fill('#password',password);await page.click('button[type=submit], input[type=submit]');await page.waitForTimeout(2500);if(/\/login/.test(page.url()))throw new Error('login failed');await page.goto(base+'/index.php/apps/eurooffice/'+id,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000);const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)throw new Error('spreadsheeteditor frame missing');const body='return ('+probeCommand.toString()+')();';const result=await frame.evaluate(commandBody=>new Promise(resolve=>window.Asc.editor.callCommand(new Function(commandBody),false,resolve)),body);let cleanup={ok:false};if(result&&result.pivotSheetName){cleanup=await frame.evaluate(name=>new Promise(resolve=>{const b='var n='+JSON.stringify(name)+';try{var s=Api.GetSheet(n);if(!s)return {ok:false};s.Delete();return {ok:Api.GetSheets().map(function(x){return x.GetName();}).indexOf(n)===-1};}catch(e){return {ok:false,error:String(e.message||e)}}';window.Asc.editor.callCommand(new Function(b),false,resolve);}),result.pivotSheetName);}console.log(JSON.stringify({milestone:'M6.4',kind:'pivot-refresh-readback-probe',result,cleanup},null,2));}finally{await browser.close();}}
main().catch(e=>{console.error(JSON.stringify({milestone:'M6.4',outcome:'launcher-error',error:String(e.message||e)},null,2));process.exitCode=1;});
