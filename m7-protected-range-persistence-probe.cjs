'use strict';
const {chromium}=require(process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright');
async function main(){
 const base=process.env.EURO_NC_BASE_URL,id=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,pw=process.env.EURO_NC_PASSWORD;
 if(!base||!id||!user||!pw) throw new Error('missing EURO_NC env');
 const browser=await chromium.launch();
 try{
  const page=await browser.newPage();
  await page.goto(base+'/login',{waitUntil:'domcontentloaded',timeout:60000});
  await page.fill('#user',user); await page.fill('#password',pw); await page.click('button[type=submit], input[type=submit]');
  await page.waitForTimeout(2500);
  const url=base+'/index.php/apps/eurooffice/'+id;
  async function open(){await page.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000);const f=page.frames().find(x=>/spreadsheeteditor/.test(x.url()));if(!f)throw new Error('editor frame missing');return f}
  async function cmd(f,s){return f.evaluate(body=>new Promise(r=>window.Asc.editor.callCommand(new Function(body),false,r)),s)}
  let f=await open();
  const created=await cmd(f,"function U(u){return u?{id:u.GetId(),name:u.GetName(),type:u.GetType()}:null} var s=Api.GetActiveSheet(),p=s.AddProtectedRange('EURO_M7_PERSISTENCE','$XFD$80:$XFD$81'); if(!p)return {ok:false}; p.AddUser('EURO_M7_PERSIST_USER','EURO M7 Persist User','CanEdit'); return {ok:true,count:(s.GetAllProtectedRanges()||[]).length,lookup:!!s.GetProtectedRange('EURO_M7_PERSISTENCE'),user:U(p.GetUser('EURO_M7_PERSIST_USER'))};");
  await page.waitForTimeout(12000);
  const before=await cmd(f,"var s=Api.GetActiveSheet(),p=null;try{p=s.GetProtectedRange('EURO_M7_PERSISTENCE')}catch(e){} return {count:(s.GetAllProtectedRanges()||[]).length,lookup:!!p,user:!!(p&&p.GetUser('EURO_M7_PERSIST_USER'))};");
  await page.reload({waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(24000);
  f=page.frames().find(x=>/spreadsheeteditor/.test(x.url()));if(!f)throw new Error('editor frame missing after reload');
  const after=await cmd(f,"function U(u){return u?{id:u.GetId(),name:u.GetName(),type:u.GetType()}:null} var s=Api.GetActiveSheet(),p=null;try{p=s.GetProtectedRange('EURO_M7_PERSISTENCE')}catch(e){} return {count:(s.GetAllProtectedRanges()||[]).length,lookup:!!p,user:p?U(p.GetUser('EURO_M7_PERSIST_USER')):null};");
  const pass=!!(created&&created.ok&&before&&before.lookup&&after&&after.lookup&&after.user&&after.user.id==='EURO_M7_PERSIST_USER');
  console.log(JSON.stringify({milestone:'M7',probe:'protected-range-persistence',source:'live-coedit-editor',humanObservationRequired:false,created,beforeReload:before,afterReload:after,verification:{status:pass?'PASS':'UNKNOWN',expected:'protected range and ACL survive editor reload',actual:after}},null,2));
 }finally{await browser.close()}
}
main().catch(e=>{console.error(JSON.stringify({milestone:'M7',probe:'protected-range-persistence',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1});
