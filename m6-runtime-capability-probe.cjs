'use strict'

function loadPlaywright(){
  for(const p of [process.env.EURO_PLAYWRIGHT_PATH,'playwright','/home/user/marveen/node_modules/playwright'].filter(Boolean)){
    try{return require(p)}catch(_){}
  }
  throw new Error('Playwright is unavailable')
}

async function call(frame,apiHely,timeout=15000){
  const body=`return (${probeCommand.toString()})();`
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{
    const e=u==='window.editor'?window.editor:(window.Asc||{}).editor
    let done=false
    const finish=v=>{if(!done){done=true;resolve(v)}}
    try{e.callCommand(new Function(body),false,v=>finish(v))}
    catch(err){finish({ok:false,error:String(err&&err.message?err.message:err)})}
    setTimeout(()=>finish({ok:false,error:'timeout'}),timeout)
  }),{u:apiHely,body,timeout})
}

function probeCommand(){
  function methods(o){
    var out=[],seen={},p=o
    for(var depth=0;p&&depth<8;depth++,p=Object.getPrototypeOf(p)){
      var names=[]
      try{names=Object.getOwnPropertyNames(p)}catch(_){}
      for(var i=0;i<names.length;i++){
        var n=names[i]
        if(seen[n])continue
        seen[n]=1
        try{if(typeof o[n]==='function')out.push(n)}catch(_){}
      }
    }
    return out.sort()
  }
  function pick(list,re){return list.filter(function(n){return re.test(n)})}
  function safe(fn){try{return fn()}catch(e){return null}}
  try{
    var sheets=Api.GetSheets?Api.GetSheets():[]
    var sh=(sheets&&sheets.length)?sheets[0]:(Api.GetActiveSheet?Api.GetActiveSheet():null)
    if(!sh)return {ok:false,error:'worksheet unavailable'}
    var range=sh.GetRange?sh.GetRange('A1:C4'):null
    var af=sh.GetAutoFilter?safe(function(){return sh.GetAutoFilter()}):null
    var validation=range&&range.GetValidation?safe(function(){return range.GetValidation()}):null
    var dns=Api.GetDefNames?safe(function(){return Api.GetDefNames()}):null
    var dn=(dns&&dns.length)?dns[0]:null
    var fp=sh.GetFreezePanes?safe(function(){return sh.GetFreezePanes()}):null
    var apiMethods=methods(Api),sheetMethods=methods(sh),rangeMethods=methods(range),afMethods=methods(af),validationMethods=methods(validation),defNameMethods=methods(dn),freezeMethods=methods(fp)
    var interesting=/Filter|Sort|Valid|DefName|Defined|Name|Freeze|Protect|Pivot|Hyperlink/i
    return {
      ok:true,
      source:'live-coedit-editor',
      objects:{
        Api:{present:true,interesting:pick(apiMethods,interesting)},
        ApiWorksheet:{present:!!sh,interesting:pick(sheetMethods,interesting)},
        ApiRange:{present:!!range,interesting:pick(rangeMethods,interesting)},
        ApiAutoFilter:{present:!!af,methods:afMethods},
        ApiValidation:{present:!!validation,methods:validationMethods},
        ApiDefName:{present:!!dn,methods:defNameMethods},
        ApiFreezePanes:{present:!!fp,interesting:pick(freezeMethods,interesting)}
      },
      candidates:{
        sort:{range:pick(rangeMethods,/Sort/i),worksheet:pick(sheetMethods,/Sort/i),autoFilter:pick(afMethods,/Sort/i)},
        filter:{worksheet:pick(sheetMethods,/Filter/i),range:pick(rangeMethods,/Filter/i),autoFilter:pick(afMethods,/Filter|Criteria|Range|Column|Clear|Remove/i)},
        validation:{worksheet:pick(sheetMethods,/Valid/i),range:pick(rangeMethods,/Valid/i),api:pick(apiMethods,/Valid/i),object:validationMethods},
        definedNames:{api:pick(apiMethods,/DefName|DefinedName/i),worksheet:pick(sheetMethods,/DefName|DefinedName/i),object:defNameMethods},
        freeze:{api:pick(apiMethods,/Freeze/i),worksheet:pick(sheetMethods,/Freeze/i),object:pick(freezeMethods,/./)},
        protection:{api:pick(apiMethods,/Protect/i),worksheet:pick(sheetMethods,/Protect/i),range:pick(rangeMethods,/Protect|Locked/i)},
        pivot:{api:pick(apiMethods,/Pivot/i),worksheet:pick(sheetMethods,/Pivot/i),range:pick(rangeMethods,/Pivot/i)}
      }
    }
  }catch(e){return {ok:false,error:String(e&&e.message?e.message:e)}}
}

async function main(){
  const base=process.env.EURO_NC_BASE_URL,fileId=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD
  if(!base||!fileId||!user||!password)throw new Error('required EURO_NC_* env missing')
  const {chromium}=loadPlaywright(),browser=await chromium.launch()
  try{
    const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage()
    await page.goto(`${base}/login`,{waitUntil:'domcontentloaded',timeout:60000})
    await page.fill('#user',user);await page.fill('#password',password)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:60000}).catch(()=>null),page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if(/\/login/.test(page.url()))throw new Error('login failed')
    await page.goto(`${base}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:60000})
    await page.waitForTimeout(22000)
    const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()))
    if(!frame)throw new Error('spreadsheeteditor frame missing')
    const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':null)
    if(!apiHely)throw new Error('callCommand unavailable')
    const probe=await call(frame,apiHely)
    console.log(JSON.stringify({milestone:'M6',kind:'runtime-capability-probe',editor:'spreadsheeteditor',apiHely,probe},null,2))
    if(!probe||!probe.ok)process.exitCode=2
  }finally{await browser.close()}
}
main().catch(e=>{console.error(JSON.stringify({milestone:'M6',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1})
