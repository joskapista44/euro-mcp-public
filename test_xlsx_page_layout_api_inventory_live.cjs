'use strict'
const assert=require('assert/strict')
const persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-page-layout-api-inventory');if(!r)throw new Error('vault secret not found: '+id);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const result=await persistent.withPersistentXlsxSession(options,async api=>{
  const inventory=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{
   const e=where==='window.editor'?window.editor:(window.Asc||{}).editor
   const body=`try{
    var s=Api.GetActiveSheet();if(!s)return {ok:false,outcome:'no-active-sheet'}
    var candidates=['GetPageOrientation','SetPageOrientation','GetTopMargin','SetTopMargin','GetBottomMargin','SetBottomMargin','GetLeftMargin','SetLeftMargin','GetRightMargin','SetRightMargin','GetPrintGridlines','SetPrintGridlines','GetPrintHeadings','SetPrintHeadings','GetPageSetup','SetPageSetup','GetPrintArea','SetPrintArea','ClearPrintArea','GetFitToWidth','SetFitToWidth','GetFitToHeight','SetFitToHeight','GetScale','SetScale','GetPaperSize','SetPaperSize','GetPageSize','SetPageSize','GetPrintTitles','SetPrintTitles','GetRowsToRepeatAtTop','SetRowsToRepeatAtTop','GetColumnsToRepeatAtLeft','SetColumnsToRepeatAtLeft','GetHeader','SetHeader','GetFooter','SetFooter','GetPageBreaks','AddPageBreak','ResetAllPageBreaks']
    var methods={};for(var i=0;i<candidates.length;i++)methods[candidates[i]]=typeof s[candidates[i]]
    var getters={};[['orientation','GetPageOrientation'],['topMargin','GetTopMargin'],['bottomMargin','GetBottomMargin'],['leftMargin','GetLeftMargin'],['rightMargin','GetRightMargin'],['printGridlines','GetPrintGridlines'],['printHeadings','GetPrintHeadings']].forEach(function(x){try{getters[x[0]]=typeof s[x[1]]==='function'?{ok:true,value:s[x[1]]()}:{ok:false,outcome:'unavailable'}}catch(err){getters[x[0]]={ok:false,error:String(err&&err.message||err)}}})
    var names=[],seen={};var o=s;for(var depth=0;o&&depth<6;depth++,o=Object.getPrototypeOf(o)){try{Object.getOwnPropertyNames(o).forEach(function(n){if(!seen[n]&&/(page|print|margin|header|footer|break|fit|scale|orient|paper)/i.test(n)){seen[n]=1;names.push(n)}})}catch(_){}}
    names.sort()
    return {ok:true,sheet:typeof s.GetName==='function'?s.GetName():null,methods:methods,getters:getters,reflected:names}
   }catch(err){return {ok:false,outcome:'inventory-error',error:String(err&&err.stack||err)}}`
   e.callCommand(new Function(body),false,resolve)
  }),{where:api.session.apiWhere})
  return inventory?.ok?{ok:true,outcome:'xlsx-page-layout-api-inventory',authority:'LIVE_VERIFY',noOp:true,inventory}:{ok:false,outcome:'xlsx-page-layout-api-inventory-failed',authority:'LIVE_READ',noOp:true,inventory}
 })
 console.log('XLSX PAGE LAYOUT API INVENTORY',JSON.stringify(result,null,2))
 assert.equal(result.ok,true);assert.equal(result.authority,'LIVE_VERIFY');assert.equal(result.noOp,true);assert.equal(result.persistentSession?.writes,0);assert.equal(result.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PAGE LAYOUT API INVENTORY: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
