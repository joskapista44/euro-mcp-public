'use strict'
function pivotInventoryCommand(sourceSheetName,pivotName){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function safe(fn){try{return fn()}catch(e){return {error:String(e&&e.message||e)}}}
 function nameOf(o){return o&&has(o,'GetName')?safe(()=>o.GetName()):null}
 const sheets=has(Api,'GetSheets')?(safe(()=>Api.GetSheets())||[]):[]
 const pivots=has(Api,'GetAllPivotTables')?(safe(()=>Api.GetAllPivotTables())||[]):[]
 const source=has(Api,'GetSheet')?safe(()=>Api.GetSheet(sourceSheetName)):null
 const named=has(Api,'GetPivotByName')?safe(()=>Api.GetPivotByName(pivotName)):null
 function describe(v){return {isNull:v==null,type:typeof v,error:v&&v.error,primitive:typeof v==='string'||typeof v==='number'?v:undefined,methods:['GetAddress','GetName','GetWorksheet'].filter(n=>has(v,n)),address:has(v,'GetAddress')?safe(()=>v.GetAddress()):null,name:has(v,'GetName')?safe(()=>v.GetName()):null}}
 const namedIdentity={present:!!named,error:named&&named.error,getters:['GetName','GetSource','GetParent','GetRowFields','GetColumnFields','GetDataFields','GetStyleName','GetData'].filter(n=>has(named,n)),source:has(named,'GetSource')?describe(safe(()=>named.GetSource())):null,parent:has(named,'GetParent')?describe(safe(()=>named.GetParent())):null}
 return {ok:true,outcome:'pivot-inventory-probed',source:'live-coedit-editor',sourceSheet:sourceSheetName,
  namedIdentity,sheets:Array.isArray(sheets)?sheets.map((s,index)=>({index,name:nameOf(s)})):[],
  sourceState:source&&typeof source==='object'?{present:true,usedRange:has(source,'GetUsedRange')?safe(()=>{const r=source.GetUsedRange();return r&&has(r,'GetAddress')?r.GetAddress():null}):null,autoFilter:has(source,'GetAutoFilter')?safe(()=>{const f=source.GetAutoFilter();return f&&has(f,'GetRange')?{present:true,range:f.GetRange().GetAddress()}:null}):null}:{present:false},
  pivots:Array.isArray(pivots)?pivots.map((p,index)=>({index,name:nameOf(p),parent:has(p,'GetParent')?safe(()=>nameOf(p.GetParent())):null,source:has(p,'GetSource')?safe(()=>{const r=p.GetSource();return r&&has(r,'GetAddress')?r.GetAddress():null}):null,rowFields:has(p,'GetRowFields')?safe(()=>p.GetRowFields().length):null,columnFields:has(p,'GetColumnFields')?safe(()=>p.GetColumnFields().length):null,dataFields:has(p,'GetDataFields')?safe(()=>p.GetDataFields().length):null,style:has(p,'GetStyleName')?safe(()=>p.GetStyleName()):null})):[]}
}
async function run(){
 const caller=require('./coedit.cjs').detectCallerId();if(!caller.ok||caller.id!=='elliot')throw Error('allowlisted elliot required')
 const {getSecret}=require('/home/user/marveen/dist/web/vault.js'),pass=getSecret('Elliot_nc_pass','xlsx-showcase-pivot-inventory-probe');if(!pass)throw Error('vault credential missing')
 process.env.EURO_PLAYWRIGHT_PATH||='/home/user/marveen/node_modules/playwright'
 const {openMinimalXlsxSession}=require('./xlsx-minimal-editor-session.cjs')
 const task=require('./xlsx-visual-showcase-contract.cjs').buildShowcaseTask(String(process.env.EURO_XLSX_SHOWCASE_RUN_ID||'768SVZ').toUpperCase())
 const spec={...task.operations.find(op=>op.intent==='create_pivot'),apply:false}
 const {pivotObserveCommand}=require('./xlsx-pivot-observer.cjs')
 const session=await openMinimalXlsxSession({url:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',user:'elliot',pass,fileId:process.env.EURO_XLSX_FILE_ID||'1236770',timeoutMs:30000})
 let result
 try{
  const body=`var inventory=(${pivotInventoryCommand.toString()})(${JSON.stringify(spec.sourceSheet)},${JSON.stringify(spec.name)});inventory.observation=(${pivotObserveCommand.toString()})(${JSON.stringify(spec)});return inventory;`
  result=await session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;clearTimeout(timer);resolve(v)}};const timer=setTimeout(()=>finish({ok:false,outcome:'callback-timeout'}),10000);try{e.callCommand(new Function(body),false,v=>finish(v||{ok:false,outcome:'empty-callback'}))}catch(_){finish({ok:false,outcome:'dispatch-failed'})}}),{where:session.apiWhere,body})
 }finally{await session.close()}
 console.error('PIVOT INVENTORY',JSON.stringify({...result,oneEditorSession:true,writes:session.writes,closed:session.closed}))
 if(!result?.ok||session.writes!==0||!session.closed)process.exitCode=1
}
module.exports={pivotInventoryCommand}
if(require.main===module)run().catch(e=>{console.error('Pivot inventory probe failed:',e.message);process.exitCode=1})
