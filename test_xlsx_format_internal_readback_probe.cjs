'use strict'
const {withPersistentXlsxSession}=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-format-internal-readback-probe');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
function command(){
 function simple(v,d){if(d>4)return '[depth]';if(v==null||typeof v==='string'||typeof v==='number'||typeof v==='boolean')return v;if(Array.isArray(v))return v.slice(0,20).map(function(x){return simple(x,d+1)});var o={},names=[];try{names=Object.getOwnPropertyNames(v)}catch(_){return String(v)}for(var i=0;i<names.length&&i<80;i++){var n=names[i];if(typeof v[n]==='function')continue;if(!/(font|bold|italic|name|size|color|align|hor|vert|wrap|num|format|fill|xf|style)/i.test(n)&&d>0)continue;try{o[n]=simple(v[n],d+1)}catch(e){o[n]='[error]'}}return o}
 function call(o,n){try{return typeof o[n]==='function'?{ok:true,value:simple(o[n](),0)}:{ok:false,missing:true}}catch(e){return {ok:false,error:String(e&&e.message||e)}}}
 var sh=Api.GetSheets()[0],ar=sh.GetRange('A1'),r=ar.range;
 var methods=['getFont','getAlign','getFill','getFillColor','getNumFormat','getNumFormatStr','getStyle','getXfs','getCompiledStyleCustom','getStyleName','getXfId'];var calls={};methods.forEach(function(n){calls[n]=call(r,n)});
 var cell=null;try{cell=typeof r.getLeftTopCell==='function'?r.getLeftTopCell():null}catch(_){}
 var cellMethods={};if(cell){['getFont','getAlign','getFill','getFillColor','getNumFormat','getNumFormatStr','getStyle','getXfs','getCompiledStyleCustom','getStyleName','getXfId'].forEach(function(n){cellMethods[n]=call(cell,n)})}
 return {ok:true,outcome:'internal-format-readback-probe',source:'live-coedit-editor',sheet:sh.GetName(),range:'A1',rangeMethods:calls,leftTopCell:simple(cell,0),leftTopCellMethods:cellMethods}
}
;(async()=>{const r=await withPersistentXlsxSession({url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50},async api=>{const body=`return (${command.toString()})();`;const x=await api.session.frame.evaluate(({u,body})=>new Promise(resolve=>{const e=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const f=v=>{if(!done){done=true;resolve(v)}};try{e.callCommand(new Function(body),false,v=>f(v||{ok:false,outcome:'empty'}))}catch(err){f({ok:false,outcome:'error',error:String(err&&err.message||err)})}setTimeout(()=>f({ok:false,outcome:'timeout'}),10000)}),{u:api.session.apiWhere,body});return {...x,authority:x.ok?'LIVE_READ':'PLAN_ONLY',writeAllowed:false,noOp:true}});console.log(JSON.stringify(r,null,2));if(!r.ok)process.exitCode=1})().catch(e=>{console.error(e);process.exitCode=1})
