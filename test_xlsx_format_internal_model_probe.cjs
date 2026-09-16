'use strict'
const {withPersistentXlsxSession}=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-format-internal-model-probe');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
function cmd(){
 function val(v){if(v==null||['string','number','boolean'].indexOf(typeof v)>=0)return v;if(Array.isArray(v))return '[array '+v.length+']';return Object.prototype.toString.call(v)}
 function inspect(o,re){var out={};if(!o)return out;var seen={};for(var cur=o,d=0;cur&&d<5;cur=Object.getPrototypeOf(cur),d++){try{Object.getOwnPropertyNames(cur).forEach(function(n){if(seen[n]||!re.test(n))return;seen[n]=1;try{out[n]={type:typeof o[n],value:typeof o[n]==='function'?'[function]':val(o[n])}}catch(e){out[n]={error:String(e&&e.message||e)}}})}catch(_){}}return out}
 var sh=Api.GetSheets()[0],r=sh.GetRange('A1');
 var result={apiRange:inspect(r,/(range|worksheet|ws|font|style|format|align|bold|italic|wrap|color|fill|text|bbox|address)/i)};
 var candidates=['range','_range','worksheet','_worksheet','ws','_ws','Api','api'];
 candidates.forEach(function(n){try{if(r[n]&&typeof r[n]==='object')result['range.'+n]=inspect(r[n],/(font|style|format|align|bold|italic|wrap|color|fill|text|cell|range|value|name|size|hor|vert|xf)/i)}catch(_){}});
 try{var ed=(typeof Asc!=='undefined'&&Asc.editor)?Asc.editor:null;result.editor=inspect(ed,/(workbook|model|worksheet|cell|range|font|style|format|api)/i);if(ed){['wb','workbook','wbModel','model'].forEach(function(n){try{if(ed[n])result['editor.'+n]=inspect(ed[n],/(worksheet|cell|range|font|style|format|model|name|active)/i)}catch(_){}})}}catch(e){result.editorError=String(e&&e.message||e)}
 return {ok:true,outcome:'internal-format-model-probe',source:'live-coedit-editor',sheet:sh.GetName(),range:'A1',result:result}
}
;(async()=>{const r=await withPersistentXlsxSession({url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50},async api=>{const body=`return (${cmd.toString()})();`;const x=await api.session.frame.evaluate(({u,body})=>new Promise(resolve=>{const e=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const f=v=>{if(!done){done=true;resolve(v)}};try{e.callCommand(new Function(body),false,v=>f(v||{ok:false,outcome:'empty'}))}catch(err){f({ok:false,outcome:'error',error:String(err&&err.message||err)})}setTimeout(()=>f({ok:false,outcome:'timeout'}),10000)}),{u:api.session.apiWhere,body});return {...x,authority:x.ok?'LIVE_READ':'PLAN_ONLY',writeAllowed:false,noOp:true}});console.log(JSON.stringify(r,null,2));if(!r.ok)process.exitCode=1})().catch(e=>{console.error(e);process.exitCode=1})
