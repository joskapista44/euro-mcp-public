'use strict'

const {openMinimalXlsxSession}=require('./xlsx-minimal-editor-session.cjs')
const {readRangeInFrame}=require('./range-reader.cjs')

async function run(){
 const caller=require('./coedit.cjs').detectCallerId();if(!caller.ok||caller.id!=='elliot')throw Error('allowlisted elliot required')
 const {getSecret}=require('/home/user/marveen/dist/web/vault.js'),pass=getSecret('Elliot_nc_pass','xlsx-showcase-formula-readback-probe');if(!pass)throw Error('vault credential missing')
 process.env.EURO_PLAYWRIGHT_PATH||='/home/user/marveen/node_modules/playwright'
 const runId=String(process.env.EURO_XLSX_SHOWCASE_RUN_ID||'768SVZ').toUpperCase()
 const session=await openMinimalXlsxSession({url:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',user:'elliot',pass,fileId:process.env.EURO_XLSX_FILE_ID||'1236770',timeoutMs:30000})
 let dash,data
 try{
  dash=await readRangeInFrame(session.frame,session.apiWhere,{sheet:`MCP_Dash_${runId}`,range:'B4',maxCells:1,callbackTimeoutMs:10000})
  data=await readRangeInFrame(session.frame,session.apiWhere,{sheet:`MCP_Data_${runId}`,range:'F2',maxCells:1,callbackTimeoutMs:10000})
 }finally{await session.close()}
 const result={ok:dash?.ok===true&&data?.ok===true,outcome:'formula-readback-probed',runId,dashB4:dash?.cells?.[0]?.[0]||dash,dataF2:data?.cells?.[0]?.[0]||data,oneEditorSession:true,writes:session.writes,closed:session.closed}
 console.error('FORMULA READBACK',JSON.stringify(result,null,2))
 if(!result.ok||session.writes!==0||!session.closed)process.exitCode=1
}

if(require.main===module)run().catch(e=>{console.error('Formula readback probe failed:',e.stack||e);process.exitCode=1})
