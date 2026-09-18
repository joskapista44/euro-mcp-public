'use strict'
async function run(){
 const caller=require('./coedit.cjs').detectCallerId();if(!caller.ok||caller.id!=='elliot')throw Error('allowlisted elliot required')
 const {getSecret}=require('/home/user/marveen/dist/web/vault.js'),pass=getSecret('Elliot_nc_pass','xlsx-showcase-chart-identity-probe');if(!pass)throw Error('vault credential missing')
 process.env.EURO_PLAYWRIGHT_PATH||='/home/user/marveen/node_modules/playwright'
 const {openMinimalXlsxSession}=require('./xlsx-minimal-editor-session.cjs')
 const task=require('./xlsx-visual-showcase-contract.cjs').buildShowcaseTask(String(process.env.EURO_XLSX_SHOWCASE_RUN_ID||'768SVZ').toUpperCase())
 const planner=require('./xlsx-agent-chart-task.cjs'),charts=require('./xlsx-persistent-chart.cjs')
 const operations=task.operations.filter(op=>op.intent==='set_chart').map(op=>{const plan=planner.planTask({operations:[op]});if(!plan.ok)throw Error(plan.outcome);return plan.operation})
 const session=await openMinimalXlsxSession({url:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',user:'elliot',pass,fileId:process.env.EURO_XLSX_FILE_ID||'1236770',timeoutMs:30000})
 let observations=[]
 try{for(const operation of operations){const read=await charts.semanticState(session,operation);observations.push({name:operation.name,geometryIdentityName:operation.geometryIdentityName,geometryIdentityRef:operation.geometryIdentityRef,ok:read?.ok,state:read?.state||null,match:read?.ok?charts.match(read.state,operation):false})}}
 finally{await session.close()}
 const result={ok:observations.length===2,observations,oneEditorSession:true,writes:session.writes,closed:session.closed}
 console.error('CHART IDENTITY PROBE',JSON.stringify(result,null,2))
 if(!result.ok||session.writes!==0||!session.closed)process.exitCode=1
}
if(require.main===module)run().catch(e=>{console.error('Chart identity probe failed:',e.stack||e);process.exitCode=1})
