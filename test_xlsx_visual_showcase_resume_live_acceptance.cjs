'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')
const {buildShowcaseTask}=require('./xlsx-visual-showcase-contract.cjs')
const {normalizeFormula}=require('./verification-contract.cjs')
function scalar(cell){for(const value of [cell?.rawValue,cell?.value])if(typeof value==='number'&&Number.isFinite(value)||typeof value==='string'&&value.trim()!==''&&Number.isFinite(Number(value)))return Number(value);return null}
function at(readback,address){for(const row of readback?.cells||[])for(const cell of row||[])if(cell?.address===address)return cell;return null}
function formula(cell,expected){return normalizeFormula(cell?.formula)===normalizeFormula(expected)}
;(async()=>{
 const caller=require('./coedit.cjs').detectCallerId();if(!caller.ok||caller.id!=='elliot')throw Error('allowlisted elliot required')
 const {getSecret}=require('/home/user/marveen/dist/web/vault.js'),pass=getSecret('Elliot_nc_pass','xlsx-visual-showcase-resume-live-acceptance');if(!pass)throw Error('vault credential missing')
 const runId=String(process.env.EURO_XLSX_SHOWCASE_RUN_ID||'768SVZ').toUpperCase(),task=buildShowcaseTask(runId),file_id=String(process.env.EURO_XLSX_FILE_ID||'1236770')
 const client=new Client({name:'xlsx-visual-showcase-resume',version:'1'})
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env:{...process.env,EURO_COEDIT_NC_URL:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',ELLIOT_NEXTCLOUD_USER:'elliot',ELLIOT_NEXTCLOUD_APP_PASSWORD:pass,EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'},stderr:'inherit'})
 try{
  await client.connect(transport)
  const reply=await client.callTool({name:'office_xlsx_batch',arguments:{file_id,operations:task.operations,readbacks:task.readbacks}},undefined,{timeout:120000})
  const r=JSON.parse(reply.content.find(c=>c.type==='text').text)
  if(!r.ok){console.error('SHOWCASE RESUME FAIL',JSON.stringify({ok:r.ok,outcome:r.outcome,authority:r.authority,session:r.persistentSession,diagnostic:{stepOutcome:r.steps?.at(-1)?.result?.outcome,applied:r.steps?.at(-1)?.result?.applied,failed:r.failed,failedReadback:r.failedReadback}},null,2));assert.fail(r.outcome)}
  assert.equal(reply.isError,false);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(r.persistentSession.oneEditorSession,true);assert.equal(r.persistentSession.closedByWrapper,true)
  assert.equal(r.wholeTaskVerification.readOnly,true);assert.equal(r.wholeTaskVerification.checks.length,53);assert(r.wholeTaskVerification.checks.every(x=>x.ok))
  const reads=r.wholeTaskVerification.readbacks;assert.equal(reads.length,3)
  const dash=reads.find(x=>x.sheet===task.names.dash),plan=reads.find(x=>x.sheet===task.names.plan),data=reads.find(x=>x.sheet===task.names.data)
  assert(dash&&plan&&data);assert.equal(scalar(at(dash,'B4')),task.expected.totalRevenue);assert.equal(scalar(at(dash,'B5')),task.expected.targetRevenue);assert.equal(scalar(at(dash,'B6')),task.expected.variance)
  assert(Math.abs(scalar(at(dash,'B7'))-task.expected.attainment)<1e-12)
  assert(formula(at(dash,'B4'),`=SUM(${task.names.plan}!B4:B9)`));assert(formula(at(data,'F2'),'=D2*E2'))
  const writes=r.persistentSession.writes;assert(Number.isInteger(writes)&&writes>=0)
  if(writes===0)assert.equal(r.persistentSession.persistenceBarrier,null);else assert.equal(r.persistentSession.persistenceBarrier?.ok,true)
  console.error('XLSX VISUAL SHOWCASE RESUME: PASS',JSON.stringify({runId,file_id,noOp:r.noOp,writes,oneEditorSession:true,barrier:writes?true:false,checks:53,readbacks:3,kpi:{totalRevenue:scalar(at(dash,'B4')),targetRevenue:scalar(at(dash,'B5')),variance:scalar(at(dash,'B6')),attainment:scalar(at(dash,'B7'))}}))
 }finally{await client.close()}
})().catch(e=>{console.error(e.stack||e);process.exitCode=1})
