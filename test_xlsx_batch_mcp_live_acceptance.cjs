'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')
;(async()=>{
 const caller=require('./coedit.cjs').detectCallerId()
 if(!caller.ok)throw Error('MCP caller is not configured/allowlisted')
 if(caller.id!=='elliot')throw Error('This acceptance fixture requires elliot')
 const vault=require('/home/user/marveen/dist/web/vault.js')
 const pass=vault.getSecret('Elliot_nc_pass','xlsx-batch-mcp-live-acceptance')
 if(!pass)throw Error('vault secret missing')
 const client=new Client({name:'xlsx-mcp-live-acceptance',version:'1'})
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env:{...process.env,EURO_COEDIT_NC_URL:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',ELLIOT_NEXTCLOUD_USER:'elliot',ELLIOT_NEXTCLOUD_APP_PASSWORD:pass,EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'},stderr:'inherit'})
 try{
  await client.connect(transport)
  assert((await client.listTools()).tools.some(t=>t.name==='office_xlsx_batch'))
  const name='EURO_MCP_'+Date.now(),file_id=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
  const args={file_id,operations:[{intent:'set_defined_name',name,refersTo:'=Sheet1!$XFD$30'}]}
  for(let i=0;i<2;i++){
   const reply=await client.callTool({name:'office_xlsx_batch',arguments:args},undefined,{timeout:90000})
   const r=JSON.parse(reply.content.find(c=>c.type==='text').text)
   console.log(i?'MCP PERSISTED RETRY':'MCP APPLY',JSON.stringify(r.ok?{ok:r.ok,noOp:r.noOp,callerId:r.callerId,session:r.persistentSession,wholeTaskVerification:r.wholeTaskVerification}:r,null,2))
   assert.equal(reply.isError,false);assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(r.noOp,i===1)
   assert.equal(r.persistentSession.oneEditorSession,true);assert.equal(r.persistentSession.writes,i?0:1)
   assert.equal(r.wholeTaskVerification.readOnly,true)
   if(i)assert.equal(r.persistentSession.persistenceBarrier,null)
   else assert.equal(r.persistentSession.persistenceBarrier.ok,true)
  }
  console.log('XLSX BATCH MCP LIVE ACCEPTANCE: PASS')
 }finally{await client.close()}
})().catch(e=>{console.error(e.stack||e);process.exitCode=1})
