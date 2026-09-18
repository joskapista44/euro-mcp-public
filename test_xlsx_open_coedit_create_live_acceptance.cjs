'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')

;(async()=>{
 const file_id=process.env.EURO_XLSX_COEDIT_FILE_ID
 if(!/^[1-9][0-9]*$/.test(file_id||''))throw Error('EURO_XLSX_COEDIT_FILE_ID is required')
 const caller=require('./coedit.cjs').detectCallerId()
 if(!caller.ok||caller.id!=='elliot')throw Error('This acceptance requires allowlisted caller elliot')
 const vault=require('/home/user/marveen/dist/web/vault.js'),pass=vault.getSecret('Elliot_nc_pass','xlsx-open-coedit-create-live-acceptance')
 if(!pass)throw Error('vault secret missing')
 const client=new Client({name:'xlsx-open-coedit-create-live-acceptance',version:'1'})
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env:{...process.env,EURO_COEDIT_NC_URL:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',ELLIOT_NEXTCLOUD_USER:'elliot',ELLIOT_NEXTCLOUD_APP_PASSWORD:pass,EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'},stderr:'inherit'})
 try{
  await client.connect(transport)
  const tag=String(Date.now()).slice(-7),sheet='EURO_COEDIT_'+tag
  const reply=await client.callTool({name:'office_xlsx_batch',arguments:{file_id,operations:[
   {intent:'create_sheet',name:sheet},
   {intent:'write_range',sheet,range:'A1:B2',values:[['Live co-editing','verified'],['Sheet structure',tag]]},
   {intent:'move_sheet',sheet,referenceSheet:'Sheet1',position:'before'}
  ],readbacks:[{sheet,range:'A1:B2'}]}},undefined,{timeout:90000})
  const result=JSON.parse(reply.content.find(item=>item.type==='text').text)
  console.log('OPEN COEDIT CREATE/MOVE',JSON.stringify(result,null,2))
  assert.equal(reply.isError,false);assert.equal(result.ok,true);assert.equal(result.authority,'LIVE_VERIFY');assert.equal(result.noOp,false)
  assert.equal(result.persistentSession?.oneEditorSession,true);assert.equal(result.persistentSession?.writes,3);assert.equal(result.persistentSession?.persistenceBarrier?.ok,true)
  assert.equal(result.wholeTaskVerification?.readOnly,true);assert.equal(result.wholeTaskVerification?.checks?.length,3)
  assert.equal(result.wholeTaskVerification?.readbacks?.[0]?.authority,'LIVE_READ')
  console.log('XLSX OPEN COEDIT CREATE/MOVE LIVE ACCEPTANCE: PASS')
 }finally{await client.close()}
})().catch(error=>{console.error(error.stack||error);process.exitCode=1})
