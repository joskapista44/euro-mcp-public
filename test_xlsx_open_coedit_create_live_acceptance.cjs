'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')

async function resolveFileId({url,user,pass,name}){
 const auth='Basic '+Buffer.from(`${user}:${pass}`).toString('base64')
 const body='<?xml version="1.0"?><d:propfind xmlns:d="DAV:" xmlns:oc="http://owncloud.org/ns"><d:prop><oc:fileid/></d:prop></d:propfind>'
 const response=await fetch(url.replace(/\/$/,'')+'/remote.php/dav/files/'+encodeURIComponent(user)+'/',{method:'PROPFIND',headers:{Authorization:auth,Depth:'infinity','Content-Type':'application/xml'},body})
 if(!response.ok)throw Error(`PROPFIND HTTP ${response.status}`)
 const xml=await response.text(),hits=[]
 for(const block of xml.match(/<d:response[\s\S]*?<\/d:response>/g)||[]){
  const href=/<d:href>([^<]+)<\/d:href>/.exec(block)?.[1],fileId=/<oc:fileid>([^<]+)<\/oc:fileid>/.exec(block)?.[1]
  if(!href||!fileId)continue
  const basename=decodeURIComponent(href.split('/').filter(Boolean).at(-1)||'')
  if(basename===name)hits.push({href,fileId})
 }
 if(hits.length!==1)throw Error(`${name}: exactly one DAV match required, found ${hits.length}`)
 return hits[0].fileId
}

;(async()=>{
 const caller=require('./coedit.cjs').detectCallerId()
 if(!caller.ok||caller.id!=='elliot')throw Error('This acceptance requires allowlisted caller elliot')
 const vault=require('/home/user/marveen/dist/web/vault.js'),pass=vault.getSecret('Elliot_nc_pass','xlsx-open-coedit-create-live-acceptance')
 if(!pass)throw Error('vault secret missing')
 const url=process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',user='elliot'
 const file_id=process.env.EURO_XLSX_COEDIT_FILE_ID||await resolveFileId({url,user,pass,name:process.env.EURO_XLSX_COEDIT_FILE_NAME||'mcp_test.xlsx'})
 if(!/^[1-9][0-9]*$/.test(file_id))throw Error('resolved file_id is invalid')
 console.log(`TARGET file_id=${file_id}`)
 const client=new Client({name:'xlsx-open-coedit-create-live-acceptance',version:'1'})
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env:{...process.env,EURO_COEDIT_NC_URL:url,ELLIOT_NEXTCLOUD_USER:user,ELLIOT_NEXTCLOUD_APP_PASSWORD:pass,EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'},stderr:'inherit'})
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
