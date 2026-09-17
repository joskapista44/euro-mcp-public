'use strict'
const assert=require('assert/strict')
const {McpServer}=require('@modelcontextprotocol/sdk/server/mcp.js')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {InMemoryTransport}=require('@modelcontextprotocol/sdk/inMemory.js')
const {register}=require('./xlsx-batch-mcp.cjs')
;(async()=>{
 let credentials=0,calls=0,mode='ok'
 const auth={detectCallerId:()=>({ok:mode!=='caller',id:'agent'}),credentialsFor:async()=>{credentials++;return {ok:mode!=='credentials',url:'https://example.invalid',user:'agent',pass:'secret-test'}}}
 const server=new McpServer({name:'batch-test',version:'1'})
 register(server,{coedit:auth,execute:async options=>{calls++;assert.equal(options.pass,'secret-test');assert.equal(options.fileId,'123');assert.equal(options.task.operations[0].intent,'set_defined_name');if(mode==='throw')throw Error('secret-test');return {ok:mode!=='failure',authority:'LIVE_VERIFY',noOp:true}}})
 const client=new Client({name:'test',version:'1'}),[a,b]=InMemoryTransport.createLinkedPair()
 await Promise.all([server.connect(a),client.connect(b)])
 const request={file_id:'123',operations:[{intent:'set_defined_name',name:'Report',refersTo:'=Sheet1!$A$1'}]}
 const call=args=>client.callTool({name:'office_xlsx_batch',arguments:args})
 try{
  assert((await client.listTools()).tools.some(t=>t.name==='office_xlsx_batch'))
  assert.equal((await call({...request,operations:[{intent:'unknown'}]})).isError,true)
  assert.equal((await call({...request,operations:[{intent:'sort_range',sheet:'S',range:'A1:B3',keyRange:'C1:C3'}]})).isError,true)
  assert.equal(credentials,0);assert.equal(calls,0)
  mode='caller';assert.equal((await call(request)).isError,true);assert.equal(credentials,0)
  mode='credentials';assert.equal((await call(request)).isError,true);assert.equal(calls,0)
  mode='ok';const good=await call(request);assert.equal(good.isError,false);assert.equal(calls,1);assert(!JSON.stringify(good).includes('secret-test'))
  mode='failure';assert.equal((await call(request)).isError,true)
  mode='throw';const failed=await call(request);assert.equal(failed.isError,true);assert(!JSON.stringify(failed).includes('secret-test'))
 }finally{await client.close();await server.close()}
 console.log('XLSX BATCH MCP STATIC: PASS (real MCP transport, schema, preflight, identity, credentials, dispatch, errors)')
})().catch(e=>{console.error(e);process.exitCode=1})
