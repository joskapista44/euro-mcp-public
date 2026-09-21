'use strict'
const assert=require('assert/strict')
const {McpServer}=require('@modelcontextprotocol/sdk/server/mcp.js')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {InMemoryTransport}=require('@modelcontextprotocol/sdk/inMemory.js')
const {register,schema}=require('./xlsx-batch-mcp.cjs')
;(async()=>{
 let credentials=0,calls=0,mode='ok'
 const auth={detectCallerId:()=>({ok:mode!=='caller',id:'agent'}),credentialsFor:async()=>{credentials++;return {ok:mode!=='credentials',url:'https://example.invalid',user:'agent',pass:'secret-test'}}}
 const server=new McpServer({name:'batch-test',version:'1'})
 register(server,{coedit:auth,execute:async options=>{calls++;assert.equal(options.pass,'secret-test');assert.equal(options.fileId,'123');assert(['set_defined_name','create_sheet','set_chart','set_print_area'].includes(options.task.operations[0].intent));if(options.task.operations[0].intent==='set_print_area')assert.deepEqual(options.task.operations[0],{intent:'set_print_area',sheet:'S',range:'A1:H20',mode:'set'});if(options.task.readbacks)assert.deepEqual(options.task.readbacks,[{sheet:'S',range:'A1:B2'}]);if(mode==='throw')throw Error('secret-test');return {ok:mode!=='failure',authority:'LIVE_VERIFY',noOp:true}}})
 const client=new Client({name:'test',version:'1'}),[a,b]=InMemoryTransport.createLinkedPair()
 await Promise.all([server.connect(a),client.connect(b)])
 const request={file_id:'123',operations:[{intent:'set_defined_name',name:'Report',refersTo:'=Sheet1!$A$1'}]}
 const call=args=>client.callTool({name:'office_xlsx_batch',arguments:args})
 try{
  assert((await client.listTools()).tools.some(t=>t.name==='office_xlsx_batch'))
  assert.equal(schema.safeParse({file_id:'123',operations:[{intent:'rename_chart',sheet:'S',name:'Old',newName:'New',chartType:'bar',title:'T',width:10,height:20,expectedSeriesCount:1}]}).success,true)
  assert.equal(schema.safeParse({file_id:'123',operations:[{intent:'refresh_pivot',name:'P',sourceSheet:'Source',sourceRange:'A1:C5',rowField:'Region',columnField:'Style',dataField:'Price',styleName:'PivotStyleMedium2',assertions:[{items:['East','A'],expected:15}]}]}).success,true)
  assert.equal(schema.safeParse({file_id:'123',operations:[{intent:'create_pivot',name:'P',sourceSheet:'Source',sourceRange:'A1:C5',sourceIdentityName:'P_SOURCE',rowField:'Region',columnField:'Style',dataField:'Price',styleName:'PivotStyleMedium2',assertions:[{items:['East','A'],expected:15}]}]}).success,true)
  assert.equal(schema.safeParse({...request,readbacks:[{sheet:'S',range:'A1:B2'}]}).success,true)
  assert.equal((await call({...request,operations:[{intent:'unknown'}]})).isError,true)
  assert.equal((await call({...request,operations:[{intent:'sort_range',sheet:'S',range:'A1:B3',keyRange:'C1:C3'}]})).isError,true)
  assert.equal(credentials,0);assert.equal(calls,0)
  mode='caller';assert.equal((await call(request)).isError,true);assert.equal(credentials,0)
  mode='credentials';assert.equal((await call(request)).isError,true);assert.equal(calls,0)
  mode='ok';const good=await call(request);assert.equal(good.isError,false);assert.equal(calls,1);assert(!JSON.stringify(good).includes('secret-test'))
  const withReadback=await call({...request,readbacks:[{sheet:'S',range:'A1:B2'}]});assert.equal(withReadback.isError,false);assert.equal(calls,2)
  assert.equal((await call({...request,operations:[{intent:'set_chart',sheet:'S',name:'C',range:'A1:B3',chartType:'bar',title:'T',expectedSeriesCount:1}]})).isError,false);assert.equal(calls,3)
  const core=await call({file_id:'123',operations:[{intent:'create_sheet',name:'Input'},{intent:'write_range',sheet:'Input',range:'A1:B1',values:[['x',1]]}]})
  assert.equal(core.isError,false);assert.equal(calls,4)
  const printArea=await call({file_id:'123',operations:[{intent:'set_print_area',sheet:'S',printAreaMode:'set',range:'A1:H20'}]})
  assert.equal(printArea.isError,false);assert.equal(calls,5)
  mode='failure';assert.equal((await call(request)).isError,true)
  mode='throw';const failed=await call(request);assert.equal(failed.isError,true);assert(!JSON.stringify(failed).includes('secret-test'))
 }finally{await client.close();await server.close()}
 console.log('XLSX BATCH MCP STATIC: PASS (real MCP transport, schema, preflight, identity, credentials, dispatch, errors)')
})().catch(e=>{console.error(e);process.exitCode=1})
