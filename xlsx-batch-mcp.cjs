'use strict'
const {z}=require('zod')
const batch=require('./xlsx-persistent-batch.cjs')
const coedit=require('./coedit.cjs')
const str=z.string().min(1)
const scalar=z.union([z.string(),z.number(),z.boolean()])
// Strict outer operation fields prevent silently discarded options. Family
// planners validate the semantic combinations before credentials or browser use.
const operation=z.object({
 intent:z.enum(['format_range','layout_range','merge_range','unmerge_range','freeze_panes','unfreeze_panes','sort_range','filter_range','clear_filter','set_validation','clear_validation','set_defined_name','delete_defined_name']),
 sheet:str.optional(),range:str.optional(),format:z.record(z.string(),z.unknown()).optional(),
 type:z.enum(['column.width','row.height','columns.hidden','rows.hidden']).optional(),
 width:z.number().positive().optional(),height:z.number().positive().optional(),hidden:z.boolean().optional(),
 mode:z.enum(['rows','columns','at']).optional(),count:z.number().int().positive().optional(),
 keyRange:str.optional(),order:z.enum(['asc','desc']).optional(),hasHeaders:z.boolean().optional(),
 field:z.number().int().positive().optional(),criteria1:z.union([scalar,z.array(scalar)]).optional(),criteria2:scalar.nullable().optional(),operator:str.optional(),
 validationType:str.optional(),alertStyle:str.optional(),formula1:scalar.optional(),formula2:scalar.optional(),
 name:str.optional(),refersTo:str.optional()
}).strict().superRefine((op,ctx)=>{
 const fields={format_range:['sheet','range','format'],layout_range:['sheet','range','type','width','height','hidden'],merge_range:['sheet','range'],unmerge_range:['sheet','range'],freeze_panes:['sheet','mode','range','count'],unfreeze_panes:['sheet'],sort_range:['sheet','range','keyRange','order','hasHeaders'],filter_range:['sheet','range','field','criteria1','criteria2','operator'],clear_filter:['sheet','range'],set_validation:['sheet','range','validationType','alertStyle','operator','formula1','formula2'],clear_validation:['sheet','range'],set_defined_name:['name','refersTo'],delete_defined_name:['name']}
 for(const key of Object.keys(op))if(key!=='intent'&&!fields[op.intent].includes(key))ctx.addIssue({code:'custom',path:[key],message:'Field not supported by this intent'})
})
const input={file_id:z.string().regex(/^[1-9][0-9]*$/),operations:z.array(operation).min(1).max(100)}
const schema=z.object(input).strict()
function result(payload){return {isError:payload.ok!==true,content:[{type:'text',text:JSON.stringify(payload)}]}}
function makeHandler(deps={}){
 const auth=deps.coedit||coedit,execute=deps.execute||batch.executeBatchTaskInPersistentSession
 return async args=>{
  const parsed=schema.safeParse(args)
  if(!parsed.success)return result({ok:false,outcome:'xlsx-batch-invalid-input',authority:'PLAN_ONLY',writeAllowed:false})
  const {file_id,operations}=parsed.data,task={operations},plan=batch.planTask(task)
  if(!plan.ok)return result({...plan,writeAllowed:false})
  try{
   const caller=auth.detectCallerId()
   if(!caller.ok)return result({ok:false,outcome:'azonossag-hiany',authority:'PLAN_ONLY',writeAllowed:false})
   const cred=await auth.credentialsFor(caller.id)
   if(!cred.ok)return result({ok:false,outcome:'konfig-hiany',authority:'PLAN_ONLY',writeAllowed:false,callerId:caller.id})
   const r=await execute({url:cred.url,user:cred.user,pass:cred.pass,fileId:file_id,task})
   return result({...r,callerId:caller.id,ncUser:cred.user})
  }catch(_err){
   // Exceptions can embed authenticated URLs. Never serialize them to MCP.
   return result({ok:false,outcome:'xlsx-batch-execution-error',authority:'UNKNOWN',writeAllowed:false,partialChangesPossible:true})
  }
 }
}
function register(server,deps){
 server.tool('office_xlsx_batch',
  'Execute final-state Excel operations in ONE persistent ONLYOFFICE editor session per call. '+
  'Supports format, fixed layout, merge/unmerge, freeze, sort, filter, durable validation, workbook name set/delete. '+
  'Read/edit/read, read-only whole-task verification, then one save barrier and close. '+
  'Submit all compatible goals together; only one goal per family per sheet/name. Not transactional. '+
  'No AutoFit, charts, pivots, conditional formatting or sheet creation in this tool. '+
  'Examples: {intent:"format_range",sheet:"Sheet1",range:"A1:B2",format:{bold:true}}; '+
  '{intent:"set_defined_name",name:"ReportRange",refersTo:"=Sheet1!$A$1:$B$2"}.',input,makeHandler(deps))
}
module.exports={register,makeHandler,schema}
