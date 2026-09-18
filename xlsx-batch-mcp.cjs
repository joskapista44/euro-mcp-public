'use strict'
const {z}=require('zod')
const batch=require('./xlsx-persistent-batch.cjs')
const coedit=require('./coedit.cjs')
const str=z.string().min(1)
const scalar=z.union([z.string(),z.number(),z.boolean()])
const chartPresentation=z.object({legendPosition:str.optional(),horizontalAxisTitle:str.optional(),verticalAxisTitle:str.optional(),dataLabels:z.object({showSeriesName:z.boolean(),showCategoryName:z.boolean(),showValue:z.boolean(),showPercent:z.boolean()}).strict().optional(),style:z.number().int().nonnegative().optional()}).strict()
const chartPosition=z.object({fromCol:z.number(),colOffset:z.number(),fromRow:z.number(),rowOffset:z.number()}).strict()
const chartSeries=z.object({index:z.number().int().nonnegative(),name:str.optional(),valuesRange:str.optional(),xValuesRange:str.optional(),categoryRange:str.optional()}).strict()
// Strict outer operation fields prevent silently discarded options. Family
// planners validate the semantic combinations before credentials or browser use.
const operation=z.object({
 intent:z.enum(['create_sheet','write_range','copy_sheet','rename_sheet','delete_sheet','move_sheet','clear_range','format_range','layout_range','merge_range','unmerge_range','freeze_panes','unfreeze_panes','sort_range','filter_range','clear_filter','set_validation','clear_validation','set_defined_name','rename_defined_name','delete_defined_name','add_conditional_format','delete_conditional_format','create_pivot','refresh_pivot','delete_pivot_sheet','set_chart','rename_chart','delete_chart']),
 sheet:str.optional(),range:str.optional(),format:z.record(z.string(),z.unknown()).optional(),
 values:z.array(z.array(z.union([scalar,z.null()]))).optional(),formulas:z.array(z.array(z.string().nullable())).optional(),
 type:z.enum(['column.width','row.height','columns.hidden','rows.hidden']).optional(),
 width:z.number().positive().optional(),height:z.number().positive().optional(),hidden:z.boolean().optional(),
 mode:z.enum(['rows','columns','at']).optional(),count:z.number().int().positive().optional(),
 referenceSheet:str.optional(),position:z.enum(['before','after']).optional(),
 keyRange:str.optional(),order:z.enum(['asc','desc']).optional(),hasHeaders:z.boolean().optional(),
 field:z.number().int().positive().optional(),criteria1:z.union([scalar,z.array(scalar)]).optional(),criteria2:scalar.nullable().optional(),operator:str.optional(),
 validationType:str.optional(),alertStyle:str.optional(),formula1:scalar.optional(),formula2:scalar.optional(),
 name:str.optional(),newName:str.optional(),refersTo:str.optional(),rule:z.record(z.string(),z.unknown()).optional(),
 sourceSheet:str.optional(),sourceRange:str.optional(),rowField:str.optional(),columnField:str.optional(),dataField:str.optional(),styleName:str.optional(),pivotSheet:str.optional(),destinationRange:str.optional(),repairIncompletePivot:z.boolean().optional(),
 chartType:str.optional(),title:str.optional(),expectedSeriesCount:z.number().int().positive().optional(),inRows:z.boolean().optional(),
 presentation:chartPresentation.optional(),chartPosition:chartPosition.optional(),series:z.array(chartSeries).min(1).optional(),
 assertions:z.array(z.object({items:z.array(scalar).min(1),expected:scalar}).strict()).optional()
}).strict().superRefine((op,ctx)=>{
 const fields={create_sheet:['name'],write_range:['sheet','range','values','formulas'],copy_sheet:['sheet','name'],rename_sheet:['sheet','name'],delete_sheet:['sheet'],move_sheet:['sheet','referenceSheet','position'],clear_range:['sheet','range'],format_range:['sheet','range','format'],layout_range:['sheet','range','type','width','height','hidden'],merge_range:['sheet','range'],unmerge_range:['sheet','range'],freeze_panes:['sheet','mode','range','count'],unfreeze_panes:['sheet'],sort_range:['sheet','range','keyRange','order','hasHeaders'],filter_range:['sheet','range','field','criteria1','criteria2','operator'],clear_filter:['sheet','range'],set_validation:['sheet','range','validationType','alertStyle','operator','formula1','formula2'],clear_validation:['sheet','range'],set_defined_name:['name','refersTo'],rename_defined_name:['name','newName','refersTo'],delete_defined_name:['name'],add_conditional_format:['sheet','range','rule'],delete_conditional_format:['sheet','range','rule'],create_pivot:['name','sourceSheet','sourceRange','rowField','columnField','dataField','styleName','assertions','pivotSheet','destinationRange','repairIncompletePivot'],refresh_pivot:['name','sourceSheet','sourceRange','rowField','columnField','dataField','styleName','assertions'],delete_pivot_sheet:['name','pivotSheet'],set_chart:['sheet','name','range','chartType','title','width','height','expectedSeriesCount','inRows','presentation','chartPosition','series'],rename_chart:['sheet','name','newName','chartType','title','width','height','expectedSeriesCount','presentation','chartPosition','series'],delete_chart:['sheet','name']}
 for(const key of Object.keys(op))if(key!=='intent'&&!fields[op.intent].includes(key))ctx.addIssue({code:'custom',path:[key],message:'Field not supported by this intent'})
})
const readback=z.object({sheet:str,range:str}).strict()
const input={file_id:z.string().regex(/^[1-9][0-9]*$/),operations:z.array(operation).min(1).max(100),readbacks:z.array(readback).max(10).optional()}
const schema=z.object(input).strict()
function result(payload){return {isError:payload.ok!==true,content:[{type:'text',text:JSON.stringify(payload)}]}}
function makeHandler(deps={}){
 const auth=deps.coedit||coedit,execute=deps.execute||batch.executeBatchTaskInPersistentSession
 return async args=>{
  const parsed=schema.safeParse(args)
  if(!parsed.success)return result({ok:false,outcome:'xlsx-batch-invalid-input',authority:'PLAN_ONLY',writeAllowed:false})
  const {file_id,operations,readbacks}=parsed.data,task={operations,readbacks},plan=batch.planTask(task)
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
  'Supports sheet create/copy/rename/delete, range writes, format, fixed layout, merge/unmerge, freeze, sort, filter, durable validation, workbook names, conditional formatting, tested pivot create on new or pre-created worksheet, refresh and delete, and named chart set/rename/delete with measurable presentation, position and series state. '+
  'Read/edit/read, read-only whole-task verification, then one save barrier and close. '+
  'Optional readbacks return selected final ranges from that SAME session after whole-task verification. '+
  'Submit all compatible final-state goals together; exact duplicate targets conflict, while independent ranges and named charts may share a worksheet. Not transactional. '+
  'Core sheet/write operations must precede enhanced operations; all create_sheet operations are executed before write_range so same-batch cross-sheet formulas resolve. No AutoFit. '+
  'Examples: {intent:"format_range",sheet:"Sheet1",range:"A1:B2",format:{bold:true}}; '+
  '{intent:"set_defined_name",name:"ReportRange",refersTo:"=Sheet1!$A$1:$B$2"}.',input,makeHandler(deps))
}
module.exports={register,makeHandler,schema}
