'use strict'

function blanks(rows,columns){return Array.from({length:rows},()=>Array(columns).fill(null))}

function buildCrossFamilyTask(runId){
 const id=String(runId||'').toUpperCase()
 if(!/^[A-Z0-9]{7}$/.test(id))throw new Error('runId must be seven uppercase alphanumeric characters')
 const input=`CF Input ${id}`,plan=`CF Plan ${id}`,staging=`CF Notes ${id}`,report=`CF Report ${id}`
 const inputValues=[
  ['Account','Region','Actual','Target','Variance','Status'],
  ['Atlas','North',82000,80000,null,'Review'],
  ['Beacon','South',61000,65000,null,'Review'],
  ['Cobalt','North',97000,90000,null,'Review'],
  ['Delta','West',74000,70000,null,'Review']
 ],inputFormulas=blanks(5,6)
 for(let row=2;row<=5;row++)inputFormulas[row-1][4]=`=C${row}-D${row}`
 const planValues=[['Quarter plan','Owner','Due'],['North follow-up','Anna','2026-10-05'],['South recovery','Bela','2026-10-08'],['West renewal','Csilla','2026-10-12']]
 const notes=[['Control','Result'],['Source','LIVE input'],['Method','Persistent batch']]
 const operations=[
  {intent:'create_sheet',name:input},{intent:'create_sheet',name:plan},{intent:'create_sheet',name:staging},{intent:'create_sheet',name:report},
  {intent:'write_range',sheet:input,range:'A1:F5',values:inputValues,formulas:inputFormulas},
  {intent:'write_range',sheet:plan,range:'A1:C4',values:planValues},
  {intent:'write_range',sheet:staging,range:'A1:B3',values:notes},
  {intent:'insert_rows',sheet:plan,range:'A2:C2'},
  {intent:'sort_range',sheet:input,range:'A1:F5',keyRange:'C1:C5',order:'desc',hasHeaders:true},
  {intent:'filter_range',sheet:input,range:'A1:F5',field:6,criteria1:'Review',operator:'xlOr'},
  {intent:'copy_range',sheet:input,range:'A1:F5',targetSheet:report,targetRange:'A1:F5'},
  {intent:'move_range',sheet:staging,range:'A1:B3',targetSheet:report,targetRange:'G2:H4'},
  {intent:'format_range',sheet:report,range:'A1:H1',format:{bold:true,fontName:'Arial',fontColor:[255,255,255],fillColor:[31,56,100],alignHorizontal:'center'}},
  {intent:'layout_range',sheet:report,range:'A1:H5',type:'columns.autofit'},
  {intent:'set_page_layout',sheet:report,orientation:'xlLandscape',topMargin:10,bottomMargin:10,leftMargin:8,rightMargin:8,printGridlines:false,printHeadings:false},
  {intent:'set_print_setup',sheet:report,mode:'fit_to_pages',fitToWidth:1,fitToHeight:1},
  {intent:'set_print_area',sheet:report,mode:'set',range:'A1:H5'},
  {intent:'set_header_footer',sheet:report,slot:'oddHeader',value:'&LCROSS-FAMILY ACCEPTANCE&C&P / &N'}
 ]
 return {operations,readbacks:[{sheet:input,range:'A1:F5'},{sheet:plan,range:'A1:C5'},{sheet:staging,range:'A1:B3'},{sheet:report,range:'A1:H5'}],names:{input,plan,staging,report},expected:{operationCount:operations.length,topAccount:'Cobalt',topActual:97000,movedControl:'Control',insertedPlanRow:'A2:C2'}}
}

function withRetryReceipts(task,result){
 const tokenByIntent=new Map()
 for(const step of result?.steps||[])if(step?.result?.retryToken)tokenByIntent.set(step.intent,step.result.retryToken)
 return {...task,operations:task.operations.map(op=>tokenByIntent.has(op.intent)?{...op,retryToken:tokenByIntent.get(op.intent)}:op)}
}

module.exports={buildCrossFamilyTask,withRetryReceipts}
