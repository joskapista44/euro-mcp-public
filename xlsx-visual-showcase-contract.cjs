'use strict'

const COLORS={navy:[31,56,100],blue:[47,117,181],paleBlue:[221,235,247],dark:[31,41,55],white:[255,255,255],paleGreen:[226,239,218]}
const CURRENCY='"$"#,##0'
function blanks(rows,columns){return Array.from({length:rows},()=>Array(columns).fill(null))}
function buildShowcaseTask(runId){
 const id=String(runId||'').toUpperCase()
 if(!/^[A-Z0-9]{6}$/.test(id))throw new Error('runId must be six uppercase alphanumeric characters')
 const dash=`MCP_Dash_${id}`,plan=`MCP_Plan_${id}`,data=`MCP_Data_${id}`
 const definedName=`MCP_Sales_${id}`,pivotSourceName=`MCP_PivotSource_${id}`,pivotName=`MCP_Pivot_${id}`,pivotSheet=`MCP_PivotView_${id}`,trendChart=`MCP_Trend_${id}`,regionChart=`MCP_Region_${id}`,trendGeometry=`MCP_TrendSize_${id}`,regionGeometry=`MCP_RegionSize_${id}`
 const dataValues=[
  ['Month','Region','Product','Units','Unit price','Revenue','Status','Review score'],
  ['Jan','North','Core',120,190,null,null,5],['Feb','North','Plus',75,320,null,null,4],['Mar','North','Core',135,195,null,null,5],
  ['Apr','North','Plus',82,330,null,null,4],['May','North','Core',150,200,null,null,5],['Jun','North','Plus',95,340,null,null,5],
  ['Jan','South','Core',90,190,null,null,3],['Feb','South','Plus',60,320,null,null,3],['Mar','South','Core',100,195,null,null,4],
  ['Apr','South','Plus',68,330,null,null,4],['May','South','Core',110,200,null,null,4],['Jun','South','Plus',72,340,null,null,4]
 ],dataFormulas=blanks(13,8)
 for(let row=2;row<=13;row++){dataFormulas[row-1][5]=`=D${row}*E${row}`;dataFormulas[row-1][6]=`=IF(F${row}>=20000,"On track","At risk")`}
 const planValues=blanks(9,5),planFormulas=blanks(9,5)
 planValues[0][0]='Monthly sales plan';planValues[1][0]='Actual revenue is calculated from the live transaction sheet.'
 planValues[2]=['Month','Actual revenue','Target revenue','Variance','Attainment']
 const months=['Jan','Feb','Mar','Apr','May','Jun'],targets=[40000,42000,45000,48000,52000,55000]
 for(let i=0;i<6;i++){const row=i+4,north=i+2,south=i+8;planValues[row-1][0]=months[i];planValues[row-1][2]=targets[i];planFormulas[row-1][1]=`=${data}!F${north}+${data}!F${south}`;planFormulas[row-1][3]=`=B${row}-C${row}`;planFormulas[row-1][4]=`=B${row}/C${row}`}
 const dashValues=blanks(16,6),dashFormulas=blanks(16,6)
 dashValues[0][0]='Commercial performance overview';dashValues[1][0]='Six-month actuals, targets and regional mix'
 ;['Total revenue','Target revenue','Variance','Attainment'].forEach((v,i)=>dashValues[i+3][0]=v)
 dashFormulas[3][1]=`=SUM(${plan}!B4:B9)`;dashFormulas[4][1]=`=SUM(${plan}!C4:C9)`;dashFormulas[5][1]='=B4-B5';dashFormulas[6][1]='=B4/B5'
 ;['North revenue','South revenue','Best month','Best month revenue'].forEach((v,i)=>dashValues[i+3][4]=v)
 dashFormulas[3][5]=`=SUM(${data}!F2:F7)`;dashFormulas[4][5]=`=SUM(${data}!F8:F13)`;dashValues[5][5]='Jun';dashFormulas[6][5]=`=${plan}!B9`
 dashValues[9]=['Month','Actual revenue','Target revenue',null,'Region','Revenue']
 for(let i=0;i<6;i++){const row=i+11,source=i+4;dashFormulas[row-1][0]=`=${plan}!A${source}`;dashFormulas[row-1][1]=`=${plan}!B${source}`;dashFormulas[row-1][2]=`=${plan}!C${source}`}
 dashValues[10][4]='North';dashFormulas[10][5]='=F4';dashValues[11][4]='South';dashFormulas[11][5]='=F5'
 const title={bold:true,fontName:'Arial',fontSize:16,fontColor:COLORS.white,fillColor:COLORS.navy,alignHorizontal:'center',alignVertical:'center'}
 const subtitle={italic:true,fontName:'Arial',fontSize:10,fontColor:COLORS.dark}
 const header={bold:true,fontName:'Arial',fontSize:10,fontColor:COLORS.white,fillColor:COLORS.blue,alignHorizontal:'center',alignVertical:'center',wrap:true}
 const label={bold:true,fontName:'Arial',fontSize:10,fontColor:COLORS.dark,fillColor:COLORS.paleBlue}
 const valueCurrency={bold:true,fontName:'Arial',fontSize:12,fontColor:COLORS.dark,numberFormat:CURRENCY}
 const valuePercent={bold:true,fontName:'Arial',fontSize:12,fontColor:COLORS.dark,numberFormat:'0.0%'}
 const operations=[
  {intent:'create_sheet',name:dash},{intent:'create_sheet',name:plan},{intent:'create_sheet',name:data},{intent:'create_sheet',name:pivotSheet},
  {intent:'write_range',sheet:data,range:'A1:H13',values:dataValues,formulas:dataFormulas},
  {intent:'write_range',sheet:plan,range:'A1:E9',values:planValues,formulas:planFormulas},
  {intent:'write_range',sheet:dash,range:'A1:F16',values:dashValues,formulas:dashFormulas},
  {intent:'move_sheet',sheet:dash,referenceSheet:'Sheet1',position:'before'},
  {intent:'merge_range',sheet:dash,range:'A1:J1'},{intent:'merge_range',sheet:plan,range:'A1:E1'},
  {intent:'format_range',sheet:dash,range:'A1:J1',format:title},{intent:'format_range',sheet:dash,range:'A2:J2',format:subtitle},
  {intent:'format_range',sheet:dash,range:'A4:A7',format:label},{intent:'format_range',sheet:dash,range:'E4:E7',format:label},
  {intent:'format_range',sheet:dash,range:'B4:B6',format:valueCurrency},{intent:'format_range',sheet:dash,range:'B7',format:valuePercent},
  {intent:'format_range',sheet:dash,range:'F4:F5',format:valueCurrency},{intent:'format_range',sheet:dash,range:'F7',format:valueCurrency},
  {intent:'format_range',sheet:dash,range:'F6',format:{bold:true,fontName:'Arial',fontSize:12,fontColor:COLORS.dark}},
  {intent:'format_range',sheet:dash,range:'A10:C10',format:header},{intent:'format_range',sheet:dash,range:'E10:F10',format:header},
  {intent:'format_range',sheet:plan,range:'A1:E1',format:title},{intent:'format_range',sheet:plan,range:'A2:E2',format:subtitle},
  {intent:'format_range',sheet:plan,range:'A3:E3',format:header},{intent:'format_range',sheet:plan,range:'B4:D9',format:{fontName:'Arial',fontSize:10,numberFormat:CURRENCY}},
  {intent:'format_range',sheet:plan,range:'E4:E9',format:{fontName:'Arial',fontSize:10,numberFormat:'0.0%'}},
  {intent:'format_range',sheet:data,range:'A1:H1',format:header},{intent:'format_range',sheet:data,range:'E2:F13',format:{fontName:'Arial',fontSize:10,numberFormat:CURRENCY}},
  {intent:'format_range',sheet:data,range:'D2:D13',format:{fontName:'Arial',fontSize:10,numberFormat:'#,##0'}},{intent:'format_range',sheet:data,range:'G2:H13',format:{fontName:'Arial',fontSize:10,alignHorizontal:'center'}},
  {intent:'layout_range',sheet:dash,range:'A1:A16',type:'column.width',width:24},{intent:'layout_range',sheet:dash,range:'B1:C16',type:'column.width',width:15},
  {intent:'layout_range',sheet:dash,range:'E1:E16',type:'column.width',width:24},{intent:'layout_range',sheet:dash,range:'F1:F16',type:'column.width',width:15},{intent:'layout_range',sheet:dash,range:'A1:J1',type:'row.height',height:28},
  {intent:'layout_range',sheet:plan,range:'A1:A9',type:'column.width',width:14},{intent:'layout_range',sheet:plan,range:'B1:E9',type:'column.width',width:16},{intent:'layout_range',sheet:plan,range:'A1:E1',type:'row.height',height:28},
  {intent:'layout_range',sheet:data,range:'A1:A13',type:'column.width',width:12},{intent:'layout_range',sheet:data,range:'B1:C13',type:'column.width',width:16},{intent:'layout_range',sheet:data,range:'D1:H13',type:'column.width',width:14},{intent:'layout_range',sheet:data,range:'A1:H1',type:'row.height',height:22},
  {intent:'sort_range',sheet:data,range:'A1:H13',keyRange:'B1:B13',order:'asc',hasHeaders:true},
  {intent:'filter_range',sheet:data,range:'A1:H13',field:7,criteria1:'On track',operator:'xlOr'},
  {intent:'set_validation',sheet:data,range:'H2:H13',validationType:'xlValidateWholeNumber',alertStyle:'xlValidAlertStop',operator:'xlBetween',formula1:1,formula2:5},
  {intent:'set_defined_name',name:definedName,refersTo:`=${data}!$A$1:$H$13`},
  {intent:'set_defined_name',name:pivotSourceName,refersTo:`=${data}!$A$1:$E$13`},
  {intent:'add_conditional_format',sheet:data,range:'F2:F13',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'25000',fillColor:COLORS.paleGreen,priority:1}},
  {intent:'add_conditional_format',sheet:plan,range:'D4:D9',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'0',fillColor:COLORS.paleGreen,priority:1}},
  {intent:'create_pivot',name:pivotName,sourceSheet:data,sourceRange:'A1:E13',sourceIdentityName:pivotSourceName,rowField:'Region',columnField:'Product',dataField:'Units',styleName:'PivotStyleMedium2',pivotSheet:pivotSheet,destinationRange:'A1',repairIncompletePivot:true,assertions:[{items:['North','Core'],expected:405},{items:['North','Plus'],expected:252},{items:['South','Core'],expected:300},{items:['South','Plus'],expected:200}]},
  {intent:'set_chart',sheet:dash,name:trendChart,range:'A10:C16',chartType:'bar',title:'Monthly revenue vs target',width:3800000,height:2300000,expectedSeriesCount:2,geometryIdentityName:trendGeometry,presentation:{legendPosition:'bottom',horizontalAxisTitle:'Month',verticalAxisTitle:'Revenue',dataLabels:{showSeriesName:false,showCategoryName:false,showValue:true,showPercent:false},style:2},chartPosition:{fromCol:6,colOffset:0,fromRow:2,rowOffset:0},series:[{index:0,name:'Actual revenue',valuesRange:`${dash}!$B$11:$B$16`,categoryRange:`${dash}!$A$11:$A$16`},{index:1,name:'Target revenue',valuesRange:`${dash}!$C$11:$C$16`,categoryRange:`${dash}!$A$11:$A$16`}]},
  {intent:'set_chart',sheet:dash,name:regionChart,range:'E10:F12',chartType:'bar',title:'Revenue by region',width:3800000,height:2300000,expectedSeriesCount:1,geometryIdentityName:regionGeometry,presentation:{legendPosition:'bottom',horizontalAxisTitle:'Region',verticalAxisTitle:'Revenue',dataLabels:{showSeriesName:false,showCategoryName:false,showValue:true,showPercent:false},style:2},chartPosition:{fromCol:6,colOffset:0,fromRow:16,rowOffset:0},series:[{index:0,name:'Revenue',valuesRange:`${dash}!$F$11:$F$12`,categoryRange:`${dash}!$E$11:$E$12`}]},
  {intent:'freeze_panes',sheet:plan,mode:'at',range:'A4'},{intent:'freeze_panes',sheet:data,mode:'at',range:'A2'}
 ]
 return {file_id:null,operations,readbacks:[{sheet:dash,range:'A1:F16'},{sheet:plan,range:'A1:E9'},{sheet:data,range:'A1:H13'}],names:{dash,plan,data,pivotSheet,definedName,pivotSourceName,pivotName,trendChart,regionChart},expected:{totalRevenue:287205,targetRevenue:282000,variance:5205,attainment:287205/282000,operationCount:operations.length,freshWrites:operations.length-1}}
}
module.exports={COLORS,CURRENCY,buildShowcaseTask}
