'use strict'
const assert=require('assert/strict')
const fs=require('fs')
const batch=require('./xlsx-persistent-batch.cjs')
const {buildShowcaseTask}=require('./xlsx-visual-showcase-contract.cjs')
const task=fs.readFileSync('./XLSX-VISUAL-SHOWCASE-AGENT-TASK.md','utf8')
for(const required of [
 'exactly one mutating MCP call','office_xlsx_batch','reply `INDULHAT`','Do not use a timed delay',
 'one `office_xlsx_batch` request','same-session final readbacks','A1:F16','A1:E9','A1:H13',
 'persistentSession.oneEditorSession=true','Whole-task verification is read-only','LIVE_READ',
 '287205 actual revenue','282000 target revenue','Do not claim completion from callback success alone'
])assert(task.includes(required),`missing showcase contract text: ${required}`)
assert(!/waitForTimeout|sleep\(|setTimeout\(/.test(task),'agent task must not prescribe a fixed delay')
const showcase=buildShowcaseTask('ABC123'),showcasePlan=batch.planTask(showcase)
assert.equal(showcasePlan.ok,true)
assert.equal(showcase.operations.length,53)
assert.equal(showcase.expected.freshWrites,52)
assert.equal(showcasePlan.steps.length,47)
assert.equal(showcasePlan.readbacks.length,3)
assert.equal(showcase.expected.totalRevenue,287205)
assert.equal(showcase.expected.targetRevenue,282000)
assert.equal(showcase.expected.variance,5205)
const showcasePivot=showcase.operations.find(op=>op.intent==='create_pivot')
assert.deepEqual({sourceRange:showcasePivot.sourceRange,dataField:showcasePivot.dataField,pivotSheet:showcasePivot.pivotSheet,destinationRange:showcasePivot.destinationRange,assertions:showcasePivot.assertions},{sourceRange:'A1:E13',dataField:'Units',pivotSheet:'MCP_PivotView_ABC123',destinationRange:'A1',assertions:[{items:['North','Core'],expected:405},{items:['North','Plus'],expected:252},{items:['South','Core'],expected:300},{items:['South','Plus'],expected:200}]})
const plan=batch.planTask({operations:[
 {intent:'format_range',sheet:'D',range:'A1:J1',format:{bold:true}},
 {intent:'format_range',sheet:'D',range:'A4:A7',format:{bold:true}},
 {intent:'layout_range',sheet:'D',range:'A1:A16',type:'column.width',width:24},
 {intent:'layout_range',sheet:'D',range:'B1:C16',type:'column.width',width:15},
 {intent:'set_chart',sheet:'D',name:'Trend',range:'A10:C16',chartType:'bar',title:'Monthly revenue vs target',expectedSeriesCount:2},
 {intent:'set_chart',sheet:'D',name:'Region',range:'E10:F12',chartType:'bar',title:'Revenue by region',expectedSeriesCount:1}
],readbacks:[{sheet:'D',range:'A1:F16'}]})
assert.equal(plan.ok,true)
assert.equal(plan.steps.length,6)
assert.deepEqual(plan.readbacks,[{sheet:'D',range:'A1:F16'}])
console.log('XLSX VISUAL SHOWCASE AGENT TASK STATIC: PASS')
