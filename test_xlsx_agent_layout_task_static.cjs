'use strict'
const assert=require('assert');const a=require('./xlsx-agent-layout-task.cjs');
for(const op of [{intent:'layout_range',sheet:'Sheet1',range:'A:A',type:'column.width',width:18},{intent:'layout_range',sheet:'Sheet1',range:'2:2',type:'row.height',height:24},{intent:'layout_range',sheet:'Sheet1',range:'B:B',type:'columns.hidden',hidden:true},{intent:'layout_range',sheet:'Sheet1',range:'3:3',type:'rows.hidden',hidden:false}])assert.equal(a.planTask({operations:[op]}).ok,true)
assert.equal(a.planTask({operations:[{intent:'layout_range',sheet:'Sheet1',range:'A1',type:'autofit.columns'}]}).ok,false)
console.log('XLSX AGENT LAYOUT TASK STATIC: PASS')
