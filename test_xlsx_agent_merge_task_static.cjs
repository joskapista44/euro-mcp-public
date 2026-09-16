'use strict'
const assert=require('assert'),agent=require('./xlsx-agent-merge-task.cjs'),observer=require('./xlsx-merge-observer.cjs')
assert.equal(agent.planTask({operations:[{intent:'merge_range',sheet:'S',range:'A1:B2'}]}).ok,true)
assert.equal(agent.planTask({operations:[{intent:'unmerge_range',sheet:'S',range:'A1:B2'}]}).ok,true)
assert.equal(agent.planTask({operations:[{intent:'merge_range',sheet:'S',range:'A1'}]}).outcome,'xlsx-merge-task-invalid-range')
const clean={ok:true,cells:[[{formula:null,rawValue:'x',value:'x',displayText:'x'},{formula:null,rawValue:'',value:'',displayText:''}],[{formula:null,rawValue:null,value:null,displayText:''},{formula:null,rawValue:null,value:null,displayText:''}]]}
assert.equal(agent.dataLossRisk(clean),false)
const dirty=JSON.parse(JSON.stringify(clean));dirty.cells[1][1].rawValue=9;dirty.cells[1][1].value=9;dirty.cells[1][1].displayText='9';assert.equal(agent.dataLossRisk(dirty),true);assert.equal(agent.sameRange(clean,JSON.parse(JSON.stringify(clean))),true);assert.equal(agent.sameRange(clean,dirty),false)
let area='A1';const range={GetCells(){return {get MergeArea(){return {GetAddress(){return area}}}}},Merge(){area='A1:B2';return true},UnMerge(){area='A1';return true}};global.Api={GetSheet(){return {GetRange(){return range}}}}
let r=observer.mergeObserveCommand({intent:'merge_range',sheet:'S',range:'A1:B2'},false);assert.equal(r.state,'unmerged');assert.equal(r.noOp,false)
r=observer.mergeObserveCommand({intent:'merge_range',sheet:'S',range:'A1:B2'},true);assert.equal(r.ok,true);assert.equal(r.state,'merged-exact')
r=observer.mergeObserveCommand({intent:'merge_range',sheet:'S',range:'A1:B2'},false);assert.equal(r.noOp,true)
r=observer.mergeObserveCommand({intent:'unmerge_range',sheet:'S',range:'A1:B2'},true);assert.equal(r.ok,true);assert.equal(r.state,'unmerged')
delete global.Api
console.log('XLSX AGENT MERGE TASK STATIC: PASS')
