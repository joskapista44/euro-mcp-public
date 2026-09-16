'use strict'
const assert=require('assert'),{observation}=require('./xlsx-structural-observation.cjs')
const inv={sheets:[{name:'S',usedRange:'$A$1:$D$6'}]}
let r=observation(inv,{sheet:'S',axis:'rows',type:'rows.insert',anchor:2,count:2});assert.equal(r.ok,true);assert.equal(r.before,'A2:D6');assert.equal(r.after,'A2:D8');assert.equal(r.dispatch,'A2:D3')
r=observation(inv,{sheet:'S',axis:'rows',type:'rows.delete',anchor:2,count:2});assert.equal(r.before,'A2:D8');assert.equal(r.after,'A2:D6');assert.equal(r.dispatch,'A2:D3')
r=observation(inv,{sheet:'S',axis:'columns',type:'columns.insert',anchor:3,count:2});assert.equal(r.before,'C1:D6');assert.equal(r.after,'C1:F6');assert.equal(r.dispatch,'C1:D6')
r=observation(inv,{sheet:'S',axis:'columns',type:'columns.delete',anchor:3,count:2});assert.equal(r.before,'C1:F6');assert.equal(r.after,'C1:D6');assert.equal(r.dispatch,'C1:D6')
r=observation({sheets:[]},{sheet:'S',axis:'rows',type:'rows.insert',anchor:2,count:1});assert.equal(r.ok,false);assert.equal(r.authority,'PLAN_ONLY')
console.log('XLSX STRUCTURAL OBSERVATION STATIC: PASS')
