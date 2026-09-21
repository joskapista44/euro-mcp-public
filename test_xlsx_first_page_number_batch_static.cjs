'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs')
const p=batch.planTask({operations:[{intent:'set_first_page_number',sheet:'S',value:7}]});assert.equal(p.ok,true);assert.equal(p.steps[0].family,'first-page-number')
const c=batch.planTask({operations:[{intent:'set_first_page_number',sheet:'S',value:7},{intent:'set_first_page_number',sheet:'S',value:8}]});assert.equal(c.ok,false);assert.equal(c.outcome,'xlsx-batch-conflicting-goals')
console.log('XLSX FIRST PAGE NUMBER BATCH STATIC: PASS')
