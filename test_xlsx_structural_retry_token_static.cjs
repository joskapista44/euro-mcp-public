'use strict'
const assert=require('assert'),retry=require('./xlsx-structural-retry-token.cjs'),{planTask}=require('./xlsx-agent-structural-task.cjs')
const base={type:'rows.insert',intent:'insert_rows',sheet:'S',range:'A2',count:1,anchor:2,axis:'rows'},before='1'.repeat(64),post='2'.repeat(64),window={before:'A2:B2',after:'A2:B3',dispatch:'A2:B2'},token=retry.make(base,before,post,window)
assert.ok(token);assert.equal(token.version,2);assert.equal(retry.validate(token,base).ok,true);assert.equal(retry.validate(token,{...base,range:'A3',anchor:3}).outcome,'xlsx-structural-retry-token-operation-mismatch')
let r=planTask({operations:[{intent:'insert_rows',sheet:'S',range:'A2',retryToken:token}]});assert.equal(r.ok,true);assert.equal(r.operation.preconditionFingerprint,before);assert.equal(r.operation.expectedPostFingerprint,post);assert.equal(r.operation.retryBeforeRange,window.before);assert.equal(r.operation.retryPostRange,window.after);assert.equal(r.operation.retryDispatchRange,window.dispatch)
r=planTask({operations:[{intent:'insert_rows',sheet:'S',range:'A3',retryToken:token}]});assert.equal(r.ok,false);assert.equal(r.outcome,'xlsx-structural-retry-token-operation-mismatch')
r=planTask({operations:[{intent:'insert_rows',sheet:'S',range:'A2',preconditionFingerprint:'3'.repeat(64),retryToken:token}]});assert.equal(r.ok,false);assert.equal(r.outcome,'xlsx-structural-retry-token-precondition-conflict')
console.log('XLSX STRUCTURAL RETRY TOKEN STATIC: PASS')
