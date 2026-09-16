'use strict'
const assert=require('assert'),p=require('./xlsx-persistent-session.cjs')
assert.equal(typeof p.executeStructuralTaskInPersistentSession,'function')
;(async()=>{const r=await p.executeStructuralTaskInPersistentSession({fileId:1,task:{operations:[{intent:'insert_rows',sheet:'S',range:'A2'}]}});assert.equal(r.ok,false);assert.equal(r.outcome,'xlsx-persistent-credentials-required');assert.equal(r.authority,'PLAN_ONLY');console.log('XLSX PERSISTENT STRUCTURAL API STATIC: PASS')})().catch(e=>{console.error(e);process.exitCode=1})
