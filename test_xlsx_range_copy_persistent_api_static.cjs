'use strict'
const assert=require('assert'),p=require('./xlsx-persistent-session.cjs')
for(const name of ['copyRangeVerifiedInSession','executeRangeCopyTaskInPersistentSession'])assert.equal(typeof p[name],'function',name+' must be exported')
console.log('XLSX RANGE COPY PERSISTENT API STATIC: PASS')
