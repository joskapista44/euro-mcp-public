'use strict'
const assert=require('assert'),fs=require('fs')
const s=fs.readFileSync(require.resolve('./xlsx-persistent-format.cjs'),'utf8')
assert.match(s,/withPersistentXlsxSession/)
assert.match(s,/formatRangeObserved/)
assert.match(s,/session\.markWrite\(\)/)
assert.match(s,/formatObserveCommand/)
assert.doesNotMatch(s,/waitForTimeout/)
console.log('XLSX PERSISTENT FORMAT STATIC: PASS')
