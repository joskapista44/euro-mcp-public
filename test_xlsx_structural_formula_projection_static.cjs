'use strict'
const assert=require('assert/strict')
const fs=require('fs'),src=fs.readFileSync(require.resolve('./xlsx-persistent-batch.cjs'),'utf8')
assert.match(src,/shiftFormulaForStructuralInsert/)
assert.match(src,/n>=at\?col\+dollar\+\(n\+count\)/)
assert.match(src,/n>=at\?dollar\+toCol\(n\+count\)\+row/)
console.log('XLSX STRUCTURAL FORMULA PROJECTION STATIC: PASS')
