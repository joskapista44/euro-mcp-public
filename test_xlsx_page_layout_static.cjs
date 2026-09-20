'use strict'
const assert=require('assert/strict'),page=require('./xlsx-persistent-page-layout.cjs')
assert.equal(page.validSpec({sheet:'S',orientation:'xlLandscape',topMargin:10,bottomMargin:11,leftMargin:12,rightMargin:13,printGridlines:true,printHeadings:false}),true)
assert.equal(page.validSpec({sheet:'S',orientation:'sideways'}),false)
assert.equal(page.validSpec({sheet:'S',topMargin:-1}),false)
const before={orientation:'xlPortrait',topMargin:19.1,bottomMargin:19.1,leftMargin:17.8,rightMargin:17.8,printGridlines:false,printHeadings:false}
const want=page.expected({orientation:'xlLandscape',leftMargin:10,printGridlines:true},before)
assert.deepEqual(want,{orientation:'xlLandscape',topMargin:19.1,bottomMargin:19.1,leftMargin:10,rightMargin:17.8,printGridlines:true,printHeadings:false})
assert.equal(page.same(want,{...want,leftMargin:10.005}),true)
assert.equal(page.same(want,{...want,leftMargin:10.02}),false)
console.log('XLSX PAGE LAYOUT STATIC: PASS')
