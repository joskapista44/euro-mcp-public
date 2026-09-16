'use strict'
const assert=require('assert'),v=require('./xlsx-range-copy-verifier.cjs')
const read=cells=>({ok:true,authority:'LIVE_READ',cells}),cell=(formula,row,column,rawValue=null)=>({formula,row,column,rawValue,value:rawValue,displayText:String(rawValue??'')})
const source=read([[cell(null,1,1,11),cell(null,1,2,22)],[cell(null,2,1,33),cell('=A1+B1',2,2,33)]]),correct=read([[cell(null,1,4,11),cell(null,1,5,22)],[cell(null,2,4,33),cell('=D1+E1',2,5,33)]]),doubleTranslated=read([[cell(null,1,4,11),cell(null,1,5,22)],[cell(null,2,4,33),cell('=G1+H1',2,5,33)]])
assert.equal(v.verifyRangeCopySemantic(source,correct).ok,true);assert.equal(v.verifyRangeCopySemantic(source,doubleTranslated).ok,false);assert.equal(v.translateFormula('=$A1+B$1',0,3),'=$A1+E$1');console.log('XLSX RANGE COPY RETRY FORMULA REGRESSION: PASS')
