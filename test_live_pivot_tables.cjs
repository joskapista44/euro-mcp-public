'use strict';
const assert = require('assert');
const { pivotCommand } = require('./live-pivot-tables.cjs');

function pivot(name, source) {
  const state = { name, source, rows: [], cols: [], data: [], style: 'PivotStyleLight16' };
  return {
    state,
    GetName(){return state.name}, SetName(v){state.name=v}, GetSource(){return state.source},
    GetRowFields(){return state.rows}, GetColumnFields(){return state.cols}, GetDataFields(){return state.data},
    GetStyleName(){return state.style}, SetStyleName(v){state.style=v}, GetTitle(){return null}, GetDescription(){return null},
    AddFields(r,c){state.rows.push(...r);state.cols.push(...c);return true}, AddDataField(f){state.data.push(f);return true},
    RefreshTable(){return true}
  };
}
function env(){
  const all=[]; const ranges={};
  global.Api={
    GetAllPivotTables(){return all}, GetPivotByName(n){const p=all.find(x=>x.GetName()===n);if(!p)throw new Error('missing');return p},
    GetActiveSheet(){return {GetRange(a){ranges[a]=ranges[a]||{address:a};return ranges[a]}}},
    InsertPivotNewWorksheet(src){const p=pivot('PivotTable'+(all.length+1),src.address);all.push(p);return p}
  };
  return {all};
}
function run(spec){return pivotCommand(spec)}
function ok(label,fn){try{fn();console.log('OK',label)}catch(e){console.error('FAIL',label,e);process.exitCode=1}}

ok('inspect absent',()=>{env();const r=run({operation:'pivot.inspect',name:'X'});assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual,null)});
ok('create exact readback',()=>{env();const r=run({operation:'pivot.createNewWorksheet',source:'A1:C5',name:'EURO_P'});assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.count,1);assert.equal(r.verification.actual.pivot.name,'EURO_P');assert.equal(r.verification.actual.pivot.source,'A1:C5')});
ok('fields exact counts',()=>{env();run({operation:'pivot.createNewWorksheet',source:'A1:C5',name:'EURO_P'});const r=run({operation:'pivot.addFields',name:'EURO_P',rowFields:['Region'],columnFields:['Style']});assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.rowFields,1);assert.equal(r.verification.actual.columnFields,1)});
ok('data field exact count',()=>{env();run({operation:'pivot.createNewWorksheet',source:'A1:C5',name:'EURO_P'});const r=run({operation:'pivot.addDataField',name:'EURO_P',field:'Price'});assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.dataFields,1)});
ok('rename exact readback',()=>{env();run({operation:'pivot.createNewWorksheet',source:'A1:C5',name:'EURO_P'});const r=run({operation:'pivot.rename',name:'EURO_P',newName:'EURO_Q'});assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.name,'EURO_Q')});
ok('style exact readback',()=>{env();run({operation:'pivot.createNewWorksheet',source:'A1:C5',name:'EURO_P'});const r=run({operation:'pivot.style',name:'EURO_P',styleName:'PivotStyleMedium2'});assert.equal(r.verification.status,'PASS');assert.equal(r.verification.actual.styleName,'PivotStyleMedium2')});
ok('refresh public completion',()=>{env();run({operation:'pivot.createNewWorksheet',source:'A1:C5',name:'EURO_P'});const r=run({operation:'pivot.refresh',name:'EURO_P'});assert.equal(r.verification.status,'PASS')});
ok('unknown fail closed',()=>{env();const r=run({operation:'pivot.nope',name:'X'});assert.notEqual(r.verification.status,'PASS')});

if(!process.exitCode) console.log('MIND OK');
