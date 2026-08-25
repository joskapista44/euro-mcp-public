'use strict'

const assert=require('assert')
const fs=require('fs')
const f=require('./live-formatting.cjs')

function makeApi(){
  const calls=[]
  const range={
    SetBold(v){calls.push(['bold',v]);return true}, SetItalic(v){calls.push(['italic',v]);return true}, SetFontName(v){calls.push(['fontName',v]);return true}, SetFontSize(v){calls.push(['fontSize',v]);return true},
    SetFontColor(v){calls.push(['fontColor',v]);return true}, SetFillColor(v){calls.push(['fillColor',v]);return true}, SetBorders(i,s,c){calls.push(['border',i,s,c]);return true},
    SetAlignHorizontal(v){calls.push(['hAlign',v]);return true}, SetAlignVertical(v){calls.push(['vAlign',v]);return true}, SetWrap(v){calls.push(['wrap',v]);return true}, GetWrapText(){return true},
    SetNumberFormat(v){calls.push(['numberFormat',v]);return true}, GetNumberFormat(){return '#,##0.00'}, SetColumnWidth(v){calls.push(['columnWidth',v])}, SetRowHeight(v){calls.push(['rowHeight',v]);return true},
    AutoFit(r,c){calls.push(['autofit',r,c])}, Merge(a){calls.push(['merge',a]);return true}, UnMerge(){calls.push(['unmerge']);return true}, GetAddress(){return '$A$1:$D$4'}
  }
  const fp={ FreezeRows(n){calls.push(['freezeRows',n])}, FreezeColumns(n){calls.push(['freezeColumns',n])}, FreezeAt(r){calls.push(['freezeAt',r])}, Unfreeze(){calls.push(['unfreeze'])}, GetLocation(){return range} }
  const sheet={ GetRange(){return range}, GetFreezePanes(){return fp} }
  return { api:{GetSheet(){return sheet},CreateColorFromRGB(r,g,b){return {r,g,b}}},calls,range,fp,sheet }
}

// Validation is strict and fail-closed.
assert.equal(f.validateFormatSpec({bold:true}).ok,true)
assert.equal(f.validateFormatSpec({fontColor:[255,10,0]}).ok,true)
assert.equal(f.validateFormatSpec({fontColor:[256,0,0]}).ok,false)
assert.equal(f.validateFormatSpec({mystery:true}).ok,false)
assert.equal(f.validateFormatSpec({}).ok,false)

// Professional worksheet formatting acceptance: header + numeric presentation primitives.
{
  const m=makeApi(); global.Api=m.api
  const spec={bold:true,italic:false,fontName:'Arial',fontSize:12,fontColor:[255,255,255],fillColor:[31,78,121],border:{index:'Bottom',style:'Medium',color:[255,255,255]},alignHorizontal:'center',alignVertical:'center',wrap:true,numberFormat:'#,##0.00'}
  const r=f.formattingCommand('Report','A1:D4',spec)
  assert.equal(r.ok,true); assert.equal(r.source,'live-coedit-editor')
  for(const key of ['bold','italic','fontName','fontSize','fontColor','fillColor','border','alignHorizontal','alignVertical','wrap','numberFormat']) assert(r.applied.includes(key),key)
  assert(m.calls.some((x)=>x[0]==='numberFormat'&&x[1]==='#,##0.00'))
}

// Currency/percent/date are explicit number-format masks, not value rewrites.
for(const mask of ['€ #,##0.00','0.00%','yyyy-mm-dd']){
  const m=makeApi(); global.Api=m.api
  const r=f.formattingCommand('Report','B2:B5',{numberFormat:mask}); assert.equal(r.ok,true); assert(m.calls.some((x)=>x[0]==='numberFormat'&&x[1]===mask))
}

// Layout primitives.
{
  const m=makeApi(); global.Api=m.api
  const r=f.layoutCommand('Report','A1:D20',{columnWidth:18,rowHeight:24,autofitRows:true,autofitColumns:true}); assert.equal(r.ok,true)
  assert(m.calls.some((x)=>x[0]==='columnWidth')); assert(m.calls.some((x)=>x[0]==='rowHeight')); assert(m.calls.some((x)=>x[0]==='autofit'&&x[1]===true&&x[2]===true))
}

// Merge / unmerge.
{
  let m=makeApi(); global.Api=m.api; assert.equal(f.structureCommand('Report','A1:D1','merge',false).ok,true); assert(m.calls.some((x)=>x[0]==='merge'))
  m=makeApi(); global.Api=m.api; assert.equal(f.structureCommand('Report','A1:D1','unmerge',false).ok,true); assert(m.calls.some((x)=>x[0]==='unmerge'))
}

// Freeze rows, columns, arbitrary location, and unfreeze.
for(const [action,value,call] of [['rows',1,'freezeRows'],['columns',2,'freezeColumns'],['at','C3','freezeAt'],['unfreeze',undefined,'unfreeze']]){
  const m=makeApi(); global.Api=m.api; const r=f.freezeCommand('Report',action,value); assert.equal(r.ok,true); assert.equal(r.location,'$A$1:$D$4'); assert(m.calls.some((x)=>x[0]===call))
}

// Unsupported Office API methods fail closed rather than pretending success.
{
  const m=makeApi(); delete m.range.SetBorders; global.Api=m.api
  const r=f.formattingCommand('Report','A1',{border:{index:'Bottom',style:'Thin',color:[0,0,0]}}); assert.equal(r.ok,false); assert.equal(r.outcome,'unsupported')
}
{
  const m=makeApi(); delete m.fp.FreezeAt; global.Api=m.api
  const r=f.freezeCommand('Report','at','B2'); assert.equal(r.ok,false); assert.equal(r.outcome,'unsupported')
}

// Architectural regression guard: M2 stays live co-edit only.
for(const file of ['live-formatting.cjs','euro-mcp-m2.cjs']){
  const src=fs.readFileSync(file,'utf8')
  assert.doesNotMatch(src,/DocBuilder\s*\(/)
  assert.doesNotMatch(src,/OOXML.*(write|edit|save)/i)
  assert.doesNotMatch(src,/xlsx.*(writeFile|save)/i)
  assert.match(src,/live-coedit-editor|CURRENT live ONLYOFFICE/)
}
const entry=fs.readFileSync('euro-mcp-m2.cjs','utf8')
for(const tool of ['office_format_range','office_resize_range','office_merge_range','office_freeze_panes'])assert.match(entry,new RegExp(tool))
assert.match(entry,/require\('\.\/euro-mcp-m15\.cjs'\)/)

console.log('test_live_formatting: OK')
