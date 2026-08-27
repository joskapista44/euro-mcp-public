'use strict'

const assert=require('assert')
const fs=require('fs')
const f=require('./live-formatting.cjs')

function color(r,g,b){ return { GetRGB(){ return ((r<<16)|(g<<8)|b)>>>0 } } }

function makeApi(){
  const calls=[]
  const state={
    Bold:false, Italic:false, FontName:'Calibri', FontSize:11, FontColor:color(0,0,0), FillColor:'No Fill',
    AlignHorizontal:'left', AlignVertical:'bottom', WrapText:false, NumberFormat:'General'
  }
  const range={
    get Bold(){return state.Bold}, get Italic(){return state.Italic}, get FontName(){return state.FontName}, get FontSize(){return state.FontSize},
    get FontColor(){return state.FontColor}, get FillColor(){return state.FillColor}, get AlignHorizontal(){return state.AlignHorizontal}, get AlignVertical(){return state.AlignVertical},
    get WrapText(){return state.WrapText}, get NumberFormat(){return state.NumberFormat},
    SetBold(v){state.Bold=v;calls.push(['bold',v]);return true}, SetItalic(v){state.Italic=v;calls.push(['italic',v]);return true}, SetFontName(v){state.FontName=v;calls.push(['fontName',v]);return true}, SetFontSize(v){state.FontSize=v;calls.push(['fontSize',v]);return true},
    SetFontColor(v){state.FontColor=v;calls.push(['fontColor',v]);return true}, SetFillColor(v){state.FillColor=v;calls.push(['fillColor',v]);return true}, GetFillColor(){return state.FillColor}, SetBorders(i,s,c){calls.push(['border',i,s,c]);return true},
    SetAlignHorizontal(v){state.AlignHorizontal=v;calls.push(['hAlign',v]);return true}, SetAlignVertical(v){state.AlignVertical=v;calls.push(['vAlign',v]);return true}, SetWrap(v){state.WrapText=v;calls.push(['wrap',v]);return true}, GetWrapText(){return state.WrapText},
    SetNumberFormat(v){state.NumberFormat=v;calls.push(['numberFormat',v]);return true}, GetNumberFormat(){return state.NumberFormat}, SetColumnWidth(v){calls.push(['columnWidth',v])}, SetRowHeight(v){calls.push(['rowHeight',v]);return true},
    AutoFit(r,c){calls.push(['autofit',r,c])}, Merge(a){calls.push(['merge',a]);return true}, UnMerge(){calls.push(['unmerge']);return true}, GetAddress(){return '$A$1:$D$4'}
  }
  const fp={ FreezeRows(n){calls.push(['freezeRows',n])}, FreezeColumns(n){calls.push(['freezeColumns',n])}, FreezeAt(r){calls.push(['freezeAt',r])}, Unfreeze(){calls.push(['unfreeze'])}, GetLocation(){return range} }
  const sheet={ GetRange(){return range}, GetFreezePanes(){return fp} }
  return { api:{GetSheet(){return sheet},CreateColorFromRGB(r,g,b){return color(r,g,b)}},calls,range,fp,sheet,state }
}

// Validation is strict and fail-closed.
assert.equal(f.validateFormatSpec({bold:true}).ok,true)
assert.equal(f.validateFormatSpec({fontColor:[255,10,0]}).ok,true)
assert.equal(f.validateFormatSpec({fontColor:[256,0,0]}).ok,false)
assert.equal(f.validateFormatSpec({mystery:true}).ok,false)
assert.equal(f.validateFormatSpec({}).ok,false)

// M4.1 acceptance: write + same-live-range readback verifies all reliably readable fields.
{
  const m=makeApi(); global.Api=m.api
  const spec={bold:true,italic:true,fontName:'Arial',fontSize:12,fontColor:[255,255,255],fillColor:[31,78,121],alignHorizontal:'center',alignVertical:'center',wrap:true,numberFormat:'#,##0.00'}
  const r=f.formattingCommand('Report','A1:D4',spec)
  assert.equal(r.ok,true); assert.equal(r.source,'live-coedit-editor'); assert.equal(r.verification.outcome,'pass')
  for(const key of Object.keys(spec)){ assert(r.applied.includes(key),key); assert.equal(r.verification.checks[key].status,'pass',key) }
}

// Border is applied live but fail-closed UNKNOWN because ApiRange exposes no border getter.
{
  const m=makeApi(); global.Api=m.api
  const r=f.formattingCommand('Report','A1:D4',{border:{index:'Bottom',style:'Medium',color:[255,255,255]}})
  assert.equal(r.ok,true); assert.equal(r.verification.outcome,'unknown'); assert.equal(r.verification.checks.border.status,'unknown')
}

// Verification mismatch is an explicit FAIL, not an assumed success.
{
  const m=makeApi(); global.Api=m.api
  m.range.SetBold=function(v){m.calls.push(['bold',v]);return true} // intentionally inert setter
  const r=f.formattingCommand('Report','A1',{bold:true})
  assert.equal(r.ok,false); assert.equal(r.outcome,'verification-mismatch'); assert.equal(r.verification.outcome,'fail'); assert.deepEqual(r.verification.mismatches,['bold'])
}

// Missing live getter/property is UNKNOWN rather than guessed.
{
  const m=makeApi(); global.Api=m.api
  Object.defineProperty(m.range,'FontName',{get(){return undefined}})
  const r=f.formattingCommand('Report','A1',{fontName:'Arial'})
  assert.equal(r.ok,true); assert.equal(r.verification.outcome,'unknown'); assert.deepEqual(r.verification.unknown,['fontName'])
}

// Currency/percent/date are explicit number-format masks and are live-verified.
for(const mask of ['€ #,##0.00','0.00%','yyyy-mm-dd']){
  const m=makeApi(); global.Api=m.api
  const r=f.formattingCommand('Report','B2:B5',{numberFormat:mask}); assert.equal(r.ok,true); assert.equal(r.verification.outcome,'pass'); assert(m.calls.some((x)=>x[0]==='numberFormat'&&x[1]===mask))
}

// Non-first worksheet/range addressing is passed through to the live sheet API.
{
  const m=makeApi(); let requested=null
  m.api.GetSheet=(name)=>{requested=name;return m.sheet}; global.Api=m.api
  const r=f.formattingCommand('Second Sheet','C3:F9',{bold:true,wrap:true}); assert.equal(r.ok,true); assert.equal(requested,'Second Sheet'); assert.equal(r.range,'C3:F9'); assert.equal(r.verification.outcome,'pass')
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

// Architectural regression guard: formatting stays live co-edit only.
for(const file of ['live-formatting.cjs','euro-mcp-m2.cjs']){
  const src=fs.readFileSync(file,'utf8')
  assert.doesNotMatch(src,/DocBuilder\s*\(/)
  assert.doesNotMatch(src,/OOXML.*(write|edit|save)/i)
  assert.doesNotMatch(src,/xlsx.*(writeFile|save)/i)
  assert.doesNotMatch(src,/WebDAV|PROPFIND|fetch\(/i)
  assert.match(src,/live-coedit-editor|CURRENT live ONLYOFFICE/)
}
const source=fs.readFileSync('live-formatting.cjs','utf8')
assert.match(source,/verificationOutcome/)
assert.match(source,/GetFillColor/)
assert.match(source,/GetWrapText/)
assert.match(source,/GetNumberFormat/)
const entry=fs.readFileSync('euro-mcp-m2.cjs','utf8')
for(const tool of ['office_format_range','office_resize_range','office_merge_range','office_freeze_panes'])assert.match(entry,new RegExp(tool))
assert.match(entry,/require\('\.\/euro-mcp-m15\.cjs'\)/)

console.log('test_live_formatting: OK')
