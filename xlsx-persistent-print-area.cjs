'use strict'
const LIVE='live-coedit-editor'
function validSpec(s){return !!s&&typeof s.sheet==='string'&&s.sheet.trim()&&['set','clear','add'].includes(s.mode)&&((s.mode==='clear'&&!s.range)||(s.mode!=='clear'&&typeof s.range==='string'&&/^[A-Z]+[1-9][0-9]*:[A-Z]+[1-9][0-9]*$/.test(s.range)))}
function normalizeRef(v){return typeof v==='string'?v.replace(/\$/g,'').replace(/^=/,''):v}
async function immediate(session,spec,apply){
 if(!validSpec(spec))return {ok:false,outcome:'print-area-invalid-spec',source:LIVE}
 return session.frame.evaluate(({sheet,mode,range,apply})=>{try{
  const e=window.Asc&&window.Asc.editor;if(!e||!e.wb||!e.wbModel)return {ok:false,outcome:'print-area-editor-api-unavailable',source:'live-coedit-editor'}
  const count=e.asc_getWorksheetsCount?e.asc_getWorksheetsCount():0;let idx=-1;for(let i=0;i<count;i++)if(e.asc_getWorksheetName(i)===sheet){idx=i;break}if(idx<0)return {ok:false,outcome:'print-area-sheet-unavailable',source:'live-coedit-editor'}
  const model=e.wbModel.getWorksheet(idx),view=e.wb.getWorksheet(idx);if(!model||!view||typeof view.changePrintArea!=='function')return {ok:false,outcome:'print-area-view-unavailable',source:'live-coedit-editor'}
  function read(){const d=model.workbook.getDefinesNames('Print_Area',model.getId(),true);return d&&d.sheetId===model.getId()?d.ref:null}
  function esc(n){return "'"+String(n).replace(/'/g,"''")+"'"}
  const target=mode==='clear'?null:esc(sheet)+'!'+range, before=read(),norm=v=>typeof v==='string'?v.replace(/\$/g,'').replace(/^=/,''):v
  const satisfied=mode==='clear'?!before:(mode==='set'&&norm(before)===norm(target))
  if(satisfied)return {ok:true,outcome:'print-area-already-satisfied',source:'live-coedit-editor',noOp:true,applied:false,state:{ref:before},verification:{measurable:true,match:true,expected:target}}
  if(!apply)return {ok:true,outcome:'print-area-observed',source:'live-coedit-editor',noOp:false,applied:false,state:{ref:before},verification:{measurable:true,match:false,expected:target,actual:before}}
  let ranges=null;if(mode!=='clear'){const r=model.getRange2(range);if(!r||!r.bbox)return {ok:false,outcome:'print-area-target-unavailable',source:'live-coedit-editor'};ranges=[r.bbox]}
  const types=window.Asc.c_oAscChangePrintAreaType, type=mode==='set'?types.set:mode==='add'?types.add:types.clear
  view.changePrintArea(type,ranges)
  const after=read(),ok=mode==='clear'?!after:(mode==='set'?norm(after)===norm(target):!!after&&norm(after).includes(norm(target)))
  return {ok,outcome:ok?'print-area-applied':'print-area-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,state:{ref:after},verification:{measurable:true,match:ok,expected:target,actual:after}}
 }catch(err){return {ok:false,outcome:'print-area-error',source:'live-coedit-editor',error:String(err&&err.message||err)}}},{sheet:spec.sheet,mode:spec.mode,range:spec.range||null,apply:!!apply})
}
async function runCommand(session,spec,apply){return immediate(session,spec,apply)}
module.exports={validSpec,normalizeRef,immediate,runCommand}
