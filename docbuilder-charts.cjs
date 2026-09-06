'use strict'

// Explicit non-live fallback for chart operations that the deployed co-edit Api.* runtime does
// not expose. Keep this module transport-neutral: it builds a DocBuilder script and verifies the
// machine-readable marker returned from that same script. The caller/adapter is responsible for
// feeding it the actual XLSX bytes and returning the saved XLSX bytes.

function js(v){return JSON.stringify(String(v))}
function marker(name,before,after,deleted){return `__EURO_M51_CHART_DELETE__:${JSON.stringify({name,before,after,deleted})}`}

function buildChartDeleteScript({docUrl='__DOC_URL__',sheet,name,outName='eredmeny.xlsx'}){
 if(!sheet)throw new Error('chart.delete docbuilder: sheet required')
 if(!name)throw new Error('chart.delete docbuilder: name required')
 const expectedMarkerPrefix='__EURO_M51_CHART_DELETE__:'
 return [
  `builder.OpenFile(${js(docUrl)}, "xlsx");`,
  'var __wb = Api.GetActiveWorkbook();',
  `var __sheet = Api.GetSheet(${js(sheet)});`,
  'if (!__sheet) throw new Error("sheet unavailable");',
  'var __before = __sheet.GetAllCharts ? (__sheet.GetAllCharts() || []) : [];',
  `var __matches = (__wb && __wb.GetDrawingsByName) ? (__wb.GetDrawingsByName(${js(name)}) || []) : [];`,
  'var __deleted = false;',
  'for (var __i=0; __i<__matches.length; __i++) { var __d=__matches[__i]; if (__d && __d.GetParentSheet && __d.GetParentSheet().GetName() === '+js(sheet)+' && __d.Delete) { __deleted = !!__d.Delete(); break; } }',
  'var __after = __sheet.GetAllCharts ? (__sheet.GetAllCharts() || []) : [];',
  `__sheet.GetRange("XFD1048576").SetValue(${js(expectedMarkerPrefix)} + JSON.stringify({name:${js(name)},before:__before.length,after:__after.length,deleted:__deleted}));`,
  `builder.SaveFile("xlsx", ${js(outName)});`,
  'builder.CloseFile();'
 ].join('\n')
}

function parseChartDeleteMarker(text){
 const prefix='__EURO_M51_CHART_DELETE__:'
 const i=String(text||'').indexOf(prefix);if(i<0)return null
 const tail=String(text).slice(i+prefix.length)
 const end=tail.search(/[\r\n<]/);const raw=(end<0?tail:tail.slice(0,end)).trim()
 try{return JSON.parse(raw)}catch(_){return null}
}

function verifyChartDelete({name,marker:found}){
 if(!found)return {status:'UNKNOWN',reason:'DocBuilder deletion marker was not machine-readable'}
 if(found.name!==name)return {status:'FAIL',reason:'marker chart name mismatch',expected:name,actual:found.name}
 if(found.deleted!==true)return {status:'FAIL',reason:'DocBuilder runtime did not delete the requested chart',actual:found}
 if(!Number.isInteger(found.before)||!Number.isInteger(found.after)||found.after!==found.before-1)return {status:'FAIL',reason:'chart count did not decrease by exactly one',actual:found}
 return {status:'PASS',expected:{name,countDelta:-1},actual:found}
}

module.exports={buildChartDeleteScript,parseChartDeleteMarker,verifyChartDelete,marker}
