'use strict'

const { LIVE_SOURCE, runInLiveEditor } = require('./live-formatting.cjs')

function resolveNumberFormat(spec) {
  if (!spec || typeof spec !== 'object' || Array.isArray(spec)) return { ok:false, outcome:'invalid-number-format', error:'number format spec must be an object' }
  const kind = spec.kind
  if (!['currency','percent','date','datetime','decimal','custom'].includes(kind)) return { ok:false, outcome:'invalid-number-format', error:'kind must be currency, percent, date, datetime, decimal or custom' }
  if (kind === 'custom') {
    if (typeof spec.format !== 'string' || !spec.format.length) return { ok:false, outcome:'invalid-number-format', error:'custom format must be a non-empty string' }
    return { ok:true, format:spec.format }
  }
  const decimals = spec.decimals == null ? 2 : spec.decimals
  if (!Number.isInteger(decimals) || decimals < 0 || decimals > 30) return { ok:false, outcome:'invalid-number-format', error:'decimals must be an integer in 0..30' }
  const zeros = decimals ? '.' + '0'.repeat(decimals) : ''
  if (kind === 'percent') return { ok:true, format:'0' + zeros + '%' }
  if (kind === 'decimal') return { ok:true, format:'#,##0' + zeros }
  if (kind === 'date') return { ok:true, format:typeof spec.format === 'string' && spec.format ? spec.format : 'yyyy-mm-dd' }
  if (kind === 'datetime') return { ok:true, format:typeof spec.format === 'string' && spec.format ? spec.format : 'yyyy-mm-dd hh:mm' }
  const symbol = spec.symbol == null ? '$' : spec.symbol
  if (typeof symbol !== 'string' || !symbol.length || /[;\r\n]/.test(symbol)) return { ok:false, outcome:'invalid-number-format', error:'currency symbol must be a non-empty single-line string without semicolons' }
  return { ok:true, format:'"' + symbol.replace(/"/g, '""') + '"#,##0' + zeros }
}

function numberFormatCommand(sheetName, rangeAddress, format, expectations) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error,extra){ var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error}; if(extra)for(var k in extra)x[k]=extra[k]; return x }
  function read(range, method, prop) {
    try {
      if (method && has(range,method)) return { measurable:true, value:range[method]() }
      if (prop && range[prop] !== undefined) return { measurable:true, value:range[prop] }
    } catch (_) {}
    return { measurable:false, value:null }
  }
  function same(a,b){ return JSON.stringify(a) === JSON.stringify(b) }
  try {
    if (!has(Api,'GetSheet')) return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(sheetName)
    if(!sheet || !has(sheet,'GetRange')) return fail('sheet-not-found','worksheet/range API unavailable',{sheet:sheetName})
    var range=sheet.GetRange(rangeAddress)
    if(!range) return fail('range-not-found','target range could not be resolved')
    if(!has(range,'SetNumberFormat')) return fail('unsupported','ApiRange.SetNumberFormat is unavailable',{capability:'numberFormat'})

    var beforeRaw=read(range,'GetValue2','Value2')
    if(!beforeRaw.measurable) beforeRaw=read(range,'GetValue','Value')
    var applied=range.SetNumberFormat(format)
    if(applied===false) return fail('format-error','ApiRange.SetNumberFormat returned false',{capability:'numberFormat'})

    var gotFormat=read(range,'GetNumberFormat','NumberFormat')
    var afterRaw=read(range,'GetValue2','Value2')
    if(!afterRaw.measurable) afterRaw=read(range,'GetValue','Value')
    var display=read(range,'GetText','Text')
    var checks={}; var mismatches=[]; var unknown=[]

    if(gotFormat.measurable){
      var fmtMatch=gotFormat.value===format
      checks.numberFormat={status:fmtMatch?'pass':'fail',expected:format,actual:gotFormat.value}
      if(!fmtMatch)mismatches.push('numberFormat')
    } else { checks.numberFormat={status:'unknown',expected:format,actual:null,reason:'live number-format getter/property unavailable'}; unknown.push('numberFormat') }

    if(beforeRaw.measurable && afterRaw.measurable){
      var rawSame=same(beforeRaw.value,afterRaw.value)
      checks.underlyingValuePreserved={status:rawSame?'pass':'fail',before:beforeRaw.value,after:afterRaw.value}
      if(!rawSame)mismatches.push('underlyingValuePreserved')
    } else { checks.underlyingValuePreserved={status:'unknown',before:beforeRaw.value,after:afterRaw.value,reason:'live raw-value getter unavailable'}; unknown.push('underlyingValuePreserved') }

    if(display.measurable) checks.displayValue={status:'observed',actual:display.value}
    else { checks.displayValue={status:'unknown',actual:null,reason:'live display-text getter/property unavailable'}; unknown.push('displayValue') }

    expectations=expectations||{}
    if(Object.prototype.hasOwnProperty.call(expectations,'expectedValue')){
      if(afterRaw.measurable){ var vm=same(expectations.expectedValue,afterRaw.value); checks.expectedValue={status:vm?'pass':'fail',expected:expectations.expectedValue,actual:afterRaw.value}; if(!vm)mismatches.push('expectedValue') }
      else { checks.expectedValue={status:'unknown',expected:expectations.expectedValue,actual:null,reason:'live raw-value getter unavailable'}; unknown.push('expectedValue') }
    }
    if(Object.prototype.hasOwnProperty.call(expectations,'expectedDisplay')){
      if(display.measurable){ var dm=same(expectations.expectedDisplay,display.value); checks.expectedDisplay={status:dm?'pass':'fail',expected:expectations.expectedDisplay,actual:display.value}; if(!dm)mismatches.push('expectedDisplay') }
      else { checks.expectedDisplay={status:'unknown',expected:expectations.expectedDisplay,actual:null,reason:'live display-text getter unavailable'}; unknown.push('expectedDisplay') }
    }

    var verificationOutcome=mismatches.length?'fail':unknown.length?'unknown':'pass'
    return {
      ok:mismatches.length===0,
      outcome:mismatches.length?'verification-mismatch':'ok',
      source:'live-coedit-editor',
      sheet:sheetName,
      range:String(rangeAddress).toUpperCase(),
      numberFormat:format,
      underlyingValue:afterRaw.measurable?afterRaw.value:null,
      displayValue:display.measurable?display.value:null,
      verification:{outcome:verificationOutcome,checks:checks,mismatches:mismatches,unknown:unknown}
    }
  } catch(err){ return fail('format-error',String(err&&err.message?err.message:err),{capability:'numberFormat'}) }
}

function formatNumberLive(common) {
  const resolved=resolveNumberFormat(common.numberFormat)
  if(!resolved.ok)return {...resolved,source:LIVE_SOURCE}
  return runInLiveEditor({...common,command:numberFormatCommand,args:[common.sheet,common.range,resolved.format,{expectedValue:common.expectedValue,expectedDisplay:common.expectedDisplay}]})
}

module.exports={ resolveNumberFormat,numberFormatCommand,formatNumberLive }
