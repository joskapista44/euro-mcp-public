'use strict'
function formatObserveCommand(sheetName,rangeAddress,spec,apply){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
 function color(rgb){return rgb==null?null:Api.CreateColorFromRGB(rgb[0],rgb[1],rgb[2])}
 function packed(rgb){return rgb==null?null:((rgb[0]<<16)|(rgb[1]<<8)|rgb[2])>>>0}
 function colorValue(v){if(v==null||v==='No Fill')return v;try{if(has(v,'GetRGB'))return v.GetRGB();if(has(v,'GetHex'))return v.GetHex()}catch(_){}return null}
 function readProp(range,prop,getter){try{if(getter&&has(range,getter))return {measurable:true,value:range[getter]()};var v=range[prop];if(v!==undefined)return {measurable:true,value:v}}catch(_){}return {measurable:false,value:null}}
 function equal(key,expected,actual){if(key==='fontColor'||key==='fillColor'){var n=colorValue(actual);if(typeof n==='number')return n===packed(expected);if(typeof n==='string'&&/^#?[0-9a-f]{6}$/i.test(n))return n.replace('#','').toLowerCase()===expected.map(function(v){return v.toString(16).padStart(2,'0')}).join('').toLowerCase();return false}return actual===expected}
 try{
  if(!has(Api,'GetSheet'))return fail('unsupported','Api.GetSheet is unavailable');var sheet=Api.GetSheet(sheetName);if(!sheet||!has(sheet,'GetRange'))return fail('sheet-not-found','worksheet/range API unavailable');var range=sheet.GetRange(rangeAddress);if(!range)return fail('range-not-found','target range could not be resolved')
  var keys=Object.keys(spec),setters={bold:'SetBold',italic:'SetItalic',fontName:'SetFontName',fontSize:'SetFontSize',fontColor:'SetFontColor',fillColor:'SetFillColor',alignHorizontal:'SetAlignHorizontal',alignVertical:'SetAlignVertical',wrap:'SetWrap',numberFormat:'SetNumberFormat'},readers={bold:['Bold',null],italic:['Italic',null],fontName:['FontName',null],fontSize:['FontSize',null],fontColor:['FontColor',null],fillColor:['FillColor','GetFillColor'],alignHorizontal:['AlignHorizontal',null],alignVertical:['AlignVertical',null],wrap:['WrapText','GetWrapText'],numberFormat:['NumberFormat','GetNumberFormat']}
  if(keys.some(function(k){return !setters[k]}))return fail('unsupported','format property has no agent-grade semantic contract')
  function observe(){var checks={},unknown=[],mismatches=[];for(var i=0;i<keys.length;i++){var key=keys[i],rd=readers[key],got=readProp(range,rd[0],rd[1]);if(!got.measurable){checks[key]={status:'unknown',expected:spec[key],actual:null};unknown.push(key);continue}var match=equal(key,spec[key],got.value);checks[key]={status:match?'pass':'fail',expected:spec[key],actual:(key==='fontColor'||key==='fillColor')?colorValue(got.value):got.value};if(!match)mismatches.push(key)}return {checks:checks,unknown:unknown,mismatches:mismatches,outcome:unknown.length?'unknown':mismatches.length?'fail':'pass'}}
  var before=observe(),noOp=before.outcome==='pass';if(apply&&!noOp){for(var i=0;i<keys.length;i++){var key=keys[i],method=setters[key];if(!has(range,method))return fail('unsupported','ApiRange.'+method+' is unavailable',{capability:key});var value=(key==='fontColor'||key==='fillColor')?color(spec[key]):spec[key],ret=range[method](value);if(ret===false)return fail('format-error','ApiRange.'+method+' returned false',{capability:key})}}
  var after=apply?observe():before;return {ok:after.outcome==='pass',outcome:after.outcome==='pass'?'ok':'verification-unavailable-or-mismatch',source:'live-coedit-editor',sheet:sheetName,range:String(rangeAddress).toUpperCase(),applied:apply&&!noOp?keys:[],noOp:noOp,verification:after,beforeVerification:before}
 }catch(err){return fail('format-error',String(err&&err.message?err.message:err))}
}
module.exports={formatObserveCommand}
