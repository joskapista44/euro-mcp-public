'use strict'
function formatObserveCommand(sheetName,rangeAddress,spec,apply){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
 function color(rgb){return rgb==null?null:Api.CreateColorFromRGB(rgb[0],rgb[1],rgb[2])}
 function packed(rgb){return rgb==null?null:((rgb[0]<<16)|(rgb[1]<<8)|rgb[2])>>>0}
 function colorValue(v){if(v==null||v==='No Fill')return v;try{if(has(v,'GetRGB'))return v.GetRGB();if(has(v,'GetHex'))return v.GetHex()}catch(_){}try{if(typeof v.rgb==='number')return v.rgb;if(v.color&&typeof v.color.rgb==='number')return v.color.rgb}catch(_){}return null}
 function borderColorValue(v){if(!v)return null;try{if(has(v,'getRgbOrNull'))return v.getRgbOrNull();if(v.c&&has(v.c,'getRgb'))return v.c.getRgb();if(v.c&&has(v.c,'GetRGB'))return v.c.GetRGB()}catch(_){}return null}
 function internal(range,key){var r=range&&range.range;if(!r)return {measurable:false,value:null};try{if(key==='bold'||key==='italic'||key==='fontName'||key==='fontSize'||key==='fontColor'){var f=has(r,'getFont')?r.getFont():null;if(!f)return {measurable:false,value:null};var m={bold:'b',italic:'i',fontName:'fn',fontSize:'fs',fontColor:'c'},v=f[m[key]];if(key==='bold'||key==='italic')v=!!v;return {measurable:true,value:v}}if(key==='alignHorizontal'||key==='alignVertical'||key==='wrap'){var a=has(r,'getAlign')?r.getAlign():null;if(!a)return {measurable:false,value:null};return {measurable:true,value:key==='alignHorizontal'?a.hor:key==='alignVertical'?a.ver:!!a.wrap}}if(key==='numberFormat'){if(has(r,'getNumFormatStr'))return {measurable:true,value:r.getNumFormatStr()}}if(key==='fillColor'){if(has(r,'getFillColor'))return {measurable:true,value:r.getFillColor()}}}catch(_){}return {measurable:false,value:null}}
 function readProp(range,key,prop,getter){try{if(getter&&has(range,getter))return {measurable:true,value:range[getter]()};var v=range[prop];if(v!==undefined)return {measurable:true,value:v}}catch(_){}return internal(range,key)}
 function normAlign(key,v){if(v==null)return null;if(typeof v!=='number')return String(v).toLowerCase();var h={0:'none',1:'left',2:'center',3:'right',4:'justify',5:'centercontinuous',6:'distributed',7:'fill'};var ver={0:'bottom',1:'center',2:'distributed',3:'justify',4:'top'};var m=key==='alignHorizontal'?h:ver;return m[v]!==undefined?m[v]:String(v)}
 function equal(key,expected,actual){if(key==='fontColor'||key==='fillColor'){var n=colorValue(actual);if(typeof n==='number')return n===packed(expected);if(typeof n==='string'&&/^#?[0-9a-f]{6}$/i.test(n))return n.replace('#','').toLowerCase()===expected.map(function(v){return v.toString(16).padStart(2,'0')}).join('').toLowerCase();return expected==null&&actual==null}if(key==='alignHorizontal'||key==='alignVertical')return normAlign(key,actual)===normAlign(key,expected);return actual===expected}
 function observeBorder(range,borderSpec){
  var styleCodes={None:0,Double:1,Hair:2,DashDotDot:3,DashDot:4,Dotted:5,Dashed:6,Thin:7,MediumDashDotDot:8,SlantDashDot:9,MediumDashDot:10,MediumDashed:11,Medium:12,Thick:13}
  var rr=range&&range.range,bbox=rr&&rr.bbox,ws=rr&&rr.worksheet,expectedStyle=styleCodes[borderSpec.style],expectedColor=packed(borderSpec.color),index=String(borderSpec.index).toLowerCase()
  if(!bbox||!ws||!has(ws,'getRange3')||expectedStyle===undefined)return {measurable:false,match:false,actual:null}
  var checked=0,failures=[]
  function checkProp(prop,row,col,side){
   checked++
   var style=prop&&typeof prop.s==='number'?prop.s:null,colr=borderColorValue(prop),match=style===expectedStyle&&(borderSpec.color==null||colr===expectedColor)
   if(!match&&failures.length<20)failures.push({row:row+1,column:col+1,side:side,style:style,color:colr})
  }
  try{
   for(var r=bbox.r1;r<=bbox.r2;r++)for(var c=bbox.c1;c<=bbox.c2;c++){
    var cellRange=ws.getRange3(r,c,r,c),b=cellRange&&has(cellRange,'getBorderFull')?cellRange.getBorderFull():null
    if(!b)return {measurable:false,match:false,actual:{checkedCells:checked,failures:failures}}
    if(index==='top'&&r===bbox.r1)checkProp(b.t,r,c,'top')
    else if(index==='bottom'&&r===bbox.r2)checkProp(b.b,r,c,'bottom')
    else if(index==='left'&&c===bbox.c1)checkProp(b.l,r,c,'left')
    else if(index==='right'&&c===bbox.c2)checkProp(b.r,r,c,'right')
    else if(index==='insidehorizontal'){
     if(r>bbox.r1)checkProp(b.t,r,c,'top')
     if(r<bbox.r2)checkProp(b.b,r,c,'bottom')
    }else if(index==='insidevertical'){
     if(c>bbox.c1)checkProp(b.l,r,c,'left')
     if(c<bbox.c2)checkProp(b.r,r,c,'right')
    }else if(index==='diagonaldown'||index==='diagonalup'){
     checkProp(b.d,r,c,'diagonal')
     var flag=index==='diagonaldown'?!!b.dd:!!b.du
     if(!flag&&failures.length<20)failures.push({row:r+1,column:c+1,side:index,flag:false})
    }
   }
   if(checked===0)return {measurable:false,match:false,actual:{checkedCells:0,failures:failures}}
   return {measurable:true,match:failures.length===0,actual:{checkedEdges:checked,failures:failures}}
  }catch(_){return {measurable:false,match:false,actual:{checkedEdges:checked,failures:failures}}}
 }
 try{
  if(!has(Api,'GetSheet'))return fail('unsupported','Api.GetSheet is unavailable');var sheet=Api.GetSheet(sheetName);if(!sheet||!has(sheet,'GetRange'))return fail('sheet-not-found','worksheet/range API unavailable');var range=sheet.GetRange(rangeAddress);if(!range)return fail('range-not-found','target range could not be resolved')
  var keys=Object.keys(spec),setters={bold:'SetBold',italic:'SetItalic',fontName:'SetFontName',fontSize:'SetFontSize',fontColor:'SetFontColor',fillColor:'SetFillColor',alignHorizontal:'SetAlignHorizontal',alignVertical:'SetAlignVertical',wrap:'SetWrap',numberFormat:'SetNumberFormat',border:'SetBorders'},readers={bold:['Bold',null],italic:['Italic',null],fontName:['FontName',null],fontSize:['FontSize',null],fontColor:['FontColor',null],fillColor:['FillColor','GetFillColor'],alignHorizontal:['AlignHorizontal',null],alignVertical:['AlignVertical',null],wrap:['WrapText','GetWrapText'],numberFormat:['NumberFormat','GetNumberFormat']}
  if(keys.some(function(k){return !setters[k]}))return fail('unsupported','format property has no agent-grade semantic contract')
  function observe(){var checks={},unknown=[],mismatches=[];for(var i=0;i<keys.length;i++){var key=keys[i];if(key==='border'){var bg=observeBorder(range,spec.border);if(!bg.measurable){checks.border={status:'unknown',expected:spec.border,actual:bg.actual};unknown.push('border')}else{checks.border={status:bg.match?'pass':'fail',expected:spec.border,actual:bg.actual};if(!bg.match)mismatches.push('border')}continue}var rd=readers[key],got=readProp(range,key,rd[0],rd[1]);if(!got.measurable){checks[key]={status:'unknown',expected:spec[key],actual:null};unknown.push(key);continue}var match=equal(key,spec[key],got.value);checks[key]={status:match?'pass':'fail',expected:spec[key],actual:(key==='fontColor'||key==='fillColor')?colorValue(got.value):got.value};if(!match)mismatches.push(key)}return {checks:checks,unknown:unknown,mismatches:mismatches,outcome:unknown.length?'unknown':mismatches.length?'fail':'pass'}}
  var before=observe(),noOp=before.outcome==='pass';if(apply&&!noOp){for(var i=0;i<keys.length;i++){var key=keys[i],method=setters[key];if(!has(range,method))return fail('unsupported','ApiRange.'+method+' is unavailable',{capability:key});var ret;if(key==='border')ret=range.SetBorders(spec.border.index,spec.border.style,color(spec.border.color));else{var value=(key==='fontColor'||key==='fillColor')?color(spec[key]):spec[key];ret=range[method](value)}if(ret===false)return fail('format-error','ApiRange.'+method+' returned false',{capability:key})}}
  var after=apply?observe():before;return {ok:after.outcome==='pass',outcome:after.outcome==='pass'?'ok':'verification-unavailable-or-mismatch',source:'live-coedit-editor',sheet:sheetName,range:String(rangeAddress).toUpperCase(),applied:apply&&!noOp?keys:[],noOp:noOp,verification:after,beforeVerification:before}
 }catch(err){return fail('format-error',String(err&&err.message?err.message:err))}
}
module.exports={formatObserveCommand}
