'use strict'
function sortObserveCommand(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function col(s){var m=String(s||'').replace(/\$/g,'').match(/(?:^|!)([A-Z]+)\d+/i);if(!m)return null;var n=0,u=m[1].toUpperCase();for(var i=0;i<u.length;i++)n=n*26+u.charCodeAt(i)-64;return n}
 function clone(v){try{return JSON.parse(JSON.stringify(v))}catch(_){return null}}
 function cmp(a,b){if(a==null||a==='')return b==null||b===''?0:1;if(b==null||b==='')return -1;if(typeof a==='number'&&typeof b==='number')return a-b;return String(a).localeCompare(String(b),undefined,{numeric:true,sensitivity:'base'})}
 function observed(matrix,keyIndex,order,hasHeaders){if(!Array.isArray(matrix)||!matrix.length||keyIndex<0)return null;var start=hasHeaders?1:0,keys=[];for(var i=start;i<matrix.length;i++){if(!Array.isArray(matrix[i])||keyIndex>=matrix[i].length)return null;keys.push(matrix[i][keyIndex])}var match=true;for(var j=1;j<keys.length;j++){var c=cmp(keys[j-1],keys[j]);if(order==='asc'?c>0:c<0){match=false;break}}return {match:match,keys:keys}}
 try{
  if(!spec||typeof spec.sheet!=='string'||typeof spec.range!=='string'||typeof spec.keyRange!=='string')return {ok:false,outcome:'invalid-operation',source:'live-coedit-editor'}
  if(!['asc','desc'].includes(spec.order))return {ok:false,outcome:'invalid-sort-order',source:'live-coedit-editor'}
  var sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null,r=sh&&has(sh,'GetRange')?sh.GetRange(spec.range):null
  if(!r||!has(r,'GetValue'))return {ok:false,outcome:'sort-readback-unavailable',source:'live-coedit-editor'}
  var keyIndex=col(spec.keyRange)-col(spec.range),before=clone(r.GetValue()),pre=observed(before,keyIndex,spec.order,spec.hasHeaders!==false)
  if(!pre)return {ok:false,outcome:'sort-precheck-unverifiable',source:'live-coedit-editor'}
  if(!spec.apply)return {ok:true,outcome:pre.match?'sort-already-satisfied':'sort-observed',source:'live-coedit-editor',noOp:pre.match,matrix:before,keyIndex:keyIndex,verification:{measurable:true,match:pre.match,keys:pre.keys}}
  if(pre.match)return {ok:true,outcome:'sort-already-satisfied',source:'live-coedit-editor',noOp:true,matrix:before,keyIndex:keyIndex,verification:{measurable:true,match:true,keys:pre.keys}}
  if(!has(r,'SetSort'))return {ok:false,outcome:'sort-setter-unavailable',source:'live-coedit-editor'}
  r.SetSort(spec.keyRange,spec.order==='asc'?'xlAscending':'xlDescending',null,null,null,null,spec.hasHeaders===false?'xlNo':'xlYes','xlSortColumns')
  var after=clone(r.GetValue()),post=observed(after,keyIndex,spec.order,spec.hasHeaders!==false)
  if(!post)return {ok:false,outcome:'sort-postcheck-unverifiable',source:'live-coedit-editor',before:before}
  return {ok:post.match,outcome:post.match?'sort-live-verified':'sort-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before:before,matrix:after,keyIndex:keyIndex,verification:{measurable:true,match:post.match,keys:post.keys}}
 }catch(e){return {ok:false,outcome:'sort-operation-error',source:'live-coedit-editor',error:String(e&&e.message||e)}}
}
module.exports={sortObserveCommand}
