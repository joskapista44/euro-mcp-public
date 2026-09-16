'use strict'
const crypto=require('crypto')
function hash(v){return crypto.createHash('sha256').update(JSON.stringify(v)).digest('hex')}
function identity(op){return {version:1,intent:'layout_range',sheet:op.sheet,range:op.range,type:op.type}}
function operationFingerprint(op){return hash(identity(op))}
function finite(v){return Number.isFinite(v)&&v>0}
function make(op,beforeDimension,postDimension){if(!finite(beforeDimension)||!finite(postDimension))return null;return {version:1,operationFingerprint:operationFingerprint(op),beforeDimension,postDimension}}
function validate(token,op){if(!token||token.version!==1||typeof token.operationFingerprint!=='string'||!finite(token.beforeDimension)||!finite(token.postDimension))return {ok:false,outcome:'xlsx-layout-autofit-retry-token-invalid',authority:'PLAN_ONLY'};if(token.operationFingerprint!==operationFingerprint(op))return {ok:false,outcome:'xlsx-layout-autofit-retry-operation-mismatch',authority:'PLAN_ONLY'};return {ok:true,...token}}
function near(a,b){return Number.isFinite(a)&&Number.isFinite(b)&&Math.abs(a-b)<=0.05}
module.exports={identity,operationFingerprint,make,validate,near}
