'use strict'
const crypto=require('crypto'),pre=require('./xlsx-structural-precondition.cjs'),FP=/^[0-9a-f]{64}$/i
function identity(op){return {version:1,intent:'copy_range',sheet:op?.sheet||null,range:op?.range||null,targetSheet:op?.targetSheet||op?.sheet||null,targetRange:op?.targetRange||null}}
function operationFingerprint(op){return crypto.createHash('sha256').update(JSON.stringify(identity(op))).digest('hex')}
function make(op,source,target){const sourceFingerprint=pre.fingerprint(source),targetFingerprint=pre.fingerprint(target);if(!FP.test(sourceFingerprint||'')||!FP.test(targetFingerprint||''))return null;return {version:1,operationFingerprint:operationFingerprint(op),sourceFingerprint,targetFingerprint}}
function validate(token,op){if(!token||token.version!==1||![token.operationFingerprint,token.sourceFingerprint,token.targetFingerprint].every(x=>FP.test(x||'')))return {ok:false,outcome:'xlsx-range-copy-retry-token-invalid',authority:'PLAN_ONLY'};if(token.operationFingerprint.toLowerCase()!==operationFingerprint(op))return {ok:false,outcome:'xlsx-range-copy-retry-token-operation-mismatch',authority:'PLAN_ONLY'};return {ok:true,outcome:'xlsx-range-copy-retry-token-valid',authority:'PLAN_ONLY',sourceFingerprint:token.sourceFingerprint.toLowerCase(),targetFingerprint:token.targetFingerprint.toLowerCase()}}
module.exports={identity,operationFingerprint,make,validate}
