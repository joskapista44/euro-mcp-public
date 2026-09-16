'use strict'
const crypto=require('crypto')
function normFormula(v){return typeof v==='string'&&v.trim()?v.replace(/\s+/g,'').toUpperCase():null}
function cellState(c){const f=normFormula(c?.formula);if(f)return {formula:f};if(c&&Object.prototype.hasOwnProperty.call(c,'rawValue'))return {value:c.rawValue};if(c&&Object.prototype.hasOwnProperty.call(c,'value'))return {value:c.value};return {value:c?.displayText??null}}
function snapshot(read){if(!read?.ok||read.authority!=='LIVE_READ'||!Array.isArray(read.cells))return null;return read.cells.map(row=>(row||[]).map(cellState))}
function fingerprint(read){const s=snapshot(read);if(!s)return null;return crypto.createHash('sha256').update(JSON.stringify(s)).digest('hex')}
function classify(read,expected){const actual=fingerprint(read);if(!actual)return {ok:false,outcome:'xlsx-structural-precondition-unreadable',authority:'PLAN_ONLY'};if(typeof expected!=='string'||!/^[0-9a-f]{64}$/i.test(expected))return {ok:false,outcome:'xlsx-structural-precondition-invalid',authority:'PLAN_ONLY',actualFingerprint:actual};if(actual.toLowerCase()!==expected.toLowerCase())return {ok:false,outcome:'xlsx-structural-precondition-mismatch',authority:'LIVE_READ',expectedFingerprint:expected.toLowerCase(),actualFingerprint:actual};return {ok:true,outcome:'xlsx-structural-precondition-match',authority:'LIVE_VERIFY',fingerprint:actual}}
module.exports={snapshot,fingerprint,classify}
