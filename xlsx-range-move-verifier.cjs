'use strict'
const copy=require('./xlsx-range-copy-verifier.cjs'),clear=require('./xlsx-range-clear-verifier.cjs')
function verifyRangeMoveSemantic(beforeSource,afterSource,afterTarget){const copied=copy.verifyRangeCopySemantic(beforeSource,afterTarget);if(!copied.ok)return {ok:false,outcome:'range-move-target-semantic-mismatch',authority:'LIVE_READ',targetVerification:copied};const cleared=clear.verifyRangeClearSemantic(afterSource);if(!cleared.ok)return {ok:false,outcome:'range-move-source-not-cleared',authority:'LIVE_READ',sourceVerification:cleared};return {ok:true,outcome:'range-move-live-verified',authority:'LIVE_VERIFY',targetVerification:copied,sourceVerification:cleared}}
module.exports={verifyRangeMoveSemantic}
