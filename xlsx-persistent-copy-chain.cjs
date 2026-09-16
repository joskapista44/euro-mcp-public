'use strict'
// Compatibility alias. Copy-chain retry classification now lives in the canonical
// executeAgentTaskInPersistentSession path; keep this export for existing callers.
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
async function executeCopyChainTaskInPersistentSession(options={}){return executeAgentTaskInPersistentSession(options)}
module.exports={executeCopyChainTaskInPersistentSession}
