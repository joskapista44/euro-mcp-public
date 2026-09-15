'use strict'

const { openMinimalXlsxSession, LIVE_SOURCE } = require('./xlsx-minimal-editor-session.cjs')
const inspector = require('./workbook-inspector.cjs')
const rangeReader = require('./range-reader.cjs')
const bulkWriter = require('./bulk-writer.cjs')
const verification = require('./verification-contract.cjs')
const agentTask = require('./xlsx-agent-task.cjs')

async function inspectInSession(session) {
  const r = await inspector.inspectWorkbookInFrame(session.frame, session.apiWhere)
  return { ...r, authority: r?.ok && r.source === LIVE_SOURCE ? 'LIVE_READ' : 'PLAN_ONLY', session:{persistent:true,sameSession:true} }
}

async function readRangeInSession(session, spec) {
  const r = await rangeReader.readRangeInFrame(session.frame, session.apiWhere, spec)
  return { ...r, authority: r?.ok && r.source === LIVE_SOURCE ? 'LIVE_READ' : 'PLAN_ONLY', session:{persistent:true,sameSession:true} }
}

async function writeRangeVerifiedInSession(session, spec) {
  const before = await readRangeInSession(session, spec)
  if (!before.ok) return { ok:false, outcome:'prewrite-live-read-failed', authority:'PLAN_ONLY', before }
  const write = await bulkWriter.writeBulkInFrame(session.frame, session.apiWhere, spec)
  if (!write?.ok) return { ok:false, outcome:'range-write-dispatch-failed', authority:'DISPATCH_ONLY', before, write }
  session.markWrite()
  const after = await readRangeInSession(session, spec)
  if (!after.ok) return { ok:false, outcome:'postwrite-live-read-failed', authority:'DISPATCH_ONLY', before, write, after }
  const verified = verification.verifyRangeReadback(after, spec.values, spec.formulas || null)
  if (!verified.ok) return { ok:false, outcome:verified.outcome || 'verify-mismatch', authority:'LIVE_READ', before, write, after, verification:verified }
  return { ok:true, outcome:'range-write-live-verified', authority:'LIVE_VERIFY', before, write, after, verification:verified, session:{persistent:true,sameSession:true} }
}

function buildSessionApi(session){
  return {
    session,
    inspect:()=>inspectInSession(session),
    readRange:spec=>readRangeInSession(session,spec),
    writeRangeVerified:spec=>writeRangeVerifiedInSession(session,spec),
  }
}

async function withPersistentXlsxSession(options, task) {
  const session = await openMinimalXlsxSession(options)
  const started = Date.now()
  try {
    const result = await task(buildSessionApi(session))
    return { ...result, persistentSession:{oneEditorSession:true,openedMs:session.openedMs,taskMs:Date.now()-started,writes:session.writes,closedByWrapper:true} }
  } finally {
    await session.close().catch(() => {})
  }
}

async function executeAgentTaskInPersistentSession({fileId,task,credentials,timeoutMs=30000,pollMs=50}={}){
  if(!credentials?.url||!credentials?.user||!credentials?.pass)return {ok:false,outcome:'xlsx-persistent-credentials-required',authority:'PLAN_ONLY',writeAllowed:false}
  return withPersistentXlsxSession({url:credentials.url,user:credentials.user,pass:credentials.pass,fileId,timeoutMs,pollMs},api=>agentTask.executeTask({task,api}))
}

module.exports = { inspectInSession, readRangeInSession, writeRangeVerifiedInSession, buildSessionApi, withPersistentXlsxSession, executeAgentTaskInPersistentSession }
