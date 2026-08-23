// A per-call trace ID that follows one Office/DocBuilder job through every point WE control
// (caller -> office adapter -> DocBuilder/editor.callCommand -> docservice -> converter -> output
// verification), so a job that produces no output can be time-window paired against the Document
// Server's own logs without guessing.
//
// *** CONCURRENCY IS THE POINT, NOT A NICE-TO-HAVE. *** A trace ID kept in a module-level
// variable would pass its own first round of testing (one call, sequentially) and still be
// wrong: two overlapping async calls share this module's state, and the SECOND call's ID would
// leak into log lines written while the FIRST call is still in flight. So `newTraceId()` never
// stores anything here -- every function below takes the ID as an explicit parameter, and the
// caller (euro-mcp.cjs, one per external tool invocation) is the only place a trace ID is ever
// created. See test-office-trace.cjs for the concurrent-calls proof.
const fs = require('fs')
const path = require('path')
const crypto = require('crypto')

const LOG_DIR = process.env.EURO_TRACE_LOG_DIR || path.join(__dirname, 'store-office-trace')

function pad(n, len = 2) { return String(n).padStart(len, '0') }

// Format: off-YYYYMMDD-HHMMSS-XXXXXX (close to the card's own example, off-20260817-181323-a7f3,
// widened from 2 to 3 random bytes -- MEASURED, not assumed: a first version with 2 bytes (16
// bits, 65536 values) hit a real collision in test-office-trace.cjs's own tight-loop check (499
// unique of 500 draws in the same wall-clock second -- exactly the birthday-paradox math predicts
// at that draw count). A colliding ID defeats the one thing this module exists for: telling two
// calls' log lines apart. 3 bytes (24 bits, ~16.7M values) makes the same test pass at 500/500.
// The timestamp half is for a human scanning the log; the random half is what actually
// guarantees uniqueness under concurrent calls in the same second.
function newTraceId(now = new Date()) {
  const stamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`
  const rand = crypto.randomBytes(3).toString('hex')
  return `off-${stamp}-${rand}`
}

function logPathFor(now = new Date()) {
  return path.join(LOG_DIR, `office-trace-${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}.jsonl`)
}

// Appends ONE JSON line. `stage` names the point in the chain (see the card's own list); `data`
// is stage-specific detail. Never throws -- a logging failure must not take down the actual
// DocBuilder call it is trying to describe (the job itself is the thing that matters).
function logTrace(traceId, stage, data = {}) {
  if (!traceId) throw new Error('logTrace: traceId is required -- an unlogged stage is a silent gap, not a missing nicety')
  const line = JSON.stringify({ ts: new Date().toISOString(), traceId, stage, ...data })
  try {
    fs.mkdirSync(LOG_DIR, { recursive: true })
    fs.appendFileSync(logPathFor(), line + '\n')
  } catch (e) {
    // Deliberately swallowed (see the function comment) -- but not silently: stderr is not the
    // job's own stdout channel, so this cannot corrupt a JSON answer a caller is about to parse.
    process.stderr.write(`office-trace: could not write log line: ${e.message}\n`)
  }
}

module.exports = { newTraceId, logTrace, LOG_DIR, logPathFor }
