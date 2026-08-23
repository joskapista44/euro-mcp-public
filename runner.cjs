const fs = require('fs')
const os = require('os')
const path = require('path')
const crypto = require('crypto')
const { spawn } = require('child_process')
const { logTrace } = require('./office-trace.cjs')

// The single way a DocBuilder job is run. Both transports drive the SAME box-helper.py with the
// same arguments and read back the same one-line JSON, so the gate exercises the code that
// actually runs in production. Two separate code paths - one for the tests, one for real use -
// would leave the gate measuring the half nobody deploys.
//
//   EURO_EXEC=ssh    (default) run the helper on the euro-office box over ssh
//   EURO_EXEC=local            run the helper here; used by the gate against a fake server
//
// Access to that box is scoped to reaching euro-office and nothing else, so nothing is installed
// there: the helper travels on stdin per call and the temp directory it uses is removed on exit.

const HELPER = path.join(__dirname, 'box-helper.py')

const EXEC = process.env.EURO_EXEC || 'ssh'
// No default here: the original deployment's default pointed at that operator's own internal
// box and would silently "work" for nobody else while looking like a functioning value. When
// EXEC is not 'local', a missing EURO_SSH_HOST is a configuration error and must fail closed --
// see the check in runJob() below.
const SSH_HOST = process.env.EURO_SSH_HOST || ''
const SSH_USER = process.env.EURO_SSH_USER || 'user'
const VAULT_KEY = process.env.EURO_VAULT_KEY || 'ssh_host'
const VAULT_ROOT = process.env.EURO_VAULT_ROOT || ''
const BOX_IP = process.env.EURO_BOX_IP || '172.22.0.1'
const TIMEOUT_MS = Number(process.env.EURO_TIMEOUT_MS || 90000)

function run(cmd, args, { input, timeoutMs } = {}) {
  return new Promise((resolve) => {
    const proc = spawn(cmd, args)
    let stdout = ''
    let stderr = ''
    const timer = setTimeout(() => proc.kill('SIGKILL'), timeoutMs || TIMEOUT_MS)
    proc.stdout.on('data', (c) => (stdout += c))
    proc.stderr.on('data', (c) => (stderr += c))
    proc.on('error', (err) => {
      clearTimeout(timer)
      resolve({ code: -1, stdout, stderr: String(err.message) })
    })
    proc.on('close', (code) => {
      clearTimeout(timer)
      resolve({ code, stdout, stderr })
    })
    if (input !== undefined) {
      proc.stdin.write(input)
      proc.stdin.end()
    }
  })
}

// A shared literal like 'euro-mcp-runner' would only say a job read the key, not WHICH caller's
// process did -- useful for a single-caller deployment, not one where several independent
// processes share the same audit trail for that key. This mirrors coedit.cjs's
// detectCallerId() identification half (NOT its EURO_COEDIT_AGENTS allowlist check -- withKey()
// is not a co-editing gate, every euro-mcp child needs SSH-key access, not just the
// co-editing-allowlisted ones): `EURO_AGENT_ID` if explicitly set, otherwise the
// `/agents/<name>/` segment of the process's own cwd (each caller's euro-mcp child launches with
// its own directory as cwd -- measured directly via /proc/<pid>/cwd against a real multi-caller
// deployment).
function callerAgentId(env = process.env, cwd = process.cwd()) {
  const explicit = String(env.EURO_AGENT_ID || '').trim()
  const fromCwd = (cwd.match(/\/agents\/([^/]+)(?:\/|$)/) || [])[1] || ''
  return explicit || fromCwd || 'euro-mcp-runner:unresolved-cwd'
}

// Resolves the private key into a 0600 temp file, hands it to one ssh call, and removes it
// again. Nothing long-lived is written: a key that outlives the call it was unlocked for is the
// failure this shape exists to avoid, and it is the one that survives a crashed run.
//
// TWO KEY SOURCES, checked in order:
//   1. EURO_SSH_KEY_PATH -- a plain private-key file, read as-is. This is the path every
//      deployment outside the original operator's own secret-vault setup should use.
//   2. A vault module at `<EURO_VAULT_ROOT>/dist/web/vault.js`, exporting `getSecret(name,
//      callerId)`. This is a CREDENTIAL-PROVIDER INTEGRATION POINT for an operator who already
//      has their own secret store shaped this way -- not a dependency of this project, and that
//      module is not part of this repository. EURO_VAULT_ROOT has no default (a hardcoded default
//      would point at one specific deployment's path and would not resolve for anyone else). If
//      neither source is available, this throws a named, fail-closed error rather than a raw
//      import failure -- there is no fallback to any other credential source.
async function resolveKeyMaterial() {
  const explicitPath = process.env.EURO_SSH_KEY_PATH
  if (explicitPath) {
    return fs.readFileSync(explicitPath, 'utf8')
  }
  if (!VAULT_ROOT) {
    throw new Error(
      'no SSH key source configured: set EURO_SSH_KEY_PATH to a private-key file, or set ' +
      'EURO_VAULT_ROOT to point at a deployment that provides dist/web/vault.js (getSecret)'
    )
  }
  const { getSecret } = await import(`file://${path.join(VAULT_ROOT, 'dist', 'web', 'vault.js')}`)
  const key = getSecret(VAULT_KEY, callerAgentId())
  if (!key) throw new Error(`the vault has no secret named ${VAULT_KEY}`)
  return key
}

async function withKey(fn) {
  const keyFile = path.join(os.tmpdir(), `.euro-mcp-key-${crypto.randomBytes(6).toString('hex')}`)
  try {
    const key = await resolveKeyMaterial()
    fs.writeFileSync(keyFile, key.endsWith('\n') ? key : key + '\n', { mode: 0o600 })
    return await fn(keyFile)
  } finally {
    try {
      fs.writeFileSync(keyFile, crypto.randomBytes(fs.statSync(keyFile).size))
    } catch {
      /* the file may not exist if resolving failed; the unlink below is what matters */
    }
    try {
      fs.unlinkSync(keyFile)
    } catch {
      /* already gone */
    }
  }
}

// One DocBuilder job. `script` may contain __DOC_URL__, which the helper replaces with the URL
// of the document it serves - the caller cannot know that address, because the port is chosen
// per run and the host is the box, not this machine.
//
// office-diag-e6-trace-id: `traceId` is the CALLER's ID (euro-mcp.cjs generates one per external
// tool invocation, via office-trace.cjs's newTraceId() -- never here, and never stored on this
// module: two overlapping calls to runJob() must never share state, only their own parameter).
// It rides along as a 5th positional arg to box-helper.py (a fresh OS process per call, so a
// module-level global is safe THERE, unlike here), which echoes it back in every JSON line it
// prints -- so the "docservice / converter / output verification" stages, which all happen
// inside that one black-box script, are at least bracketed and named by the SAME id our own
// dispatch/result log lines carry, even though we cannot see inside that single blocking call.
async function runJob({ script, documentBase64, returnDoc, traceId }) {
  const scriptB64 = Buffer.from(script, 'utf8').toString('base64')
  const docArg = documentBase64 || '-'
  const docFlag = returnDoc ? 'return-doc' : '-'
  const traceArg = traceId || ''

  if (traceId) logTrace(traceId, 'office-adapter-dispatch', { exec: EXEC, scriptBytes: script.length, hasInputDoc: Boolean(documentBase64), returnDoc: Boolean(returnDoc) })

  let result
  if (EXEC === 'local') {
    result = await run('python3', [HELPER, scriptB64, docArg, BOX_IP, docFlag, traceArg])
  } else {
    if (!SSH_HOST) {
      throw new Error(
        'EURO_SSH_HOST is required when EURO_EXEC is not "local" -- set it to the euro-office ' +
        'box address (no default is provided)'
      )
    }
    result = await withKey((keyFile) =>
      run(
        'ssh',
        [
          '-i', keyFile,
          '-o', 'StrictHostKeyChecking=accept-new',
          '-o', 'ConnectTimeout=10',
          '-o', 'BatchMode=yes',
          `${SSH_USER}@${SSH_HOST}`,
          `python3 - '${scriptB64}' '${docArg}' '${BOX_IP}' '${docFlag}' '${traceArg}'`,
        ],
        { input: fs.readFileSync(HELPER) },
      ),
    )
  }

  // office-diag-e6-trace-id: every exit from here on is the "job-result" stage -- it collapses
  // docservice+converter+output-verification into one log line because box-helper.py is a single
  // blocking round trip, not a multi-stage reporter back to us; see the function's own comment.
  const logResult = (obj) => {
    if (traceId) logTrace(traceId, 'job-result', { ok: obj.ok, outcome: obj.outcome, dsError: obj.dsError, kind: obj.kind })
    return obj
  }

  // The helper prints its own server's access log before the JSON line, and that log is useful:
  // it says how far the other side got. So the answer is the LAST json object on stdout, not the
  // first line - and if there is none, the raw output is handed back rather than a guess.
  const lines = result.stdout.split('\n').filter((l) => l.trim().startsWith('{'))
  if (!lines.length) {
    return logResult({
      ok: false,
      outcome: 'nem-mert',
      detail: 'the helper produced no JSON answer',
      exitCode: result.code,
      stderr: result.stderr.slice(0, 600),
      stdout: result.stdout.slice(0, 600),
      traceId,
    })
  }
  try {
    const answer = JSON.parse(lines[lines.length - 1])
    // The access log lines are evidence, not noise: without them a fetch failure cannot be told
    // apart from the service never having reached us. The count excludes 127.0.0.1, which is our
    // own precondition check.
    //
    // Over the local transport it is deliberately NOT reported: there the fake server also calls
    // from 127.0.0.1, so its fetches are indistinguishable from ours. Reporting 0 would look like
    // "the service never reached us" for a run that worked - a measurement that cannot tell the
    // two apart must say so rather than pick one.
    if (EXEC === 'local') return logResult({ ...answer, serverFetches: null })
    // The access log goes to STDERR - python's http.server logs there, not on stdout. Counting
    // only stdout reported 0 fetches for a run that demonstrably worked, which is the worst shape
    // a measurement can take: a false zero that reads as "the service never reached us".
    const both = `${result.stdout}\n${result.stderr}`
    const fetched = both
      .split('\n')
      .filter((l) => /GET \/(script\.docbuilder|bemenet\.docx)/.test(l) && !l.startsWith('127.0.0.1'))
    return logResult({ ...answer, serverFetches: fetched.length })
  } catch (err) {
    return logResult({ ok: false, outcome: 'nem-mert', detail: `the helper's answer could not be parsed: ${err.message}`, traceId })
  }
}

module.exports = { runJob, EXEC }
