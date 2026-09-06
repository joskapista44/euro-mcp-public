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
const HELPER = path.join(__dirname, 'box-helper.py')
const EXEC = process.env.EURO_EXEC || 'ssh'
const SSH_HOST = process.env.EURO_SSH_HOST || ''
const SSH_USER = process.env.EURO_SSH_USER || 'user'
const VAULT_KEY = process.env.EURO_VAULT_KEY || 'ssh_host'
const VAULT_ROOT = process.env.EURO_VAULT_ROOT || ''
const BOX_IP = process.env.EURO_BOX_IP || '172.22.0.1'
const DS_URL = process.env.EO_DS_URL || 'http://127.0.0.1:8081'
const TIMEOUT_MS = Number(process.env.EURO_TIMEOUT_MS || 90000)
// On the measured EuroOffice deployment the JWT is owned by the running Document Server
// container, not by a host-side env file. The SSH transport therefore streams only that one
// variable through a 0600 FIFO into box-helper.py. The secret never appears in argv, stdout,
// stderr, Git, or a regular file. Override the container name for other deployments.
const JWT_CONTAINER = process.env.EURO_JWT_CONTAINER || 'euro-office'

function run(cmd, args, { input, timeoutMs } = {}) {
  return new Promise((resolve) => {
    const proc = spawn(cmd, args)
    let stdout = ''
    let stderr = ''
    const timer = setTimeout(() => proc.kill('SIGKILL'), timeoutMs || TIMEOUT_MS)
    proc.stdout.on('data', (c) => (stdout += c))
    proc.stderr.on('data', (c) => (stderr += c))
    proc.on('error', (err) => { clearTimeout(timer); resolve({ code: -1, stdout, stderr: String(err.message) }) })
    proc.on('close', (code) => { clearTimeout(timer); resolve({ code, stdout, stderr }) })
    if (input !== undefined) { proc.stdin.write(input); proc.stdin.end() }
  })
}

function callerAgentId(env = process.env, cwd = process.cwd()) {
  const explicit = String(env.EURO_AGENT_ID || '').trim()
  const fromCwd = (cwd.match(/\/agents\/([^/]+)(?:\/|$)/) || [])[1] || ''
  return explicit || fromCwd || 'euro-mcp-runner:unresolved-cwd'
}

async function resolveKeyMaterial() {
  const explicitPath = process.env.EURO_SSH_KEY_PATH
  if (explicitPath) return fs.readFileSync(explicitPath, 'utf8')
  if (!VAULT_ROOT) throw new Error('no SSH key source configured: set EURO_SSH_KEY_PATH to a private-key file, or set EURO_VAULT_ROOT to point at a deployment that provides dist/web/vault.js (getSecret)')
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
    try { fs.writeFileSync(keyFile, crypto.randomBytes(fs.statSync(keyFile).size)) } catch {}
    try { fs.unlinkSync(keyFile) } catch {}
  }
}

function shellQuote(value) { return `'${String(value).replace(/'/g, `'"'"'`)}'` }

function remoteHelperCommand(scriptB64, docArg, docFlag, traceArg) {
  const args = [scriptB64, docArg, BOX_IP, docFlag, traceArg].map(shellQuote).join(' ')
  const container = shellQuote(JWT_CONTAINER)
  const dsUrl = shellQuote(DS_URL)
  // mktemp creates a regular file; remove it immediately and replace the pathname with a FIFO.
  // docker writes JWT_SECRET into the pipe while Python reads it as its env-file input. No secret
  // bytes are persisted on the SSH host. `wait` also prevents a hidden docker-exec failure.
  return [
    'set -eu',
    'p=$(mktemp /tmp/euro-mcp-jwt.XXXXXX)',
    'rm -f "$p"',
    'mkfifo -m 600 "$p"',
    'trap \'rm -f "$p"\' EXIT HUP INT TERM',
    `(docker exec ${container} sh -c 'if [ -z "$JWT_SECRET" ]; then exit 41; fi; printf "JWT_SECRET=%s\\n" "$JWT_SECRET"' >"$p") & jwtpid=$!`,
    `EO_ENV_FILE="$p" EO_DS_URL=${dsUrl} python3 - ${args}`,
    'rc=$?',
    'wait "$jwtpid" || { echo "JWT source failed" >&2; exit 42; }',
    'exit "$rc"',
  ].join('; ')
}

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
    if (!SSH_HOST) throw new Error('EURO_SSH_HOST is required when EURO_EXEC is not "local" -- set it to the euro-office box address (no default is provided)')
    result = await withKey((keyFile) => run('ssh', [
      '-i', keyFile,
      '-o', 'StrictHostKeyChecking=accept-new',
      '-o', 'ConnectTimeout=10',
      '-o', 'BatchMode=yes',
      `${SSH_USER}@${SSH_HOST}`,
      remoteHelperCommand(scriptB64, docArg, docFlag, traceArg),
    ], { input: fs.readFileSync(HELPER) }))
  }

  const logResult = (obj) => {
    if (traceId) logTrace(traceId, 'job-result', { ok: obj.ok, outcome: obj.outcome, dsError: obj.dsError, kind: obj.kind })
    return obj
  }
  const lines = result.stdout.split('\n').filter((l) => l.trim().startsWith('{'))
  if (!lines.length) return logResult({ ok: false, outcome: 'nem-mert', detail: 'the helper produced no JSON answer', exitCode: result.code, stderr: result.stderr.slice(0, 600), stdout: result.stdout.slice(0, 600), traceId })
  try {
    const answer = JSON.parse(lines[lines.length - 1])
    if (EXEC === 'local') return logResult({ ...answer, serverFetches: null })
    const both = `${result.stdout}\n${result.stderr}`
    const fetched = both.split('\n').filter((l) => /GET \/(script\.docbuilder|bemenet\.docx)/.test(l) && !l.startsWith('127.0.0.1'))
    return logResult({ ...answer, serverFetches: fetched.length })
  } catch (err) {
    return logResult({ ok: false, outcome: 'nem-mert', detail: `the helper's answer could not be parsed: ${err.message}`, traceId })
  }
}

module.exports = { runJob, EXEC }
