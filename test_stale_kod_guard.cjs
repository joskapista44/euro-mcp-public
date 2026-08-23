// RED/GREEN/NEG.KONTROLL for the freshness guard on the write
// entry points (create_document, coedit_write_operations). Each module (lib.cjs, coedit.cjs,
// euro-mcp.cjs) captures its OWN mtime once, at require() time -- to prove that, the RED case
// must load a COPIED tree fresh (a real require(), not a mocked constant) and then touch the
// COPY's mtime forward, never the live tree (modify-shared-live-bridge-code discipline).

const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')

let osszes = 0
const hibak = []
function check(cimke, felteteles, reszlet = '') {
  osszes += 1
  if (felteteles) { console.log(`  ok    ${cimke}`) } else { hibak.push(cimke); console.log(`  BUKAS ${cimke}${reszlet ? ' -- ' + reszlet : ''}`) }
}

const REPO = __dirname
// this list drifted once already -- an earlier split
// ("split the OPERATIONS table (B) into per-core + multi-core modules") pulled
// lib-operations-{registry,docx,xlsx,pptx}.cjs out of lib.cjs, and this file was not updated, so
// EVERY run failed with "Cannot find module './lib-operations-registry.cjs'" -- the guard itself
// went stale. Kept as an EXPLICIT list (not a blanket directory copy): the repo root also holds
// live secrets (.env and its .bak-* siblings), and a disposable copy under /tmp must never carry
// those. FILES/DIRS below must stay in sync with lib.cjs's own top-level `require('./...')` graph.
const FILES = [
  'lib.cjs', 'coedit.cjs', 'euro-mcp.cjs', 'runner.cjs', 'package-consistency.cjs', 'euro-magok.cjs',
  'lib-operations-registry.cjs', 'lib-operations-docx.cjs', 'lib-operations-xlsx.cjs', 'lib-operations-pptx.cjs',
]
const DIRS = ['operations'] // required by the four lib-operations-*.cjs modules above

function makeDisposableCopy(tag) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), `stale-kod-guard-${tag}-`))
  for (const f of FILES) fs.copyFileSync(path.join(REPO, f), path.join(dir, f))
  for (const d of DIRS) fs.cpSync(path.join(REPO, d), path.join(dir, d), { recursive: true })
  // node_modules must resolve too (McpServer/zod/etc for euro-mcp.cjs).
  fs.symlinkSync(path.join(REPO, 'node_modules'), path.join(dir, 'node_modules'))
  return dir
}

console.log('\n[1] GREEN -- fresh copy, nothing touched, all three report fresh')
{
  const dir = makeDisposableCopy('green')
  const lib = require(path.join(dir, 'lib.cjs'))
  const coedit = require(path.join(dir, 'coedit.cjs'))
  check('lib.cjs checkFreshness() -- fresh', lib.checkFreshness().fresh === true)
  check('coedit.cjs checkFreshness() -- fresh', coedit.checkFreshness().fresh === true)
  fs.rmSync(dir, { recursive: true, force: true })
}

console.log('\n[2] RED -- touch the COPIED lib.cjs forward -> named stale error, coedit.cjs unaffected')
{
  const dir = makeDisposableCopy('red-lib')
  const lib = require(path.join(dir, 'lib.cjs'))
  const coedit = require(path.join(dir, 'coedit.cjs'))
  // Advance the file's mtime 5 minutes into the future -- simulates a comb/merge landing on
  // disk while this "process" (this require) already has the old code loaded.
  const future = new Date(Date.now() + 5 * 60 * 1000)
  fs.utimesSync(path.join(dir, 'lib.cjs'), future, future)
  const libResult = lib.checkFreshness()
  check('lib.cjs checkFreshness() -- reports stale', libResult.fresh === false)
  check('  names the file', libResult.file === 'lib.cjs')
  check('  message says "inditsd ujra a sessiont"', /inditsd ujra a sessiont/.test(libResult.message), libResult.message)
  check('coedit.cjs checkFreshness() -- STILL fresh (per-file, not a shared threshold)', coedit.checkFreshness().fresh === true)
  fs.rmSync(dir, { recursive: true, force: true })
}

console.log('\n[3] RED -- touch the COPIED coedit.cjs forward -> lib.cjs unaffected (the other direction)')
{
  const dir = makeDisposableCopy('red-coedit')
  const lib = require(path.join(dir, 'lib.cjs'))
  const coedit = require(path.join(dir, 'coedit.cjs'))
  const future = new Date(Date.now() + 5 * 60 * 1000)
  fs.utimesSync(path.join(dir, 'coedit.cjs'), future, future)
  check('coedit.cjs checkFreshness() -- reports stale', coedit.checkFreshness().fresh === false)
  check('lib.cjs checkFreshness() -- STILL fresh', lib.checkFreshness().fresh === true)
  fs.rmSync(dir, { recursive: true, force: true })
}

console.log('\n[4] staleCodeGuard() -- combines multiple checks, reports EVERY stale file at once')
{
  const dir = makeDisposableCopy('guard-combo')
  const lib = require(path.join(dir, 'lib.cjs'))
  const coedit = require(path.join(dir, 'coedit.cjs'))
  const future = new Date(Date.now() + 5 * 60 * 1000)
  fs.utimesSync(path.join(dir, 'lib.cjs'), future, future)
  fs.utimesSync(path.join(dir, 'coedit.cjs'), future, future)
  const results = [lib.checkFreshness(), coedit.checkFreshness()]
  check('both report stale independently', results.every((r) => r.fresh === false))
  // staleCodeGuard() itself is not exported from euro-mcp.cjs (only `server` is) -- its combining
  // behavior (report ALL stale files, not just the first) is proven directly against its own
  // definition text below, since reaching it through the MCP SDK's tool-registration internals
  // would make this test depend on SDK-internal shapes rather than this repo's own code.
  const src = fs.readFileSync(path.join(REPO, 'euro-mcp.cjs'), 'utf-8')
  const fnMatch = src.match(/function staleCodeGuard\(\.\.\.checks\) \{([\s\S]*?)\n\}/)
  check('staleCodeGuard() filters to ONLY the non-fresh results (does not stop at the first)', /\.filter\(\(r\) => !r\.fresh\)/.test(fnMatch && fnMatch[1]))
  check('staleCodeGuard() joins ALL stale messages, not just one', /stale\.map\(\(s\) => s\.message\)\.join/.test(fnMatch && fnMatch[1]))
  fs.rmSync(dir, { recursive: true, force: true })
}

console.log('\n[5] STRUCTURAL: the guard sits FIRST in each write handler, and ONLY there')
{
  const src = fs.readFileSync(path.join(REPO, 'euro-mcp.cjs'), 'utf-8')
  // `= staleCodeGuard(` matches an actual invocation (assignment from a call), excluding the
  // `function staleCodeGuard(...)` definition line and the one comment mentioning it by name.
  const guardCallSites = [...src.matchAll(/= staleCodeGuard\(([^)]*)\)/g)]
  check('NEG.KONTROLL: staleCodeGuard() pontosan KETSZER hivodik a forrasban (create_document + coedit_write_operations -- nem tobbszor, nem olvaso utakon)', guardCallSites.length === 2, `talalatok: ${guardCallSites.length}`)

  const createDocIdx = src.indexOf("'create_document'")
  const createDocHandlerIdx = src.indexOf('async ({ core, operations, output_path }) =>', createDocIdx)
  const createDocFirstLogicIdx = src.indexOf('if (output_path && fs.existsSync', createDocHandlerIdx)
  const createDocGuardIdx = src.indexOf('staleCodeGuard(', createDocHandlerIdx)
  check('create_document: a guard a HANDLER ELSO logikai sora, MEGELOZI a meglevo output_path-ellenorzest', createDocGuardIdx > createDocHandlerIdx && createDocGuardIdx < createDocFirstLogicIdx)
  check('  es CSAK euro-mcp.cjs + lib.cjs-t ellenorzi (coedit.cjs nincs ezen az uton)', /staleCodeGuard\(checkOwnFreshness, lib\.checkFreshness\)/.test(src.slice(createDocGuardIdx, createDocGuardIdx + 60)))

  const coeditIdx = src.indexOf("'coedit_write_operations'")
  const coeditHandlerIdx = src.indexOf('async ({ file_id, core, operations }) =>', coeditIdx)
  const coeditFirstLogicIdx = src.indexOf('coedit.loadBridgeEnv()', coeditHandlerIdx)
  const coeditGuardIdx = src.indexOf('staleCodeGuard(', coeditHandlerIdx)
  check('coedit_write_operations: a guard a HANDLER ELSO logikai sora, MEGELOZI a loadBridgeEnv()-et', coeditGuardIdx > coeditHandlerIdx && coeditGuardIdx < coeditFirstLogicIdx)
  check('  es MINDHAROM fajlt ellenorzi (euro-mcp.cjs + coedit.cjs + lib.cjs -- ez az ut mindharmon at megy)', /staleCodeGuard\(checkOwnFreshness, coedit\.checkFreshness, lib\.checkFreshness\)/.test(src.slice(coeditGuardIdx, coeditGuardIdx + 90)))
}

console.log(`\nellenorzesek: ${osszes - hibak.length} ok, ${hibak.length} bukas`)
assert.strictEqual(hibak.length, 0, `bukott: ${hibak.join(' | ')}`)
