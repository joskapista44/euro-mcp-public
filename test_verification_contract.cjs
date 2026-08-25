'use strict'

const assert = require('assert')
const fs = require('fs')
const { LIVE_SOURCE, acceptLive, verifyEditEnvelope, verifyRangeReadback, runRangeWriteContract } = require('./verification-contract.cjs')

function live(extra = {}) { return { ok: true, outcome: 'ok', source: LIVE_SOURCE, ...extra } }
function cell(address, value, formula = null) {
  return { address, rawValue: value, value, displayText: value == null ? '' : String(value), formula, dataType: formula ? 'formula' : (value == null ? 'blank' : typeof value) }
}

// M1.1 inspect must prove the live co-edit source.
{
  const r = acceptLive(live({ sheetCount: 2, sheets: [{ name: 'Sheet1' }, { name: 'Sales' }] }), 'inspect', 'M1.1 office_inspect_workbook')
  assert.equal(r.ok, true)
  assert.equal(r.source, LIVE_SOURCE)
}

// M1.2 readback verifies values, formula overlays and explicit blanks.
{
  const after = live({ rows: 2, columns: 2, cells: [
    [cell('A1', 10), cell('B1', 20, '=A1*2')],
    [cell('A2', true), cell('B2', null)],
  ] })
  const r = verifyRangeReadback(after, [[10, 0], [true, null]], [[null, '=A1*2'], [null, null]])
  assert.equal(r.ok, true)
  assert.equal(r.outcome, 'verified')
}

// M1.3 edits participate in the same fail-closed envelope.
{
  assert.equal(verifyEditEnvelope(live({ operation: 'sheet.rename' }), 'M1.3 office_workbook_operation').ok, true)
  const wrongSource = verifyEditEnvelope({ ok: true, outcome: 'ok', source: 'saved-xlsx' }, 'M1.3 office_workbook_operation')
  assert.equal(wrongSource.ok, false)
  assert.equal(wrongSource.outcome, 'non-live-source')
}

// M1.4 full inspect -> read -> edit -> verify contract.
;(async () => {
  let readCount = 0
  const before = live({ rows: 1, columns: 2, cells: [[cell('A1', 1), cell('B1', 2)]] })
  const after = live({ rows: 1, columns: 2, cells: [[cell('A1', 7), cell('B1', 14, '=A1*2')]] })
  const result = await runRangeWriteContract({
    inspect: async () => live({ sheetCount: 1, sheets: [{ name: 'Sheet1' }] }),
    read: async () => (++readCount === 1 ? before : after),
    edit: async () => live({ writtenValues: 1, writtenFormulas: 1 }),
    values: [[7, 0]],
    formulas: [[null, '=A1*2']],
  })
  assert.equal(result.ok, true)
  assert.deepEqual(result.phases, { inspect: 'pass', read: 'pass', edit: 'pass', verify: 'pass' })
  assert.equal(readCount, 2)

  // Verification mismatch must fail closed and preserve evidence.
  readCount = 0
  const mismatch = await runRangeWriteContract({
    inspect: async () => live({ sheetCount: 1 }),
    read: async () => (++readCount === 1 ? before : live({ rows: 1, columns: 2, cells: [[cell('A1', 8), cell('B1', 14, '=A1*2')]] })),
    edit: async () => live({ writtenValues: 1, writtenFormulas: 1 }),
    values: [[7, 0]],
    formulas: [[null, '=A1*2']],
  })
  assert.equal(mismatch.ok, false)
  assert.equal(mismatch.outcome, 'verify-mismatch')
  assert.equal(mismatch.phase, 'verify')
  assert.ok(mismatch.before)
  assert.ok(mismatch.after)

  // Any non-live phase stops the pipeline before edit.
  let edited = false
  const blocked = await runRangeWriteContract({
    inspect: async () => live({ sheetCount: 1 }),
    read: async () => ({ ok: true, outcome: 'ok', source: 'downloaded-xlsx' }),
    edit: async () => { edited = true; return live() },
    values: [[1]],
  })
  assert.equal(blocked.ok, false)
  assert.equal(blocked.outcome, 'non-live-source')
  assert.equal(edited, false)

  // Architectural regression guard: M1.5 implementation must remain live-only.
  for (const file of ['verification-contract.cjs', 'euro-mcp-m15.cjs']) {
    const src = fs.readFileSync(file, 'utf8')
    assert.doesNotMatch(src, /DocBuilder\s*\(/)
    assert.doesNotMatch(src, /OOXML.*(write|edit|save)/i)
    assert.doesNotMatch(src, /xlsx.*(writeFile|save)/i)
  }
  const entry = fs.readFileSync('euro-mcp-m15.cjs', 'utf8')
  assert.match(entry, /office_write_range_verified/)
  assert.match(entry, /inspectWorkbookLive/)
  assert.match(entry, /readRangeLive/)
  assert.match(entry, /writeBulkLive/)

  console.log('test_verification_contract: OK')
})().catch((err) => { console.error(err); process.exit(1) })
