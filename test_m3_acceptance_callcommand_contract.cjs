'use strict'

const assert = require('assert')
const fs = require('fs')

const modules = [
  'formula-inspector.cjs',
  'formula-writer.cjs'
]

for (const file of modules) {
  const source = fs.readFileSync(require.resolve(`./${file}`), 'utf8')
  const oldFourArg = /callCommand\s*\(\s*new Function\([^)]*\)\s*,\s*false\s*,\s*false\s*,\s*\(/m
  assert.ok(!oldFourArg.test(source), `${file}: callback is still passed as the 4th callCommand argument`)
  const callbackSecond = /callCommand\s*\(\s*new Function\([^)]*\)\s*,\s*\(/m
  assert.ok(callbackSecond.test(source), `${file}: no callback-in-second-position callCommand invocation found`)
}

console.log('M3 acceptance callCommand contract: PASS')
