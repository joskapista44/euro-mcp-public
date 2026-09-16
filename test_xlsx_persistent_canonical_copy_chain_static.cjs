'use strict'
const assert=require('assert'),fs=require('fs')
const s=fs.readFileSync(require.resolve('./xlsx-persistent-session.cjs'),'utf8')
assert.ok(s.includes("copyChainRetry=require('./xlsx-copy-chain-retry-classifier.cjs')"))
assert.ok(s.includes('copyChainRetry.classifyCopyChainRetry({plan,api,inventory:initial})'))
assert.ok(s.includes('if(retry.matched)return {...retry,plan,freshBoundaryCount:1}'))
assert.ok(s.includes('return agentTask.executeTask({task:options.task,api})'))
assert.equal((s.match(/classifyCopyChainRetry\(/g)||[]).length,1)
console.log('XLSX CANONICAL PERSISTENT COPY CHAIN STATIC: PASS')
