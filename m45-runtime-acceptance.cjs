'use strict'

const { runStructureInFrame } = require('./live-structure.cjs')
const { runTableInFrame } = require('./live-tables.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function statusOf(step) {
  return step && step.verification && step.verification.status ? step.verification.status : (step && step.ok === false ? 'FAIL' : 'UNKNOWN')
}

function overall(steps) {
  const statuses = steps.map(statusOf)
  if (statuses.includes('FAIL')) return 'FAIL'
  if (statuses.includes('UNKNOWN')) return 'UNKNOWN'
  return 'PASS'
}

async function runM45AcceptanceInFrame(frame, apiHely, suffix = Date.now().toString(36)) {
  const sheet = `M45T_${suffix}`.slice(0, 28)
  const steps = []
  const cleanup = []

  const createSheetBody = `return (function(){try{var s=Api.AddSheet(${JSON.stringify(sheet)});return {ok:!!s,outcome:s?'ok':'operation-error',source:'live-coedit-editor',verification:{status:s?'PASS':'FAIL',expected:${JSON.stringify(sheet)},actual:s&&typeof s.GetName==='function'?s.GetName():null}}}catch(e){return {ok:false,outcome:'operation-error',source:'live-coedit-editor',error:String(e&&e.message?e.message:e),verification:{status:'FAIL'}}}})();`
  const seedBody = `return (function(){try{var s=Api.GetSheet(${JSON.stringify(sheet)});if(!s)return {ok:false,outcome:'sheet-not-found',source:'live-coedit-editor',verification:{status:'FAIL'}};var r=s.GetRange('A1:B3');r.SetValue([['Product','Price'],['Apples',100],['Oranges',150]]);var a=r.GetValue();return {ok:true,outcome:'ok',source:'live-coedit-editor',verification:{status:'PASS',actual:a}}}catch(e){return {ok:false,outcome:'operation-error',source:'live-coedit-editor',error:String(e&&e.message?e.message:e),verification:{status:'FAIL'}}}})();`

  async function raw(body) {
    return frame.evaluate(({ u, body }) => new Promise((resolve) => {
      const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
      if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok:false, outcome:'nincs-api', source:'live-coedit-editor', verification:{status:'FAIL'} })
      let done=false; const finish=v=>{if(!done){done=true;resolve(v)}}
      try { editor.callCommand(new Function(body), false, v => finish(v === undefined ? {ok:false,outcome:'empty-callback',source:'live-coedit-editor',verification:{status:'FAIL'}} : v)) }
      catch(e){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(e&&e.message?e.message:e),verification:{status:'FAIL'}})}
      setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',verification:{status:'FAIL'}}),15000)
    }), { u: apiHely, body })
  }

  try {
    steps.push({ name: 'create-sheet', ...(await raw(createSheetBody)) })
    if (statusOf(steps[steps.length - 1]) !== 'PASS') return finish()
    steps.push({ name: 'seed-data', ...(await raw(seedBody)) })
    if (statusOf(steps[steps.length - 1]) !== 'PASS') return finish()
    steps.push({ name: 'create-table', ...(await runTableInFrame(frame, apiHely, { type:'table.create', sheet, range:'A1:B3', name:`T_${suffix}`.slice(0,24) })) })
    steps.push({ name: 'inspect-table', ...(await runTableInFrame(frame, apiHely, { type:'table.inspect', sheet })) })
    return finish()
  } finally {
    cleanup.push({ name: 'delete-sheet', ...(await runStructureInFrame(frame, apiHely, { type:'sheet.inspect' }).then(async () => {
      const body = `return (function(){try{var s=Api.GetSheet(${JSON.stringify(sheet)});if(!s)return {ok:true,outcome:'already-absent',source:'live-coedit-editor',verification:{status:'PASS'}};var x=s.Delete();return {ok:x!==false,outcome:x===false?'operation-error':'ok',source:'live-coedit-editor',verification:{status:x===false?'FAIL':'PASS'}}}catch(e){return {ok:false,outcome:'operation-error',source:'live-coedit-editor',error:String(e&&e.message?e.message:e),verification:{status:'FAIL'}}}})();`
      return raw(body)
    })) })
  }

  function finish() {
    const testOutcome = overall(steps)
    const cleanupOutcome = overall(cleanup)
    return { milestone:'M4.5', source:LIVE_SOURCE, outcome:testOutcome === 'PASS' && cleanupOutcome === 'PASS' ? 'PASS' : (testOutcome === 'FAIL' || cleanupOutcome === 'FAIL' ? 'FAIL' : 'UNKNOWN'), testOutcome, cleanupOutcome, humanObservationRequired:false, sheet, steps, cleanup }
  }
}

module.exports = { statusOf, overall, runM45AcceptanceInFrame }
