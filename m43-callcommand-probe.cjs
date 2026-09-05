'use strict'

// Read-only live transport probe for the spreadsheet editor.
// It does not mutate workbook content. The goal is to distinguish the historical
// editor.callCommand transport from callback-capable plugin/connector variants.

async function probeInFrame(frame, timeoutMs = 4000) {
  return frame.evaluate(async (timeout) => {
    function describe(value) {
      if (value === null) return { type: 'null' }
      const type = typeof value
      const out = { type }
      try { out.constructor = value && value.constructor ? value.constructor.name : null } catch (_) {}
      try { out.thenable = !!value && typeof value.then === 'function' } catch (_) { out.thenable = false }
      return out
    }

    const asc = window.Asc || {}
    const targets = {
      ascEditor: asc.editor || null,
      windowEditor: window.editor || null,
      ascPlugin: asc.plugin || null,
      ascEditorCapital: asc.Editor || null,
    }

    const capabilities = {}
    for (const [name, target] of Object.entries(targets)) {
      let fn = null
      try { fn = target && target.callCommand } catch (_) {}
      capabilities[name] = {
        exists: !!target,
        callCommand: typeof fn === 'function',
        arity: typeof fn === 'function' ? fn.length : null,
      }
    }

    const editor = targets.ascEditor || targets.windowEditor
    if (!editor || typeof editor.callCommand !== 'function') {
      return { ok: false, capabilities, error: 'no editor.callCommand target' }
    }

    const command = new Function('return {ok:true, marker:"M43-CALLCOMMAND-PROBE", n:7};')

    async function attempt(name, invoke) {
      let callbackValue = '__NOT_CALLED__'
      let callbackCalled = false
      let returned
      let thrown = null
      try {
        returned = invoke((value) => { callbackCalled = true; callbackValue = value })
      } catch (err) {
        thrown = String(err && err.message ? err.message : err)
      }

      const returnedDescription = describe(returned)
      let awaited = null
      if (!thrown && returnedDescription.thenable) {
        try {
          awaited = await Promise.race([
            Promise.resolve(returned).then((value) => ({ state: 'resolved', value })),
            new Promise((resolve) => setTimeout(() => resolve({ state: 'timeout' }), timeout)),
          ])
        } catch (err) {
          awaited = { state: 'rejected', error: String(err && err.message ? err.message : err) }
        }
      } else {
        await new Promise((resolve) => setTimeout(resolve, Math.min(timeout, 1000)))
      }

      return {
        name,
        thrown,
        returned: returnedDescription,
        callbackCalled,
        callbackValue: callbackCalled ? callbackValue : null,
        awaited,
      }
    }

    const attempts = []
    attempts.push(await attempt('editor.callCommand(fn)', (cb) => editor.callCommand(command)))
    attempts.push(await attempt('editor.callCommand(fn, cb)', (cb) => editor.callCommand(command, cb)))
    attempts.push(await attempt('editor.callCommand(fn, false, cb)', (cb) => editor.callCommand(command, false, cb)))
    attempts.push(await attempt('editor.callCommand(fn, false, false, cb)', (cb) => editor.callCommand(command, false, false, cb)))

    return { ok: true, capabilities, attempts }
  }, timeoutMs)
}

async function runProbeLive({ url, user, pass, fileId, loadPlaywright, timeoutMs = 60000, probeTimeoutMs = 4000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok: false, error: loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user)
    await page.fill('#password', pass)
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null),
      page.click('button[type=submit], input[type=submit]'),
    ])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { ok: false, error: 'login failed' }

    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { ok: false, error: 'spreadsheeteditor frame not found' }
    return await probeInFrame(frame, probeTimeoutMs)
  } finally {
    await browser.close().catch(() => {})
  }
}

module.exports = { probeInFrame, runProbeLive }
