'use strict'

const { loadPlaywright } = require('./coedit.cjs')
const { runM45AcceptanceInFrame } = require('./m45-runtime-acceptance.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

async function runLive({ url, user, pass, fileId, timeoutMs = 60000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { milestone:'M4.5', source:LIVE_SOURCE, outcome:'FAIL', error:loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil:'domcontentloaded', timeout:timeoutMs })
    await page.fill('#user', user)
    await page.fill('#password', pass)
    await Promise.all([
      page.waitForNavigation({ waitUntil:'domcontentloaded', timeout:timeoutMs }).catch(() => null),
      page.click('button[type=submit], input[type=submit]')
    ])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { milestone:'M4.5', source:LIVE_SOURCE, outcome:'FAIL', error:'login failed' }

    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil:'domcontentloaded', timeout:timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { milestone:'M4.5', source:LIVE_SOURCE, outcome:'FAIL', error:'spreadsheeteditor frame not found' }

    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return { milestone:'M4.5', source:LIVE_SOURCE, outcome:'FAIL', error:'callCommand is unavailable' }

    const result = await runM45AcceptanceInFrame(frame, apiHely)
    return { ...result, editor:'spreadsheeteditor', apiHely }
  } finally {
    await browser.close().catch(() => {})
  }
}

async function main() {
  const env = process.env
  const required = ['EURO_NC_URL','EURO_NC_USER','EURO_NC_PASS','EURO_NC_FILE_ID']
  const missing = required.filter((k) => !env[k])
  if (missing.length) {
    console.log(JSON.stringify({ milestone:'M4.5', source:LIVE_SOURCE, outcome:'FAIL', error:`missing environment: ${missing.join(', ')}` }, null, 2))
    process.exitCode = 2
    return
  }

  let result
  try {
    result = await runLive({
      url: env.EURO_NC_URL.replace(/\/$/, ''),
      user: env.EURO_NC_USER,
      pass: env.EURO_NC_PASS,
      fileId: env.EURO_NC_FILE_ID,
    })
  } catch (err) {
    result = { milestone:'M4.5', source:LIVE_SOURCE, outcome:'FAIL', error:String(err && err.stack ? err.stack : err) }
  }

  console.log(JSON.stringify(result, null, 2))
  process.exitCode = result.outcome === 'PASS' ? 0 : (result.outcome === 'UNKNOWN' ? 3 : 2)
}

if (require.main === module) main()

module.exports = { runLive }
