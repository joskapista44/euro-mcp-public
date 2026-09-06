'use strict'

const fs = require('fs')
const { runCapabilityProbeInFrame } = require('./m46-capability-probe.cjs')

function loadPlaywright() {
  const candidates = [
    process.env.EURO_PLAYWRIGHT_PATH,
    'playwright',
    '/home/user/marveen/node_modules/playwright',
  ].filter(Boolean)
  for (const candidate of candidates) {
    try { return require(candidate) } catch (_) {}
  }
  throw new Error('Playwright is unavailable; set EURO_PLAYWRIGHT_PATH')
}

async function main() {
  const baseUrl = process.env.EURO_NC_BASE_URL
  const fileId = process.env.EURO_NC_FILE_ID
  const user = process.env.EURO_NC_USER
  const password = process.env.EURO_NC_PASSWORD
  const sheet = process.env.EURO_M46_SHEET || 'Sheet1'
  const range = process.env.EURO_M46_RANGE || 'A1:B3'
  if (!baseUrl || !fileId || !user || !password) {
    throw new Error('EURO_NC_BASE_URL, EURO_NC_FILE_ID, EURO_NC_USER and EURO_NC_PASSWORD are required')
  }

  const { chromium } = loadPlaywright()
  const browser = await chromium.launch()
  try {
    const context = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await context.newPage()
    await page.goto(`${baseUrl}/login`, { waitUntil: 'domcontentloaded', timeout: 60000 })
    await page.fill('#user', user)
    await page.fill('#password', password)
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => null),
      page.click('button[type=submit], input[type=submit]'),
    ])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) throw new Error('Nextcloud login did not succeed')

    await page.goto(`${baseUrl}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: 60000 })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) throw new Error('spreadsheeteditor frame did not open')
    const apiHely = await frame.evaluate(() => ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function')
      ? 'window.Asc.editor'
      : (window.editor && typeof window.editor.callCommand === 'function') ? 'window.editor' : null)
    if (!apiHely) throw new Error('callCommand is unavailable')

    const result = await runCapabilityProbeInFrame(frame, apiHely, sheet, range)
    console.log(JSON.stringify({ milestone: 'M4.6', editor: 'spreadsheeteditor', apiHely, ...result }, null, 2))
  } finally {
    await browser.close()
  }
}

if (require.main === module) {
  main().catch((err) => {
    console.error(JSON.stringify({ milestone: 'M4.6', ok: false, outcome: 'launcher-error', error: String(err && err.message ? err.message : err) }, null, 2))
    process.exitCode = 1
  })
}
