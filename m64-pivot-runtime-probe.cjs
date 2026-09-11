'use strict';
const { chromium } = require(process.env.EURO_PLAYWRIGHT_PATH || '/home/user/marveen/node_modules/playwright');

function publicMethodsCommand() {
  function methods(obj) {
    if (!obj) return [];
    var out = [], p = obj;
    while (p && p !== Object.prototype) {
      Object.getOwnPropertyNames(p).forEach(function (name) {
        try {
          if (typeof obj[name] === 'function' && name !== 'constructor' && name[0] !== '_' && name.indexOf('private_') !== 0) out.push(name);
        } catch (_) {}
      });
      p = Object.getPrototypeOf(p);
    }
    return Array.from(new Set(out)).sort();
  }
  try {
    var ws = Api.GetActiveSheet();
    var workbookPivots = typeof Api.GetAllPivotTables === 'function' ? Api.GetAllPivotTables() : [];
    var sheetPivots = ws && typeof ws.GetAllPivotTables === 'function' ? ws.GetAllPivotTables() : [];
    var first = workbookPivots && workbookPivots.length ? workbookPivots[0] : (sheetPivots && sheetPivots.length ? sheetPivots[0] : null);
    return {
      ok: true,
      source: 'live-coedit-editor',
      apiPivotMethods: methods(Api).filter(function (n) { return /Pivot/i.test(n); }),
      worksheetPivotMethods: methods(ws).filter(function (n) { return /Pivot/i.test(n); }),
      workbookPivotCount: workbookPivots && typeof workbookPivots.length === 'number' ? workbookPivots.length : null,
      worksheetPivotCount: sheetPivots && typeof sheetPivots.length === 'number' ? sheetPivots.length : null,
      pivotObjectPresent: !!first,
      pivotObjectMethods: methods(first)
    };
  } catch (e) {
    return { ok: false, source: 'live-coedit-editor', error: String(e && e.message ? e.message : e) };
  }
}

async function main() {
  const base = process.env.EURO_NC_BASE_URL;
  const id = process.env.EURO_NC_FILE_ID;
  const user = process.env.EURO_NC_USER;
  const password = process.env.EURO_NC_PASSWORD;
  if (!base || !id || !user || !password) throw new Error('required EURO_NC_* env missing');
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    await page.goto(base + '/login', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.fill('#user', user);
    await page.fill('#password', password);
    await page.click('button[type=submit], input[type=submit]');
    await page.waitForTimeout(2500);
    if (/\/login/.test(page.url())) throw new Error('login failed');
    await page.goto(base + '/index.php/apps/eurooffice/' + id, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(22000);
    const frame = page.frames().find(f => /spreadsheeteditor/.test(f.url()));
    if (!frame) throw new Error('spreadsheeteditor frame missing');
    const body = 'return (' + publicMethodsCommand.toString() + ')();';
    const result = await frame.evaluate(commandBody => new Promise(resolve => {
      const editor = (window.Asc || {}).editor;
      if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, error: 'callCommand unavailable' });
      editor.callCommand(new Function(commandBody), false, resolve);
    }), body);
    console.log(JSON.stringify({ milestone: 'M6.4', kind: 'pivot-runtime-probe', result }, null, 2));
  } finally {
    await browser.close();
  }
}

main().catch(e => {
  console.error(JSON.stringify({ milestone: 'M6.4', outcome: 'launcher-error', error: String(e && e.message ? e.message : e) }, null, 2));
  process.exitCode = 1;
});
