'use strict';
const { chromium } = require(process.env.EURO_PLAYWRIGHT_PATH || '/home/user/marveen/node_modules/playwright');

function worksheetDeleteProbeCommand() {
  function publicMethods(obj) {
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
  function cleanupCandidates(methods) {
    return methods.filter(function (name) {
      return /(delete|remove|sheet|worksheet)/i.test(name);
    });
  }
  function safe(fn) {
    try { return fn(); }
    catch (e) { return { error: String(e && e.message ? e.message : e) }; }
  }

  try {
    var active = Api.GetActiveSheet();
    var apiMethods = publicMethods(Api);
    var sheetMethods = publicMethods(active);
    var result = {
      ok: true,
      source: 'live-coedit-editor',
      kind: 'public-worksheet-delete-capability-inventory',
      activeSheetName: safe(function () { return active && typeof active.GetName === 'function' ? active.GetName() : null; }),
      apiCleanupCandidates: cleanupCandidates(apiMethods),
      worksheetCleanupCandidates: cleanupCandidates(sheetMethods),
      apiMethods: apiMethods,
      worksheetMethods: sheetMethods,
      mutationAttempted: false,
      verification: {
        status: 'UNKNOWN',
        expected: 'discover a documented/public worksheet deletion method before attempting cleanup',
        actual: 'capability inventory only; no worksheet deletion attempted'
      }
    };
    return result;
  } catch (e) {
    return {
      ok: false,
      source: 'live-coedit-editor',
      kind: 'public-worksheet-delete-capability-inventory',
      mutationAttempted: false,
      error: String(e && e.message ? e.message : e),
      verification: {
        status: 'UNKNOWN',
        expected: 'public worksheet deletion capability inventory',
        actual: 'probe failed before capability could be established'
      }
    };
  }
}

async function main() {
  const base=process.env.EURO_NC_BASE_URL,id=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD;
  if(!base||!id||!user||!password) throw new Error('required EURO_NC_* env missing');
  const browser=await chromium.launch();
  try {
    const page=await browser.newPage();
    await page.goto(base+'/login',{waitUntil:'domcontentloaded',timeout:60000});
    await page.fill('#user',user); await page.fill('#password',password); await page.click('button[type=submit], input[type=submit]');
    await page.waitForTimeout(2500); if(/\/login/.test(page.url())) throw new Error('login failed');
    await page.goto(base+'/index.php/apps/eurooffice/'+id,{waitUntil:'domcontentloaded',timeout:60000}); await page.waitForTimeout(22000);
    const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url())); if(!frame) throw new Error('spreadsheeteditor frame missing');
    const body='return ('+worksheetDeleteProbeCommand.toString()+')();';
    const result=await frame.evaluate(commandBody=>new Promise(resolve=>{
      const editor=(window.Asc||{}).editor;
      if(!editor||typeof editor.callCommand!=='function') return resolve({ok:false,error:'callCommand unavailable'});
      editor.callCommand(new Function(commandBody),false,resolve);
    }),body);
    console.log(JSON.stringify({milestone:'M6.4',kind:'worksheet-delete-runtime-probe',result},null,2));
  } finally { await browser.close(); }
}

main().catch(e=>{
  console.error(JSON.stringify({milestone:'M6.4',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));
  process.exitCode=1;
});
