'use strict';
const { chromium } = require(process.env.EURO_PLAYWRIGHT_PATH || '/home/user/marveen/node_modules/playwright');

function worksheetDeleteProbeCommand() {
  function sheetNames() {
    var sheets = Api.GetSheets();
    return Array.isArray(sheets) ? sheets.map(function (s) { return s.GetName(); }) : [];
  }
  function safe(fn) {
    try { return fn(); }
    catch (e) { return { error: String(e && e.message ? e.message : e) }; }
  }

  try {
    var probeName = 'EURO_M64_DELETE_PROBE';
    var before = sheetNames();
    if (before.indexOf(probeName) !== -1) {
      return {
        ok: false,
        source: 'live-coedit-editor',
        kind: 'public-worksheet-delete-semantic-probe',
        mutationAttempted: false,
        verification: {
          status: 'UNKNOWN',
          expected: { probeSheetAbsentBeforeStart: true },
          actual: { before: before, reason: 'probe sheet already exists; refusing destructive ambiguity' }
        }
      };
    }

    var created = Api.AddSheet(probeName);
    var afterCreate = sheetNames();
    var createVerified = afterCreate.indexOf(probeName) !== -1;
    if (!createVerified) {
      return {
        ok: false,
        source: 'live-coedit-editor',
        kind: 'public-worksheet-delete-semantic-probe',
        mutationAttempted: true,
        verification: {
          status: 'UNKNOWN',
          expected: { presentAfterCreate: true },
          actual: { before: before, afterCreate: afterCreate }
        }
      };
    }

    var target = Api.GetSheet(probeName) || created;
    if (!target || typeof target.Delete !== 'function') {
      return {
        ok: false,
        source: 'live-coedit-editor',
        kind: 'public-worksheet-delete-semantic-probe',
        mutationAttempted: true,
        verification: {
          status: 'UNKNOWN',
          expected: { publicDeleteCallable: true },
          actual: { afterCreate: afterCreate, publicDeleteCallable: false }
        }
      };
    }

    var deleteResult = safe(function () { target.Delete(); return true; });
    var afterDelete = sheetNames();
    var absentAfterDelete = afterDelete.indexOf(probeName) === -1;
    var pass = deleteResult === true && absentAfterDelete;
    return {
      ok: pass,
      source: 'live-coedit-editor',
      kind: 'public-worksheet-delete-semantic-probe',
      mutationAttempted: true,
      method: 'ApiWorksheet.Delete',
      before: before,
      afterCreate: afterCreate,
      afterDelete: afterDelete,
      deleteResult: deleteResult,
      verification: {
        status: pass ? 'PASS' : 'FAIL',
        expected: { presentAfterCreate: true, absentAfterDelete: true },
        actual: { presentAfterCreate: createVerified, absentAfterDelete: absentAfterDelete }
      }
    };
  } catch (e) {
    return {
      ok: false,
      source: 'live-coedit-editor',
      kind: 'public-worksheet-delete-semantic-probe',
      mutationAttempted: true,
      error: String(e && e.message ? e.message : e),
      verification: {
        status: 'UNKNOWN',
        expected: 'create isolated probe worksheet, delete it with public ApiWorksheet.Delete, read back absence via Api.GetSheets',
        actual: 'probe raised before semantic verification completed'
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
    if (!result || !result.verification || result.verification.status !== 'PASS') process.exitCode=1;
  } finally { await browser.close(); }
}

main().catch(e=>{
  console.error(JSON.stringify({milestone:'M6.4',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));
  process.exitCode=1;
});
