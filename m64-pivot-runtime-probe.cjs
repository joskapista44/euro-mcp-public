'use strict';
const { chromium } = require(process.env.EURO_PLAYWRIGHT_PATH || '/home/user/marveen/node_modules/playwright');

function pivotProbeCommand() {
  function methods(obj) {
    if (!obj) return [];
    var out = [], p = obj;
    while (p && p !== Object.prototype) {
      Object.getOwnPropertyNames(p).forEach(function (name) {
        try { if (typeof obj[name] === 'function' && name !== 'constructor' && name[0] !== '_' && name.indexOf('private_') !== 0) out.push(name); } catch (_) {}
      });
      p = Object.getPrototypeOf(p);
    }
    return Array.from(new Set(out)).sort();
  }
  function safe(fn) { try { return fn(); } catch (e) { return { error: String(e && e.message ? e.message : e) }; } }
  try {
    var ws = Api.GetActiveSheet();
    var sourceAddress = 'XFA40:XFC44';
    var source = ws.GetRange(sourceAddress);
    source.SetValue([
      ['Region', 'Style', 'Price'],
      ['East', 'Fancy', 42.5],
      ['West', 'Fancy', 35.2],
      ['East', 'Tee', 12.3],
      ['West', 'Tee', 24.8]
    ]);
    var before = Api.GetAllPivotTables();
    var pivot = Api.InsertPivotNewWorksheet(source);
    var after = Api.GetAllPivotTables();
    var result = {
      ok: !!pivot,
      source: 'live-coedit-editor',
      sourceAddress: sourceAddress,
      beforeCount: before.length,
      afterCount: after.length,
      pivotObjectMethods: methods(pivot),
      name: safe(function(){ return typeof pivot.GetName === 'function' ? pivot.GetName() : null; }),
      sourceValue: safe(function(){ var x = typeof pivot.GetSource === 'function' ? pivot.GetSource() : null; return x && typeof x.GetAddress === 'function' ? x.GetAddress() : x; }),
      parentSheet: safe(function(){ var x = typeof pivot.GetParent === 'function' ? pivot.GetParent() : null; return x && typeof x.GetName === 'function' ? x.GetName() : null; })
    };
    if (pivot && typeof pivot.AddFields === 'function') {
      result.addFields = safe(function(){ pivot.AddFields({ rows: 'Region', columns: 'Style' }); return true; });
    }
    if (pivot && typeof pivot.AddDataField === 'function') {
      result.addDataField = safe(function(){ pivot.AddDataField('Price'); return true; });
    }
    result.rowFields = safe(function(){ var x = typeof pivot.GetRowFields === 'function' ? pivot.GetRowFields() : null; return Array.isArray(x) ? x.length : null; });
    result.columnFields = safe(function(){ var x = typeof pivot.GetColumnFields === 'function' ? pivot.GetColumnFields() : null; return Array.isArray(x) ? x.length : null; });
    result.dataFields = safe(function(){ var x = typeof pivot.GetDataFields === 'function' ? pivot.GetDataFields() : null; return Array.isArray(x) ? x.length : null; });
    return result;
  } catch (e) {
    return { ok: false, source: 'live-coedit-editor', error: String(e && e.message ? e.message : e) };
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
    const body='return ('+pivotProbeCommand.toString()+')();';
    const result=await frame.evaluate(commandBody=>new Promise(resolve=>{const editor=(window.Asc||{}).editor;if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,error:'callCommand unavailable'});editor.callCommand(new Function(commandBody),false,resolve);}),body);
    console.log(JSON.stringify({milestone:'M6.4',kind:'pivot-object-runtime-probe',result},null,2));
  } finally { await browser.close(); }
}
main().catch(e=>{console.error(JSON.stringify({milestone:'M6.4',outcome:'launcher-error',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1;});
