'use strict';
const {runChartInFrame}=require('./live-charts.cjs');
const {writeBulkInFrame}=require('./bulk-writer.cjs');

function loadPlaywright() {
  for (const p of [
    process.env.EURO_PLAYWRIGHT_PATH,
    'playwright',
    '/home/user/marveen/node_modules/playwright'
  ].filter(Boolean)) {
    try { return require(p); } catch (_) {}
  }
  throw new Error('Playwright is unavailable');
}

function publicProbe(spec) {
  function invoke(obj, name, args) {
    if (!obj || typeof obj[name] !== 'function')
      return { __error: 'public method unavailable: ' + name };
    try {
      return obj[name].apply(obj, args || []);
    } catch (e) {
      return { __error: String(e && e.message ? e.message : e) };
    }
  }

  function same(a, b) {
    return JSON.stringify(a) === JSON.stringify(b);
  }

  function trimParagraphEnd(value) {
    return typeof value === 'string'
      ? value.replace(/[\r\n]+$/, '')
      : value;
  }

  function normalizeFormula(value) {
    return typeof value === 'string' && value.charAt(0) === '='
      ? value.slice(1)
      : value;
  }

  function decodeSeriesName(value) {
    if (typeof value !== 'string')
      return value;

    if (value.slice(0, 2) === '="' && value.slice(-1) === '"')
      return value.slice(2, -1).replace(/""/g, '"');

    return value;
  }

  function findChart(sheet, name) {
    var charts = sheet.GetAllCharts() || [];
    for (var i = 0; i < charts.length; i++) {
      if (String(invoke(charts[i], 'GetName')) === String(name))
        return charts[i];
    }
    return null;
  }

  try {
    var sh = Api.GetSheet(spec.sheet);
    if (!sh) return { ok: false, error: 'sheet not found' };

    var c = findChart(sh, spec.name);
    if (!c) return { ok: false, error: 'main chart not found' };

    var pie = findChart(sh, spec.pieName);
    if (!pie) return { ok: false, error: 'pie chart not found' };

    var scatter = findChart(sh, spec.scatterName);
    if (!scatter) return { ok: false, error: 'scatter chart not found' };

    var result = {
      ok: false,
      capabilities: {}
    };

    /* ------------------------------------------------------------
       M5.3 — Legend position regression
       ------------------------------------------------------------ */
    var legendSteps = [];
    var legendPositions = ['left', 'top', 'right', 'bottom', 'none'];

    for (var li = 0; li < legendPositions.length; li++) {
      var lp = legendPositions[li];
      var ls = invoke(c, 'SetLegendPos', [lp]);
      var la = invoke(c, 'GetLegendPos');
      legendSteps.push({
        requested: lp,
        setResult: ls,
        actual: la,
        pass: la === lp
      });
    }

    result.capabilities.legendPosition = {
      available:
        typeof c.SetLegendPos === 'function' &&
        typeof c.GetLegendPos === 'function',
      steps: legendSteps,
      pass: legendSteps.every(function(x) { return x.pass; })
    };

    /* ------------------------------------------------------------
       M5.3 — Horizontal / vertical axis titles
       ------------------------------------------------------------ */
    var horTitle = 'EURO horizontal axis ' + spec.token;
    var verTitle = 'EURO vertical axis ' + spec.token;

    var hSet = invoke(c, 'SetHorAxisTitle', [horTitle]);
    var hActual = invoke(c, 'GetHorAxisTitle');

    var vSet = invoke(c, 'SetVerAxisTitle', [verTitle]);
    var vActual = invoke(c, 'GetVerAxisTitle');

    result.capabilities.axisTitles = {
      available:
        typeof c.GetHorAxisTitle === 'function' &&
        typeof c.GetVerAxisTitle === 'function',
      horizontal: {
        requested: horTitle,
        setResult: hSet,
        actual: hActual,
        normalizedActual: trimParagraphEnd(hActual),
        pass: trimParagraphEnd(hActual) === horTitle
      },
      vertical: {
        requested: verTitle,
        setResult: vSet,
        actual: vActual,
        normalizedActual: trimParagraphEnd(vActual),
        pass: trimParagraphEnd(vActual) === verTitle
      }
    };

    result.capabilities.axisTitles.pass =
      result.capabilities.axisTitles.horizontal.pass &&
      result.capabilities.axisTitles.vertical.pass;

    /* ------------------------------------------------------------
       M5.3 — Data labels on normal chart
       ------------------------------------------------------------ */
    var normalLabelsExpected = {
      showSerName: true,
      showCatName: true,
      showVal: true,
      showPercent: false
    };

    var normalLabelsSet =
      invoke(c, 'SetShowDataLabels', [true, true, true, false]);

    var normalLabelsActual = invoke(c, 'GetDataLabels');

    var normalLabelsPass =
      same(normalLabelsActual, normalLabelsExpected);

    /* Pie chart: percent=true must survive semantically. */
    var pieLabelsExpected = {
      showSerName: false,
      showCatName: true,
      showVal: false,
      showPercent: true
    };

    var pieLabelsSet =
      invoke(pie, 'SetShowDataLabels', [false, true, false, true]);

    var pieLabelsActual = invoke(pie, 'GetDataLabels');

    var pieLabelsPass =
      same(pieLabelsActual, pieLabelsExpected);

    result.capabilities.dataLabels = {
      available:
        typeof c.GetDataLabels === 'function' &&
        typeof pie.GetDataLabels === 'function',
      normal: {
        requested: normalLabelsExpected,
        setResult: normalLabelsSet,
        actual: normalLabelsActual,
        pass: normalLabelsPass
      },
      pie: {
        requested: pieLabelsExpected,
        setResult: pieLabelsSet,
        actual: pieLabelsActual,
        pass: pieLabelsPass
      },
      pass: normalLabelsPass && pieLabelsPass
    };

    /* ------------------------------------------------------------
       M5.3 — Chart style
       Test multiple style inputs.
       ------------------------------------------------------------ */
    var styleSteps = [];
    var styleInputs = [0, 1, 2];

    for (var si = 0; si < styleInputs.length; si++) {
      var styleId = styleInputs[si];
      var styleSet = invoke(c, 'ApplyChartStyle', [styleId]);
      var styleActual = invoke(c, 'GetChartStyle');

      styleSteps.push({
        requested: styleId,
        setResult: styleSet,
        actual: styleActual,
        pass: styleSet === true && styleActual === styleId
      });
    }

    result.capabilities.chartStyle = {
      available:
        typeof c.ApplyChartStyle === 'function' &&
        typeof c.GetChartStyle === 'function',
      steps: styleSteps,
      pass: styleSteps.every(function(x) { return x.pass; })
    };

    /* ------------------------------------------------------------
       M5.4 — Position regression
       ------------------------------------------------------------ */
    var positionSteps = [];
    var positions = [
      { fromCol: 2, colOffset: 360000, fromRow: 3, rowOffset: 720000 },
      { fromCol: 7, colOffset: 180000, fromRow: 9, rowOffset: 540000 },
      { fromCol: 0, colOffset: 0, fromRow: 0, rowOffset: 0 }
    ];

    for (var pi = 0; pi < positions.length; pi++) {
      var pos = positions[pi];

      var posSet = invoke(c, 'SetPosition', [
        pos.fromCol,
        pos.colOffset,
        pos.fromRow,
        pos.rowOffset
      ]);

      var posActual = invoke(c, 'GetPosition');

      positionSteps.push({
        requested: pos,
        setResult: posSet,
        actual: posActual,
        pass: same(posActual, pos)
      });
    }

    result.capabilities.position = {
      available:
        typeof c.SetPosition === 'function' &&
        typeof c.GetPosition === 'function',
      steps: positionSteps,
      pass: positionSteps.every(function(x) { return x.pass; })
    };

    /* ------------------------------------------------------------
       M5.4 — Series semantic getters

       Bar chart:
         PUBLIC SetSeriaName / SetSeriaValues / SetCatFormula
         PUBLIC ApiChartSeries getters.

       Scatter chart:
         PUBLIC SetSeriaXValues
         PUBLIC ApiChartSeries.GetXValues.

       No internal runtime state is used for acceptance.
       ------------------------------------------------------------ */

    var seriesBefore = invoke(c, 'GetAllSeries');

    if (!Array.isArray(seriesBefore) || !seriesBefore.length) {
      result.capabilities.seriesData = {
        available: false,
        pass: false,
        error: 'bar GetAllSeries returned no series'
      };
    } else {
      var seriesIndex = 0;
      var seriesName = 'EURO Series ' + spec.token;
      var valuesRange = spec.valuesRange;
      var catRange = spec.catRange;

      var setName =
        invoke(c, 'SetSeriaName', [seriesName, seriesIndex]);

      var setValues =
        invoke(c, 'SetSeriaValues', [valuesRange, seriesIndex]);

      var setCat =
        invoke(c, 'SetCatFormula', [catRange]);

      var seriesAfter = invoke(c, 'GetAllSeries');
      var s0 =
        Array.isArray(seriesAfter) && seriesAfter.length
          ? seriesAfter[0]
          : null;

      var gotName = invoke(s0, 'GetName');
      var gotValues = invoke(s0, 'GetValues');
      var gotCat = invoke(s0, 'GetCatFormula');

      var normalizedName = decodeSeriesName(gotName);
      var normalizedValues = normalizeFormula(gotValues);
      var normalizedCat = normalizeFormula(gotCat);

      var namePass = normalizedName === seriesName;
      var valuesPass = normalizedValues === normalizeFormula(valuesRange);
      var catPass = normalizedCat === normalizeFormula(catRange);

      var scatterSeriesBefore = invoke(scatter, 'GetAllSeries');
      var scatterS0 =
        Array.isArray(scatterSeriesBefore) && scatterSeriesBefore.length
          ? scatterSeriesBefore[0]
          : null;

      var scatterSetX =
        typeof scatter.SetSeriaXValues === 'function'
          ? invoke(scatter, 'SetSeriaXValues', [spec.xValuesRange, 0])
          : { __error: 'public method unavailable: SetSeriaXValues' };

      var scatterSeriesAfter = invoke(scatter, 'GetAllSeries');
      var scatterS0After =
        Array.isArray(scatterSeriesAfter) && scatterSeriesAfter.length
          ? scatterSeriesAfter[0]
          : null;

      var scatterGotX = invoke(scatterS0After, 'GetXValues');
      var normalizedScatterX = normalizeFormula(scatterGotX);
      var expectedScatterX = normalizeFormula(spec.xValuesRange);

      var scatterXPass =
        !!scatterS0 &&
        !!scatterS0After &&
        typeof scatter.SetSeriaXValues === 'function' &&
        typeof scatterS0After.GetXValues === 'function' &&
        normalizedScatterX === expectedScatterX;

      result.capabilities.seriesData = {
        available:
          !!s0 &&
          !!scatterS0After &&
          typeof s0.GetName === 'function' &&
          typeof s0.GetValues === 'function' &&
          typeof s0.GetCatFormula === 'function' &&
          typeof scatter.SetSeriaXValues === 'function' &&
          typeof scatterS0After.GetXValues === 'function',

        name: {
          requested: seriesName,
          setResult: setName,
          actual: gotName,
          normalizedActual: normalizedName,
          pass: namePass
        },

        values: {
          requested: valuesRange,
          setResult: setValues,
          actual: gotValues,
          normalizedActual: normalizedValues,
          pass: valuesPass
        },

        category: {
          requested: catRange,
          setResult: setCat,
          actual: gotCat,
          normalizedActual: normalizedCat,
          pass: catPass
        },

        scatterXValues: {
          requested: spec.xValuesRange,
          setResult: scatterSetX,
          actual: scatterGotX,
          normalizedActual: normalizedScatterX,
          pass: scatterXPass
        },

        pass:
          namePass &&
          valuesPass &&
          catPass &&
          scatterXPass
      };
    }

    var keys = Object.keys(result.capabilities);
    result.ok = keys.every(function(k) {
      return result.capabilities[k].pass === true;
    });

    return result;

  } catch (e) {
    return {
      ok: false,
      error: String(e && e.message ? e.message : e)
    };
  }
}

async function call(frame, apiHely, spec, timeout = 20000) {
  const body =
    `return (${publicProbe.toString()})(${JSON.stringify(spec)});`;

  return frame.evaluate(
    ({ u, body, timeout }) =>
      new Promise(resolve => {
        const e =
          u === 'window.editor'
            ? window.editor
            : (window.Asc || {}).editor;

        let done = false;

        const finish = value => {
          if (!done) {
            done = true;
            resolve(value);
          }
        };

        try {
          /* EXACT accepted transport: 3 arguments. */
          e.callCommand(
            new Function(body),
            false,
            value => finish(value)
          );
        } catch (err) {
          finish({
            ok: false,
            error: String(err && err.message ? err.message : err)
          });
        }

        setTimeout(
          () => finish({ ok: false, error: 'timeout' }),
          timeout
        );
      }),
    { u: apiHely, body, timeout }
  );
}

async function editorCommand(frame, apiHely, fn, spec, timeout = 20000) {
  const body =
    `return (${fn.toString()})(${JSON.stringify(spec)});`;

  return frame.evaluate(
    ({ u, body, timeout }) =>
      new Promise(resolve => {
        const e =
          u === 'window.editor'
            ? window.editor
            : (window.Asc || {}).editor;

        let done = false;

        const finish = value => {
          if (!done) {
            done = true;
            resolve(value);
          }
        };

        try {
          e.callCommand(
            new Function(body),
            false,
            value => finish(value)
          );
        } catch (err) {
          finish({
            ok: false,
            error: String(err && err.message ? err.message : err)
          });
        }

        setTimeout(
          () => finish({ ok: false, error: 'timeout' }),
          timeout
        );
      }),
    { u: apiHely, body, timeout }
  );
}

function seedCommand(spec) {
  try {
    var sh = Api.GetSheet(spec.sheet);
    if (!sh) return { ok: false, error: 'sheet not found' };

    sh.GetRange(spec.range).SetValue(spec.values);

    return {
      ok: true,
      range: spec.range
    };
  } catch (e) {
    return {
      ok: false,
      error: String(e && e.message ? e.message : e)
    };
  }
}

function createChartCommand(spec) {
  try {
    var sh = Api.GetSheet(spec.sheet);
    if (!sh) return { ok: false, error: 'sheet not found' };

    var chart = sh.AddChart(
      spec.range,
      spec.chartType,
      spec.style,
      spec.width,
      spec.height,
      spec.fromCol,
      spec.colOffset,
      spec.fromRow,
      spec.rowOffset
    );

    if (!chart)
      return { ok: false, error: 'AddChart returned no chart' };

    chart.SetName(spec.name);

    if (spec.title)
      chart.SetTitle(spec.title);

    var charts = sh.GetAllCharts() || [];
    var found = false;

    for (var i = 0; i < charts.length; i++) {
      if (
        typeof charts[i].GetName === 'function' &&
        String(charts[i].GetName()) === String(spec.name)
      ) {
        found = true;
        break;
      }
    }

    return {
      ok: found,
      name: spec.name,
      count: charts.length
    };
  } catch (e) {
    return {
      ok: false,
      error: String(e && e.message ? e.message : e)
    };
  }
}

function deleteChartCommand(spec) {
  try {
    var sh = Api.GetSheet(spec.sheet);
    if (!sh) return { ok: false, error: 'sheet not found' };

    var charts = sh.GetAllCharts() || [];
    var target = null;

    for (var i = 0; i < charts.length; i++) {
      if (
        typeof charts[i].GetName === 'function' &&
        String(charts[i].GetName()) === String(spec.name)
      ) {
        target = charts[i];
        break;
      }
    }

    if (!target)
      return { ok: true, alreadyAbsent: true };

    /*
     * Public chart deletion is intentionally not faked here.
     * Cleanup is performed by the proven UI Delete path outside
     * this command when needed.
     */
    return {
      ok: false,
      needsUiDelete: true
    };

  } catch (e) {
    return {
      ok: false,
      error: String(e && e.message ? e.message : e)
    };
  }
}

async function deleteChartViaUi(frame, name) {
  const selected = await frame.evaluate(
    name =>
      new Promise(resolve => {
        const e = (window.Asc || {}).editor;
        let done = false;

        const finish = v => {
          if (!done) {
            done = true;
            resolve(v);
          }
        };

        try {
          e.callCommand(
            new Function(`
              try {
                var sh=Api.GetActiveSheet();
                var charts=sh.GetAllCharts()||[];
                for(var i=0;i<charts.length;i++){
                  if(typeof charts[i].GetName==="function" &&
                     String(charts[i].GetName())===${JSON.stringify(name)}){
                    charts[i].Select();
                    return {ok:true};
                  }
                }
                return {ok:false,error:"chart not found"};
              } catch(e) {
                return {ok:false,error:String(e&&e.message?e.message:e)};
              }
            `),
            false,
            v => finish(v)
          );
        } catch (err) {
          finish({
            ok: false,
            error: String(err && err.message ? err.message : err)
          });
        }

        setTimeout(
          () => finish({ ok: false, error: 'timeout' }),
          10000
        );
      }),
    name
  );

  if (!selected || !selected.ok)
    return { ok: false, selected };

  await frame.locator('body').focus().catch(() => {});
  await frame.page().keyboard.press('Delete');
  await frame.page().waitForTimeout(1200);

  const verify = await frame.evaluate(
    name =>
      new Promise(resolve => {
        const e = (window.Asc || {}).editor;

        try {
          e.callCommand(
            new Function(`
              var sh=Api.GetActiveSheet();
              var charts=sh.GetAllCharts()||[];
              for(var i=0;i<charts.length;i++){
                if(typeof charts[i].GetName==="function" &&
                   String(charts[i].GetName())===${JSON.stringify(name)})
                  return {ok:false,stillPresent:true};
              }
              return {ok:true};
            `),
            false,
            resolve
          );
        } catch (err) {
          resolve({
            ok: false,
            error: String(err && err.message ? err.message : err)
          });
        }
      }),
    name
  );

  return verify;
}

async function main() {
  const base = process.env.EURO_NC_BASE_URL;
  const fileId = process.env.EURO_NC_FILE_ID;
  const user = process.env.EURO_NC_USER;
  const password = process.env.EURO_NC_PASSWORD;

  if (!base || !fileId || !user || !password)
    throw new Error('required EURO_NC_* env missing');

  const sheet = 'Sheet1';
  const token = Date.now().toString(36);

  /*
   * Keep fixture away from normal workbook content.
   */
  const range = 'XFA20:XFC23';

  const data = [
    ['Quarter', 'Revenue', 'Cost'],
    ['Q1', 100, 70],
    ['Q2', 140, 90],
    ['Q3', 125, 85]
  ];

  const mainName = 'M53_M54_COMBINED_' + token;
  const pieName = 'M53_LABEL_PIE_' + token;
  const scatterName = 'M54_SERIES_SCATTER_' + token;

  const scatterRange = 'XEY20:XEZ23';
  const scatterData = [
    ['X', 'Y'],
    [10, 100],
    [20, 140],
    [30, 125]
  ];

  const { chromium } = loadPlaywright();
  const browser = await chromium.launch();

  let cleanupMain = null;
  let cleanupPie = null;
  let cleanupScatter = null;

  try {
    const ctx = await browser.newContext({
      viewport: { width: 1400, height: 900 }
    });

    const page = await ctx.newPage();

    await page.goto(`${base}/login`, {
      waitUntil: 'domcontentloaded',
      timeout: 60000
    });

    await page.fill('#user', user);
    await page.fill('#password', password);

    await Promise.all([
      page.waitForNavigation({
        waitUntil: 'domcontentloaded',
        timeout: 60000
      }).catch(() => null),
      page.click('button[type=submit], input[type=submit]')
    ]);

    await page.waitForTimeout(2500);

    if (/\/login/.test(page.url()))
      throw new Error('login failed');

    await page.goto(
      `${base}/index.php/apps/eurooffice/${fileId}`,
      {
        waitUntil: 'domcontentloaded',
        timeout: 60000
      }
    );

    await page.waitForTimeout(22000);

    const frame =
      page.frames().find(f => /spreadsheeteditor/.test(f.url()));

    if (!frame)
      throw new Error('spreadsheeteditor frame missing');

    const apiHely = await frame.evaluate(() =>
      (window.Asc || {}).editor &&
      typeof window.Asc.editor.callCommand === 'function'
        ? 'window.Asc.editor'
        : null
    );

    if (!apiHely)
      throw new Error('callCommand unavailable');

    const seed = await writeBulkInFrame(
      frame,
      apiHely,
      {
        sheet,
        range,
        values: data
      }
    );

    const scatterSeed = seed.ok
      ? await writeBulkInFrame(
          frame,
          apiHely,
          {
            sheet,
            range: scatterRange,
            values: scatterData
          }
        )
      : null;

    const createdMain = seed.ok && scatterSeed && scatterSeed.ok
      ? await runChartInFrame(
          frame,
          apiHely,
          {
            type: 'chart.create',
            sheet,
            range,
            chartType: 'bar',
            style: 2,
            width: 4000000,
            height: 2400000,
            name: mainName,
            title: 'M5.3/M5.4 combined acceptance'
          }
        )
      : null;

    const createdPie =
      seed.ok && createdMain && createdMain.ok
        ? await runChartInFrame(
            frame,
            apiHely,
            {
              type: 'chart.create',
              sheet,
              range: 'XFA20:XFB23',
              chartType: 'pie',
              style: 2,
              width: 3200000,
              height: 2400000,
              name: pieName,
              title: 'M5.3 labels acceptance'
            }
          )
        : null;

    const createdScatter =
      createdPie && createdPie.ok
        ? await runChartInFrame(
            frame,
            apiHely,
            {
              type: 'chart.create',
              sheet,
              range: scatterRange,
              chartType: 'scatter',
              style: 2,
              width: 3200000,
              height: 2400000,
              name: scatterName,
              title: 'M5.4 scatter X-values acceptance'
            }
          )
        : null;

    const probe =
      createdScatter && createdScatter.ok
        ? await call(
            frame,
            apiHely,
            {
              sheet,
              name: mainName,
              pieName,
              scatterName,
              token,

              /*
               * Public getters may prefix formulas with exactly one "=".
               * Acceptance normalizes that representation only.
               */
              valuesRange: 'Sheet1!$XFB$21:$XFB$23',
              xValuesRange: 'Sheet1!$XEY$21:$XEY$23',
              catRange: 'Sheet1!$XFA$21:$XFA$23'
            }
          )
        : null;

    if (createdScatter && createdScatter.ok)
      cleanupScatter = await runChartInFrame(
        frame,
        apiHely,
        { type: 'chart.delete', sheet, name: scatterName }
      );

    if (createdPie && createdPie.ok)
      cleanupPie = await runChartInFrame(
        frame,
        apiHely,
        { type: 'chart.delete', sheet, name: pieName }
      );

    if (createdMain && createdMain.ok)
      cleanupMain = await runChartInFrame(
        frame,
        apiHely,
        { type: 'chart.delete', sheet, name: mainName }
      );

    const cleanupPass =
      (!createdScatter || !createdScatter.ok || cleanupScatter?.ok) &&
      (!createdPie || !createdPie.ok || cleanupPie?.ok) &&
      (!createdMain || !createdMain.ok || cleanupMain?.ok);

    const outcome =
      seed.ok &&
      scatterSeed &&
      scatterSeed.ok &&
      createdMain &&
      createdMain.ok &&
      createdPie &&
      createdPie.ok &&
      createdScatter &&
      createdScatter.ok &&
      probe &&
      probe.ok &&
      cleanupPass
        ? 'PASS'
        : 'UNKNOWN';

    console.log(JSON.stringify({
      milestone: 'M5.3+M5.4',
      kind: 'combined-public-getter-acceptance',
      source: 'live-coedit-editor',
      outcome,
      seed,
      scatterSeed,
      createdMain,
      createdPie,
      createdScatter,
      probe,
      cleanup: {
        scatter: cleanupScatter,
        pie: cleanupPie,
        main: cleanupMain,
        pass: cleanupPass
      },
      editor: 'spreadsheeteditor',
      apiHely
    }, null, 2));

    if (outcome !== 'PASS')
      process.exitCode = 2;

  } finally {
    await browser.close();
  }
}

main().catch(e => {
  console.error(JSON.stringify({
    milestone: 'M5.3+M5.4',
    outcome: 'launcher-error',
    error: String(e && e.message ? e.message : e)
  }, null, 2));

  process.exitCode = 1;
});
