'use strict'

const {runJob}=require('./runner.cjs')

function buildProbeScript(){
 return [
  'builder.OpenFile("__DOC_URL__", "docx");',
  'builder.CloseFile();',
  'builder.CreateFile("xlsx");',
  'var ws = Api.GetActiveSheet();',
  'ws.SetName("M51_PROBE");',
  'ws.GetRange("A1").SetValue("Quarter");',
  'ws.GetRange("B1").SetValue("Revenue");',
  'ws.GetRange("A2").SetValue("Q1");',
  'ws.GetRange("B2").SetValue(100);',
  'ws.GetRange("A3").SetValue("Q2");',
  'ws.GetRange("B3").SetValue(140);',
  'ws.GetRange("A4").SetValue("Q3");',
  'ws.GetRange("B4").SetValue(125);',
  'var c = ws.AddChart("A1:B4", false, "bar", 2, 3600000, 2200000, 4, 0, 0, 0);',
  'var before = ws.GetAllCharts ? (ws.GetAllCharts() || []).length : -1;',
  'var hasDelete = !!(c && typeof c.Delete === "function");',
  'var deleted = false;',
  'if (hasDelete) deleted = !!c.Delete();',
  'var after = ws.GetAllCharts ? (ws.GetAllCharts() || []).length : -1;',
  'var pass = !!c && hasDelete && deleted && before === 1 && after === 0;',
  'ws.SetName(pass ? "M51_DELETE_PASS_1_0" : ("M51_DELETE_FAIL_" + (hasDelete ? "HAS" : "NO") + "_" + before + "_" + after));',
  'builder.SaveFile("xlsx", "m51-delete-probe.xlsx");',
  'builder.CloseFile();'
 ].join('\n')
}

function verify(job){
 const text=String(job&&job.documentText||'')
 if(/M51_DELETE_PASS_1_0/.test(text))return {status:'PASS',evidence:'M51_DELETE_PASS_1_0'}
 const fail=(text.match(/M51_DELETE_FAIL_[A-Z]+_-?\d+_-?\d+/)||[])[0]
 if(fail)return {status:'FAIL',evidence:fail}
 return {status:'UNKNOWN',evidence:null,documentText:text}
}

async function main(){
 const job=await runJob({script:buildProbeScript(),returnDoc:true,traceId:`m51-docbuilder-delete-cap-${Date.now()}`})
 const verification=verify(job)
 const result={milestone:'M5.1 DocBuilder chart Delete capability',source:'docbuilder',outcome:verification.status,verification,job:{ok:job.ok,outcome:job.outcome,kind:job.kind,serverFetches:job.serverFetches,savedBytes:job.savedBytes,hasReturnedDocument:!!job.savedBase64}}
 console.log(JSON.stringify(result,null,2))
 if(result.outcome!=='PASS')process.exitCode=2
}

if(require.main===module)main().catch(err=>{console.error(JSON.stringify({milestone:'M5.1 DocBuilder chart Delete capability',source:'docbuilder',outcome:'ERROR',error:String(err&&err.message?err.message:err)},null,2));process.exitCode=1})
module.exports={buildProbeScript,verify}
