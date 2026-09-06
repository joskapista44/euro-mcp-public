'use strict'

const zlib=require('zlib')
const {runJob}=require('./runner.cjs')
const {xlsxMag}=require('./euro-magok.cjs')

// The deployed DocBuilder route requires a type-correct OpenFile input. Reuse the same minimal
// XLSX seed that the historical, package-verified XLSX DocBuilder route used; CreateFile("xlsx")
// is deliberately not used here because that path is known to kill this deployment's job.
function buildProbeScript(){
 return [
  'builder.OpenFile("__DOC_URL__");',
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
  // Exact XLSX AddChart call shape recovered and package-verified on this DocBuilder deployment.
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

// Tiny ZIP reader for the one STORED/DEFLATED XML part we need. Verification is read-only: the
// returned XLSX bytes are never rewritten locally. This avoids relying on box-helper's deliberately
// lossy `documentText` field for spreadsheets.
function workbookXmlFromXlsx(raw){
 const buf=Buffer.isBuffer(raw)?raw:Buffer.from(raw)
 let off=0
 while(off+30<=buf.length && buf.readUInt32LE(off)===0x04034b50){
  const method=buf.readUInt16LE(off+8)
  const csize=buf.readUInt32LE(off+18)
  const nlen=buf.readUInt16LE(off+26)
  const xlen=buf.readUInt16LE(off+28)
  const name=buf.subarray(off+30,off+30+nlen).toString('utf8')
  const start=off+30+nlen+xlen
  const packed=buf.subarray(start,start+csize)
  if(name==='xl/workbook.xml'){
   if(method===0)return packed.toString('utf8')
   if(method===8)return zlib.inflateRawSync(packed).toString('utf8')
   throw new Error(`unsupported XLSX compression method ${method}`)
  }
  off=start+csize
 }
 throw new Error('xl/workbook.xml not found in returned XLSX')
}

function verifyReturnedXlsx(savedBase64){
 if(!savedBase64)return {status:'UNKNOWN',evidence:null,reason:'DocBuilder returned no XLSX bytes'}
 let xml
 try{xml=workbookXmlFromXlsx(Buffer.from(savedBase64,'base64'))}
 catch(err){return {status:'UNKNOWN',evidence:null,reason:`returned XLSX could not be inspected: ${err.message}`}}
 if(/name="M51_DELETE_PASS_1_0"/.test(xml))return {status:'PASS',evidence:'M51_DELETE_PASS_1_0'}
 const m=xml.match(/name="(M51_DELETE_FAIL_(?:HAS|NO)_-?\d+_-?\d+)"/)
 if(m)return {status:'FAIL',evidence:m[1]}
 return {status:'UNKNOWN',evidence:null,reason:'probe worksheet marker missing from returned workbook.xml'}
}

async function main(){
 const job=await runJob({
  script:buildProbeScript(),
  documentBase64:xlsxMag(['M51_SEED']).toString('base64'),
  returnDoc:true,
  traceId:`m51-docbuilder-delete-cap-${Date.now()}`
 })
 const verification=verifyReturnedXlsx(job.savedBase64)
 const result={milestone:'M5.1 DocBuilder chart Delete capability',source:'docbuilder',outcome:verification.status,verification,job:{ok:job.ok,outcome:job.outcome,detail:job.detail,dsError:job.dsError,kind:job.kind,serverFetches:job.serverFetches,savedBytes:job.savedBytes,hasReturnedDocument:!!job.savedBase64,exitCode:job.exitCode,stderr:job.stderr}}
 console.log(JSON.stringify(result,null,2))
 if(result.outcome!=='PASS')process.exitCode=2
}

if(require.main===module)main().catch(err=>{console.error(JSON.stringify({milestone:'M5.1 DocBuilder chart Delete capability',source:'docbuilder',outcome:'ERROR',error:String(err&&err.message?err.message:err)},null,2));process.exitCode=1})
module.exports={buildProbeScript,workbookXmlFromXlsx,verifyReturnedXlsx}
