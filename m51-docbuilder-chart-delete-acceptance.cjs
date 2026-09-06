'use strict'
const fs=require('fs')
const {runJob}=require('./runner.cjs')
const {buildChartDeleteScript,parseChartDeleteMarker,verifyChartDelete}=require('./docbuilder-charts.cjs')

function markerFromJob(job){
 const candidates=[job.documentText,job.text,job.outputText,job.markerText].filter(v=>typeof v==='string')
 for(const s of candidates){const m=parseChartDeleteMarker(s);if(m)return m}
 return null
}

async function main(){
 const input=process.env.EURO_M51_INPUT_XLSX,sheet=process.env.EURO_M51_SHEET||'Sheet1',name=process.env.EURO_M51_CHART_NAME
 if(!input)throw new Error('EURO_M51_INPUT_XLSX is required')
 if(!name)throw new Error('EURO_M51_CHART_NAME is required')
 const documentBase64=fs.readFileSync(input).toString('base64')
 const script=buildChartDeleteScript({sheet,name})
 const job=await runJob({script,documentBase64,returnDoc:true,traceId:`m51-chart-delete-${Date.now()}`})
 const found=markerFromJob(job)
 const verification=verifyChartDelete({name,marker:found})
 const returned=job.documentBase64||job.outputBase64||job.returnedDocumentBase64||null
 const result={milestone:'M5.1 DocBuilder chart.delete acceptance',source:'docbuilder',outcome:verification.status==='PASS'&&job.ok?'PASS':'FAIL',sheet,name,job:{ok:job.ok,outcome:job.outcome,kind:job.kind,serverFetches:job.serverFetches,hasReturnedDocument:!!returned},marker:found,verification}
 console.log(JSON.stringify(result,null,2))
 if(result.outcome!=='PASS')process.exitCode=2
}
if(require.main===module)main().catch(e=>{console.error(JSON.stringify({milestone:'M5.1 DocBuilder chart.delete acceptance',source:'docbuilder',outcome:'ERROR',error:String(e&&e.message?e.message:e)},null,2));process.exitCode=1})
module.exports={markerFromJob}
