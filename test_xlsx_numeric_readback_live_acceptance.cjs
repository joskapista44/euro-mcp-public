'use strict'
// Read-only acceptance of the 36 known numeric constants in showcase 768SVZ.
async function run() {
  const caller=require('./coedit.cjs').detectCallerId()
  if(!caller.ok||caller.id!=='elliot')throw Error('allowlisted elliot required')
  const {getSecret}=require('/home/user/marveen/dist/web/vault.js')
  const pass=getSecret('Elliot_nc_pass','xlsx-numeric-type-readonly-probe')
  if(!pass)throw Error('vault credential missing')
  process.env.EURO_PLAYWRIGHT_PATH ||= '/home/user/marveen/node_modules/playwright'
  const {openMinimalXlsxSession}=require('./xlsx-minimal-editor-session.cjs')
  const session=await openMinimalXlsxSession({url:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',user:'elliot',pass,fileId:process.env.EURO_XLSX_FILE_ID||'1236770',timeoutMs:30000})
  let result
  try {
    const {readRangeInFrame}=require('./range-reader.cjs')
    const task=require('./xlsx-visual-showcase-contract.cjs').buildShowcaseTask('768SVZ')
    const expected=task.operations.find(op=>op.intent==='write_range'&&op.sheet===task.names.data)
    const read=await readRangeInFrame(session.frame,session.apiWhere,{sheet:task.names.data,range:'A1:H13'})
    const {cellMatches}=require('./verification-contract.cjs')
    const mismatches=[];let checked=0
    if(!read.ok)throw Error('live read failed')
    for(let row=1;row<13;row++)for(const col of [3,4,7]){
      const cell=read.cells?.[row]?.[col],value=expected.values[row][col];checked++
      if(!cellMatches(cell,{value}))mismatches.push({address:cell?.address,expected:value,actual:cell})
    }
    // Read-only reference controls: headers must not satisfy a numeric expectation.
    const headersRejectNumbers=[3,4,7].every(col=>!cellMatches(read.cells?.[0]?.[col],{value:0}))
    result={ok:checked===36&&mismatches.length===0&&headersRejectNumbers,checked,mismatches,headersRejectNumbers}
  } finally { await session.close() }
  console.error('NUMERIC READBACK',JSON.stringify({...result,oneEditorSession:true,writes:session.writes,closed:session.closed}))
  if(!result?.ok||session.writes!==0||!session.closed)process.exitCode=1
}
if(require.main===module)run().catch(error=>{console.error('Numeric readback acceptance failed:',error.message);process.exitCode=1})
