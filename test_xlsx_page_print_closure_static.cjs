'use strict'
const {execFileSync}=require('child_process')
const staticTests=[
'test_xlsx_first_page_number_mcp_schema_static.cjs',
'test_xlsx_print_page_intents_schema_static.cjs',
'test_xlsx_print_setup_batch_static.cjs',
'test_xlsx_print_titles_batch_static.cjs',
'test_xlsx_print_area_batch_static.cjs',
'test_xlsx_header_footer_batch_static.cjs',
'test_xlsx_page_break_batch_static.cjs'
]
for(const t of staticTests){console.log('===',t,'===');execFileSync(process.execPath,[t],{stdio:'inherit',env:process.env})}
console.log('XLSX PAGE PRINT CLOSURE STATIC: PASS')
