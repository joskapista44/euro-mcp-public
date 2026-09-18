# Excel Power User visual showcase — agent task

## Assignment

Build a polished, business-readable sales performance workbook in the supplied existing XLSX file while the operator watches it update in ONLYOFFICE co-editing mode.

Use the EURO-MCP `office_xlsx_batch` tool. The workbook must demonstrate that you do not merely dispatch mutations: you must request final range readbacks in the same batch session and report the values you actually observed.

## Inputs

- `file_id`: supplied by the operator.
- `run_id`: choose one six-character uppercase alphanumeric identifier and keep it unchanged throughout the task.
- Derived worksheet names:
  - `MCP_Dash_<run_id>`
  - `MCP_Plan_<run_id>`
  - `MCP_Data_<run_id>`
  - `MCP_PivotView_<run_id>`
- Derived object names:
  - defined name: `MCP_Sales_<run_id>`
  - pivot: `MCP_Pivot_<run_id>`
  - charts: `MCP_Trend_<run_id>` and `MCP_Region_<run_id>`

If `file_id` is missing, ask only for that value. Before calling the tool, tell the operator the four worksheet names, ask them to open the workbook and reply `INDULHAT`, then wait. Do not use a timed delay.

## Execution contract

- Make exactly one mutating MCP call: one `office_xlsx_batch` request containing the complete final-state plan.
- Do not invoke Chrome, Playwright, DocBuilder, downloaded-XLSX or OOXML tooling yourself.
- Do not use standalone write primitives and do not split construction into multiple MCP calls.
- Begin with exactly the four `create_sheet` operations (dashboard, plan, data, pivot view). Only after all four sheets exist, put the three corresponding `write_range` operations. All remaining operations follow them. Do not interleave sheet creation and range writing: dashboard and plan formulas reference sheets created later in the request.
- Do not alter or delete any pre-existing worksheet or object. Only use the run-specific names above.
- Leave the finished workbook in place. Do not clean it up.
- Include these same-session final readbacks in the batch request:
  - dashboard `A1:F16`
  - plan `A1:E9`
  - data `A1:H13`

## Workbook design

Use Arial throughout the authored presentation areas. Use dark navy `[31, 56, 100]` for title bands, blue `[47, 117, 181]` for table headers, pale blue `[221, 235, 247]` for label blocks, dark text `[31, 41, 55]`, white title/header text, and pale green `[226, 239, 218]` for positive conditional formatting. Use restrained formatting: no decorative icons, no 3-D charts and no dense full-cell borders.

### 1. Data worksheet

Write this table to `A1:H13`. Store numbers as numeric values. Columns F and G must be formulas, not pasted results.

| Month | Region | Product | Units | Unit price | Revenue | Status | Review score |
| --- | --- | --- | ---: | ---: | ---: | --- | ---: |
| Jan | North | Core | 120 | 190 | `=D2*E2` | `=IF(F2>=20000,"On track","At risk")` | 5 |
| Feb | North | Plus | 75 | 320 | `=D3*E3` | `=IF(F3>=20000,"On track","At risk")` | 4 |
| Mar | North | Core | 135 | 195 | `=D4*E4` | `=IF(F4>=20000,"On track","At risk")` | 5 |
| Apr | North | Plus | 82 | 330 | `=D5*E5` | `=IF(F5>=20000,"On track","At risk")` | 4 |
| May | North | Core | 150 | 200 | `=D6*E6` | `=IF(F6>=20000,"On track","At risk")` | 5 |
| Jun | North | Plus | 95 | 340 | `=D7*E7` | `=IF(F7>=20000,"On track","At risk")` | 5 |
| Jan | South | Core | 90 | 190 | `=D8*E8` | `=IF(F8>=20000,"On track","At risk")` | 3 |
| Feb | South | Plus | 60 | 320 | `=D9*E9` | `=IF(F9>=20000,"On track","At risk")` | 3 |
| Mar | South | Core | 100 | 195 | `=D10*E10` | `=IF(F10>=20000,"On track","At risk")` | 4 |
| Apr | South | Plus | 68 | 330 | `=D11*E11` | `=IF(F11>=20000,"On track","At risk")` | 4 |
| May | South | Core | 110 | 200 | `=D12*E12` | `=IF(F12>=20000,"On track","At risk")` | 4 |
| Jun | South | Plus | 72 | 340 | `=D13*E13` | `=IF(F13>=20000,"On track","At risk")` | 4 |

Apply these features:

- Header `A1:H1`: blue fill, white bold Arial text, centered vertically and horizontally.
- Currency number format `"$"#,##0` on `E2:F13`; integer format `#,##0` on `D2:D13`; center `G2:H13`.
- Column widths: A 12, B:C 16, D:H 14. Row 1 height 22.
- Freeze at `A2`.
- Sort `A1:H13` ascending by `B1:B13`, with headers. The supplied data is already in that final order, so this must verify as a safe no-op rather than rewrite rows.
- Apply an active filter to `A1:H13`: field 7 equals `On track` with operator `xlOr`.
- Add whole-number validation to `H2:H13`: stop alert, between 1 and 5.
- Add conditional formatting to `F2:F13`: cell value greater than `25000`, pale-green fill, priority 1.
- Create the run-specific workbook defined name referring to `=<data-sheet>!$A$1:$H$13`.

### 2. Plan worksheet

Create this layout:

- `A1`: `Monthly sales plan`; merge `A1:E1`.
- `A2`: `Actual revenue is calculated from the live transaction sheet.`
- Headers in `A3:E3`: `Month`, `Actual revenue`, `Target revenue`, `Variance`, `Attainment`.
- Months Jan–Jun in `A4:A9`.
- `B4:B9`: formulas adding the matching North and South revenue cells from the data sheet:
  - Jan: data `F2+F8`
  - Feb: data `F3+F9`
  - Mar: data `F4+F10`
  - Apr: data `F5+F11`
  - May: data `F6+F12`
  - Jun: data `F7+F13`
- Targets in `C4:C9`: 40000, 42000, 45000, 48000, 52000, 55000.
- `D4:D9`: formula `Actual revenue - Target revenue` for each row.
- `E4:E9`: formula `Actual revenue / Target revenue` for each row.

Apply these features:

- Title `A1:E1`: navy fill, white bold 16-point Arial, centered.
- Subtitle `A2:E2`: italic dark text.
- Header `A3:E3`: blue fill, white bold text, centered.
- Currency format `"$"#,##0` on `B4:D9`; percent format `0.0%` on `E4:E9`.
- Conditional formatting on `D4:D9`: values greater than 0 use pale-green fill, priority 1.
- Column widths: A 14, B:E 16. Row 1 height 28.
- Freeze at `A4`.

### 3. Dashboard worksheet

Create this layout:

- `A1`: `Commercial performance overview`; merge `A1:J1`.
- `A2`: `Six-month actuals, targets and regional mix`.
- KPI labels in `A4:A7`: `Total revenue`, `Target revenue`, `Variance`, `Attainment`.
- KPI formulas in `B4:B7`: plan actual total, plan target total, their difference, and actual divided by target.
- Secondary labels in `E4:E7`: `North revenue`, `South revenue`, `Best month`, `Best month revenue`.
- Secondary results in `F4:F7`: sums of the North and South data revenue ranges, text `Jun`, and the June plan revenue.
- Monthly chart table in `A10:C16`: headers `Month`, `Actual revenue`, `Target revenue`; rows link to plan `A4:C9`.
- Regional chart table in `E10:F12`: headers `Region`, `Revenue`; rows `North` and `South`, linked to dashboard cells F4 and F5.

Apply these features:

- Title `A1:J1`: navy fill, white bold 16-point Arial, centered.
- Subtitle `A2:J2`: italic dark text.
- Label blocks `A4:A7` and `E4:E7`: pale-blue fill, bold dark text.
- KPI values: bold dark text; currency `"$"#,##0` on `B4:B6`, `F4:F5` and `F7`; percentage `0.0%` on `B7`.
- Headers `A10:C10` and `E10:F10`: blue fill, white bold centered text.
- Column widths: A and E 24, B:C and F 15. Row 1 height 28.
- Move this worksheet immediately before `Sheet1`.

Add two native editable charts to the dashboard:

1. Monthly chart using `A10:C16`, type `bar`, style 2, title `Monthly revenue vs target`, two series, legend at bottom, horizontal title `Month`, vertical title `Revenue`, value labels enabled, positioned from column 7 / row 3, width 3,800,000 and height 2,300,000. Set explicit series names `Actual revenue` and `Target revenue`, values from B11:B16 and C11:C16, categories from A11:A16.
2. Regional chart using `E10:F12`, type `bar`, style 2, title `Revenue by region`, one series, legend at bottom, horizontal title `Region`, vertical title `Revenue`, value labels enabled, positioned from column 7 / row 17, width 3,800,000 and height 2,300,000. Set series name `Revenue`, values from F11:F12 and categories from E11:E12.

### 4. Pivot analysis

Create the run-specific pivot from data `A1:E13` on the pre-created `MCP_PivotView_<run_id>` worksheet at `A1` with:

- row field `Region`
- column field `Product`
- data field `Units`
- style `PivotStyleMedium2`

Require these semantic `GetData` assertions:

| Items | Expected |
| --- | ---: |
| North, Core | 405 |
| North, Plus | 252 |
| South, Core | 300 |
| South, Plus | 200 |

## Required final evidence

Treat the task as successful only if the single MCP response proves all of the following:

- `ok=true`, `authority=LIVE_VERIFY`, `noOp=false`.
- `persistentSession.oneEditorSession=true`.
- Exactly one successful persistence barrier is present.
- Whole-task verification is read-only and every operation check passes.
- All three requested final readbacks are present with `LIVE_READ` cell data from the same session.
- From those readbacks, report the observed dashboard values for total revenue, target revenue, variance and attainment. Expected displayed business values are based on 287205 actual revenue and 282000 target revenue, but report the live observed values rather than merely repeating these expectations.
- Report the pivot assertion values and the two verified chart identities from the MCP result.

If any evidence is missing or mismatched, fail closed and show the relevant diagnostic. Do not claim completion from callback success alone.
