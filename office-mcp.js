const fs = require('fs')
const path = require('path')
const crypto = require('crypto')
const { spawn } = require('child_process')

const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js')
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')

// Office MCP - scoped tools for driving the euro-office/OnlyOffice Document Server.
// SCAFFOLD. Wiring into an agent session is a separate, access-gated step.
//
// Capability split, and the reason for it:
//   editing a cell  -> local OOXML rewrite, needs NO Document Server;
//   recalculating   -> needs the DS docbuilder API, which is licence-gated (advanced_api).
// They are deliberately separate tools. ConvertService is NOT a recalc path: it returns the
// cached value for formula cells (or the formula text when there is no cache), so routing
// recalc through it would silently hand back stale numbers - the worst possible failure for
// a spreadsheet tool. `office_status` exists so a caller can find out, before it edits
// anything, whether recalc is actually available on this deployment.

const envPath = path.join(__dirname, '.env')

if (fs.existsSync(envPath)) {
  for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
    const trimmed = line.trim()
    if (!trimmed || trimmed.startsWith('#')) continue

    const pos = trimmed.indexOf('=')
    if (pos < 1) continue

    const key = trimmed.slice(0, pos).trim()
    // .env values may be quoted; the bridge convention is to store them unquoted, but strip
    // defensively so a quoted secret never ends up inside the signature.
    const value = trimmed.slice(pos + 1).trim().replace(/^(['"])(.*)\1$/, '$2')

    if (!(key in process.env)) {
      process.env[key] = value
    }
  }
}

// Endpoint is not a secret and may be logged. The JWT secret is live and must never be
// printed, echoed into an error message, or sent anywhere except the DS signature.
const DS_URL = (process.env.OFFICE_DS_URL || '').replace(/\/$/, '')
const DS_JWT = process.env.OFFICE_DS_JWT || ''
const DS_TIMEOUT_MS = Number(process.env.OFFICE_DS_TIMEOUT_MS || 30000)

function configError() {
  const missing = []
  if (!DS_URL) missing.push('OFFICE_DS_URL')
  if (!DS_JWT) missing.push('OFFICE_DS_JWT')
  if (!missing.length) return null
  return (
    `Document Server is not configured: ${missing.join(', ')} missing. ` +
    'Both are injected from the deployment config; the JWT secret is never stored in this file.'
  )
}

function b64url(buf) {
  return Buffer.from(buf).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_')
}

// HS256, per the Document Server contract: the same token goes in the request body as
// `token` AND in the Authorization header.
function signJwt(payload) {
  const header = b64url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }))
  const body = b64url(JSON.stringify(payload))
  const signature = b64url(crypto.createHmac('sha256', DS_JWT).update(`${header}.${body}`).digest())
  return `${header}.${body}.${signature}`
}

async function dsPost(endpoint, payload) {
  const token = signJwt(payload)
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), DS_TIMEOUT_MS)
  try {
    const res = await fetch(`${DS_URL}${endpoint}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        // ConvertService answers XML unless JSON is requested explicitly.
        Accept: 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({ ...payload, token }),
      signal: controller.signal,
    })
    const text = await res.text()
    try {
      return { ok: res.ok, status: res.status, body: JSON.parse(text) }
    } catch {
      return { ok: res.ok, status: res.status, body: { raw: text.slice(0, 400) } }
    }
  } finally {
    clearTimeout(timer)
  }
}

// Whether this deployment can recalculate at all. The docbuilder API is gated behind the
// advanced_api licence feature; without it every docbuilder call fails with error -3, which
// on its own looks like a broken script rather than a missing entitlement.
async function probeCapabilities() {
  const version = await dsPost('/coauthoring/CommandService.ashx', { c: 'version' })
  const reachable = version.ok && version.body && version.body.error === 0

  let advancedApi = null
  let licenceError = null
  if (reachable) {
    const licence = await dsPost('/coauthoring/CommandService.ashx', { c: 'license' })
    const info = licence.body && licence.body.license
    if (info && typeof info.advanced_api === 'boolean') {
      advancedApi = info.advanced_api
    } else {
      licenceError = 'licence response did not report advanced_api'
    }
  }

  return {
    endpoint: DS_URL,
    reachable,
    authOk: reachable, // CommandService error:0 means the signature was accepted
    serverVersion: reachable ? version.body.version || null : null,
    recalcAvailable: advancedApi === true,
    advancedApi,
    licenceError,
    // error 6 / -8 mean the token was rejected - config problem, not a network problem.
    lastError: reachable ? null : (version.body && version.body.error) ?? `HTTP ${version.status}`,
  }
}

// The workbook rewrite runs in a Python helper: xlsx is a ZIP container, Node has no bundled
// ZIP support and this host has no `zip` binary, so the alternative would be hand-rolling
// central-directory writing - a reliable way to produce a subtly corrupt workbook. `zipfile`
// is stdlib, so no dependency is added by going this route.
const XLSX_WRITER = path.join(__dirname, 'office_xlsx.py')

// Every helper speaks the same contract - one JSON request on stdin, one JSON object on
// stdout, its own errors included as {ok:false,error} - so they all spawn the same way.
// A non-zero exit or unparseable stdout is a helper defect and is reported as one clean
// message; the caller must never receive a Python traceback.
function runHelper(helper, request, label) {
  return new Promise((resolve, reject) => {
    const child = spawn('python3', [helper], { stdio: ['pipe', 'pipe', 'pipe'] })
    let out = ''
    let err = ''
    child.stdout.on('data', (d) => (out += d))
    child.stderr.on('data', (d) => (err += d))
    child.on('error', (e) => reject(new Error(`could not run the ${label} helper: ${e.message}`)))
    child.on('close', (code) => {
      if (code !== 0) {
        return reject(new Error(err.trim() || `${label} helper exited with code ${code}`))
      }
      try {
        resolve(JSON.parse(out))
      } catch {
        reject(new Error(`${label} helper returned unreadable output: ${out.slice(0, 200)}`))
      }
    })
    child.stdin.end(JSON.stringify(request))
  })
}

// Recalculation runs through headless LibreOffice, in a separate Python helper that owns the
// verified recipe (per-job profile with recalc-on-load forced) and the process handling
// (timeout, kill-on-hang, trust-the-file-not-the-exit-code, one clean error, temp cleanup).
// Deliberately NOT the Document Server / ConvertService, which returns stale cached values.
// The helper always answers a single JSON line on stdout, including for its own errors.
const RECALC_HELPER = path.join(__dirname, 'office_recalc.py')
const PDF_HELPER = path.join(__dirname, 'office_pdf.py')
const DOCX_HELPER = path.join(__dirname, 'office_docx.py')

// Building a docx/pptx from scratch runs in its own helper for the same reason the cell
// write does: both formats are ZIP containers of XML, and `zipfile` is stdlib while a Node
// ZIP writer would have to be hand-rolled. python-docx / python-pptx are NOT installed
// here (no pip, no root), so the builders own the verified by-hand OOXML recipe instead.
const DOCX_BUILDER = path.join(__dirname, 'office_docx_build.py')
const PPTX_BUILDER = path.join(__dirname, 'office_pptx_build.py')
const PPTX_HELPER = path.join(__dirname, 'office_pptx.py')

const runXlsxWriter = (request) => runHelper(XLSX_WRITER, request, 'xlsx')
const runRecalc = (request) => runHelper(RECALC_HELPER, request, 'recalc')
const runPdf = (request) => runHelper(PDF_HELPER, request, 'pdf')
const runDocx = (request) => runHelper(DOCX_HELPER, request, 'docx')
const runDocxBuild = (request) => runHelper(DOCX_BUILDER, request, 'docx build')
const runPptxBuild = (request) => runHelper(PPTX_BUILDER, request, 'pptx build')
const runPptx = (request) => runHelper(PPTX_HELPER, request, 'pptx')

const server = new McpServer({ name: 'office-mcp', version: '0.1.0' })

server.tool(
  'office_status',
  'Report whether the Document Server is reachable, whether the JWT is accepted, and ' +
    'whether formula recalculation is available on this deployment. Call this before ' +
    'relying on recalc: editing works without the Document Server, recalculating does not.',
  {},
  async () => {
    const problem = configError()
    if (problem) {
      return { content: [{ type: 'text', text: problem }], isError: true }
    }
    try {
      const status = await probeCapabilities()
      const lines = [
        `endpoint: ${status.endpoint}`,
        `reachable: ${status.reachable}`,
        `server version: ${status.serverVersion ?? 'unknown'}`,
        `recalc available: ${status.recalcAvailable}`,
      ]
      if (!status.recalcAvailable) {
        lines.push(
          status.advancedApi === false
            ? 'recalc is unavailable: the docbuilder API is not licensed here (advanced_api=false). ' +
              'Editing cells still works; recalculation needs this resolved on the Document Server side.'
            : `recalc availability unknown: ${status.licenceError || status.lastError}`,
        )
      }
      return { content: [{ type: 'text', text: lines.join('\n') }] }
    } catch (err) {
      return {
        content: [{ type: 'text', text: `Document Server probe failed: ${err.message}` }],
        isError: true,
      }
    }
  },
)

server.tool(
  'xlsx_set_cells',
  'Write values or formulas into cells of an xlsx workbook. Operates on the file directly ' +
    'and does not need the Document Server. Formula cells written here are stored without a ' +
    'cached result, so a reader that does not recalculate will see the formula, not a number ' +
    '- call xlsx_recalc afterwards when a computed value is required.',
  {
    file: z.string().describe('Absolute path of the .xlsx workbook to modify'),
    sheet: z.string().describe('Worksheet name, e.g. "Sheet1"'),
    cells: z
      .array(
        z.object({
          ref: z.string().describe('Cell reference, e.g. "B7"'),
          value: z.union([z.string(), z.number()]).optional().describe('Literal cell value'),
          formula: z.string().optional().describe('Formula without the leading "=", e.g. "A1+B1"'),
        }),
      )
      .min(1)
      .describe('Cells to write; give either value or formula per cell'),
  },
  async ({ file, sheet, cells }) => {
    try {
      const result = await runXlsxWriter({ file, sheet, cells })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      const written = result.applied.map((c) => `${c.ref} (${c.kind})`).join(', ')
      const note = result.applied.some((c) => c.kind === 'formula')
        ? '\nFormula cells were written without a cached result, and the workbook is marked to ' +
          'recalculate on open. A reader that does not recalculate will show the formula rather ' +
          'than a number - use xlsx_recalc when a computed value has to be stored.'
        : ''
      return {
        content: [{ type: 'text', text: `Wrote ${result.applied.length} cell(s) to ${sheet}: ${written}.${note}` }],
      }
    } catch (err) {
      return { content: [{ type: 'text', text: `Cell write failed: ${err.message}` }], isError: true }
    }
  },
)

server.tool(
  'xlsx_recalc',
  'Recalculate the formulas of an xlsx workbook through headless LibreOffice and write the ' +
    'computed values back into the file. LibreOffice does NOT recalculate xlsx on load by ' +
    'default and --convert-to does not override that, so this forces recalc-on-load in a ' +
    'throwaway per-job profile; without it the stale cached value is returned unchanged. ' +
    'Runs on the container engine by default (auto when docker and the image are available, ' +
    'or OFFICE_ENGINE=docker:<image>), falling back visibly to host soffice; the result names ' +
    'which engine ran. On any failure the original file is left untouched.',
  {
    file: z.string().describe('Absolute path of the .xlsx workbook to recalculate'),
  },
  async ({ file }) => {
    try {
      const result = await runRecalc({ file })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      return {
        content: [
          {
            type: 'text',
            text:
              `Recalculated ${result.file} via ${result.engine} (recalc-on-load forced). ` +
              'Formula cells now carry freshly computed values.',
          },
        ],
      }
    } catch (err) {
      return { content: [{ type: 'text', text: `Recalc failed: ${err.message}` }], isError: true }
    }
  },
)

async function main() {
  await server.connect(new StdioServerTransport())
}

server.tool(
  'office_to_pdf',
  'Export a spreadsheet, document or presentation (xlsx/docx/pptx and their older and ' +
    'OpenDocument equivalents) to PDF through the same headless LibreOffice engine the ' +
    'recalculation uses. The source file is not modified. Writes <name>.pdf next to the ' +
    'source unless out_file says otherwise, and names which engine ran. A spreadsheet is ' +
    'exported as it stands: if its formula cells carry stale cached values, call xlsx_recalc ' +
    'first, or the PDF will faithfully show the stale numbers. A workbook with several ' +
    'sheets must say WHICH sheet to export: LibreOffice puts every sheet in the PDF, hidden ' +
    'ones included, so a whole-workbook export of a customer file can carry personal data ' +
    'nobody asked for. Without sheet (or an explicit all_sheets) the call is refused and the ' +
    'available sheets are listed.',
  {
    file: z.string().describe('Absolute path of the source file'),
    out_file: z.string().optional().describe('Absolute path for the PDF; defaults to the source with a .pdf extension'),
    sheet: z.string().optional().describe('For a spreadsheet: the ONE sheet to export, by name. Only that sheet ends up in the PDF.'),
    all_sheets: z.boolean().optional().describe('Export every sheet of a workbook. Ask for this deliberately: it puts all sheets, hidden ones included, into the PDF.'),
  },
  async ({ file, out_file, sheet, all_sheets }) => {
    try {
      const result = await runPdf({ file, out_file, sheet, all_sheets })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      return {
        content: [{ type: 'text', text: `Exported ${file} to ${result.file} via ${result.engine}.` }],
      }
    } catch (err) {
      return { content: [{ type: 'text', text: `PDF export failed: ${err.message}` }], isError: true }
    }
  },
)

server.tool(
  'docx_replace_text',
  'Replace text in a Word document, in place. Searches what a READER sees, not the raw ' +
    'markup: Word stores a sentence as several runs, so a value like "12 500 000" is often ' +
    'split mid-number and a plain search-and-replace over the file finds nothing while ' +
    'reporting success. Each replacement comes back with the number of occurrences changed - ' +
    'a count of 0 means the text was not in the document, which is the answer you want to see ' +
    'rather than a silent no-op. Where a replacement spans two differently formatted runs, ' +
    'the result takes the first run\'s formatting.',
  {
    file: z.string().describe('Absolute path of the .docx to modify'),
    replacements: z
      .array(z.object({
        find: z.string().describe('Text to look for, as it reads on the page'),
        replace: z.string().describe('Text to put in its place'),
      }))
      .min(1)
      .describe('Replacements, applied in order to the whole document'),
  },
  async ({ file, replacements }) => {
    try {
      const result = await runDocx({ file, replacements })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      const summary = result.applied
        .map((a) => `"${a.find}": ${a.replaced} occurrence(s)`)
        .join('; ')
      const missed = result.applied.filter((a) => a.replaced === 0)
      const note = missed.length
        ? `\nNothing matched for: ${missed.map((a) => `"${a.find}"`).join(', ')}. ` +
          'The document was still rewritten for the other replacements; check the wording, ' +
          'including spaces and non-breaking spaces.'
        : ''
      return { content: [{ type: 'text', text: `Updated ${file}. ${summary}.${note}` }] }
    } catch (err) {
      return { content: [{ type: 'text', text: `Text replacement failed: ${err.message}` }], isError: true }
    }
  },
)

server.tool(
  'docx_set_table_cell',
  'Write one table cell of a Word document, addressed by position (all zero-based: table ' +
    'index in the document, row, column). Any other text already in that cell is cleared, so ' +
    'the old value cannot end up sitting next to the new one on the page.',
  {
    file: z.string().describe('Absolute path of the .docx to modify'),
    table: z.number().int().min(0).describe('Which table in the document, counting from 0'),
    row: z.number().int().min(0).describe('Row index within the table, counting from 0'),
    column: z.number().int().min(0).describe('Cell index within the row, counting from 0'),
    text: z.string().describe('Text to place in the cell'),
  },
  async ({ file, table, row, column, text }) => {
    try {
      const result = await runDocx({ file, cells: [{ table, row, column, text }] })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      return {
        content: [{ type: 'text', text: `Wrote table ${table}, row ${row}, column ${column} of ${file}.` }],
      }
    } catch (err) {
      return { content: [{ type: 'text', text: `Table cell write failed: ${err.message}` }], isError: true }
    }
  },
)

server.tool(
  'pptx_replace_text',
  'Replace text in an existing PowerPoint deck, in place. Searches what a VIEWER sees, not ' +
    'the raw markup: PowerPoint stores a line as several runs, so a value like "12 500 000" ' +
    'is often split mid-number and a plain search-and-replace over the file finds nothing ' +
    'while reporting success. Each replacement comes back with the number of occurrences ' +
    'changed AND which slides they were on - a count of 0 means the text was not in the deck, ' +
    'which is the answer you want to see rather than a silent no-op. Slides are numbered the ' +
    'way a viewer counts them (the first slide is 1), which is NOT necessarily the order of ' +
    'the slide files inside the package. Speaker notes, the master and layouts, and text ' +
    'inside charts or SmartArt are NOT searched. Where a replacement spans two differently ' +
    'formatted runs, the result takes the first run\'s formatting.',
  {
    file: z.string().describe('Absolute path of the .pptx to modify'),
    replacements: z
      .array(z.object({
        find: z.string().describe('Text to look for, as it reads on the slide'),
        replace: z.string().describe('Text to put in its place'),
      }))
      .min(1)
      .describe('Replacements, applied in order'),
    slide: z
      .number()
      .int()
      .min(1)
      .optional()
      .describe('Limit the replacement to ONE slide, counting from 1. Omit to visit every slide.'),
  },
  async ({ file, replacements, slide }) => {
    try {
      const result = await runPptx({ file, replacements, slide })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      const summary = result.applied
        .map((a) => `"${a.find}": ${a.replaced} occurrence(s)` +
          (a.slides && a.slides.length ? ` on slide ${a.slides.join(', ')}` : ''))
        .join('; ')
      const missed = result.applied.filter((a) => a.replaced === 0)
      const note = missed.length
        ? `\nNothing matched for: ${missed.map((a) => `"${a.find}"`).join(', ')}. ` +
          'The deck was still rewritten for the other replacements; check the wording, and ' +
          'remember that speaker notes, the master/layouts and chart text are not searched.'
        : ''
      return { content: [{ type: 'text', text: `Updated ${file}. ${summary}.${note}` }] }
    } catch (err) {
      return { content: [{ type: 'text', text: `Text replacement failed: ${err.message}` }], isError: true }
    }
  },
)

server.tool(
  'pptx_set_table_cell',
  'Write one table cell of a PowerPoint slide, addressed by position. The slide is counted ' +
    'the way a viewer counts it (the first slide is 1) - deliberately NOT zero-based, because ' +
    'the file order inside a .pptx does not have to match the display order, and an ' +
    'off-by-one there edits a different slide and still looks like success. Table, row and ' +
    'column are zero-based, matching docx_set_table_cell. Any other text already in the cell ' +
    'is cleared, so the old value cannot end up next to the new one; a cell that is empty ' +
    '(no text run at all, which is normal in PowerPoint) gets a new run, which takes its ' +
    'formatting from the table style.',
  {
    file: z.string().describe('Absolute path of the .pptx to modify'),
    slide: z.number().int().min(1).describe('Which slide, counting from 1 as a viewer does'),
    table: z.number().int().min(0).describe('Which table on that slide, counting from 0'),
    row: z.number().int().min(0).describe('Row index within the table, counting from 0'),
    column: z.number().int().min(0).describe('Cell index within the row, counting from 0'),
    text: z.string().describe('Text to place in the cell'),
  },
  async ({ file, slide, table, row, column, text }) => {
    try {
      const result = await runPptx({ file, cells: [{ slide, table, row, column, text }] })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      return {
        content: [{
          type: 'text',
          text: `Wrote slide ${slide}, table ${table}, row ${row}, column ${column} of ${file}.`,
        }],
      }
    } catch (err) {
      return { content: [{ type: 'text', text: `Table cell write failed: ${err.message}` }], isError: true }
    }
  },
)

// Run formatting, shared by both builders. `type` is a plain string rather than an enum so
// the REFUSAL of an unknown block type comes from the builder itself, with its message
// naming what is supported - one source of truth instead of two that can drift apart.
const runShape = {
  text: z.string().describe('The run text'),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: z.string().optional().describe('Hex RGB without "#", e.g. "1F4E79"'),
  size: z.number().optional().describe('Font size in points'),
}

const docxRun = z.object({
  ...runShape,
  highlight: z.string().optional().describe('Highlight name, e.g. "yellow"'),
})

const pptxRun = z.object(runShape)

const cellValue = z.union([z.string(), z.number()])

server.tool(
  'docx_build',
  'Build a Word document from a structured description - a NEW file, not an edit of an ' +
    'existing one (use docx_replace_text / docx_set_table_cell for that). Handles exactly ' +
    'these block types: heading (level 1-3), paragraph (with inline formatting per run), ' +
    'list (bullet or ordered), table (styled repeating header, banded rows), quote (a ' +
    'callout with a coloured left bar), image (PNG, scaled to the text width, optional ' +
    'caption) and page_break. A header and a footer with "page / total" come from the top ' +
    'level. Anything richer - columns, footnotes, native charts - is REFUSED by name rather ' +
    'than silently dropped, because a document that is missing a block the caller asked for ' +
    'is worse than a call that failed. Needs no Document Server and no third-party library.',
  {
    file: z.string().describe('Absolute path of the .docx to create (overwritten if it exists)'),
    title: z.string().optional().describe('Document title, rendered above the first block'),
    header: z.string().optional().describe('Page header text'),
    footer: z.string().optional().describe('Page footer text; the page number is appended as "n / m"'),
    blocks: z
      .array(
        z.object({
          type: z.string().describe('heading | paragraph | list | table | quote | image | page_break'),
          text: z.string().optional().describe('heading/paragraph/quote text'),
          level: z.number().int().min(1).max(3).optional().describe('heading level'),
          runs: z.array(docxRun).optional().describe('paragraph: runs with inline formatting'),
          items: z
            .array(z.union([z.string(), z.array(docxRun)]))
            .optional()
            .describe('list: items, each a string or an array of runs'),
          ordered: z.boolean().optional().describe('list: numbered instead of bulleted'),
          header: z.array(z.string()).optional().describe('table: header cells'),
          rows: z.array(z.array(cellValue)).optional().describe('table: rows, each matching the header width'),
          banded: z.boolean().optional().describe('table: alternate row shading (default true)'),
          author: z.string().optional().describe('quote: attribution'),
          path: z.string().optional().describe('image: absolute path of a PNG'),
          caption: z.string().optional().describe('image: caption below the picture'),
        }),
      )
      .optional()
      .describe('Document content, in order. A title alone is also accepted.'),
  },
  async ({ file, title, header, footer, blocks }) => {
    try {
      const result = await runDocxBuild({ file, title, header, footer, blocks })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      const counts = Object.entries(result.built.blocks || {})
        .map(([kind, n]) => `${n} ${kind}`)
        .join(', ')
      return {
        content: [
          {
            type: 'text',
            text:
              `Built ${result.file} (${counts || 'title only'}` +
              `${result.built.images ? `, ${result.built.images} image(s) embedded` : ''}). ` +
              'Render it with office_to_pdf if the layout has to be checked.',
          },
        ],
      }
    } catch (err) {
      return { content: [{ type: 'text', text: `Document build failed: ${err.message}` }], isError: true }
    }
  },
)

server.tool(
  'pptx_build',
  'Build a PowerPoint deck from a structured description - a NEW file, not an edit of an ' +
    'existing one. One slide per entry, each with a title and a stack of blocks: text (with ' +
    'inline formatting per run), list (bullet or ordered), table (a REAL table with a styled ' +
    'header and banded rows, not a picture of one), callout (a boxed one-liner) and image ' +
    '(PNG, aspect ratio read from the file). A top-level title (with optional subtitle) adds ' +
    'a cover slide. Blocks that would run past the bottom edge are REFUSED with the slide ' +
    'named, so the caller splits the slide instead of shipping a deck with half a table ' +
    'hanging off it; animation, native charts and speaker notes are likewise refused rather ' +
    'than half-supported. Needs no Document Server and no third-party library.',
  {
    file: z.string().describe('Absolute path of the .pptx to create (overwritten if it exists)'),
    title: z.string().optional().describe('Cover slide heading; omit for no cover slide'),
    subtitle: z.string().optional().describe('Cover slide subtitle'),
    slides: z
      .array(
        z.object({
          title: z.string().optional().describe('Slide title'),
          blocks: z
            .array(
              z.object({
                type: z.string().describe('text | list | table | callout | image'),
                text: z.string().optional().describe('text/callout text'),
                runs: z.array(pptxRun).optional().describe('text: runs with inline formatting'),
                items: z
                  .array(z.union([z.string(), z.array(pptxRun)]))
                  .optional()
                  .describe('list: items, each a string or an array of runs'),
                ordered: z.boolean().optional().describe('list: numbered instead of bulleted'),
                header: z.array(z.string()).optional().describe('table: header cells'),
                rows: z.array(z.array(cellValue)).optional().describe('table: rows, each matching the header width'),
                path: z.string().optional().describe('image: absolute path of a PNG'),
                scale: z.number().optional().describe('image: fraction of the content width (default 1.0)'),
              }),
            )
            .optional()
            .describe('Slide content, stacked from the top'),
        }),
      )
      .optional()
      .describe('The slides, in order. A title alone builds a one-slide cover deck.'),
  },
  async ({ file, title, subtitle, slides }) => {
    try {
      const result = await runPptxBuild({ file, title, subtitle, slides })
      if (!result.ok) {
        return { content: [{ type: 'text', text: result.error }], isError: true }
      }
      return {
        content: [
          {
            type: 'text',
            text:
              `Built ${result.file}: ${result.built.slides} slide(s)` +
              `${result.built.images ? `, ${result.built.images} image(s) embedded` : ''}. ` +
              'Render it with office_to_pdf if the layout has to be checked.',
          },
        ],
      }
    } catch (err) {
      return { content: [{ type: 'text', text: `Deck build failed: ${err.message}` }], isError: true }
    }
  },
)

// Only start the transport when run as a server; requiring the file (tests) must not block
// on stdio. The signing helper is exported so it can be verified without a Document Server.
if (require.main === module) {
  main().catch((err) => {
    console.error('office-mcp failed to start:', err.message)
    process.exit(1)
  })
}

module.exports = { signJwt, probeCapabilities, configError }
