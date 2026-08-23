const fs = require('fs')
const os = require('os')
const path = require('path')
const { execFileSync } = require('child_process')

const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js')
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js')
const { z } = require('zod')

const lib = require('./lib.cjs')
const { runJob, EXEC } = require('./runner.cjs')
const coedit = require('./coedit.cjs')
const { checkPackageConsistency } = require('./package-consistency.cjs')
const magok = require('./euro-magok.cjs')
const trace = require('./office-trace.cjs')

// this file's OWN mtime, same pattern as lib.cjs/coedit.cjs's
// checkFreshness() -- a third, independent axis (the entry file itself, not just the modules it
// requires). staleCodeGuard() below combines this with whichever of lib.cjs/coedit.cjs a given
// write path actually depends on, and reports EVERY stale file at once, not just the first --
// a long-idle child can have all three combed out from under it simultaneously.
const __EURO_MCP_LOAD_MTIME_MS = fs.statSync(__filename).mtimeMs
function checkOwnFreshness() {
  const diskMtimeMs = fs.statSync(__filename).mtimeMs
  if (diskMtimeMs <= __EURO_MCP_LOAD_MTIME_MS) return { fresh: true }
  return {
    fresh: false,
    file: 'euro-mcp.cjs',
    message:
      `euro-mcp.cjs: a betoltott kod regebbi, mint a lemezen levo (betoltve: ${new Date(__EURO_MCP_LOAD_MTIME_MS).toISOString()}, ` +
      `lemezen: ${new Date(diskMtimeMs).toISOString()}) -- inditsd ujra a sessiont`,
  }
}

// Runs the given checkFreshness()-shaped functions and, if ANY report stale, returns a
// protokoll-tilt-shaped refusal naming every stale file (not just the first). Returns null when
// everything is fresh -- the caller proceeds unguarded in that case. A JELZES, nem cselekves
// (MI NEM TARTOZIK BELE): nothing here reloads a module or restarts the process.
function staleCodeGuard(...checks) {
  const stale = checks.map((fn) => fn()).filter((r) => !r.fresh)
  if (!stale.length) return null
  return {
    ok: false,
    outcome: 'protokoll-tilt',
    error:
      'a betoltott kod regebbi, mint a lemezen levo -- inditsd ujra a sessiont, mielott irsz. ' +
      stale.map((s) => s.message).join(' | '),
    staleFiles: stale.map((s) => s.file),
  }
}

// EURO-MCP - document editing through the Document Server's DocBuilder API.
//
// The DocBuilder path drives the real document model (paragraph objects, layout, recalculation).
// That is what separates it from the local OOXML rewriting office-mcp.js already does: for plain
// text substitution office-mcp is the cheaper tool and needs no Document Server at all.
//
// WHY THERE IS NO open/edit/save TRIPLE: the DocBuilder API keeps no session across HTTP calls.
// One POST carries one whole script, which the server runs start to finish and then discards, so
// an `open_document` tool would have nothing to hand to an `edit_document` tool - the name would
// promise state the protocol does not keep. The split here is by the decision a caller actually
// makes: ask what the service can do, edit a document, or drive the builder directly.
//
// WHERE THE WORK HAPPENS: on the euro-office box, always. The Document Server fetches the
// builder script and the input document itself, over HTTP, from an address IT can reach - so
// both the call and the file serving have to sit next to it. See runner.cjs.

function textResult(payload) {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] }
}

const server = new McpServer({ name: 'euro-mcp', version: '0.2.0' })

// ---------------------------------------------------------------------------
server.tool(
  'service_status',
  'Runs one minimal real job end to end and reports whether document editing works right now. ' +
    'It answers the question an edit failure actually raises: is the token accepted, is the ' +
    'DocBuilder API available on this instance, and can the server reach our files - three ' +
    'causes an edit error cannot tell apart on its own.',
  {},
  async () => {
    // office-diag-e6-trace-id: generated HERE, per call, never stored on any module -- passed
    // down as a plain parameter so concurrent calls cannot cross-contaminate each other's log
    // lines (see office-trace.cjs's own comment for why a module-level ID would be wrong).
    const traceId = trace.newTraceId()
    trace.logTrace(traceId, 'mcp-tool-invoked', { tool: 'service_status' })
    const marker = lib.makeMarker()
    const script = lib.buildEditScript({ docUrl: '__DOC_URL__', operations: [], marker })
    const answer = await runJob({ script, traceId })
    return textResult({
      ok: answer.ok === true,
      transport: EXEC,
      outcome: answer.outcome,
      traceId,
      // The number of fetches the server made is what separates "it never reached us" from
      // "it reached us and failed afterwards"; the error code alone says neither.
      serverFetches: answer.serverFetches ?? null,
      detail: answer.detail || (answer.ok ? 'a full edit round trip completed and was verified' : 'the round trip did not complete'),
    })
  },
)

// ---------------------------------------------------------------------------
server.tool(
  'edit_document',
  'Opens a docx, applies edit operations, saves it, and verifies that the document really ' +
    'changed by reading a per-call marker back out of the saved file. Without document_path a ' +
    'minimal document is generated, which is what makes this usable as a smoke test.',
  {
    operations: z
      .array(
        z.union([
          z.object({ type: z.literal('append_paragraph'), text: z.string() }),
          z.object({ type: z.literal('replace_text'), search: z.string(), replace: z.string() }),
        ]),
      )
      .describe('Edit operations, applied in order'),
    document_path: z.string().optional().describe('Local path of the docx to edit; omit to edit a generated minimal document'),
    output_path: z.string().optional().describe('Local path to write the edited document to'),
  },
  async ({ operations, document_path, output_path }) => {
    // *** PROTOKOLL-KAPU (a tulajdonos rendelkezese, 2026-08-15 11:0x): ***
    //   a fajl MEG NEM LETEZIK -> DocBuilder, es CSAK a letrehozasra
    //   a fajl MAR LETEZIK     -> DocBuilder TILOS, a szerkesztes a co-editing uton megy
    // `edit_document` egy MEGLEVO dokumentumot szerkeszt, tehat pontosan a tiltott eset -- a
    // `document_path` jelenlete maga a "mar letezik" jel. A tilalom KODBAN all, nem a leirasban:
    // egy szandek, amit semmi nem tart be, ugyanugy viselkedik, mint a regi csendes null.
    if (document_path) {
      // *** ES A MEGTAGADAS MEGMONDJA, HOGY AZ AJTO, AMIRE MUTAT, NYITVA VAN-E. *** Ez a tilalom
      // egy MASIK utra kuldi a hivot -- ha az szinten zart (a co-editing per-agens engedelyezett),
      // akkor a ket kapu egyutt KEPESSEG-HIANYT csinal, es a hivo ezt csak ket hibauzenetbol
      // rakna ossze. Merve 2026-08-15 14:2x: ures allowlisten pontosan ez allt elo.
      const coeditUt = coedit.coeditUtAllapota()
      return textResult({
        ok: false,
        outcome: 'protokoll-tilt',
        error:
          'editing an EXISTING document with the DocBuilder is no longer allowed: that route is ' +
          'reserved for CREATING a file that does not exist yet. Open the file and edit it through ' +
          'the co-editing session instead. (Calling edit_document without document_path still ' +
          'works: it generates a new minimal document, which is a creation, not an edit.)',
        coeditUtElerheto: coeditUt.elerheto,
        coeditUtAllapot: coeditUt.indok,
      })
    }
    const documentBase64 = undefined
    const traceId = trace.newTraceId()
    trace.logTrace(traceId, 'mcp-tool-invoked', { tool: 'edit_document' })

    const marker = lib.makeMarker()
    let script
    try {
      script = lib.buildEditScript({ docUrl: '__DOC_URL__', operations, marker })
    } catch (err) {
      return textResult({ ok: false, error: `the edit could not be expressed as a builder script: ${err.message}`, traceId })
    }

    const answer = await runJob({ script, documentBase64, returnDoc: Boolean(output_path), traceId })

    let written = null
    if (answer.savedBase64 && output_path) {
      fs.writeFileSync(output_path, Buffer.from(answer.savedBase64, 'base64'))
      written = output_path
    }

    return textResult({
      // The service answering says it ran; the marker says the document changed. Only the second
      // one is what the caller asked for, so only the second one decides `ok`.
      ok: answer.ok === true,
      outcome: answer.outcome,
      traceId,
      // A `kind` a KIMENET tenyleges tipusa, a csomagbol olvasva. Enelkul a hivo csak azt latta,
      // hogy "ok" vagy "nem" -- azt nem, hogy MIT kapott. (A regi ut ezen a ponton nemult el.)
      kind: answer.kind ?? null,
      contentVerified: Array.isArray(answer.markersFound) ? answer.markersFound.includes(marker) : null,
      savedBytes: answer.savedBytes ?? null,
      outputPath: written,
      documentText: answer.documentText ?? null,
      serverFetches: answer.serverFetches ?? null,
      detail: answer.detail || null,
    })
  },
)

// ---------------------------------------------------------------------------
server.tool(
  'run_builder_script',
  'Runs a raw DocBuilder script. This is the primitive the edit tool is built on; use it for ' +
    'anything edit_document does not express. Use __DOC_URL__ where the script needs the URL of ' +
    'the input document - the port is chosen per run, so the caller cannot know that address. ' +
    'No content verification is performed: the script decides what it produces, so only the ' +
    'caller knows what a correct result looks like.',
  {
    script: z.string().describe('The DocBuilder script body; __DOC_URL__ is replaced with the input document URL'),
    document_path: z.string().optional().describe('Local path of the input docx; omit to use a generated minimal document'),
    output_path: z.string().optional().describe('Local path to write the produced document to'),
  },
  async ({ script, document_path, output_path }) => {
    let documentBase64
    if (document_path) {
      if (!fs.existsSync(document_path)) {
        return textResult({ ok: false, error: `the input document does not exist: ${document_path}` })
      }
      documentBase64 = fs.readFileSync(document_path).toString('base64')
    }
    if (!script.includes('builder.OpenFile(')) {
      return textResult({
        ok: false,
        error: 'the script does not call builder.OpenFile(). CreateFile is deliberately not offered here: upstream #321 reports it failing on missing templates.',
      })
    }
    // *** PROTOKOLL-KAPU, MASODIK FELE -- ES ITT SZANDEKOSAN NEM A `document_path`-RA NEZUNK. ***
    // A rendelkezes szerint a DocBuilder CSAK letrehozasra hasznalhato. Kesertes volna a
    // `document_path` jelenletet venni a "mar letezik" jelnek, de az HAMIS lenne: a CreateFile
    // nincs kinalva, ezert egy UJ xlsx/pptx letrehozasahoz is KELL egy tipushelyes mag-fajl.
    // A mag nem a cel-fajl. Amit viszont elleorizni tudunk es kell: a DocBuilder ne IRJON FELUL
    // semmit -- egy letezo kimenetre iranyulo hivas mar nem letrehozas, hanem szerkesztes.
    if (output_path && fs.existsSync(output_path)) {
      // *** KIKOTES 1: eddig
      // csak a bool ment ki, az INDOK nem -- a hivo tudta hogy a masik ut is ZART, de nem tudta
      // MIERT, es egy HARMADIK eszkozzel (edit_document) kellett volna ujraprobalnia az indokert.
      // Most ugyanabbol a `coeditUtAllapota()` hivasbol mindket mezo egyutt megy ki, ahogy az
      // `edit_document` mar teszi.
      const coeditUt = coedit.coeditUtAllapota()
      return textResult({
        ok: false,
        outcome: 'protokoll-tilt',
        error:
          `the output already exists: ${output_path}. The DocBuilder route may only CREATE a file ` +
          'that does not exist yet; changing one that does is the co-editing session\'s job. ' +
          'Choose a new output path, or edit the existing file through the editor.',
        // Ugyanaz az indok, mint az `edit_document` tilalmanal: ha a masik ut sincs nyitva ennek
        // a hivonak, azt ITT kell megmondani, ne egy masodik hibauzenetbol alljon ossze.
        coeditUtElerheto: coeditUt.elerheto,
        coeditUtAllapot: coeditUt.indok,
      })
    }
    // returnDoc only when asked: until this existed, the strongest primitive
    // could not hand back what it produced -- runJob was called without it and there was no
    // output_path at all, so a caller who needed the FILE had to bypass the tool and drive
    // runner.cjs directly. Same semantics as edit_document: the flag follows output_path, so
    // callers who only want the outcome do not pay for shipping a document back.
    const traceId = trace.newTraceId()
    trace.logTrace(traceId, 'mcp-tool-invoked', { tool: 'run_builder_script' })
    const answer = await runJob({ script, documentBase64, returnDoc: Boolean(output_path), traceId })

    let written = null
    let savedBuffer = null
    if (answer.savedBase64 && output_path) {
      savedBuffer = Buffer.from(answer.savedBase64, 'base64')
      fs.writeFileSync(output_path, savedBuffer)
      written = output_path
    }
    const negyallapotu = negyallapotuMezok(answer, { savedBuffer })

    return textResult({
      // `ok` mostantol a negy fuggetlen mezo ES-e, nem
      // az `answer.ok` kozvetlen atvetele -- ez a tool NEM hordoz olyan nevesitett, teszttel
      // rogzitett dontest, mint a create_document `ok`-ja (lasd a create_document `ok` mezojenek
      // sajat magyarazatat lent), ezert a valtoztatas itt biztonsagos.
      ok: negyallapotu.negyallapotuSiker,
      outcome: answer.outcome,
      traceId,
      kind: answer.kind ?? null,
      parts: answer.parts ?? null,
      documentText: answer.documentText ?? null,
      serverFetches: answer.serverFetches ?? null,
      // Named explicitly rather than left implicit: an output_path that was asked for but not
      // written comes back as null here, instead of the caller inferring success from `ok`.
      written,
      detail: answer.detail || null,
      transportOk: negyallapotu.transportOk,
      executionOk: negyallapotu.executionOk,
      outputProduced: negyallapotu.outputProduced,
      contentVerified: negyallapotu.contentVerified,
    })
  },
)

// ---------------------------------------------------------------------------
// NEGYALLAPOTU KIMENET (a tulajdonos 4. pontja): `ok:true` NEM
// SIKER -- a co-editing ut honapokig `ok:true`-t adott az `Api.CreateNumbering` csendes
// tartalom-vesztesenel, es epp ez tette lathatatlanna. Negy FUGGETLEN mezo, osszevonas nelkul:
//   transportOk       jott-e VALAMILYEN ertelmezheto valasz a helpertol (SSH/local elerte-e)
//   executionOk       a Document Server ELFOGADTA es LEFUTTATTA-e a jobot (nem blocked/auth/fetch)
//   outputProduced    keszult-e FELISMERHETO OOXML csomag (a `kind` docx/xlsx/pptx, nem ismeretlen)
//   contentVerified   a csomag NEVESITETT RESZE valoban TARTALMAZ-e valamit (nem csak letezik) --
//                     MERCEJE (korabbi elo meres alapjan): sem az, hogy a hivas nem
//                     dobott (az executionOk), sem egy marker-iras sikere nem eleg -- a KIMENETI
//                     ARTEFAKTUM nevesitett zip-reszenek LETEZESE ES TARTALMA a merce.
// Az OSSZESITETT siker csak akkor all, ha MIND A NEGY igaz -- egyetlen `false` FAILURE-t jelent,
// meg akkor is, ha a tobbi harom igaz (a tulajdonos peldaja a kartyan: contentVerified:false eseten).
//
// A `savedBuffer` opcionalis: CSAK akkor tudunk zip-tartalmat ellenorizni, ha a hivo tenylegesen
// visszakerte a bajtokat (`output_path` + `returnDoc`). Enelkul a contentVerified a MEGLEVO,
// marker-alapu jelre esik vissza (docx: `answer.markersFound`), es xlsx/pptx-re `false` marad --
// a `box-helper.py` sajat kommentje szerint MA is "kesz-tartalom-nem-ellenorizve": egy nem
// ellenorzott allitas nem szamit siker-mezonek, csak mert nem dobott hibat.
function listZipEntries(buffer) {
  // python3 stdlib zipfile-t hasznal, ugyanugy, ahogy a fake-ds.cjs sajat buildDocx()-e es a
  // box-helper.py detect_kind()-je is teszi -- kezzel irt zip-olvaso konnyen csendben hibas.
  const py = `
import base64, io, json, sys, zipfile
data = base64.b64decode(sys.stdin.read())
try:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        print(json.dumps([{"name": i.filename, "size": i.file_size} for i in z.infolist()]))
except Exception as e:
    print(json.dumps({"error": str(e)}))
`
  try {
    const out = execFileSync('python3', ['-c', py], { input: buffer.toString('base64') })
    const parsed = JSON.parse(out.toString())
    return Array.isArray(parsed) ? parsed : null
  } catch {
    return null
  }
}

// A regi `DS_FAILURE_OUTCOMES` (blocked/auth/fetch/unknown/nem-mert/ismeretlen-csomag) ket kulon
// hibaosztalyt kevert -- a negyallapotu bontashoz KULON halmaz kell a transport/execution
// hatarra, mert 'ismeretlen-csomag' mar TULJUTOTT az execution-on (a DS lefuttatta, csak a
// csomag nem ismerheto fel), tehat az outputProduced agba tartozik, nem az executionOk-ba; a
// 'nem-mert' pedig a transport-hatarba (lasd alapMezok() `transportOk` sora). A regi konstans
// innentol felesleges: az `ok` a create_document handlerben mar nem hasznalja kozvetlenul.
const EXECUTION_FAILURE_OUTCOMES = new Set(['blocked', 'auth', 'fetch', 'unknown'])

// A HAROM, MINDEN TOOLBAN AZONOS MODON SZAMITHATO MEZO -- a negyedik (contentVerified) tool-
// specifikus marad, mert MI SZAMIT "tartalom"-nak mashol mast jelent (marker-echo egy
// szerkesztesnel, zip-resz egy letrehozasnal, `tartalomIgazolva` egy coedit-irasnal).
function alapMezok(answer) {
  // transportOk: a runJob() NO-JSON-ANSWER es PARSE-ERROR agai a sajat, nevesitett mezoiket
  // hordozzak (exitCode/stderr/stdout, illetve a "could not be parsed" detail) -- ezek jelenlete
  // maga a jel, nem kell kulon zaszlot bevezetni ra.
  const transportOk = !(
    answer.outcome === 'nem-mert' &&
    (answer.exitCode !== undefined || /could not be parsed/.test(answer.detail || ''))
  )
  const executionOk = transportOk && !EXECUTION_FAILURE_OUTCOMES.has(answer.outcome)
  const outputProduced = executionOk && typeof answer.kind === 'string' && answer.kind !== 'ismeretlen'
  return { transportOk, executionOk, outputProduced }
}

// A markerekbol/media-reszekbol szarmaztatott tartalom-jel, `run_builder_script`/`edit_document`
// hasznalja: a szoveg-marker (ha van) VAGY barmely `word/media/` | `xl/media/` | `ppt/media/`
// zip-resz NEMNULLA merettel. Ez fedi az AddImage-piros fixture-t, ahol a marker-mechanizmus
// szerkezetileg vak (kepbeszuras nem ir szoveg-markert), es a resz LETEZIK, de URES.
function markerVagyMediaTartalomVerifikalt(answer, savedBuffer) {
  if (savedBuffer) {
    const entries = listZipEntries(savedBuffer)
    if (!entries) return false // a zip nem olvashato -- NEM-VERIFIKALT, nem hallgatunk el
    const hasMarker = Array.isArray(answer.markersFound) && answer.markersFound.length > 0
    const hasMedia = entries.some((e) => /^(word|xl|ppt)\/media\//.test(e.name) && e.size > 0)
    return hasMarker || hasMedia
  }
  return Array.isArray(answer.markersFound) && answer.markersFound.length > 0
  // Nincs bajt ES nincs marker (xlsx/pptx raw script, bajtok nelkul): MARAD false, ugyanaz az
  // allitas, mint a box-helper.py sajat "kesz-tartalom-nem-ellenorizve" kommentje.
}

function negyallapotuMezok(answer, { savedBuffer } = {}) {
  const alap = alapMezok(answer)
  const contentVerified = alap.outputProduced ? markerVagyMediaTartalomVerifikalt(answer, savedBuffer) : false
  return {
    ...alap,
    contentVerified,
    negyallapotuSiker: alap.transportOk && alap.executionOk && alap.outputProduced && contentVerified,
  }
}

// ---------------------------------------------------------------------------
// CREATE_DOCUMENT -- the operations-table (lib.cjs buildCreateScript/OPERATIONS) exposed as a
// tool for the first time. E0..E7 each added a capability
// to the OPERATIONS table and proved it against buildCreateScript directly, in the core's own
// test file -- none of them wired a caller-facing tool (measured: `git log --all` shows no
// euro-mcp.cjs change in any of those commits). E8 comes last in the darabolás precisely because
// answering "what did NOT apply, and why" needs to already know the whole capability set.

// `ok` a DS/csomag-szintu verdikt marad (SZANDEKOS, teszttel
// rogzitve -- test-tools.cjs [7]), es EZ a fuggveny a mukodesehez nem nyul. Amit `ok` NEM mond meg:
// hogy a KERT muveletek KOZUL mennyi maradt ki (`ok` mar akkor is true, ha csak EGY alkalmazodott).
// Ket kulon mezo adja ezt a hivonak, a muveletek tombjebol szarmaztatva, KOZLES-javitaskent.
//
// *** FRISSITVE: `ok` KEPLETE valtozott (mostantol a
// negy fuggetlen mezo ES-e), DE AZ ERTEKE a nem-kep agon minden
// eddigi tesztesetben (test-tools.cjs [7]/[8]/[9]/[10]) byte-azonos a regivel -- a levezetes ott
// all, ahol az uj keplet szuletik (lasd create_document handler, `ok` sora). A "viselkedese nem
// valtozik" allitas tehat MEGMERT, nem feltetelezett -- de a KEPLET maga mar nem ez a fuggveny.
function muveletOsszegzes(muveletek) {
  const lista = Array.isArray(muveletek) ? muveletek : []
  return {
    mindenMuveletAlkalmazva: lista.length > 0 && lista.every((m) => m.outcome === 'alkalmazva'),
    nemAlkalmazottMuveletSzam: lista.filter((m) => m.outcome !== 'alkalmazva').length,
  }
}

server.tool(
  'create_document',
  'Creates a NEW document (docx/xlsx/pptx) from a caller-described operations list -- text, ' +
    'table, image, chart, and whatever else the OPERATIONS table in lib.cjs knows. Never edits an ' +
    'existing file (use coedit_write for that). Every requested operation is reported ' +
    'individually as applied / not-supported / error, so a partial result can never be mistaken ' +
    'for a complete one, and the SAVED package is checked for internal consistency (do its own ' +
    "internal references resolve?) rather than trusting the service's own success answer alone. " +
    "IMPORTANT: the top-level `ok` field is a SERVER/PACKAGE verdict (did the Document Server " +
    'call and the saved package check out) -- it can be true even when SOME requested operations ' +
    'were not-supported or errored, as long as at least one applied. Check `mindenMuveletAlkalmazva` ' +
    '(boolean) and `nemAlkalmazottMuveletSzam` (count) to see whether the requested BATCH was ' +
    'complete, not just whether the call itself succeeded.',
  {
    core: z.enum(['docx', 'xlsx', 'pptx']),
    operations: z
      .array(z.object({ type: z.string() }).passthrough())
      .min(1)
      .describe(
        'Operation objects, e.g. {type:"text",...}; each is validated by lib.cjs and reported ' +
          'by name, never silently dropped. For xlsx, `op.sheet` (which sheet an operation ' +
          'targets) accepts ONLY a non-negative integer index on this route -- this route builds ' +
          'its own sheets and cannot resolve a name against them. A sheet NAME is only accepted ' +
          'on the co-editing route (coedit_write_operations), which edits an already-existing file.',
      ),
    output_path: z.string().optional().describe('Local path to write the created document to; omit to only get the verdict'),
  },
  async ({ core, operations, output_path }) => {
    // FIRST, before any write work -- this route depends on
    // this file's own code (euro-mcp.cjs) and lib.cjs's buildCreateScript (below); coedit.cjs is
    // not on this path, so it is not checked here.
    const staleGuard = staleCodeGuard(checkOwnFreshness, lib.checkFreshness)
    if (staleGuard) return textResult(staleGuard)
    // Same protocol gate as run_builder_script's own (a tulajdonos rendelkezese, 2026-08-15 11:0x): the
    // DocBuilder route may only CREATE a file that does not exist yet.
    if (output_path && fs.existsSync(output_path)) {
      // KIKOTES 1: az indok is
      // megy, nem csak a bool -- lasd a run_builder_script azonos gate-jenel levo kommentet.
      const coeditUt = coedit.coeditUtAllapota()
      return textResult({
        ok: false,
        outcome: 'protokoll-tilt',
        error:
          `the output already exists: ${output_path}. The DocBuilder route may only CREATE a file ` +
          'that does not exist yet; changing one that does is the co-editing session\'s job. ' +
          'Choose a new output path, or edit the existing file through the editor.',
        coeditUtElerheto: coeditUt.elerheto,
        coeditUtAllapot: coeditUt.indok,
      })
    }

    // *** A MAG: a
    // create_document korabban SOSE adott at documentBase64-et a runJob-nak -- az implicit
    // kiindulo dokumentum docx-kompatibilis csak, ezert xlsx/pptx-nel a preambulum
    // (Api.GetActiveSheet()/Api.GetPresentation()) minden operacio elott elbukott (MERVE:
    // korabban itt allt egy elore-kapu, ami ezt NEVESITVE visszautasitotta -- most mar nem kell,
    // mert a mag megvan). A korabban landolt "mag-hianyzik" elore-kapu
    // ITT KERULT ELTAVOLITASRA, semmi mas nem valtozott azon a helyen.
    //
    // A mag docx-nel NEM kell (az implicit dokumentum mar docx-kompatibilis, ez valtozatlan).
    // xlsx-nel es pptx-nel EGYFORMA minta (ez xlsx-re is
    // kiterjesztette azt, amit a pptx oldal E6 ota mar tud): az operaciok `op.sheet`/`op.slide`
    // indexszel hivatkozhatnak barmelyik lapra/diara (lib.cjs buildCreateScript), a lap-/dia-
    // szamot a hivonak kell megadnia, mert a script NEM tudja sajat magat lekerdezni futas
    // kozben. Itt ez a keresztol vezetett SZAMITOTT ertek, NEM uj sema-mezo: a legnagyobb
    // hivatkozott index + 1 (alapertelmezes 1 lap/dia, ha egyetlen operacio sem hivatkozik
    // explicit indexre -- ez a korabbi, egy-lapos viselkedessel byte-azonos).
    let documentBase64
    let pptxSlideCount
    let xlsxSheetCount
    if (core === 'xlsx') {
      const maxSheetIdx = operations.reduce((max, op) => {
        const idx = Number(op && op.sheet !== undefined && op.sheet !== null ? op.sheet : 0)
        return Number.isFinite(idx) && idx > max ? idx : max
      }, 0)
      xlsxSheetCount = Math.min(maxSheetIdx + 1, 200)
      const sheetNames = Array.from({ length: xlsxSheetCount }, (_, i) => `Munka${i + 1}`)
      documentBase64 = magok.xlsxMag(sheetNames).toString('base64')
    } else if (core === 'pptx') {
      const maxSlideIdx = operations.reduce((max, op) => {
        const idx = Number(op && op.slide !== undefined && op.slide !== null ? op.slide : 0)
        return Number.isFinite(idx) && idx > max ? idx : max
      }, 0)
      pptxSlideCount = Math.min(maxSlideIdx + 1, 200)
      documentBase64 = magok.pptxMag(pptxSlideCount).toString('base64')
    }

    let built
    try {
      built = lib.buildCreateScript({
        core,
        operations,
        outName: output_path ? path.basename(output_path) : undefined,
        slideCount: pptxSlideCount,
        sheetCount: xlsxSheetCount,
      })
    } catch (err) {
      // No operation could be applied at all (MI A KESZ's degenerate case --
      // the mixed-batch example always has at least one applied operation, but an all-refused
      // batch is not a different KIND of failure, just the same report with zero survivors).
      const errMuveletek = Array.isArray(err.report) ? err.report : []
      return textResult({
        ok: false,
        outcome: 'nincs-alkalmazhato-muvelet',
        error: err.message,
        muveletek: errMuveletek,
        ...muveletOsszegzes(errMuveletek),
      })
    }

    // A package-consistency check needs the actual bytes, whether or not the caller asked to
    // keep a copy -- so returnDoc is unconditional here (unlike run_builder_script, where it
    // follows output_path because most callers there only want the verdict).
    const traceId = trace.newTraceId()
    trace.logTrace(traceId, 'mcp-tool-invoked', { tool: 'create_document' })
    const answer = await runJob({ script: built.script, documentBase64, returnDoc: true, traceId })

    let written = null
    let tempPath = null
    let savedBuffer = null
    if (answer.savedBase64) {
      const buf = Buffer.from(answer.savedBase64, 'base64')
      savedBuffer = buf
      if (output_path) {
        fs.writeFileSync(output_path, buf)
        written = output_path
      } else {
        // TRANZIT (CLAUDE.md "Tranzit-fájl higiénia"): csak a konzisztencia-ellenőrzés idejére
        // él, utána azonnal törlődik -- a hívó nem kért tartós másolatot.
        tempPath = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'euro-mcp-create-')), `eredmeny.${core}`)
        fs.writeFileSync(tempPath, buf)
      }
    }

    // A csomag-konzisztencia csak akkor mérhető, ha VAN mit mérni -- a szerver-oldali hiba
    // (blocked/auth/fetch/...) esetén nincs kiírt fájl, ezt nem keverjük össze egy "hibás
    // csomag" verdikttel.
    let csomagKonzisztens = null
    const checkPath = written || tempPath
    if (checkPath) {
      try {
        csomagKonzisztens = checkPackageConsistency(checkPath, core)
      } catch (err) {
        csomagKonzisztens = { ok: false, issues: [`az ellenőrzés maga nem futott le: ${err.message}`] }
      }
    }
    if (tempPath) fs.rmSync(path.dirname(tempPath), { recursive: true, force: true })

    // box-helper.py marker-alapú `ok`/`outcome`-ja a docx ágon a szerkesztésre épül (edit_document
    // maga helyezi el a markert) -- egy create_document hívás SOHA nem tesz markert, tehát ott a
    // nyers `ok` mindig false, `outcome` mindig "nem-valtozott" LENNE, akkor is, ha a létrehozás
    // tökéletesen sikerült. Ezért a tool SAJÁT verdiktet számol, a szerver-szintű hibaágakból
    // (blocked/auth/fetch/unknown/nem-mert/ismeretlen-csomag) és a fenti konzisztencia-ellenőrzésből,
    // nem a nyers `answer.ok`-ból.
    const vanAlkalmazott = built.applied.length > 0

    // NEGYALLAPOTU MEZOK (a tulajdonos 4. pontja: "a regi `ok` mezo MARAD, A NEGY UJ MEZOBOL
    // SZARMAZTATVA"). A `contentVerified` itt
    // NEM a markerVagyMediaTartalomVerifikalt() altalanos szabalyat kapja -- create_document SOHA
    // nem tesz szoveg-markert (lasd lent), es egy tiszta szoveges/tablazatos letrehozasnal a
    // csomag-konzisztencia MAR bizonyitja a tartalmat, media-resz nelkul is. A merce ezert KETAGU:
    // ha a koteg tartalmazott ALKALMAZOTT `image` muveletet, a media-zip-reszt is meg KELL nezni
    // (ez az AddImage-piros fixture sajat esete); egyebkent a MEGLEVO
    // konzisztencia-jel marad a merce.
    //
    // EGYENERTEKUSEG A REGI `ok`-VAL, MERT NEM TALALGATVA: a regi keplet
    // `!dsHiba && vanAlkalmazott && (csomagKonzisztens===null||csomagKonzisztens.ok)` volt.
    // `dsHiba` (DS_FAILURE_OUTCOMES) minden eleme vagy executionOk-ot, vagy outputProduced-et
    // nullazza itt -- tehat `!dsHiba` levezetheto `executionOk && outputProduced`-bol, es a nem-
    // kepes agon `contentVerified` PONTOSAN a regi `vanAlkalmazott && (csomagKonzisztens===null||
    // csomagKonzisztens.ok)`. A negy mezo ES-e ezert a NEM-kepes agon BYTE-AZONOS a regi `ok`-val --
    // ezt a test-tools.cjs [7]/[8]/[9]/[10] (mind a negy eddig VALTOZATLAN) igazolja vissza.
    const alap = alapMezok(answer)
    const hasImageOp = built.report.some((m) => m.type === 'image' && m.outcome === 'alkalmazva')
    let contentVerified
    if (!alap.outputProduced) {
      contentVerified = false
    } else if (hasImageOp) {
      const entries = savedBuffer ? listZipEntries(savedBuffer) : null
      contentVerified = Boolean(entries && entries.some((e) => /^(word|xl|ppt)\/media\//.test(e.name) && e.size > 0))
    } else {
      contentVerified = vanAlkalmazott && (csomagKonzisztens === null || csomagKonzisztens.ok)
    }
    const ok = alap.transportOk && alap.executionOk && alap.outputProduced && contentVerified

    return textResult({
      ok,
      traceId,
      // A negy uj mezo -- lasd a mezo-szerkezet fejlec-kommentjet a fajl elejen. `ok` fentebb
      // MOSTANTOL a negy mezo ES-e (korabbi kikotes szerint); egyenkent is szerepelnek, hogy a hivo
      // lassa, MELYIK ag bukott, ne csak azt, hogy bukott-e.
      transportOk: alap.transportOk,
      executionOk: alap.executionOk,
      outputProduced: alap.outputProduced,
      contentVerified,
      // MI A KIMENET: műveletenként ALKALMAZVA / NEM-TAMOGATOTT / HIBA, és
      // a LÉTREHOZÁS útjának neve mindegyiken -- lib.cjs `report`-ja pontosan ezt adja, változtatás
      // nélkül. Csupa támogatott kérésnél minden bejegyzés 'alkalmazva', tehát a hiánylista
      // szerkezetileg ÜRES, nem hiányzik -- így látszik, hogy a kapu nézte.
      muveletek: built.report,
      // `ok` a DS/csomag-verdikt, es EGY nem-tamogatott
      // muvelet mellett is true tud maradni (lasd fent) -- ez a ket mezo a KOTEG teljesseget
      // adja at kulon, hogy a hivonak ne kelljen a `muveletek` tombot maga vegigszamolnia.
      ...muveletOsszegzes(built.report),
      csomagKonzisztens,
      dsOutcome: answer.outcome,
      dsHibaJelentesEMinositesreNemHasznalhato: answer.outcome === 'nem-valtozott'
        ? 'ez a create_document uton MINDIG igy jelentkezik, marker hianyaban -- lasd fent, ne ez dontse el a sikert'
        : null,
      kind: answer.kind ?? null,
      savedBytes: answer.savedBytes ?? null,
      outputPath: written,
      serverFetches: answer.serverFetches ?? null,
      detail: answer.detail || null,
      // the `export`/`documentStats`
      // operations' results, extracted by box-helper.py from their own marker paragraphs -- null
      // when the batch used neither. UNTRUNCATED, unlike `documentText` above (which stays a
      // 400-char preview for the unrelated EURO-MCP-marker/edit-route convention it already served).
      exportResult: answer.exportResult ?? null,
      stats: answer.stats ?? null,
      customProperties: answer.customProperties ?? null,
    })
  },
)

// ---------------------------------------------------------------------------
// CO-EDITING IRAS -- a MEGLEVO fajlok utja (a tulajdonos rendelkezese, 2026-08-15 10:2x).
//
// A DocBuilder a lemezen ir, session nelkul: ha valaki eppen nyitva tartja a fajlt, az irasunk
// elveszik vagy utkozik, es ezt ELORE NEM TUDJUK MEGNEZNI. Ez az eszkoz ugyanabba a sessionbe ir,
// amit a felhasznalo lat -- ezert a MEGLEVO fajlok ide tartoznak, es a DocBuilder csak letrehozasra.
server.tool(
  'coedit_write',
  'Writes formatted HTML into an ALREADY EXISTING Nextcloud document by opening a real editor ' +
    'session as the CALLING agent. This is the route for files that exist; the DocBuilder tools ' +
    'are for creating a file that does not. Works in all three cores (docx, xlsx, pptx), but the ' +
    'content can only be verified back in the text core -- the answer says which case applies.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    html: z.string().describe('The HTML to insert at the cursor; formatting (bold, colour, tables, lists) is carried through'),
  },
  async ({ file_id, html }) => {
    // *** AZ AZONOSSAG AZ ELSO KAPU, ES FAIL-CLOSED. *** Egy dokumentum-szerkesztesnel a rossz
    // azonossag nem hibauzenet: a dokumentumon MASVALAKI fog szerzokent latszani.
    // A .env betoltese LUSTAN, itt: a tobbi eszkoznek nem kell, es egy modul-szintu betoltes
    // minden hivasnal beolvasna egy titkokat tartalmazo fajlt olyan kodutakon is, amiknek semmi
    // koze hozza. *A titok akkor a legbiztonsagosabb, ha a legkevesebb kodut latja.*
    const envAllapot = coedit.loadBridgeEnv()
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) {
      return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    }
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) {
      // Ha a .env sem toltodott be, azt kulon megnevezzuk: a "hianyzo kulcs" es a "nem is
      // olvastuk a fajlt" ket kulon ok, es a hivo mast tesz a kettore.
      return textResult({
        ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok,
        envBetoltes: envAllapot.ok ? `${envAllapot.betoltve} kulcs` : envAllapot.indok,
      })
    }
    const r = await coedit.writeToDocument({
      url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, html,
    })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      // A HIVO es az IDENTITAS FORRASA a valaszban all: egy kesobbi auditban ez mondja meg,
      // KI irt es MIBOL kovetkezett a neve -- nem egy allitott nev.
      callerId: hivo.id,
      identitasForras: hivo.forras || null,
      // A HITELESITO FORRASA is a valaszban all (a NEVE, sosem az erteke): ket ut van (.env / vault),
      // es egy kesobbi hibanal ez mondja meg, MELYIKET kellene javitani.
      hitelesitoForras: cred.forras || null,
      ncUser: cred.user,
      mag: r.mag ?? null,
      apiHely: r.apiHely ?? null,
      tartalomIgazolva: r.tartalomIgazolva ?? null,
      igazolasIndok: r.igazolasIndok ?? null,
      error: r.indok ?? null,
    })
  },
)

// ---------------------------------------------------------------------------
// The owner's requested `getSelection`/`replaceSelection` capability, listed as
// office.get_selection / office.replace_selection: reads or replaces whatever text is
// CURRENTLY selected in the live, already-open editor, as the CALLING agent -- proven
// cross-core via `Asc.editor.callMethod`, a different
// surface from coedit_write's pluginMethod_PasteHtml and coedit_write_operations'
// editor.callCommand. This tool does not make a selection itself; it reads/replaces whatever
// selection already exists (the human user's, or another co-editing client's).
server.tool(
  'coedit_get_selection',
  "Reads the text currently selected in an ALREADY EXISTING, already-open Nextcloud document, " +
    "as the CALLING agent. Does not select anything itself -- an empty result legitimately " +
    "means nothing is selected right now, not an error. Works across all three cores " +
    "(docx/xlsx/pptx).",
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to read the selection from'),
  },
  async ({ file_id }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await coedit.getSelectionFromDocument({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      identitasForras: hivo.forras || null,
      ncUser: cred.user,
      mag: r.mag ?? null,
      apiHely: r.apiHely ?? null,
      text: r.text ?? null,
      error: r.indok ?? null,
    })
  },
)

server.tool(
  'coedit_replace_selection',
  "Replaces the text currently selected in an ALREADY EXISTING, already-open Nextcloud " +
    "document with the given plain text, as the CALLING agent. Only the SAVED file is trusted " +
    "as proof of the write; the call itself can report success while changing nothing. " +
    "Content verification (a text search in the re-downloaded package) is available in the " +
    "docx core only -- other cores report `tartalomIgazolva: null`.",
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    text: z.string().describe('Plain text to replace the current selection with'),
  },
  async ({ file_id, text }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await coedit.replaceSelectionInDocument({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, text })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      identitasForras: hivo.forras || null,
      ncUser: cred.user,
      mag: r.mag ?? null,
      apiHely: r.apiHely ?? null,
      bytesElotte: r.bytesElotte ?? null,
      bytesUtana: r.bytesUtana ?? null,
      tartalomIgazolva: r.tartalomIgazolva ?? null,
      igazolasIndok: r.igazolasIndok ?? null,
      error: r.indok ?? null,
    })
  },
)

// ---------------------------------------------------------------------------
// The SAME caller-described `operations` list `create_document`
// takes, translated by lib.buildCoeditScript() into an editor.callCommand(...) body and run
// against an ALREADY-OPEN document -- the co-editing route's answer to "what euro-mcp can
// express", not just the DocBuilder create route's. Own tool, not folded into coedit_write's
// `html` parameter: a different input shape (operations vs. HTML) deserves a different schema,
// not a union that would make either caller guess which fields apply.
server.tool(
  'coedit_write_operations',
  'Applies a caller-described operations list (the SAME shape create_document takes -- text, ' +
    'table, image, chart, ...) to an ALREADY EXISTING Nextcloud document, via the live editor\'s ' +
    "own Api (editor.callCommand), as the CALLING agent. Unlike coedit_write's HTML insert, this " +
    'goes through the same operation translator as document creation. Only the SAVED file is ' +
    'trusted as proof -- the call itself can report success while changing nothing. UNLIKE ' +
    'create_document, this route is ALL-OR-NOTHING: every operation is validated before any of ' +
    'them is written (an already-open, co-edited document cannot be safely rolled back if a ' +
    'partial batch had to be undone), so if ANY operation fails, NONE of them are applied. Every ' +
    'requested operation is still reported individually in `muveletek` -- a failed batch names ' +
    'each bad operation (not-supported / error) and marks the operations that would have applied ' +
    'as `nem-alkalmazva` (validated fine on their own, but nothing was written because a sibling ' +
    'failed); check `mindenMuveletAlkalmazva` / `nemAlkalmazottMuveletSzam` to see whether the ' +
    'requested BATCH was complete, not just whether the call itself succeeded.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    core: z.enum(['docx', 'xlsx', 'pptx']),
    operations: z
      .array(z.object({ type: z.string() }).passthrough())
      .min(1)
      .describe('Operation objects, e.g. {type:"text",...} -- same schema as create_document'),
  },
  async ({ file_id, core, operations }) => {
    // FIRST, before any write work -- this route depends on
    // euro-mcp.cjs (this file), coedit.cjs (writeOperationsToDocument), AND lib.cjs
    // (buildCoeditScript, called from inside coedit.cjs) -- all three checked, per-file, because
    // they comb at different times (measured: lib.cjs/euro-mcp.cjs same moment, coedit.cjs over
    // an hour earlier on a real fleet checkout).
    const staleGuard = staleCodeGuard(checkOwnFreshness, coedit.checkFreshness, lib.checkFreshness)
    if (staleGuard) return textResult(staleGuard)
    const envAllapot = coedit.loadBridgeEnv()
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) {
      return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    }
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) {
      return textResult({
        ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok,
        envBetoltes: envAllapot.ok ? `${envAllapot.betoltve} kulcs` : envAllapot.indok,
      })
    }
    const r = await coedit.writeOperationsToDocument({
      url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core, operations,
    })
    return textResult({
      ok: r.ok === true && r.outcome === 'meret-valtozott',
      outcome: r.outcome,
      callerId: hivo.id,
      identitasForras: hivo.forras || null,
      hitelesitoForras: cred.forras || null,
      ncUser: cred.user,
      mag: r.mag ?? null,
      apiHely: r.apiHely ?? null,
      // UNLIKE
      // create_document, this route is all-or-nothing -- buildCoeditScript validates every
      // operation before writing any of them, and if one fails, NONE are applied. A failed batch's
      // `r.report` therefore mixes real failures (nem-tamogatott/hiba) with entries that validated
      // fine but were rejected anyway (`nem-alkalmazva`, because the call never wrote them), so
      // `muveletOsszegzes` on a rejected batch correctly reports `mindenMuveletAlkalmazva: false`
      // for the WHOLE batch, not just the operations that were individually wrong. `r.report` is
      // only present once buildCoeditScript actually ran; a pre-translation failure (missing
      // identity/config/fileId) has no operations to report on, so `muveletek` stays null there
      // rather than a misleading empty/all-false summary.
      muveletek: r.report ?? null,
      ...(r.report ? muveletOsszegzes(r.report) : {}),
      // A MENTETT CSOMAG a bizonyitek, nem a hivas visszaterese (a szulo kartya sajat, ketszer
      // igazolt lelete) -- ezert a valasz a MERT bajt-adatokat is hordozza, nem csak egy verdiktet.
      bytesElotte: r.bytesElotte ?? null,
      bytesUtana: r.bytesUtana ?? null,
      tartalomIgazolva: r.tartalomIgazolva ?? null,
      igazolasIndok: r.igazolasIndok ?? null,
      // The reason this check exists: a `formula` (or
      // `table`) xlsx operation can report `alkalmazva` above while its cell silently never
      // landed (SetValue on a syntactically-"=" but semantically-invalid formula creates nothing,
      // no exception). `cellaEllenorzes` re-reads the SAVED sheet and names, PER OPERATION,
      // whether its requested cell(s) actually exist -- check this, not just `muveletek`, before
      // trusting an xlsx write. null when no operation in this batch targeted a cell (not the
      // same as "verified, nothing missing"); a per-entry `nemMertIndok` means the re-read itself
      // failed -- that is NOT a claim the write failed, only that it could not be confirmed.
      cellaEllenorzes: r.cellaEllenorzes ?? null,
      error: r.indok ?? null,
    })
  },
)

// ---------------------------------------------------------------------------
// office_get_text / office_get_comments (a tulajdonos kerese, egy tool-reteg az Euro-Office
// muveletek fole): getDocumentText es getComments a 15-os listabol. Nem callCommand-on
// mennek (lasd coedit.cjs readDocumentContent fejlec-kommentjet): a MENTETT csomagot olvassak
// WebDAV-on at, egyszeru unzip+szoveg-kinyeressel -- se bongeszo, se Playwright, se elo session
// nem kell hozza. Ket kulon tool, mert a tulajdonos listaja is ket kulon muveletkent nevezi oket, bar a
// belso megvalositas ugyanazt az egy letoltest hasznalja mindkettonel.
//
// *** A HOZZAFERES-KAPU MA UGYANAZ, MINT A COEDIT-IRAS UTON, SZANDEKOSAN, NEM SZUKSEGSZERUEN: ***
// a `detectCallerId()` (coedit.cjs) az EURO_COEDIT_AGENTS allowlistet nezi FUGGETLENUL attol,
// hogy a hivo irni vagy csak olvasni akar-e -- ez a fuggveny kozos minden coedit*-uttal. Egy
// WebDAV-olvasas strukturalisan NEM hordozza azt a kockazatot (tobb egyideju szerkeszto
// utkozese), amiert ez az allowlist letezik -- DE ennek a kapunak a lazitasa UGYANAZ az
// owner-dontes, ami a masik 2 (getSelection/replaceSelection) muvelet miatt ma is fut. Ezert ez a ket uj tool MA tudatosan a SZIGORUBB, meglevo kaput
// oroklte, nem egy kulon, lazabb utat -- ha ez felesleges korlatozasnak bizonyul, az KULON,
// nevesitett kerdes az owner fele, nem valami, amit ez az egyseg csendben eldontott.
server.tool(
  'office_get_text',
  'Reads the full text content of an ALREADY EXISTING Nextcloud document (docx and pptx today ' +
    '-- xlsx is not yet supported, see `outcome`). For ' +
    'docx, `bekezdesek` is one string per paragraph in document order. For pptx, `bekezdesek` is ' +
    'the SAME flat shape (every paragraph across every slide, in slide DISPLAY order -- a moved ' +
    'slide keeps its underlying part name but this list follows its new position), and `diak` ' +
    'additionally breaks it out per slide ({index, bekezdesek}) since a pptx caller usually cares ' +
    'which slide text is on. docx also gets `labjegyzetek`/`vegjegyzetek` ({id, text} footnotes/ ' +
    'endnotes), `konyvjelzok` ({name, text} bookmarks), and `tablazatok` ({index, parentIndex, ' +
    'sorok} tables including nested ones; all four are null on ' +
    'pptx, same as `diak` is null on docx). Reads the LAST-SAVED state via WebDAV, not live in-editor keystrokes ' +
    '-- if the document is open in a co-editing session, a save may lag by up to ~1 minute behind ' +
    'what a live user is typing (same lag coedit_write_operations already documents for its own writes).',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to read'),
    core: z.enum(['docx', 'xlsx', 'pptx']),
  },
  async ({ file_id, core }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await coedit.readDocumentContent({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      bekezdesek: r.bekezdesek ?? null,
      // Null on docx (no per-slide concept there), the per-slide
      // breakdown on pptx -- not merged into `bekezdesek` itself, so an existing docx-only caller
      // reading that one field sees no shape change from this addition.
      diak: r.diak ?? null,
      // Docx-only, null on pptx (readPptxContent does not set these
      // fields at all -- same "new field, old caller sees no shape change" reasoning as `diak`
      // above, just the other core). `konyvjelzok[].text` is null when the matching bookmarkEnd
      // was not found (a malformed/externally-edited package), not an empty-bookmark signal.
      // `tablazatok[].parentIndex` is null for a top-level table, or another entry's own `index`
      // for a table nested inside that entry -- a cell that only WRAPS a nested table reports an
      // EMPTY string for its own text (the nested table's text lives at its own array entry, not
      // duplicated onto the parent).
      labjegyzetek: r.labjegyzetek ?? null,
      vegjegyzetek: r.vegjegyzetek ?? null,
      konyvjelzok: r.konyvjelzok ?? null,
      tablazatok: r.tablazatok ?? null,
      // "mukodik" here means "reads the LAST-SAVED
      // package", not "reads what a live co-editor is typing right now" -- a caller assuming live
      // state would otherwise get a SUCCESSFUL call with STALE data, silently. This field and the
      // freshness note travel WITH every answer, not just in the tool description prose (a caller
      // reading only the response, not the schema text, must still see the limit).
      olvasasForrasa: r.ok ? 'mentett-csomag' : null,
      frissessegKikotes: r.ok
        ? 'ez a MENTETT csomag allapota, nem a live szerkeszto EPP MOST allapota -- ha valaki masik ' +
          'egyideju szerkesztoje az utolso mentes ota irt, azt EZ a valasz meg nem latja'
        : null,
      bytesOlvasva: r.bytesOlvasva ?? null,
      error: r.indok ?? null,
    })
  },
)

server.tool(
  'office_get_comments',
  'Reads all comments on an ALREADY EXISTING Nextcloud document (docx and pptx today, same core ' +
    'limit as office_get_text) -- each as {id/idx, author, date, text} (docx) or ' +
    '{idx, authorId, authorName, date, text, slideIndex} (pptx: pptx ' +
    'comments have no document-wide id, only a per-slide `idx`, and the author is a numeric ' +
    '`authorId` resolved against a separate part -- `authorName` is null if that part could not ' +
    'be read, which is NOT the same as an anonymous comment). `id`/`idx` is the OOXML value from ' +
    'the saved package, NOT the live editor Comment.GetId() value (the two are different ' +
    'numbering schemes, measured) -- do not use this id to target ' +
    'a comment through a different, live-session API.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to read'),
    core: z.enum(['docx', 'xlsx', 'pptx']),
  },
  async ({ file_id, core }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await coedit.readDocumentContent({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      kommentek: r.kommentek ?? null,
      // Same caveat as office_get_text's own field
      // -- see that tool's comment for why this travels in the response, not just the schema text.
      olvasasForrasa: r.ok ? 'mentett-csomag' : null,
      frissessegKikotes: r.ok
        ? 'ez a MENTETT csomag allapota, nem a live szerkeszto EPP MOST allapota -- egy epp most ' +
          'hozzaadott komment, amit meg nem mentettek, ITT MEG NEM latszik'
        : null,
      error: r.indok ?? null,
    })
  },
)

server.tool(
  'office_find',
  'Searches the text of an ALREADY EXISTING Nextcloud document (docx and pptx today, same core ' +
    'limit as office_get_text) for a literal, case-SENSITIVE substring -- reports how many ' +
    'times it occurs and in which paragraphs (by index, with that paragraph\'s own text), not ' +
    'a character offset (there is no stable offset across saves/reflows to report). On pptx, the ' +
    '`paragraphIndex` is a GLOBAL index across every slide\'s flattened text, ' +
    'not a per-slide one -- use office_get_text\'s `diak` breakdown first ' +
    'if you need to know which slide a match is on. Reads the LAST-SAVED state via WebDAV, same ' +
    'freshness limit as office_get_text/office_get_comments.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to search'),
    core: z.enum(['docx', 'xlsx', 'pptx']),
    query: z.string().min(1).describe('Literal text to search for (case-sensitive, no regex)'),
  },
  async ({ file_id, core, query }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await coedit.readDocumentContent({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core })
    const found = r.ok ? lib.findMatchesInParagraphs(r.bekezdesek, query) : null
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      totalCount: found ? found.totalCount : null,
      matches: found ? found.matches : null,
      olvasasForrasa: r.ok ? 'mentett-csomag' : null,
      frissessegKikotes: r.ok
        ? 'ez a MENTETT csomag allapota, nem a live szerkeszto EPP MOST allapota'
        : null,
      error: r.indok ?? null,
    })
  },
)

server.tool(
  'office_get_slide_contents',
  'pptx-only: reports what objects (shapes, images, tables, charts, groups) already exist on ' +
    'each slide of an ALREADY EXISTING Nextcloud presentation (the ' +
    '"dia-tartalom visszaolvasasa" item from the owner\'s request list -- reports the SAME information Slide.GetAllShapes/' +
    'GetAllCharts/GetAllTables/GetAllImages would, but NOT by calling those DocBuilder Api ' +
    'methods: this route parses the SAVED PACKAGE\'s raw slide XML directly (no live editor ' +
    'session, no engine call), same architecture and freshness limit as office_get_text). ' +
    '`diak[i].tartalom` gives the per-slide breakdown; `tartalomOsszesen` gives ' +
    'the SAME objects flattened across every slide, each tagged with its own `slideIndex` -- use ' +
    'whichever shape fits the question. Each shape/image/table/chart entry carries {id, name, x, ' +
    'y, cx, cy} (EMU units, same as every other position/size field in this tool layer); shapes ' +
    'additionally carry `shapeType`. `oleObjects` is ALWAYS null with a reason: this route has no ' +
    'way yet to create an OLE object to measure the detection against, so a claimed zero-count ' +
    'would look identical to "not attempted" -- do not read null as "none present". A shape\'s ' +
    'x/y INSIDE a group is relative to the GROUP\'s own child coordinate space, not the slide -- ' +
    'comparing a grouped shape\'s position against an ungrouped one without accounting for its ' +
    'enclosing group will give the wrong answer. `diak[i].elrendezesEsTema` covers the ' +
    '"elrendezes/tema lekerdezese" item (the same information Slide.GetLayout/GetTheme would ' +
    'give, again without calling them) -- {layoutName, layoutType, themeName}, resolved by ' +
    'following the slide -> layout -> master -> theme rels chain through the saved package\'s ' +
    'own XML parts; any field is null if that hop could not be resolved (a slide with no layout ' +
    'link, or a layout part that could not be read), not a thrown error.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the presentation to read'),
  },
  async ({ file_id }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await coedit.readDocumentContent({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core: 'pptx' })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      diak: r.diak ? r.diak.map((d) => ({ index: d.index, tartalom: d.tartalom, elrendezesEsTema: d.elrendezesEsTema })) : null,
      tartalomOsszesen: r.tartalomOsszesen ?? null,
      oleObjects: null,
      oleObjectsNemMertIndok: r.ok
        ? 'nincs bekotott muvelet, ami OLE-objektumot hoz letre ezen az uton -- a felismeres SOHA nem lett elo bemeneten futtatva, egy 0-t allito valasz megkulonboztethetetlen lenne egy "nem probaltam"-tol'
        : null,
      olvasasForrasa: r.ok ? 'mentett-csomag' : null,
      frissessegKikotes: r.ok
        ? 'ez a MENTETT csomag allapota, nem a live szerkeszto EPP MOST allapota'
        : null,
      error: r.indok ?? null,
    })
  },
)

server.tool(
  'office_get_document_metadata',
  'pptx-only today: reads document-level metadata (title/subject/creator/lastModifiedBy/created/' +
    'modified/revision/category from docProps/core.xml, application/slides/hiddenSlides/notes/' +
    'words/totalTime from docProps/app.xml) of an ALREADY EXISTING Nextcloud presentation, from ' +
    'the SAVED PACKAGE (the "dokumentum-metaadat" item). *** THIS ROUTE ' +
    'EXISTS BECAUSE THE LIVE API ROUTE IS BROKEN, MEASURED: *** oPresentation.GetDocumentInfo() ' +
    'THROWS on this Document Server instance -- this tool reads the same underlying XML parts ' +
    'directly instead. A field that is null means the tag was ABSENT from the package (never ' +
    'set); an empty string means the tag was PRESENT but empty (explicitly cleared, or the ' +
    'default this Document Server writes for an author-less save) -- the two are different facts, ' +
    'not collapsed into one. GetCustomProperties/GetCustomXmlParts are NOT covered by this tool ' +
    '(unmeasured -- no bound operation creates a custom property to test the ' +
    'populated shape against).',
  {
    file_id: z.string().describe('The Nextcloud fileId of the presentation to read'),
  },
  async ({ file_id }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const r = await coedit.readDocumentContent({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core: 'pptx' })
    return textResult({
      ok: r.ok === true,
      outcome: r.outcome,
      callerId: hivo.id,
      metaadat: r.metaadat ?? null,
      olvasasForrasa: r.ok ? 'mentett-csomag' : null,
      frissessegKikotes: r.ok
        ? 'ez a MENTETT csomag allapota, nem a live szerkeszto EPP MOST allapota'
        : null,
      error: r.indok ?? null,
    })
  },
)

// CALCULATES, never writes -- named office_* like its read
// siblings above, not folded into lib.cjs's OPERATIONS table (that table is uniformly a
// document-WRITE translator, 62 entries, shared by buildCreateScript/buildCoeditScript; this
// would be its first non-writing entry, touching three call routes for one operation --
// a deliberate design decision). MEASURED (live DocBuilder probe against a
// disposable seed document, run_builder_script): Api.Intersect(oWorksheet.GetRange(range1),
// oWorksheet.GetRange(range2)).GetAddress() returns the rectangular overlap address (e.g.
// "A1:C3" ∩ "B2:D4" -> "B2:C3"); when the two ranges do not overlap at all, the call THROWS a
// generic `Error: Ranges do not intersect.` (name "Error", not a distinguishing subclass).
// Since the computation is pure address geometry -- it never reads or writes actual cell DATA,
// only the two range addresses -- it does not need the CALLER's document at all (an earlier design
// stipulation: do not ask for file_id "for symmetry" when nothing ties the call to one). This
// tool instead opens its OWN throwaway seed (magok.xlsxMag) on every call, exactly the way
// run_builder_script's own "omit document_path" default was measured NOT to work for xlsx (a
// bare generated minimal document is not xlsx-shaped -- Api.GetActiveSheet() has nothing to
// return on it, and the whole job comes back `outcome:"blocked"`, unhandled, before any
// try/catch in the script even runs). Still a REAL engine call, not a local reimplementation of
// rectangle math -- if the Document Server's own Intersect ever diverges from naive geometry on
// some address form (non-contiguous selections, whole-column/-row ranges, etc.), this tool
// reports what the engine actually said, not what a caller might assume.
server.tool(
  'office_range_intersect',
  'CALCULATES the rectangular overlap of two xlsx cell-range addresses (e.g. "A1:C3" and ' +
    '"B2:D4" -> "B2:C3") via the real Document Server engine (Api.Intersect) -- it does NOT ' +
    'read or write any actual document; the two addresses are the entire input, and no Nextcloud ' +
    'file is touched or required. When the ranges do NOT overlap, that is a valid, EXPECTED ' +
    'result, not a tool failure -- it comes back as `ok:false, outcome:"nincs-metszet"` (the ' +
    "engine's own message travels in `error`), not a thrown exception, so check `outcome` rather " +
    'than treating any ok:false as broken. Both addresses must name ranges on the SAME sheet ' +
    '(this route only ever opens one single-sheet throwaway document) -- multi-sheet references ' +
    'are unmeasured and refused before the engine is even asked.',
  {
    range1: z.string().min(1).describe('First xlsx range address, e.g. "A1:C3"'),
    range2: z.string().min(1).describe('Second xlsx range address, e.g. "B2:D4"'),
  },
  async ({ range1, range2 }) => {
    if (/!/.test(range1) || /!/.test(range2)) {
      return textResult({
        ok: false,
        outcome: 'tobb-lap-nem-tamogatott',
        error: 'sheet-qualified references (a "!" in the address) are not supported -- this route only ever opens one single-sheet throwaway document, so a cross-sheet intersection has never been measured against it',
        address: null,
      })
    }
    const seed = magok.xlsxMag(['Munka1'])
    const marker = lib.makeMarker()
    const script = [
      'builder.OpenFile("__DOC_URL__");',
      'var oWorksheet = Api.GetActiveSheet();',
      'try {',
      `  var r1 = oWorksheet.GetRange(${JSON.stringify(range1)});`,
      `  var r2 = oWorksheet.GetRange(${JSON.stringify(range2)});`,
      '  var inter = Api.Intersect(r1, r2);',
      `  oWorksheet.GetRange("Z1").SetValue(${JSON.stringify(marker)} + "|OK|" + inter.GetAddress());`,
      '} catch (e) {',
      `  oWorksheet.GetRange("Z1").SetValue(${JSON.stringify(marker)} + "|ERR|" + e.message);`,
      '}',
      'builder.SaveFile("xlsx", "eredmeny.xlsx");',
      'builder.CloseFile();',
    ].join('\n')

    const traceId = trace.newTraceId()
    trace.logTrace(traceId, 'mcp-tool-invoked', { tool: 'office_range_intersect' })
    const answer = await runJob({ script, documentBase64: seed.toString('base64'), returnDoc: true, traceId })
    if (answer.ok !== true || !answer.savedBase64) {
      return textResult({
        ok: false,
        outcome: answer.outcome || 'nem-mert',
        error: answer.detail || 'the job did not complete or returned no document',
        address: null,
        traceId,
      })
    }
    // lib.cjs stays pure (no fs/zip access, per its own header) -- the unzip happens here, the
    // SAME temp-file + `unzip -p` pattern coedit.cjs already uses for its own xlsx cell reads.
    const tmpFile = path.join(os.tmpdir(), `euro-mcp-intersect-${process.pid}-${Date.now()}.xlsx`)
    fs.writeFileSync(tmpFile, Buffer.from(answer.savedBase64, 'base64'))
    let sheetXml = null
    let sharedStringsXml = null
    try {
      sheetXml = execFileSync('unzip', ['-p', tmpFile, 'xl/worksheets/sheet1.xml'], { encoding: 'utf8' })
      try {
        sharedStringsXml = execFileSync('unzip', ['-p', tmpFile, 'xl/sharedStrings.xml'], { encoding: 'utf8' })
      } catch { /* no string cells at all -- the marker write itself would then be absent too */ }
    } catch {
      sheetXml = null
    } finally {
      fs.rmSync(tmpFile, { force: true })
    }
    const cell = lib.resolveXlsxCellText(sheetXml, sharedStringsXml, 'Z1')
    if (cell === null || !cell.startsWith(marker + '|')) {
      // The marker itself is missing or does not match -- the script ran, saved a document, but
      // this specific call's own write cannot be told apart from a stale/unrelated one.
      return textResult({ ok: false, outcome: 'nem-mert', error: 'the marker cell (Z1) did not contain this call\'s own marker -- cannot attribute the saved content to this request', address: null, traceId })
    }
    const rest = cell.slice(marker.length + 1)
    if (rest.startsWith('OK|')) {
      return textResult({ ok: true, outcome: 'metszet-van', address: rest.slice(3), error: null, traceId })
    }
    return textResult({ ok: false, outcome: 'nincs-metszet', address: null, error: rest.slice(4), traceId })
  },
)

// The shared body office_insert_text / office_insert_table run -- identity,
// credentials, ONE coedit_write_operations-shaped call, ONE response shape. Not a new
// capability: `text`/`table` are EXISTING entries in lib.cjs's OPERATIONS table, already
// reachable through coedit_write_operations today. the owner's request was ergonomics ("the agent
// should not have to handle the raw Euro-Office API directly") -- a named tool with named
// parameters, not a new underlying write path. Extracted once other office_* write tools would
// otherwise repeat this same five-line block a third and fourth time.
async function egyMuveletesCoeditHivas({ file_id, core, operation }) {
  const hivo = coedit.detectCallerId()
  if (!hivo.ok) return { ok: false, outcome: 'azonossag-hiany', callerId: null, error: hivo.indok }
  const cred = await coedit.credentialsFor(hivo.id)
  if (!cred.ok) return { ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok }
  const r = await coedit.writeOperationsToDocument({
    url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core, operations: [operation],
  })
  return {
    ok: r.ok === true && r.outcome === 'meret-valtozott',
    outcome: r.outcome,
    callerId: hivo.id,
    muveletek: r.report ?? null,
    bytesElotte: r.bytesElotte ?? null,
    bytesUtana: r.bytesUtana ?? null,
    tartalomIgazolva: r.tartalomIgazolva ?? null,
    igazolasIndok: r.igazolasIndok ?? null,
    // xlsx `table`/`formula` writes get a PER-CELL
    // presence check (a client-side "ok:true" alone does not prove a cell landed).
    // null when not applicable (e.g. a docx/pptx write), NOT the same
    // as "verified nothing missing".
    cellaEllenorzes: r.cellaEllenorzes ?? null,
    error: r.indok ?? null,
  }
}

server.tool(
  'office_insert_text',
  'Inserts a new paragraph of text into an ALREADY EXISTING Nextcloud document (docx: a real ' +
    'paragraph; pptx: a text box on the given slide -- xlsx is not supported, use ' +
    'coedit_write_operations with a table operation for cell values). Covers BOTH insertText ' +
    "and insertParagraph from the owner's operation list: on this route a paragraph IS the unit of " +
    'inserted text, there is no separate inline-without-a-paragraph form. Same all-or-nothing ' +
    'write contract as coedit_write_operations: nothing is ' +
    'written if the operation is refused.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    core: z.enum(['docx', 'pptx']),
    text: z.string().describe('The text to insert'),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    align: z.enum(['left', 'right', 'center', 'both']).optional().describe('docx only'),
    heading: z.number().int().min(1).max(9).optional().describe('docx only -- maps to the built-in "Heading N" style'),
    slide: z.number().int().min(0).optional().describe('pptx only -- target slide index, default 0'),
  },
  async ({ file_id, core, text, bold, italic, align, heading, slide }) => {
    const operation = { type: 'text', text, bold, italic, align, heading, slide }
    return textResult(await egyMuveletesCoeditHivas({ file_id, core, operation }))
  },
)

server.tool(
  'office_insert_table',
  'Inserts a new table into an ALREADY EXISTING Nextcloud document (docx, pptx, xlsx all ' +
    'supported). Same all-or-nothing write contract as coedit_write_operations: nothing is ' +
    'written if the operation is refused.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    core: z.enum(['docx', 'pptx', 'xlsx']),
    rows: z.array(z.array(z.string())).min(1).describe('Row-major cell text, e.g. [["a","b"],["c","d"]]'),
    header: z.boolean().optional().describe('Style the first row as a header (default true)'),
    at: z.string().optional().describe('xlsx only -- top-left cell, e.g. "B3" (default A1)'),
    slide: z.number().int().min(0).optional().describe('pptx only -- target slide index, default 0'),
  },
  async ({ file_id, core, rows, header, at, slide }) => {
    const operation = { type: 'table', rows, header, at, slide }
    return textResult(await egyMuveletesCoeditHivas({ file_id, core, operation }))
  },
)

server.tool(
  'office_replace',
  'Replaces EVERY occurrence of a literal, case-sensitive text in an ALREADY EXISTING Nextcloud ' +
    'document (docx only -- MEASURED absent on xlsx/pptx: neither ' +
    'has the underlying SearchAndReplace/Search methods this operation needs). Same ' +
    'all-or-nothing write contract as coedit_write_operations: nothing is written if the ' +
    'operation is refused. Use office_find first if you need to know how many occurrences exist ' +
    'before replacing them.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    core: z.literal('docx'),
    search: z.string().min(1).describe('Literal text to find (case-sensitive, no regex)'),
    replace: z.string().optional().describe('Replacement text (default: empty string, i.e. delete the match)'),
  },
  async ({ file_id, core, search, replace }) => {
    const operation = { type: 'replaceText', search, replace }
    return textResult(await egyMuveletesCoeditHivas({ file_id, core, operation }))
  },
)

server.tool(
  'office_add_comment',
  'Adds a comment anchored to the FIRST occurrence of a literal, case-sensitive text in an ' +
    'ALREADY EXISTING Nextcloud document (docx only). *** UNLIKE the other office_* write tools, ' +
    'a successful `ok:true` here does NOT by itself prove the comment landed: whether ' +
    '`anchorText` exists in this document can only be known once the script runs inside the ' +
    'live editor, so a missing anchor is a silent no-op at the translation level, ' +
    'not a batch failure. *** This tool ' +
    'compensates by re-reading the SAVED document after the write and reporting whether the ' +
    'comment text is now actually present (`ellenorizve`) -- check that field, not just `ok`, ' +
    'before trusting the comment was added.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    core: z.literal('docx'),
    anchorText: z.string().min(1).describe('Literal text to anchor the comment to (case-sensitive, first match only)'),
    text: z.string().min(1).describe('The comment text'),
    author: z.string().optional().describe('Comment author name (default: "euro-mcp")'),
  },
  async ({ file_id, core, anchorText, text, author }) => {
    const hivo = coedit.detectCallerId()
    if (!hivo.ok) return textResult({ ok: false, outcome: 'azonossag-hiany', error: hivo.indok })
    const cred = await coedit.credentialsFor(hivo.id)
    if (!cred.ok) return textResult({ ok: false, outcome: 'konfig-hiany', callerId: hivo.id, error: cred.indok })
    const operation = { type: 'addComment', anchorText, text, author }
    const write = await coedit.writeOperationsToDocument({
      url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core, operations: [operation],
    })
    const ok = write.ok === true && write.outcome === 'meret-valtozott'
    // The readback must
    // NEVER become a precondition for reporting the write's own outcome -- if the readback
    // itself fails, that is a NEM-MERT verification, not a failed write (the write already
    // happened, or didn't, independently of whether we can currently re-read the file).
    let ellenorizve = null
    let ellenorizesIndok = 'a beolvasas kihagyva, mert az iras maga nem sikerult'
    if (ok) {
      const readBack = await coedit.readDocumentContent({ url: cred.url, user: cred.user, pass: cred.pass, fileId: file_id, core })
      if (readBack.ok) {
        ellenorizve = readBack.kommentek.some((k) => k.text === text)
        ellenorizesIndok = ellenorizve
          ? 'a mentett komment-lista tartalmazza a kuldott szoveget'
          : 'a mentett komment-listaban NEM talalhato a kuldott szoveg -- valoszinu ok: az anchorText nem volt megtalalhato a dokumentumban'
      } else {
        ellenorizesIndok = `NEM-MERT -- az iras utani visszaolvasas maga bukott: ${readBack.indok}`
      }
    }
    return textResult({
      ok,
      outcome: write.outcome,
      callerId: hivo.id,
      ellenorizve,
      ellenorizesIndok,
      muveletek: write.report ?? null,
      bytesElotte: write.bytesElotte ?? null,
      bytesUtana: write.bytesUtana ?? null,
      error: write.indok ?? null,
    })
  },
)

server.tool(
  'office_set_tracked_changes',
  'Turns Track Changes (review mode) on or off for an ALREADY EXISTING Nextcloud document (docx ' +
    'only) -- MEASURED: `oDocument.SetTrackRevisions(...)` produces ' +
    'real <w:ins>/<w:del> elements in the saved package for subsequent edits. This is a ' +
    'document-wide FLAG, not a per-edit choice -- turning it on here affects every edit made ' +
    '(by any editor, human or agent) AFTER this call, until it is turned off again. Same ' +
    'all-or-nothing write contract as coedit_write_operations: nothing is written if the ' +
    'operation is refused.',
  {
    file_id: z.string().describe('The Nextcloud fileId of the document to write into'),
    core: z.literal('docx'),
    enabled: z.boolean().optional().describe('Turn tracking on (default) or off (false)'),
  },
  async ({ file_id, core, enabled }) => {
    const operation = { type: 'trackedChanges', enabled }
    return textResult(await egyMuveletesCoeditHivas({ file_id, core, operation }))
  },
)

if (require.main === module) {
  const transport = new StdioServerTransport()
  server.connect(transport)
}

module.exports = { server }
