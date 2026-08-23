#!/usr/bin/env node
// PURE formatter/gate for euro-office-capabilities.json.
// Deliberately split from the live-reflection GATHERING step (that needs a running Document
// Server session, only reachable through an agent's own MCP tool calls -- coedit_write_operations
// / run_builder_script -- not from a plain Node script). This file does the part that CAN be pure
// and self-tested without a live instance: shape the registry, and refuse to mark anything
// "verified" that the reflection did not actually confirm exists.
//
// the owner's requirement: "Ne maradhasson olyan
// komment, hogy MERVE MUKODIK, ha kozben nem ugyanazzal a szignaturaval futott a teszt." -- the
// concrete case that motivated this: lib.cjs used to call AddProtectedRange(range, title), backwards,
// while its own comment said "MEASURED, working". This gate exists so that mistake has a mechanical
// check, not just a human re-reading the comment next to the code.
//
// HOW TO GATHER FRESH RAW SIGNATURES (the live half this file does NOT do):
//   1. mcp__euro-mcp__create_document (core: xlsx, a trivial `table` op) -> local seed .xlsx
//   2. mcp__euro-mcp__create_document (core: pptx, a trivial `text` op) -> local seed .pptx
//   3. mcp__euro-mcp__run_builder_script against the xlsx seed:
//        builder.OpenFile("__DOC_URL__");
//        var oWorksheet = Api.GetActiveSheet();
//        oWorksheet.GetRange("Z1").SetValue(String(oWorksheet.Move));
//        oWorksheet.GetRange("Z2").SetValue(String(oWorksheet.AddProtectedRange));
//        oWorksheet.GetRange("Z3").SetValue(String(oWorksheet.GetRange("A1").Paste));
//        builder.SaveFile("xlsx", "eredmeny.xlsx");
//        builder.CloseFile();
//      then unzip the OUTPUT (not the tool's own JSON report -- "a callCommand/DocBuilder valasza
//      semmit nem bizonyit") and read xl/sharedStrings.xml for the reflected function source.
//   4. Same idea for pptx/oSlide.ApplyLayout, reading ppt/slides/slideN.xml <a:t> text instead.
//   5. ALWAYS include one KNOWN-GOOD control method (e.g. GetRange) alongside the disputed ones,
//      and one DELIBERATELY NONEXISTENT method name -- reflecting a missing method gives the
//      literal string "undefined", not an exception; a generator that does not check for this
//      would mark a typo'd/removed method "verified: true".
//
// This module takes the raw `String(obj.Method)` outputs as input (already gathered live) and
// produces the registry -- it never calls a network endpoint itself.

const UNDEFINED_MARKER = 'undefined'

// A raw reflection string counts as "verified" only if it looks like an actual function body --
// not the literal "undefined" (JS's own signal for "no such property"), not empty/whitespace, and
// not some other falsy stringification (null -> "null", NaN -> "NaN" -- neither is a function).
export function isGenuineFunctionSource(raw) {
  if (typeof raw !== 'string') return false
  const trimmed = raw.trim()
  if (!trimmed) return false
  if (trimmed === UNDEFINED_MARKER || trimmed === 'null' || trimmed === 'NaN') return false
  // A real reflected method body starts with "function" (this engine's Function.prototype.toString
  // never returns an arrow-function form for these Api methods -- measured on all 5 reflected
  // methods today, all "function(...){...}"). A stricter check than "non-empty" on purpose: an
  // engine change that started returning e.g. "[native code]" for a built-in should NOT pass this
  // as if it were a real, inspectable signature.
  return /^function\s*\(/.test(trimmed)
}

// `rawSignatures`: { "ApiWorksheet.Move": "<raw String(fn) output>", ... }
// `nowIso`: caller-supplied timestamp (this module must not call Date.now()/new Date() itself --
// callers running inside a workflow script may be in a context where those throw; keeping the
// clock OUTSIDE this pure function also makes it trivially testable without a live clock).
export function buildMethodsRegistry(rawSignatures, nowIso) {
  if (!rawSignatures || typeof rawSignatures !== 'object') {
    throw new Error('buildMethodsRegistry: rawSignatures must be an object of {methodName: rawString}')
  }
  if (!nowIso) throw new Error('buildMethodsRegistry: nowIso is required (caller-supplied, not generated here)')
  const methods = {}
  for (const [name, raw] of Object.entries(rawSignatures)) {
    const genuine = isGenuineFunctionSource(raw)
    methods[name] = {
      signature: raw,
      verified: genuine,
      verifiedAt: genuine ? nowIso : null,
    }
  }
  return methods
}

// `buildFields`: { packageVersion: 'NEM-MERT'|<string>, ... } -- each of the 5 fields the owner named.
// `nemMertOkok`: { packageVersion: "reason", ... } for any field whose value is 'NEM-MERT'.
// Every 'NEM-MERT' field MUST carry a reason -- a three-state field (real value | NEM-MERT+reason |
// missing entirely) is what this project's own convention requires (see office-diag-e5-statusz-
// taxonomia's fingerprint-field rule: "a `null` mint 'nincs adat' NEM megengedett ertek, mert egy
// osszehasonlitas egyenlonek latja" -- the same reasoning applies here, so 'NEM-MERT' is spelled
// out as a literal, not left as `null`).
export function buildBuildFingerprint(buildFields, nemMertOkok) {
  const FIELDS = ['packageVersion', 'imageId', 'wordSdkHash', 'cellSdkHash', 'slideSdkHash']
  const out = {}
  for (const f of FIELDS) {
    if (!(f in buildFields)) throw new Error(`buildBuildFingerprint: missing required field '${f}'`)
    const v = buildFields[f]
    out[f] = v
    if (v === 'NEM-MERT') {
      const reason = nemMertOkok && nemMertOkok[f]
      if (!reason) throw new Error(`buildBuildFingerprint: '${f}' is NEM-MERT but has no named reason`)
    }
  }
  return out
}

export function buildRegistry({ rawSignatures, buildFields, nemMertOkok, nowIso, generatedBy, source }) {
  return {
    generatedAt: nowIso,
    generatedBy: generatedBy || 'unknown',
    source: source || 'unknown',
    build: buildBuildFingerprint(buildFields, nemMertOkok),
    buildNemMertOkok: nemMertOkok || {},
    methods: buildMethodsRegistry(rawSignatures, nowIso),
  }
}
