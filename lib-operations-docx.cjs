'use strict'
// This file owns the DOCX-only OPERATIONS entries (extracted verbatim -- same emit()/validate()
// bodies, same comments, same order). The multi-core entries that also touch docx (per each
// one's own `cores` array) are pulled in from their own individual modules under operations/,
// not duplicated here. lib.cjs remains the aggregator: it populates
// lib-operations-registry.cjs with every helper function/const these files may need BEFORE
// requiring any of them, so this file (and its siblings) can destructure from a plain,
// already-populated object -- no circular require back into lib.cjs, and no function wrapper
// around the entries themselves (a wrapper would roll every nested emit()'s complexity up into
// one artificially huge factory-function score, which is exactly what the first draft of this
// split did before qlty smells caught it).
const { jsString, notSupportedError, validateRgbColor } = require('./lib-operations-registry.cjs')
const text = require('./operations/text.cjs')
const table = require('./operations/table.cjs')
const image = require('./operations/image.cjs')
const shape = require('./operations/shape.cjs')
const wordArt = require('./operations/wordArt.cjs')
const chart = require('./operations/chart.cjs')
const pageSetup = require('./operations/pageSetup.cjs')
const runs = require('./operations/runs.cjs')
module.exports = {
  text,
  table,
  image,
  shape,
  wordArt,
  chart,
  pageSetup,
  runs,

  save: {
    cores: ['docx'],
    // Explicit save trigger, separate from the
    // implicit one every write already performs. MEASURED 2026-08-17 (live co-editing
    // route, coedit_write_operations): `Api.Save` IS a function in this callCommand context
    // (typeof-probed: `Api.asc_Save` and `oDocument.Save` do NOT exist here -- only `Api.Save`
    // does, confirmed via `Object.keys(Api)` filtered on /save/i, package-verified in the saved
    // document text), and calling it does not throw (also package-verified: a marker paragraph
    // written immediately after the call landed in the saved package).
    //
    // WHAT THIS DOES NOT CLAIM: every write on this route already forces a save as part of its
    // OWN success contract -- coedit.writeOperationsToDocument closes the browser and waits
    // `postSaveWaitMs` (70s) before re-downloading, and that disconnect-triggered autosave is
    // what every other operation's package-verification already relies on. Whether an explicit
    // `Api.Save()` call makes a change visible to a DIFFERENT concurrent reader/session FASTER
    // than that existing wait was NOT measured -- this operation exists for a caller who wants
    // to trigger it explicitly (e.g. mid-batch, before continuing in the same session), not as
    // a replacement for the existing wait. xlsx/pptx not measured -- refused here rather than
    // assumed to behave the same as docx.
    emit() {
      return ['Api.Save();']
    },
  },

  tableEdit: {
    cores: ['docx'],
    // The `table` operation above only ever
    // CREATES a new table -- this is the first operation that targets one ALREADY in the
    // document. MEASURED 2026-08-17: GetAllTables()[tableIndex].GetCell(row,col)
    // .GetContent().GetElement(0).AddText(text) applies cleanly and is package-verified, on
    // BOTH the docbuilder-create route (bisected against a negative control -- a script that
    // stops after `oDocument.GetAllTables()` alone still saves fine, so nothing about calling
    // it kills the job) and the live co-editing route (coedit_write_operations, editor.
    // callCommand). Appends to the cell's existing text; does not clear it first -- an
    // overwrite mode was not measured, so it is not offered here rather than guessed.
    //
    // MEASURED NEG. CONTROL 2026-08-17 (live co-editing route, out-of-range tableIndex): the
    // generated script's own `throw` DOES stop the write -- the saved package was byte-identical
    // to before the call and did not contain the requested text. BUT the coedit_write_operations
    // wrapper's own report still said `outcome: "alkalmazva"` for it (bytesElotte===bytesUtana,
    // so a caller checking the BYTE COUNT would catch it, but the per-operation report alone
    // would not) -- the same "the callCommand return proves nothing" class already documented
    // elsewhere in this file for other operations. A caller must check the SAVED PACKAGE, not
    // this operation's own report entry, to know whether an out-of-range index actually wrote.
    emit(op) {
      const tableIndex = Number(op.tableIndex ?? 0)
      if (!Number.isInteger(tableIndex) || tableIndex < 0) throw new Error('tableEdit: `tableIndex` must be a non-negative integer')
      const row = Number(op.row)
      const col = Number(op.col)
      if (!Number.isInteger(row) || row < 0) throw new Error('tableEdit: `row` must be a non-negative integer')
      if (!Number.isInteger(col) || col < 0) throw new Error('tableEdit: `col` must be a non-negative integer')
      if (typeof op.text !== 'string' || !op.text.length) throw new Error('tableEdit: `text` is required and must be non-empty')
      return [
        'var oTables = oDocument.GetAllTables();',
        `if (!oTables || oTables.length <= ${tableIndex}) throw new Error("tableEdit: no table at index ${tableIndex}");`,
        `var oTable = oTables[${tableIndex}];`,
        `var oCell = oTable.GetCell(${row}, ${col});`,
        `if (!oCell) throw new Error("tableEdit: no cell at row ${row}, col ${col}");`,
        'var oContent = oCell.GetContent();',
        'var oPara = oContent.GetElement(0);',
        `oPara.AddText(${jsString(op.text)});`,
      ]
    },
  },

  toc: {
    cores: ['docx'],
    // MEASURED 2026-08-17: oDocument.AddTableOfContents() inserts
    // a real Word TOC FIELD (instrText `TOC \o "1-9" \h`, one PAGEREF per heading it finds) at
    // the document's CURRENT CURSOR position -- same MoveCursorTo*() dependency as footnotes/
    // endnotes above. Package-verified against a doc with one "Heading 1" paragraph: the field
    // referenced that heading's text and a PAGEREF, present in the saved word/document.xml
    // (6 <w:fldChar> elements, the standard begin/separate/end triple twice: once for the TOC
    // field itself, once for its inner PAGEREF). Negative control (no AddTableOfContents call)
    // has none. `\o "1-9"` covers every built-in heading level `text`'s own `heading` field can
    // produce (1-9), so no level range is exposed as a caller option -- there is nothing to pick
    // narrower than what already gets built.
    //
    // `position` picks where the cursor goes first: "start" (default -- the common case, a TOC
    // page before the body) or "end". Both package-verified individually.
    //
    // *** OPERATION ORDER TRAP, MEASURED: put this `toc` operation LAST in the caller's
    // `operations` array, after every `text` with a `heading`. *** AddTableOfContents() only
    // sees headings that exist in the document MODEL at the moment it runs -- a `toc` operation
    // placed FIRST (the naive instinct: "the TOC should read first") builds against an empty
    // document and comes back "No table of contents entries found.", even with `position:
    // "start"` correctly moving the cursor there first. `position` controls WHERE the field
    // lands, not WHEN the scan happens. (oDocument.UpdateAllTOC() can refresh an early-built TOC
    // afterward -- package-verified working -- but is not wired here: it is simpler to state the
    // ordering requirement than to add a second call every `toc` use would carry.)
    emit(op) {
      if (op.position !== undefined && op.position !== 'start' && op.position !== 'end') {
        throw notSupportedError(`toc: unknown position ${JSON.stringify(op.position)} (known: start, end)`)
      }
      return [
        `oDocument.MoveCursorTo${op.position === 'end' ? 'End' : 'Start'}();`,
        'oDocument.AddTableOfContents();',
      ]
    },
  },

  // `oDocument.AddTableOfFigures(oTofPr, bReplace)` -- real
  // signature MEASURED via `.toString()`. Package-verified end to end (2026-08-17): a table with
  // a caption added via the `table` operation's own `caption` field above (`Table.AddCaption`),
  // followed by this operation, produced a real `TOC \h \c "Figure"` field with a genuine
  // PAGEREF entry -- an empty document (no captions) instead produces the builder's own "No table
  // of figures entries found." placeholder text, which this tool does NOT try to detect or
  // refuse: telling the two apart from inside the saved package is the caller's job (same
  // division of responsibility `toc` above already has for its own empty-heading case).
  //
  // *** SAME ORDERING CONTRACT AS `toc`: put this operation AFTER every `table` with a `caption`
  // whose figures it should list, and it only sees captions that exist in the document MODEL at
  // the moment it runs. *** `buildFrom` is the caption LABEL to scan for (matches the `table`
  // operation's own `caption.label`, default "Table" there vs. "Figure" here -- deliberately
  // different defaults, matching each operation's own most common use, not a shared default that
  // would silently mismatch one of them). Only `buildFrom` and `position` are exposed; the other
  // `oTofPr` fields (ShowPageNums/RightAlgn/FormatAsLinks/LabelNumber/LeaderType/TofStyle) are
  // left at the builder's own measured-working defaults rather than offered unmeasured.
  tableOfFigures: {
    cores: ['docx'],
    emit(op) {
      if (op.position !== undefined && op.position !== 'start' && op.position !== 'end') {
        throw notSupportedError(`tableOfFigures: unknown position ${JSON.stringify(op.position)} (known: start, end)`)
      }
      const buildFrom = op.buildFrom ?? 'Figure'
      return [
        `oDocument.MoveCursorTo${op.position === 'end' ? 'End' : 'Start'}();`,
        `oDocument.AddTableOfFigures({BuildFrom: ${jsString(String(buildFrom))}}, true);`,
      ]
    },
  },

  // Signatures MEASURED via the live functions' own
  // `.toString()` (2026-08-17), the same technique that found `AddBookmarkCrossRef`'s real
  // argument order: `Section.GetHeader(sType, isCreate)` / `GetFooter(sType, isCreate)`, where
  // `sType` is one of "default"/"even"/"title" (title = Word's own "first page" header/footer,
  // NOT the header's own text -- confusingly named on the API side, kept as "title" here rather
  // than remapped, since that IS the string this route accepts). The returned object has its own
  // `Push()`, separate from `oDocument.Push()` -- content goes into the header/footer, never the
  // body.
  //
  // Package-verified: `headerReference`/`footerReference` wired into document.xml's sectPr, AND
  // the matching relationship registered in word/_rels/document.xml.rels (the same two-place
  // check `hyperlink` already established for its own rels entry) -- both checked, not just the
  // header/footer part's own text. Negative control: a document with no `headerFooter` operation
  // has zero header*.xml/footer*.xml parts and no such rels entries.
  //
  // `AddPageNumber()` takes NO arguments (measured via `.toString()`: `function(){...}`, an extra
  // argument is silently ignored, not an error) and inserts a real `PAGE` field
  // (`<w:fldSimple w:instr="PAGE \* MERGEFORMAT">`); `AddPagesCount()` inserts `NUMPAGES` the same
  // way. Both package-verified against a real DocBuilder round-trip building "Page {N} of {M}".
  //
  // *** COMBINED WITH `comments` ON THE MAIN BODY (the mandatory order-trap gate this session's
  // established): MEASURED, no conflict found. *** Two early runs came back `outcome: "blocked"`
  // (the whole-job-killed class documented elsewhere in this file), but FIVE subsequent runs of
  // the identical script all succeeded with both the header text and the comment anchor present
  // and correct -- this reads as Document Server flakiness on those two runs, not a deterministic
  // incompatibility, and is NOT treated as a refusal-worthy finding (unlike the genuine, 100%-
  // reproducible comments/footnotes ordering trap from koteg04).
  headerFooter: {
    cores: ['docx'],
    emit(op) {
      if (op.target !== 'header' && op.target !== 'footer') {
        throw new Error(`headerFooter: \`target\` must be "header" or "footer", got ${JSON.stringify(op.target)}`)
      }
      const variant = op.variant ?? 'default'
      if (!['default', 'even', 'title'].includes(variant)) {
        throw notSupportedError(`headerFooter: unknown variant ${JSON.stringify(variant)} (known: default, even, title) -- this is the Section.GetHeader/GetFooter \`sType\` argument itself; an unrecognised value returns null with no throw, so this tool refuses it here instead`)
      }
      const parts = Array.isArray(op.parts) ? op.parts : (op.text !== undefined && op.text !== null ? [String(op.text)] : null)
      if (!parts || !parts.length) {
        throw new Error('headerFooter: `parts` (non-empty array) or `text` is required')
      }
      const getter = op.target === 'header' ? 'GetHeader' : 'GetFooter'
      const varName = op.target === 'header' ? 'oHeader' : 'oFooter'
      const lines = [
        'var oSection = oDocument.GetFinalSection();',
        `var ${varName} = oSection.${getter}(${jsString(variant)}, true);`,
        'var oHFParagraph = Api.CreateParagraph();',
      ]
      parts.forEach((part, idx) => {
        if (typeof part === 'string') {
          lines.push(`oHFParagraph.AddText(${jsString(part)});`)
        } else if (part && part.pageNumber === true) {
          lines.push('oHFParagraph.AddPageNumber();')
        } else if (part && part.pagesCount === true) {
          lines.push('oHFParagraph.AddPagesCount();')
        } else {
          throw new Error(`headerFooter: parts[${idx}] must be a string, {pageNumber: true}, or {pagesCount: true} -- got ${JSON.stringify(part)}`)
        }
      })
      lines.push(`${varName}.Push(oHFParagraph);`)
      return lines
    },
  },

  formField: {
    cores: ['docx'],
    // All five Create*Form calls MEASURED against a real
    // DocBuilder round-trip (2026-08-17), unzipped and read back from word/document.xml. Every
    // form is a real <w:sdt> content control, built with Api.Create<Kind>Form({...}), inserted
    // via a fresh paragraph's AddElement() + oDocument.Push() -- the same "build detached, then
    // attach" shape as `shape`/`table` elsewhere in this file. Negative control (no formField
    // operation) has zero <w:sdt> anywhere.
    //   text     -> <w:formPr key/helpText/required> + <w:textFormPr>, placeholder text visible
    //               as the SDT's own content run until filled in
    //   checkbox -> <w14:checkbox><w14:checked val="0|1">, rendered as an actual ☐/☑ glyph run
    //   combobox -> <w:dropDownList> with one <w:listItem> per `items` entry (plus the DocBuilder's
    //               own default "Choose an item" first entry -- not caller-suppressible, MEASURED)
    //   date     -> <w:date><w:dateFormat val="..."/> -- `format` passed straight through, no
    //               validation against a known-good token set (not measured which tokens the
    //               Document Server actually renders; this tool does not guess one)
    //   picture  -> a picture content control; MEASURED to produce TWO <w:sdt> per call (a
    //               placeholder run plus the picture frame itself) -- both package-verified
    //               present, this tool does not try to collapse them to one
    // `key` is the field's programmatic name (GetFormsByKey/SetFormsData use it later); `tip` is
    // the on-hover help text; `required` defaults to false in every kind, matching the Document
    // Server's own default.
    emit(op) {
      if (!op.key) throw new Error('formField: `key` is required (the field\'s programmatic name)')
      const required = op.required ? 'true' : 'false'
      const tip = jsString(op.tip ?? '')
      const key = jsString(op.key)
      let createCall
      if (op.kind === 'text') {
        createCall = `Api.CreateTextForm({key: ${key}, tip: ${tip}, required: ${required}, placeholder: ${jsString(op.placeholder ?? '')}, comb: false})`
      } else if (op.kind === 'checkbox') {
        createCall = `Api.CreateCheckBoxForm({key: ${key}, tip: ${tip}, required: ${required}, checked: ${op.checked ? 'true' : 'false'}})`
      } else if (op.kind === 'combobox') {
        if (!Array.isArray(op.items) || !op.items.length) throw new Error('formField: kind "combobox" needs a non-empty `items` array')
        const items = `[${op.items.map((i) => jsString(String(i))).join(', ')}]`
        createCall = `Api.CreateComboBoxForm({key: ${key}, tip: ${tip}, required: ${required}, format: "", items: ${items}, editable: false})`
      } else if (op.kind === 'date') {
        createCall = `Api.CreateDateForm({key: ${key}, tip: ${tip}, required: ${required}, format: ${jsString(op.format ?? 'MM/DD/YYYY')}})`
      } else if (op.kind === 'picture') {
        createCall = `Api.CreatePictureForm({key: ${key}, tip: ${tip}, required: ${required}, scaleFlag: "always", lockAspectRatio: true})`
      } else {
        throw notSupportedError(`formField: unknown kind ${JSON.stringify(op.kind)} (known: text, checkbox, combobox, date, picture)`)
      }
      return [
        `var oForm = ${createCall};`,
        'var oFormParagraph = Api.CreateParagraph();',
        'oFormParagraph.AddElement(oForm);',
        'oDocument.Push(oFormParagraph);',
      ]
    },
  },

  // Document-level operations (watermark, track-changes,
  // search/replace), all MEASURED against a real DocBuilder round-trip (2026-08-17). Unlike every
  // OPERATIONS entry above these three do not build a new element to Push() -- they call straight
  // through to `oDocument`, the same document-level object `pageSetup`/`toc` already use.
  watermark: {
    cores: ['docx'],
    // `oDocument.InsertWatermark({...})` writes a REAL Word watermark: a WordArt textbox
    // ("PowerPlusWaterMarkObject..." naming, matching Word's own convention) anchored behind the
    // text in the section's header (word/header1.xml), wired via a fresh headerReference in
    // document.xml's sectPr -- package-verified: `text` shows up literally inside the header
    // part's txbx content, and a document with NO watermark call has no header part at all
    // (negative control: zero header*.xml files in the saved package). `RemoveWatermark()` is
    // exposed via `remove: true` -- not independently package-verified in this batch (no prior
    // watermark existed in the fixture to remove), refused-nothing rather than guessed silent.
    emit(op) {
      if (op.remove) return ['oDocument.RemoveWatermark();']
      if (!op.text) throw new Error('watermark: `text` is required unless `remove: true`')
      const diagonal = op.diagonal === false ? 'false' : 'true'
      const transparent = op.transparent === false ? 'false' : 'true'
      const color = jsString(op.color ?? '000000')
      return [
        `oDocument.InsertWatermark({type: "text", text: ${jsString(String(op.text))}, isDiagonal: ${diagonal}, isAuto: true, color: ${color}, transparent: ${transparent}});`,
      ]
    },
  },

  trackChanges: {
    cores: ['docx'],
    // `oDocument.SetTrackRevisions(bool)`. Package-verified: text added to the document AFTER
    // `enabled: true` gets real Word revision marks -- `<w:ins w:id="..." w:date="...">` on the
    // inserted run/paragraph-mark, the same markup Word itself writes when a human types with
    // Track Changes on. Negative control (no trackChanges operation) has zero `<w:ins`/`<w:del`
    // anywhere. This does not touch existing content -- there is nothing to mark as inserted
    // before this operation runs, since every OPERATIONS entry in this file only ADDS content;
    // an `enabled: true` therefore governs everything emitted AFTER it in the caller's
    // `operations` array, same ordering contract as `toc`/`pageSetup`.
    emit(op) {
      if (op.enabled === undefined || op.enabled === null) throw new Error('trackChanges: `enabled` is required')
      return [`oDocument.SetTrackRevisions(${op.enabled ? 'true' : 'false'});`]
    },
  },

  searchReplace: {
    cores: ['docx'],
    // `oDocument.SearchAndReplace({searchString, replaceString, matchCase})` -- MEASURED on the
    // create route (builder.OpenFile -> script -> SaveFile -> CloseFile), which is this file's
    // ONLY route. an earlier finding that a `SearchAndReplace_Run` call threw on a different
    // route is NOT this call and NOT this route -- a different name, most likely the co-editing
    // surface (editor.callCommand), out of scope here; do not read that finding as this call
    // being refused. Package-verified: replaces EVERY matching occurrence in the document (not
    // just the first), `matchCase: true` skips differently-cased occurrences, `matchCase: false`
    // replaces regardless of case using the LITERAL replacement text (no case-adaptation to the
    // matched occurrence, unlike some word processors' "smart" replace).
    emit(op) {
      if (!op.search) throw new Error('searchReplace: `search` is required')
      if (op.replace === undefined || op.replace === null) throw new Error('searchReplace: `replace` is required')
      const matchCase = op.matchCase === true
      return [
        `oDocument.SearchAndReplace({searchString: ${jsString(String(op.search))}, replaceString: ${jsString(String(op.replace))}, matchCase: ${matchCase}});`,
      ]
    },
  },

  // `Api.ReplaceTextSmart(textStrings, tab, newLine)` -- NOT on
  // `oDocument` (a first probe there found `undefined`); it lives on `Api` itself (confirmed by
  // the API's own reflection inventory, euro-api-leltar-docx.md's `===Api===` list) and replaces
  // whatever paragraph(s) are CURRENTLY SELECTED, diffing the new text against the existing runs
  // (`AscCommon.getTextDelta`) rather than a blunt full replace -- this is what "smart" means
  // here: formatting on unchanged characters survives, unlike `searchReplace`/`replaceText` above.
  //
  // *** SELECTION GRANULARITY MATTERS, MEASURED: selecting only a MATCHED WORD (`oDocument.
  // Search(text)[0].Select()`, the same pattern `addComment` above uses) and then calling this
  // with a full replacement sentence does NOT do a clean word-for-word swap -- it diffs the
  // SELECTED paragraph's WHOLE text against the replacement and inserts the new string alongside
  // the old one (measured: "first paragraph with " + "<replacement>" + " inside it" surrounding
  // an unrelated leftover fragment). Selecting the WHOLE PARAGRAPH (`oParagraph.Select()`) and
  // replacing it wholesale is clean and package-verified: an "original sentence here" paragraph
  // became "replaced sentence now" with nothing left over. This operation only offers the
  // paragraph-level shape, since the sub-range shape produced a surprising result that would need
  // its own separate investigation to use safely. ***
  //
  // *** ORDERING CONTRACT, same class as `toc`/`mathEquation` above: this targets `oParagraph`,
  // the variable the MOST RECENT `text` operation declared with `var` (so a later `text` op
  // reassigns it) -- put a `smartReplace` operation immediately after the `text` operation whose
  // paragraph it should replace, not after some other paragraph-producing operation. ***
  smartReplace: {
    cores: ['docx'],
    emit(op) {
      if (op.text === undefined || op.text === null) throw new Error('smartReplace: `text` is required')
      return [
        'oParagraph.Select();',
        `Api.ReplaceTextSmart([${jsString(String(op.text))}]);`,
      ]
    },
  },

  // `oDocument.AcceptAllRevisionChanges()` / `RejectAllRevisionChanges()`
  // -- both zero-argument, oDocument-level, MEASURED via `.toString()` (both a plain call-through
  // to the same-named internal Document method, always `return true`) and package-verified with a
  // real tracked insertion (`trackChanges`/`trackedChanges` above enable tracking; content added
  // afterward gets real `<w:ins>` markup, same mechanism those operations already document):
  // Accept resolves the markers and KEEPS the content (zero `<w:ins>`/`<w:del>` afterward, text
  // still present); Reject resolves the markers and REVERTS the content (zero markers, text
  // GONE). A document with no tracked changes at all is untested here but is not expected to
  // throw -- both calls are unconditional sweeps over whatever revisions exist, none is a
  // reasonable case for either call to fail on.
  resolveRevisions: {
    cores: ['docx'],
    emit(op) {
      if (op.action !== 'accept' && op.action !== 'reject') {
        throw new Error(`resolveRevisions: \`action\` must be "accept" or "reject", got ${JSON.stringify(op.action)}`)
      }
      return [`oDocument.${op.action === 'accept' ? 'AcceptAllRevisionChanges' : 'RejectAllRevisionChanges'}();`]
    },
  },

  // `oDocument.DeleteBookmark(sName)` -- MEASURED via `.toString()`:
  // a thin call-through to `this.Document.RemoveBookmark(sName)`, returning `false` (not
  // throwing) when `sName` is `undefined`. Package-verified: a bookmark added via the `bookmark`
  // field on `text` above (koteg09) and then named here disappears from
  // `oDocument.GetAllBookmarksNames()` (`["myMarker"]` before, `[]` after) and its
  // `<w:bookmarkStart>`/`<w:bookmarkEnd>` pair leaves the saved package. Naming a bookmark that
  // does not exist is UNMEASURED (not run against this instance) -- refused nowhere client-side,
  // but nothing here claims it is safe either, same disclosure style as `Table.Split`'s own
  // unmeasured multi-entry case above.
  deleteBookmark: {
    cores: ['docx'],
    emit(op) {
      if (!op.name) throw new Error('deleteBookmark: `name` is required')
      return [`oDocument.DeleteBookmark(${jsString(String(op.name))});`]
    },
  },

  // The THIRD targeting mode this file's docx operations use,
  // distinct from the other two (`text`/`table`/`shape` always Push() to the document END;
  // `searchReplace`/`replaceText`/`addComment` target by SEARCH PATTERN): these five target
  // "wherever the document's own cursor is". MEASURED (2026-08-17), and the finding that shapes
  // every field below: on this route (a fresh builder.OpenFile with no interactive session) the
  // cursor's DEFAULT position is the very START of the document -- and "the document" includes
  // whatever content the CALLER's own input document already had there (in the common case with
  // no caller-supplied documentBase64, that is this route's own minimal-doc placeholder
  // paragraph). A caller who does not explicitly position the cursor first will touch THAT
  // content, not their own newly-Push()'d paragraphs -- package-verified directly:
  // `oDocument.GetCurrentWord()` at the untouched default position returned "EURO" (the first
  // word of the minimal doc's own placeholder text), not anything from a same-script `text`
  // operation run beforehand. `position` is therefore REQUIRED on every operation below, not
  // optional-with-a-guessed-default -- the same reasoning `toc`/`mathEquation` already apply to
  // their own MoveCursorToEnd() dependency, just non-optional here because the wrong default is
  // not merely unhelpful, it is a different, caller-invisible target entirely.
  enterText: {
    cores: ['docx'],
    // `oDocument.EnterText(sText)` -- MEASURED: with `position: "end"` (MoveCursorToEnd() first),
    // the text lands as its own run immediately after the last Push()'d content, package-
    // verified. With `position: "start"`, it lands before EVERYTHING, including the input
    // document's own pre-existing first paragraph (same caveat as the class comment above).
    emit(op) {
      if (op.text === undefined || op.text === null) throw new Error('enterText: `text` is required')
      if (op.position !== 'start' && op.position !== 'end') {
        throw new Error(`enterText: \`position\` must be "start" or "end", got ${JSON.stringify(op.position)}`)
      }
      return [
        `oDocument.MoveCursorTo${op.position === 'end' ? 'End' : 'Start'}();`,
        `oDocument.EnterText(${jsString(String(op.text))});`,
      ]
    },
  },

  // `oDocument.InsertContent(arrContent, isInline, oPr)` -- MEASURED via `.toString()`: `arrContent`
  // is an array of ALREADY-BUILT ApiParagraph/ApiTable/ApiBlockLvlSdt objects (NOT a path or a
  // second document's content, despite the name) inserted at the cursor via the document's own
  // `GetCurrentParagraph()`. This binding builds one plain paragraph per string in `paragraphs`,
  // matching the `text` operation's own AddText field name/shape rather than inventing a new
  // element-description schema -- richer per-paragraph formatting is exactly what the existing
  // `text` operation already offers via Push(), so this stays minimal on purpose. Package-
  // verified with `position: "end"`: the new paragraph(s) land immediately after the last
  // Push()'d content, in array order.
  insertContent: {
    cores: ['docx'],
    emit(op) {
      if (!Array.isArray(op.paragraphs) || !op.paragraphs.length) {
        throw new Error('insertContent: `paragraphs` must be a non-empty array of strings')
      }
      if (op.position !== 'start' && op.position !== 'end') {
        throw new Error(`insertContent: \`position\` must be "start" or "end", got ${JSON.stringify(op.position)}`)
      }
      const lines = [`oDocument.MoveCursorTo${op.position === 'end' ? 'End' : 'Start'}();`]
      const varNames = op.paragraphs.map((text, i) => {
        const v = `oInsertContentPara${i}`
        lines.push(`var ${v} = Api.CreateParagraph();`, `${v}.AddText(${jsString(String(text))});`)
        return v
      })
      lines.push(`oDocument.InsertContent([${varNames.join(', ')}], false);`)
      return lines
    },
  },

  // `oDocument.ReplaceCurrentWord(sReplace, sPart)` / `ReplaceCurrentSentence` -- MEASURED via
  // `.toString()`, both delegate to `this.Document.Replace{Word,Sentence}(dir, replace)` where
  // `dir` comes from `sPart` ("before"/"after" -> -1/1, omitted -> 0). *** THIS OPERATION IS
  // BLIND BY CONSTRUCTION, NAMED NOT SOLVED: neither this route nor this file has a way to read
  // back "what word/sentence is at the cursor" before replacing it -- the caller must already
  // know the target document's content at `position`, e.g. a known placeholder word at the start
  // of a caller-supplied template (documentBase64). On the route's own synthetic minimal
  // document there is nothing meaningful to target; do not use this operation without a
  // caller-supplied document whose content at `position` is known. ***
  replaceCurrent: {
    cores: ['docx'],
    emit(op) {
      if (op.scope !== 'word' && op.scope !== 'sentence') {
        throw new Error(`replaceCurrent: \`scope\` must be "word" or "sentence", got ${JSON.stringify(op.scope)}`)
      }
      if (op.replace === undefined || op.replace === null) throw new Error('replaceCurrent: `replace` is required')
      if (op.position !== 'start' && op.position !== 'end') {
        throw new Error(`replaceCurrent: \`position\` must be "start" or "end", got ${JSON.stringify(op.position)}`)
      }
      if (op.part !== undefined && op.part !== null && op.part !== 'before' && op.part !== 'after') {
        throw new Error(`replaceCurrent: \`part\` must be "before" or "after" if given, got ${JSON.stringify(op.part)}`)
      }
      const method = op.scope === 'word' ? 'ReplaceCurrentWord' : 'ReplaceCurrentSentence'
      const args = [jsString(String(op.replace))]
      if (op.part) args.push(jsString(op.part))
      return [
        `oDocument.MoveCursorTo${op.position === 'end' ? 'End' : 'Start'}();`,
        `oDocument.${method}(${args.join(', ')});`,
      ]
    },
  },

  // `oDocument.InsertBlankPage()` -- zero-arg, MEASURED to insert at the cursor (page count
  // confirmed incrementing after `MoveCursorToEnd()`). Same `position`-required discipline as
  // the rest of this class.
  insertBlankPage: {
    cores: ['docx'],
    emit(op) {
      if (op.position !== 'start' && op.position !== 'end') {
        throw new Error(`insertBlankPage: \`position\` must be "start" or "end", got ${JSON.stringify(op.position)}`)
      }
      return [
        `oDocument.MoveCursorTo${op.position === 'end' ? 'End' : 'Start'}();`,
        'oDocument.InsertBlankPage();',
      ]
    },
  },

  // `oDocument.ClearAllFields()` -- MEASURED document-WIDE (`GetFormsManager().GetAllForms()` +
  // `ClearAllSpecialForms`), NOT cursor-related despite living in this file's "cursor" class --
  // grouped here because it arrived from the same source sweep, not because it shares the
  // targeting mechanism. Package-verified: a text form field with a caller-set value reverts to
  // OnlyOffice's own default placeholder text ("Your text here"), not to empty and not to the
  // caller's original placeholder string -- name this to a caller rather than let them assume
  // either of those two more intuitive outcomes.
  clearAllFields: {
    cores: ['docx'],
    emit() {
      return ['oDocument.ClearAllFields();']
    },
  },

  // `oDocument.GoToPage(index)` -- REFUSED, MEASURED rather than assumed: on this route it is a
  // confirmed NO-OP. Three independent checks, all against a real two-page document: (1) a
  // same-script `EnterText()` right after `GoToPage(0)`/`GoToPage(1)` landed at the document
  // START regardless of which page index was passed -- identical to never having called it; (2)
  // `oDocument.GetCurrentPage()` read back 0 both BEFORE and AFTER `GoToPage(0)`, so the call does
  // not even move whatever it is that GetCurrentPage reports; (3) the call itself returns `true`
  // in every case (a valid index), so nothing here throws or signals failure -- this is the
  // established silent-no-op class this file refuses elsewhere (`align`, `highlight`,
  // `SetTextDirection`), not a route-wide "blocked" failure (the surrounding script still ran and
  // saved normally). Likely a page/view-navigation concept the create-route's headless
  // DocBuilder session has nothing to navigate (no visible viewport), distinct from the document
  // MODEL cursor `MoveCursorToStart/End/ToPos` and `EnterText` etc. actually read.
  goToPage: {
    cores: ['docx'],
    emit() {
      throw notSupportedError('goToPage: `oDocument.GoToPage()` is refused -- MEASURED (2026-08-17) to be a no-op on this route: it does not move the cursor `EnterText`/`ReplaceCurrentWord`/etc. read (verified against a real two-page document, same result regardless of page index), and `oDocument.GetCurrentPage()` itself reports the same value before and after the call. It returns `true` and the job completes normally either way, so this is the silent-no-op class this file already refuses elsewhere, not a blocked-job class.')
    },
  },

  mathEquation: {
    cores: ['docx'],
    // `oDocument.AddMathEquation(sText, sFormat)` -- `sFormat` one
    // of "latex"/"unicode"/"mathml" (read off the live function's own .toString(); "mathml" is
    // the third branch the .toString() dump truncated before reaching, named here on that basis
    // but NOT independently package-verified -- "latex" and "unicode" both were). Same
    // MoveCursorToEnd() dependency as footnotes/endnotes/toc/comments above: without it, the
    // equation lands at the document's default cursor position (its START), not after whatever
    // was last Push()'d -- measured directly (a plain string call with no MoveCursorToEnd put the
    // equation's flattened text BEFORE the earlier paragraph's own text in document order).
    // Package-verified: a real OOXML math structure (<m:oMath>, with <m:sSup>/<m:e>/<m:sup> for
    // the "x^2+y^2=z^2" fixture's superscripts) in word/document.xml -- not just that SOME XML
    // changed. Negative control: no `mathEquation` operation, zero <m:oMath> anywhere.
    emit(op) {
      if (!op.text) throw new Error('mathEquation: `text` is required')
      const format = op.format ?? 'unicode'
      if (!['latex', 'unicode', 'mathml'].includes(format)) {
        throw notSupportedError(`mathEquation: unknown format ${JSON.stringify(format)} (known: latex, unicode, mathml)`)
      }
      return [
        'oDocument.MoveCursorToEnd();',
        `oDocument.AddMathEquation(${jsString(String(op.text))}, ${jsString(format)});`,
      ]
    },
  },

  // `export`/`documentStats`. Both are READ operations
  // (ToMarkdown/ToHtml/GetStatistics), and this route (builder.OpenFile -> script -> SaveFile ->
  // CloseFile) has EXACTLY ONE channel back to the caller: the saved document's own content --
  // there is no console.log, no return value, no out-of-band channel between the DocBuilder
  // script and the Node-side caller. The preferred shape is a docProps/custom-property field,
  // NOT the visible document body -- but MEASURED (2026-08-17, `for...in` over both `oDocument`
  // and `Api` for /custom/i-matching methods): only `GetCustomProperties`/`GetCustomXmlParts`
  // exist, there is NO write API (`SetCustomProperty` or equivalent) anywhere on this instance.
  // *** THIS IS THEREFORE A FALLBACK SHAPE, NOT THE PREFERRED ONE -- TRANSITIONAL, UNTIL A REAL
  // WRITE API EXISTS: *** a mandatory, greppable marker prefix (below); the real custom-property
  // channel is separate follow-up work, at the euro-mcp tool layer this file's OPERATIONS table
  // does not own.
  export: {
    cores: ['docx'],
    // `oDocument.ToMarkdown()`/`ToHtml()`, both MEASURED with no throw and real content returned
    // (a two-paragraph fixture round-tripped through both formats correctly). The result is
    // appended as a NEW paragraph, `__EXPORT_MARKDOWN__:<result>` or `__EXPORT_HTML__:<result>` --
    // a caller extracts it by reading the saved package's own text and matching that prefix.
    emit(op) {
      const format = op.format ?? 'markdown'
      if (format !== 'markdown' && format !== 'html') {
        throw notSupportedError(`export: unknown format ${JSON.stringify(format)} (known: markdown, html)`)
      }
      const call = format === 'markdown' ? 'ToMarkdown' : 'ToHtml'
      const marker = format === 'markdown' ? '__EXPORT_MARKDOWN__:' : '__EXPORT_HTML__:'
      return [
        `var oExportResult = oDocument.${call}();`,
        'var oExportParagraph = Api.CreateParagraph();',
        `oExportParagraph.AddText(${jsString(marker)} + oExportResult);`,
        'oDocument.Push(oExportParagraph);',
      ]
    },
  },

  documentStats: {
    cores: ['docx'],
    // `oDocument.GetStatistics()` -- MEASURED working, returns {PageCount, WordsCount,
    // ParagraphCount, SymbolsCount, SymbolsWSCount}. `oDocument.GetCustomProperties()` -- MEASURED
    // working (returns an object even when the document has none set). *** `GetDocumentInfo()` IS
    // REFUSED, NOT OFFERED: MEASURED (2x, reproducible) to THROW `Cannot read property
    // 'asc_getApplication' of null` on this route -- an editor-Application-level dependency this
    // headless DocBuilder-create context does not provide, the same class of route-mismatch as
    // `LoadMailMergeData`'s `editor.asc_StartMailMergeByList` elsewhere in this file's own
    // measurement notes. An uncaught throw here kills the WHOLE job (measured: "blocked" outcome
    // when this call was not wrapped), so it is refused client-side rather than left to crash a
    // caller's script. *** Same marker convention as `export`: `__STATS_JSON__:<JSON>` and,
    // if requested, `__CUSTOM_PROPERTIES_JSON__:<JSON>` as separate paragraphs.
    emit(op) {
      const lines = [
        'var oStatsResult = oDocument.GetStatistics();',
        'var oStatsParagraph = Api.CreateParagraph();',
        'oStatsParagraph.AddText("__STATS_JSON__:" + JSON.stringify(oStatsResult));',
        'oDocument.Push(oStatsParagraph);',
      ]
      if (op.includeCustomProperties) {
        lines.push(
          'var oCustomPropsResult = oDocument.GetCustomProperties();',
          'var oCustomPropsParagraph = Api.CreateParagraph();',
          'oCustomPropsParagraph.AddText("__CUSTOM_PROPERTIES_JSON__:" + JSON.stringify(oCustomPropsResult));',
          'oDocument.Push(oCustomPropsParagraph);',
        )
      }
      return lines
    },
  },

  // `oDocument.SearchAndReplace` is
  // NOT a general capability of this route -- it exists on the DOCUMENT object, and MEASURED
  // absent on the worksheet/presentation equivalents (diag round 5, live
  // Document Server: `oWorksheet.SearchAndReplace` and `oWorksheet.Search` both `typeof ...
  // undefined` on xlsx). docx-only, not a placeholder restriction -- the xlsx/pptx absence is a
  // measured fact, not an unmeasured gap this operation happens to skip.
  //
  // Already proven in production BEFORE this operations-table entry existed: `buildEditScript`'s
  // `replace_text` op type has used this exact call since before the OPERATIONS-table refactor
  // (this entry gives it the SAME translator/report/route parity every other operation already
  // has -- buildCreateScript AND buildCoeditScript, not just the older, edit_document-only path).
  // addComment: docx-only -- Search()/AddComment()
  // are DOCUMENT-level methods, and the same absence measured for replaceText's
  // SearchAndReplace (diag round 5) applies to Search too: not tested on
  // xlsx/pptx, no equivalent assumed. Live-proven end to end (diag round 6, real Document
  // Server, save+reopen): `oDocument.Search(anchorText)[0].AddComment(text, author, initials)`
  // produces a real word/comments.xml <w:comment>, anchored to the FIRST match of `anchorText`.
  // Anchors on the FIRST match only, by design -- an operation that commented on every match of
  // a short anchor string would surprise a caller who meant "this one specific sentence".
  //
  // *** DELIBERATELY DOES NOT THROW WHEN THE ANCHOR IS MISSING -- this is a SCRIPT-LEVEL
  // decision, not an oversight, and it matters more here than in any other operation in this
  // table: every OTHER operation's validity is knowable from its own PARAMETERS, checked in
  // Node.js before any script line is even generated (unknown type, wrong core, missing
  // required field). Whether `anchorText` exists in THIS document is knowable ONLY once the
  // script is actually running inside the live editor -- there is no client-side way to check
  // it first. Because buildCoeditScript emits every operation's lines into ONE sequential
  // script body, a THROW here would abort the DocBuilder job mid-script, on a live document,
  // potentially AFTER earlier operations in the same batch already mutated it -- exactly the
  // partial-application failure mode this whole route was redesigned to prevent. A missing
  // anchor is therefore a SILENT no-op at the script level, on purpose -- and the caller-facing
  // verification (does the comment now exist in the saved file?) is the tool layer's job
  // (office_add_comment), not this operation's. ***
  addComment: {
    cores: ['docx'],
    emit(op) {
      if (!op.anchorText) throw new Error('addComment: `anchorText` is required')
      if (!op.text) throw new Error('addComment: `text` is required')
      const author = op.author ?? 'euro-mcp'
      // Initials default to the first two letters of author, uppercased -- OOXML wants SOME
      // value here (an empty w:initials is legal but renders as a blank avatar tag in most
      // editors), not because any specific value was measured to matter functionally.
      const initials = op.initials ?? String(author).slice(0, 2).toUpperCase()
      return [
        `var oCommentResults = oDocument.Search(${jsString(String(op.anchorText))});`,
        `if (oCommentResults && oCommentResults.length > 0) {`,
        `  oCommentResults[0].AddComment(${jsString(String(op.text))}, ${jsString(String(author))}, ${jsString(initials)});`,
        `}`,
      ]
    },
  },

  replaceText: {
    cores: ['docx'],
    emit(op) {
      // `!search` refuses BOTH a missing `search` and an empty string in one check: an empty
      // searchString is undefined/unmeasured behaviour on this API, not a documented no-op --
      // better a named refusal than a guess at what the server would do with it.
      if (!op.search) throw new Error('replaceText: `search` is required and must not be empty')
      const replace = op.replace ?? ''
      return [`oDocument.SearchAndReplace({ searchString: ${jsString(String(op.search))}, replaceString: ${jsString(String(replace))} });`]
    },
  },

  // `oDocument.SetTrackRevisions`
  // is a document-wide toggle, MEASURED confirmed working (handoff "What Worked": run against the
  // live Document Server via the DocBuilder create route, the saved word/document.xml carried
  // real <w:ins>/<w:del> elements afterwards -- not taken from documentation alone). Deliberately
  // ONLY this one call: no object lookup, no insertion-position guessing, nothing this operation
  // needs beyond what was already run and confirmed. `enabled` defaults to true (asking for
  // "trackedChanges" and getting nothing is the more surprising outcome of the two defaults).
  trackedChanges: {
    cores: ['docx'],
    emit(op) {
      const enabled = op.enabled !== false
      return [`oDocument.SetTrackRevisions(${enabled ? 'true' : 'false'});`]
    },
  },

  // MEASURED via toString() -- `Api.CreateOleObject(imageSrc,
  // width, height, data, appId)`, a POSITIONAL signature. The function's own internal guard
  // (recovered via toString()) silently returns `null` unless imageSrc/data are non-empty
  // strings, appId is a non-empty string, and width/height are real numbers -- replicated here
  // client-side so a caller gets a NAMED error instead of a silently-dropped drawing. Unlike the
  // xlsx `sheetDrawing.ole` entry (ApiWorksheet.AddOleObject, CONFIRMED BROKEN -- not a real OLE
  // object in the saved package), this is a DIFFERENT call on a DIFFERENT object: measured
  // package-verified working here (`word/embeddings/oleObjectN.bin` + `<o:OLEObject>` present).
  // The two findings do not transfer between cores -- different object, different engine path.
  oleObject: {
    cores: ['docx'],
    emit(op) {
      if (!op.imageSrc) throw new Error('oleObject: `imageSrc` (base64 preview image data) is required')
      if (!op.data) throw new Error('oleObject: `data` (the embedded OLE payload string) is required')
      if (!op.appId) throw new Error('oleObject: `appId` (e.g. "Word.Document") is required')
      const w = Number(op.width ?? 1000000)
      const h = Number(op.height ?? 1000000)
      if (!Number.isFinite(w) || !Number.isFinite(h)) throw new Error('oleObject: `width`/`height` must be real numbers (EMU)')
      return [
        `var oOle = Api.CreateOleObject(${jsString(op.imageSrc)}, ${w}, ${h}, ${jsString(op.data)}, ${jsString(op.appId)});`,
        'var oOlePar = Api.CreateParagraph();',
        'oOlePar.AddDrawing(oOle);',
        'oDocument.Push(oOlePar);',
      ]
    },
  },

  // MEASURED -- `Api.CreateGroup([drawings])` ALWAYS throws
  // "All drawings must be in document" here, regardless of whether the drawings were already
  // pushed (tried: separate paragraphs, same paragraph -- both fail identically). The call that
  // ACTUALLY works is `oDocument.GroupDrawings([drawings])` (recovered via toString(): it checks
  // `IsUseInDocument()` on each drawing first, i.e. it expects them already in the document) --
  // package-verified present (`<wpg:wgp>` grouped-shape marker, drawing count drops from N to 1).
  // Each shape is therefore built as its OWN paragraph and pushed individually before grouping,
  // same discipline as `wordArt`/`oleObject` above (AddDrawing, never a raw Push of the drawing).
  drawingGroup: {
    cores: ['docx'],
    emit(op) {
      if (!Array.isArray(op.shapes) || op.shapes.length < 2) {
        throw new Error('drawingGroup: `shapes` must be an array of at least 2 shape descriptors (grouping fewer than 2 shapes is not a group)')
      }
      const lines = []
      const varNames = []
      op.shapes.forEach((shapeSpec, idx) => {
        if (!shapeSpec || typeof shapeSpec !== 'object') throw new Error(`drawingGroup: shapes[${idx}] must be an object`)
        const shapeType = String(shapeSpec.shapeType ?? 'rect')
        const w = Number(shapeSpec.width ?? 1000000)
        const h = Number(shapeSpec.height ?? 1000000)
        let fill = 'Api.CreateNoFill()'
        if (shapeSpec.fill) {
          validateRgbColor(`drawingGroup: shapes[${idx}].fill`, shapeSpec.fill)
          fill = `Api.CreateSolidFill(Api.CreateRGBColor(${shapeSpec.fill.map(Number).join(', ')}))`
        }
        const v = `oDrawGroupShape${idx}`
        varNames.push(v)
        lines.push(
          `var ${v} = Api.CreateShape(${jsString(shapeType)}, ${w}, ${h}, ${fill}, Api.CreateStroke(0, Api.CreateNoFill()));`,
          `var ${v}Par = Api.CreateParagraph();`,
          `${v}Par.AddDrawing(${v});`,
          `oDocument.Push(${v}Par);`,
        )
      })
      lines.push(`oDocument.GroupDrawings([${varNames.join(', ')}]);`)
      return lines
    },
  },
}
