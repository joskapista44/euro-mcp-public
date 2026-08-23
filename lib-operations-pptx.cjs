'use strict'
// This file owns the PPTX-only OPERATIONS entries (extracted verbatim -- same emit()/validate()
// bodies, same comments, same order). The multi-core entries that also touch pptx (per each
// one's own `cores` array) are pulled in from their own individual modules under operations/,
// not duplicated here. lib.cjs remains the aggregator: it populates
// lib-operations-registry.cjs with every helper function/const these files may need BEFORE
// requiring any of them, so this file (and its siblings) can destructure from a plain,
// already-populated object -- no circular require back into lib.cjs, and no function wrapper
// around the entries themselves (a wrapper would roll every nested emit()'s complexity up into
// one artificially huge factory-function score, which is exactly what the first draft of this
// split did before qlty smells caught it).
const { buildSlideGroupLines, buildSlideRemoveObjectLine, buildSlideTransitionLines, jsString, resolveSlideBackgroundLine } = require('./lib-operations-registry.cjs')
const text = require('./operations/text.cjs')
const table = require('./operations/table.cjs')
const image = require('./operations/image.cjs')
const shape = require('./operations/shape.cjs')
const wordArt = require('./operations/wordArt.cjs')
const chart = require('./operations/chart.cjs')
const runs = require('./operations/runs.cjs')
module.exports = {
  text,
  table,
  image,
  shape,
  wordArt,
  chart,
  runs,

  // SLIDE LAYOUT: registered (not "unknown type") so a caller asking for it gets the
  // measured reason, not a guess. Measured: `oSlide.ApplyLayout()` silently kills the whole job
  // (outcome "blocked", zero output, zero error message) -- a structural gap on this DocBuilder
  // instance, not a bug in this tool. Whether the SAME call also fails through the co-editing
  // route (editor.callCommand) is UNMEASURED -- this entry is named ONLY for the
  // DocBuilder-create route this file builds for.
  layout: {
    cores: ['pptx'],
    emit() {
      throw new Error('layout: not supported on the docbuilder-create route -- measured: oSlide.ApplyLayout() silently kills the whole job (outcome "blocked", zero output, zero error message). Whether the co-editing route (editor.callCommand) has the same gap is not measured.')
    },
  },

  // MEASURED via full method enumeration on the Slide
  // object (same technique as pageSetup on the xlsx side) -- exactly two note-related methods
  // exist: GetNotesPage and AddNotesText. AddNotesText(text) WORKS, confirmed on all THREE parts
  // a speaker note actually needs (a notesSlide XML part can exist
  // without its slide-rels relationship or its [Content_Types].xml Override, and PowerPoint never
  // shows it then -- checking only the XML part would be a false green):
  //   ppt/notesSlides/notesSlideN.xml           -- created, with the exact requested text
  //   ppt/slides/_rels/slideN.xml.rels          -- carries a notesSlide relationship to it
  //   [Content_Types].xml                        -- carries the matching Override entry
  // Also confirmed: the note text does NOT leak onto the visible slide (slide1.xml's own <a:t>
  // stayed empty), and an out-of-range slide index is refused before this operation even runs
  // (the shared oSlide/pptxSlideCount check above, same as every other pptx operation here).
  speakerNotes: {
    cores: ['pptx'],
    emit(op) {
      if (!op.text) throw new Error('speakerNotes: `text` is required')
      return [`oSlide.AddNotesText(${jsString(String(op.text))});`]
    },
  },

  // Slide.AddComment(posX, posY, text, author, userId) --
  // toString()-recovered, package-verified: writes ppt/comments/comment1.xml + ppt/
  // commentAuthors.xml, a slide-rels `comments` relationship, and the matching [Content_Types].xml
  // Override for both parts -- same three/four-part discipline as speakerNotes' own comment above
  // (a comment missing any one of these is invisible in a real viewer, checking only the XML part
  // would be a false green).
  //
  // TWO traps found live, neither visible from the call's own return value or an "ok:true" job:
  //   1. `text` falsy (e.g. "") -> the call returns `false`, no exception, NO comment written --
  //      confirmed: a second AddComment call with "" in the same script produced only ONE <p:cm>
  //      entry (the other, non-empty call), not two. Refused here rather than a silent no-op.
  //   2. `author` omitted -> the call still returns `true` and DOES write a comment, but
  //      `AscCommon.UserInfoParser.getCurrentName()` (the source's own fallback) returns EMPTY in
  //      this headless DocBuilder context -- package-verified: commentAuthors.xml lands with
  //      `name=""` `initials=""`, an anonymous-looking artifact, not a sensible default identity.
  //      Treated as required here rather than silently producing that.
  comment: {
    cores: ['pptx'],
    emit(op) {
      if (!op.text) throw new Error('comment: `text` is required (an empty/falsy text silently returns false from AddComment, no comment gets written)')
      if (!op.author) throw new Error('comment: `author` is required (omitting it does not fall back to a real name in this environment -- measured: the saved commentAuthors.xml lands with an empty name/initials)')
      const x = Number(op.x ?? 500000)
      const y = Number(op.y ?? 500000)
      return [`oSlide.AddComment(${x}, ${y}, ${jsString(op.text)}, ${jsString(op.author)});`]
    },
  },

  // Slide-level operations -- background, slide-instance
  // management (delete/duplicateTo/moveTo), transition, visibility, content removal, object
  // grouping. All call shapes toString()-recovered and package-verified against a live,
  // multi-slide Document Server session -- not taken from
  // documentation alone. `delete`/`duplicateTo`/`moveTo` mutate the SLIDE COLLECTION and are
  // therefore refused in combination with each other or with any content field (background/
  // visible/removeAllObjects/removeObject/transition/group) in the SAME operation call: there is
  // no sensible execution order for "delete this slide and also recolor it".
  //
  // *** INDEX-DRIFT CAVEAT, NAMED NOT SOLVED: *** each pptx operation resolves its own `slide`
  // index against the presentation's CURRENT state at the point it runs (same mechanism as every
  // other pptx operation in this table) -- but `delete`/`duplicateTo`/`moveTo` change what index
  // N refers to for every operation AFTER them in the same batch (sequential-script semantics,
  // not batch-atomic ones). A caller mixing a structural slide op with later ops addressing
  // slides by a pre-computed index should expect drift. Not enforced here; keep structural slide
  // edits to one per create_document call until a stronger guard is built.
  //
  // Deleting the presentation's LAST remaining slide is UNMEASURED (not run against this
  // instance) -- refused nowhere client-side, but nothing here claims it is safe either.
  slide: {
    cores: ['pptx'],
    emit(op) {
      const structural = ['delete', 'duplicateTo', 'moveTo'].filter((k) => op[k] !== undefined && op[k] !== null && op[k] !== false)
      if (structural.length > 1) {
        throw new Error(`slide: only one of delete/duplicateTo/moveTo may be requested per operation -- got ${structural.join(', ')}`)
      }
      // `visible` is a real boolean value (false = hide, not "unset") -- unlike removeAllObjects,
      // which is a fire-or-not trigger, so it is checked for definedness, not truthiness.
      const hasContent = (op.background !== undefined && op.background !== null)
        || (op.visible !== undefined && op.visible !== null)
        || Boolean(op.removeAllObjects)
        || (op.removeObject !== undefined && op.removeObject !== null)
        || (op.transition !== undefined && op.transition !== null)
        || (op.group !== undefined && op.group !== null)
      if (structural.length === 1 && hasContent) {
        throw new Error(`slide: ${structural[0]} cannot be combined with content fields (background/visible/removeAllObjects/removeObject/transition/group) in the same operation -- the execution order would be arbitrary`)
      }
      if (structural.length === 0 && !hasContent) {
        throw new Error('slide: at least one field is required (background/visible/removeAllObjects/removeObject/transition/group/delete/duplicateTo/moveTo)')
      }

      if (op.delete) {
        return ['oSlide.Delete();']
      }
      if (op.duplicateTo !== undefined && op.duplicateTo !== null) {
        const pos = Number(op.duplicateTo)
        if (!Number.isInteger(pos) || pos < 0) throw new Error(`slide: duplicateTo must be a non-negative integer, got ${JSON.stringify(op.duplicateTo)}`)
        return [`oSlide.Duplicate(${pos});`]
      }
      if (op.moveTo !== undefined && op.moveTo !== null) {
        const pos = Number(op.moveTo)
        if (!Number.isInteger(pos) || pos < 0) throw new Error(`slide: moveTo must be a non-negative integer, got ${JSON.stringify(op.moveTo)}`)
        return [`oSlide.MoveTo(${pos});`]
      }

      const lines = []
      if (op.background !== undefined && op.background !== null) {
        lines.push(resolveSlideBackgroundLine(op.background))
      }
      if (op.visible !== undefined && op.visible !== null) {
        lines.push(`oSlide.SetVisible(${Boolean(op.visible)});`)
      }
      if (op.removeAllObjects) {
        lines.push('oSlide.RemoveAllObjects();')
      }
      if (op.removeObject !== undefined && op.removeObject !== null) {
        lines.push(buildSlideRemoveObjectLine(op.removeObject))
      }
      if (op.transition !== undefined && op.transition !== null) {
        lines.push(...buildSlideTransitionLines(op.transition))
      }
      // `group` -- Slide.GroupDrawings(aDrawings) needs live shape-object REFERENCES (each with
      // a `.Drawing` on the current slide), not indices or ids -- there is no bound operation yet
      // for looking up an EXISTING, previously-added shape by position (that is the Slide.
      // GetAllShapes/GetAllDrawings read-back surface, K7 territory, not bound in this unit). So
      // `group` here creates its own shapes fresh and groups them in the same call -- grouping
      // shapes added by an earlier, separate operation in the same batch is NOT supported and
      // would need that read-back binding first.
      if (op.group !== undefined && op.group !== null) {
        lines.push(...buildSlideGroupLines(op.group))
      }
      return lines
    },
  },
}
