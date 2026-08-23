// The KEPESSEG-STATUSZ MODEL (the owner's 6th/7th/11th point) --
// the ten-state classification + the promotion rule that decides when a capability may be marked
// a STABLE blocked/ENGINE_BUG, as opposed to the weaker SUSPECT ceiling every capability sits
// under today (see ceilingFor below).
//
// Entry point matters (E1 spec, "belepesi pont" section): the SAME method
// can behave differently through create_document (DocBuilder) vs coedit_write_operations
// (co-editing callCommand) vs run_builder_script (raw DocBuilder) -- CreateNumbering's own case
// tonight (fails via create_document, works via run_builder_script+SetBullet) is the concrete
// example that forced this field into the spec. A status without its entry point names nothing.

export const STATUS = Object.freeze({
  SUPPORTED_VERIFIED: 'SUPPORTED_VERIFIED',
  SUPPORTED_UNVERIFIED: 'SUPPORTED_UNVERIFIED',
  SUSPECT: 'SUSPECT',
  SIGNATURE_MISMATCH: 'SIGNATURE_MISMATCH',
  NO_OP: 'NO_OP',
  THROWS: 'THROWS',
  JOB_ABORT: 'JOB_ABORT',
  OUTPUT_CORRUPT: 'OUTPUT_CORRUPT',
  REQUIRES_EDITOR_CONTEXT: 'REQUIRES_EDITOR_CONTEXT',
  ENGINE_BUG: 'ENGINE_BUG',
})

const STABLE_NEGATIVE_STATUSES = new Set([STATUS.JOB_ABORT, STATUS.ENGINE_BUG])

// the owner's 7. pont, SIX conditions, all required for a STABLE blocked/ENGINE_BUG (or JOB_ABORT)
// classification -- a single strange result caps at SUSPECT, no matter how strange.
export function sixConditionsMet(evidence) {
  const e = evidence || {}
  const missing = []
  if (!e.minimalReproducerRan) missing.push('minimalReproducerRan')
  if (!e.runtimeVerifiedSignature) missing.push('runtimeVerifiedSignature')
  if (!(Number(e.repeatCount) >= 3)) missing.push('repeatCount>=3')
  if (!e.sameSdkFingerprint) missing.push('sameSdkFingerprint')
  if (!e.controlOperationsRun) missing.push('controlOperationsRun')
  // "lehetoleg kulon fixture-on" (preferably) -- the ONE condition of the six that is a SHOULD,
  // not a MUST ("lehetoleg"). Tracked but does not block promotion.
  return { met: missing.length === 0, missing, separateFixtureRecommendedButMissing: !e.separateFixture }
}

// The ceiling rule: a STABLE
// blocked/ENGINE_BUG classification additionally requires `sameSdkFingerprint` to be TRUE, which
// itself requires the fingerprint to have been MEASURED at all (E1's own 11 required engine
// fields are ALL NEM-MERT as of e41978ab) -- so today, sixConditionsMet() can never
// return met:true for a case whose fingerprint is unmeasured, and this function is the place that
// makes that ceiling explicit and testable on its own, independent of the other five conditions.
export function ceilingFor(evidence) {
  const e = evidence || {}
  if (e.sdkFingerprintMeasured !== true) {
    return { maxStatus: STATUS.SUSPECT, reason: 'sdkSha256 NEM-MERT (E1, measured as of e41978ab) -- the owner\'s 7th point stabil blocked/ENGINE_BUG-hoz UGYANAZON SDK fingerprint melletti megerositest kovetel; amig ez nincs merve, a feltetel nem teljesithetheto, a plafon SUSPECT' }
  }
  return { maxStatus: null, reason: null } // no ceiling from this rule -- other rules may still apply
}

// Combines sixConditionsMet + ceilingFor into the actual classify-or-refuse decision for the two
// STABLE NEGATIVE statuses (JOB_ABORT, ENGINE_BUG). Every OTHER status (SUPPORTED_VERIFIED,
// NO_OP, THROWS, SIGNATURE_MISMATCH, REQUIRES_EDITOR_CONTEXT, OUTPUT_CORRUPT,
// SUPPORTED_UNVERIFIED) is NOT gated by this rule -- the owner's 7. pont names "stabil ENGINE_BUG/
// blocked" specifically, not a positive verified-working classification (measured differently:
// package-verified content is its own proof, it does not need a stable-failure repeatability bar).
export function classify(requestedStatus, evidence) {
  if (!STABLE_NEGATIVE_STATUSES.has(requestedStatus)) {
    return { status: requestedStatus, downgraded: false, reason: null }
  }
  const ceiling = ceilingFor(evidence)
  if (ceiling.maxStatus) {
    return { status: ceiling.maxStatus, downgraded: true, reason: ceiling.reason }
  }
  const six = sixConditionsMet(evidence)
  if (!six.met) {
    return { status: STATUS.SUSPECT, downgraded: true, reason: `the owner's 7th point's hat feltetelebol hianyzik: ${six.missing.join(', ')} -- egyetlen furcsa eredmeny utan a maximum SUSPECT` }
  }
  return { status: requestedStatus, downgraded: false, reason: null }
}

// 3. kikotes: two capability-status records are
// comparable ONLY if they share the SAME entry point (create_document | run_builder_script |
// coedit | egyeb) -- a status measured on one path is NEM-MERT on another, not inherited (the
// NEGATIVE half of the rule, without which the three schema stipulations are only
// half-valid). `a`/`b`: { entryPoint, sdkSha256 (string | 'NEM-MERT' | undefined) }.
export function comparable(a, b) {
  if (!a || !b) return { comparable: false, reason: 'hianyzo bejegyzes' }
  if (a.entryPoint !== b.entryPoint) {
    return { comparable: false, reason: `kulonbozo belepesi pont (${a.entryPoint} vs ${b.entryPoint}) -- egy uton mert allapot a masikra NEM ervenyes es NEM oroklodik` }
  }
  return { comparable: true, reason: null }
}

// (A) vs (B) szetvalasztas (egy korabbi felvetes korrigalva):
// (A) szabad megallapitani, hogy VALTOZOTT-e egy statusz akkor is, ha az egyik/mindket
// oldal fingerprintje hianyzik -- a valtozas TENYE all, csak az OKA (build-e) nem attribualhato.
// (B) egy statusz SOHA nem allithatja, hogy egy MASIK futasra is all, ha a fingerprint hianyzik --
// se elore (kesobbi futas), se visszafele (korabbi futas).
export function flipDetectable(a, b) {
  const cmp = comparable(a, b)
  if (!cmp.comparable) return { flipKnown: false, reason: cmp.reason }
  const aKnown = typeof a.sdkSha256 === 'string' && a.sdkSha256 !== 'NEM-MERT'
  const bKnown = typeof b.sdkSha256 === 'string' && b.sdkSha256 !== 'NEM-MERT'
  return {
    // (A): a flip TENYE megallapithato a statusz-ertekek osszevetesevel FUGGETLENUL attol, hogy a
    // fingerprintek ismertek-e -- csak az OK-attribucio (FLIP_FINGERPRINT_UNKNOWN vs valodi build-
    // valtozas) fugg tole, azt ez a fuggveny nem donti el.
    flipKnown: a.status !== b.status,
    attributableToBuildChange: aKnown && bKnown && a.sdkSha256 !== b.sdkSha256,
    reason: (!aKnown || !bKnown)
      ? 'legalabb az egyik oldal fingerprintje NEM-MERT -- a valtozas TENYE megallapithato, az OKA (build-e) nem attribualhato'
      : null,
  }
}
