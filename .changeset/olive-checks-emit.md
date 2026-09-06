---
"@stll/folio-core": minor
---

`compareDocx` now reports what its self-check found instead of only acting on
it, and can be asked for its best attempt when it cannot prove one.

Every successful result carries `verification`: `{ status: "verified" }`, or
`{ status: "unverified", failures }` where each failure names the invariant that
did not hold (`accept-reproduces-target` or `reject-reproduces-base`), the
projection field that diverged (`container`, `block-count`, `style`,
`list-level`, `invisible-structure`, `whitespace`, `text`), the story it
happened in, and a structural detail carrying no phrase of either document.

The default is unchanged: an unproven redline is refused, because a reader
cannot tell one that lost something from one that did not.
`CompareDocxRoundTripError` now names the invariant and the cause and carries
the whole failure list, in place of the two block-text arrays it used to hold.

`onUnverified: "emit"` is the opt-in for the other trade — the redline it could
build, plus the typed list of what it could not represent. A parse, apply or
serialize failure is still an error under either setting: there is no redline
to emit.

Both directions of the round trip are now checked. The self-check previously
proved only that accepting reproduces the target; it also proves that rejecting
reproduces the base it was written against.
