---
"@stll/folio-core": minor
---

`fromProseDoc` takes the reuse it is asked for. A third argument,
`{ reuse?: ProjectionReuse }`, names whether a record the editor did not change
may come back from the base document by reference. `"none"` is the default and
is what the function has always done: every record is rebuilt out of
ProseMirror. `"matched"` is the merge against a matched base record, and it
panics until it is implemented rather than falling back to a rebuild, so a
caller cannot believe it asked for a merge and get today's behaviour.

The option ships before the merge because the measurement has to. The corpus
gate's new `editor-projection` invariant is `editor-round-trip` with
`{ reuse: "none" }` forced, the way `reserialize` strips the capture slots so
the serializers must run. Once reuse lands, `editor-round-trip` measures the
merge and `editor-projection` measures the projection; until then the two are
the same measurement, which is the point of adding the second one first.
