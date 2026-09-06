# stella-text-shaper

`stella-text-shaper` turns a run of text and a font into positioned glyphs.

Shaping is the part of text layout that cannot be derived from code points. An
Arabic letter takes its form from the letters beside it, lam followed by alef
ligates, a Devanagari cluster reorders and forms a conjunct, a Hebrew point
hangs off the letter it belongs to. One code point does not select one glyph in
any of these scripts, and a renderer that assumes it does produces text a reader
of the script sees as broken at a glance.

One entry takes font bytes, a run of text, a direction, an optional script and
language, and the OpenType features to force on or off. It returns glyph ids,
clusters, advances and offsets in the font's own design units, so the caller
scales them: a measurement in CSS pixels and a PDF text matrix in points come
from the same numbers without either rounding the other's.

The crate supports native Rust consumers directly. Browser and server consumers
enable the `wasm` feature; `@stll/folio-core`'s `src/shaping/shaper.ts` provides
the corresponding TypeScript binding, and is the only module that imports the
generated artifact.

## The artifact

The WebAssembly artifact is committed under `packages/core/src/generated` with
its own size budget, separate from the DOCX kernel's. It is fetched the first
time a document actually contains a run that shapes: a document in Latin,
Cyrillic or Greek never loads it.

Regenerate it with `bun --filter @stll/folio-core wasm:generate`. The committed
bytes are compared against a fresh build, and CI builds on linux/amd64, so on
any other platform regenerate through
`scripts/regenerate-wasm-canonically.sh @stll/folio-core`, which runs the same
step in that image. CI is still the arbiter of the bytes: when its drift check
fails it uploads what it built, and that is what to commit.

## What it does not do

It shapes one run. Splitting text into runs, resolving the bidirectional
algorithm, and choosing which face covers which character are the caller's, and
a run handed here must not span a direction change: the producer has already
split at every boundary and knows which way each run reads.
