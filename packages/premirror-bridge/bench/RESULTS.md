# Segment-fit measurement: benchmark results

The segment-fit seam routes plain-text line breaking through a prepare-once
segmentation + pure-arithmetic line fitter (`@chenglou/pretext`) instead of the
legacy word-walk's per-call word re-measurement and `findMaxFittingLength`
slice-probe binary search. The metric that matters is the number of canvas
`measureText` calls: real-canvas `measureText` (font shaping) is the expensive
operation, and avoiding it is the point of the seam.

Three harnesses, from most deterministic to most realistic.

## 1. Deterministic parity (frozen in `src/pretextParity.test.ts`)

Both engines measured through a fixed 5px/char fake canvas (linear, kerning-free
— widths agree by construction, so any divergence is algorithmic). Frozen,
probe-verified:

| case | legacy calls | segment-fit calls |
| --- | ---: | ---: |
| first-pass 100-word paragraph (w120) | 199 | 127 |
| overlong 400-char token (w50) | 82 | 3 |
| repeat measure (both paths) | 0 | 0 |

Line-break and width parity is **exact** on spaced text, trailing-space edges,
overlong tokens, and space-less CJK. The engine declines offset-unsafe text
(CR/FF) and measures legacy-identically there.

## 2. Cold per-paragraph micro-benchmark (`bench/measure-engine.mjs`)

Real `measureParagraph` hot path, deterministic fake canvas, cache cleared before
each measure (true first-paint), warmup dropped, 200 reps. Cold canvas
`measureText` calls per paragraph, legacy vs segment-fit:

| archetype | width | legacy | segment-fit | saved |
| --- | ---: | ---: | ---: | ---: |
| prose 100w | 120 | 201 | 38 | 81% |
| prose 100w | 600 | 201 | 8 | 96% |
| prose 400w | 120 | 801 | 163 | 80% |
| overlong 400c | 120 | 55 | 1 | 98% |
| CJK 120c | any | 23 | 1 | 96% |
| mixed 16-run paragraph | any | ~196 | ~186 | ~5% |

Warm (steady-state) re-measurement: **0 extra calls on both paths** — folio's
width cache already covers repeats, so the honest win is first-pass and
changed-text measurement, not steady state. Multi-run paragraphs barely benefit
(the cross-run word walk is unchanged).

Run: `bun bench/measure-engine.mjs`

## 3. Real-browser corpus (`bench/measure-browser.mjs`)

The synthetic single-paragraph reductions above do **not** transfer uniformly to
real documents. This harness loads real `.docx` into the actual editor in real
headless Chromium with real fonts (Carlito/Arimo), a fresh browser context per
(document, engine) so the measure caches start cold, and patches the **real**
`CanvasRenderingContext2D.measureText` to count calls. It compares the pretext
path (default) against the legacy walk (`?segmentfit=off`).

A corpus of 500 real `.docx` (2 KB – 884 KB, median ~12 KB), 492 paired cleanly:

| metric | result |
| --- | ---: |
| aggregate canvas `measureText` calls | 92,129 → 84,517 (**8.3% fewer**) |
| documents with fewer calls | 166 / 492 (34%) |
| documents unchanged | 280 (57%) |
| documents +1 call (noise) | 46 |
| per-document reduction, median (all docs) | 0% |
| per-document reduction, median (among the 166 it engages) | ~87% (up to 97%) |

Engagement is concentrated in small, simple, plain-text documents:

| size | docs | engaged | calls saved |
| --- | ---: | ---: | ---: |
| < 5 KB | 104 | 98% | 82.6% |
| 5–15 KB | 261 | 10% | 4.7% |
| 15–50 KB | 83 | 23% | 1.2% |
| > 50 KB | 44 | 41% | 7.4% |

**Layout parity (correctness):** 490 / 492 documents paginate to an identical
page count with the seam on vs off. **2 large documents shifted by one page.**
This is the seam's documented break-rule divergence from the legacy walk
(trailing whitespace hangs past the line edge; CJK/Thai breaks between
characters) — both OOXML/CSS-defensible on the segment-fit side, but not
byte-identical to the legacy walk on those two documents.

Run (needs a served playground on `:4200` and a TSV of `size<TAB>path` rows):
`FIXTURES=corpus.tsv bun bench/measure-browser.mjs`

## Reading

Where the seam engages it is a large win (80–98% fewer measurement calls on the
paragraphs the legacy walk re-measures word-by-word or slice-probes), it is a
no-op at steady state, and its cost when it does not help is +1 call. The
aggregate real-document win is modest (~8%) because most real paragraphs are
short or otherwise ineligible, but the win is real, concentrated, and never
negative beyond noise. The two one-page pagination shifts are the one behavior
change to weigh against enabling it by default.
