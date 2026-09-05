# Compare benchmark results

Baseline for `compareDocx`, taken before any optimization work so every later
change can report its effect against it. Reproduce with `bun run bench:compare`;
`README.md` says what each column means.

## How to read these numbers

Absolute milliseconds belong to one machine on one day. What survives a change
of hardware is the shape: which stage dominates, how cost scales with block
count, and which invariants fail. Read the shape; re-measure the milliseconds.

Every run records the host's one-minute load average, because a run taken on a
busy machine is not comparable with one taken idle.

## Run

- Apple arm64, 8 logical CPUs, Bun 1.4.2, macOS.
- Sizes `s` and `m`: 4 warm-ups, 9 iterations, 144 configurations, host load
  21.6 falling to 10.0. **The machine was busy.** Treat the absolute
  milliseconds as an upper bound; the ratios below are what to read.
- Size `l`, class `prose`: same sampling, measured separately under comparable
  load.
- Schema validation was skipped: this machine has no .NET toolchain. CI has
  one, so the gate runs there.

## Where the time goes

Summed medians over all 127 measured configurations at `s` and `m`:

| Stage       | Share |
| ----------- | ----- |
| `parse`     | 35.5% |
| `align`     | 0.4%  |
| `apply`     | 41.7% |
| `serialize` | 22.3% |

**The comparison algorithm is not the cost.** Alignment — segment, row and cell
pairing, move detection, the whole planning pass — is 0.4% of the work. The
other 99.6% is converting two packages into editor models, writing tracked
changes through the applier, and serializing one package back out. Any
optimization aimed at the diff itself is aimed at the wrong stage.

## Size `m` (320 blocks), by class

Wall time is the median of the whole comparison; the stage columns are medians
of the same runs measured separately.

| Configuration           | Blocks | Changes | Wall     | parse   | align  | apply    | serialize |
| ----------------------- | ------ | ------- | -------- | ------- | ------ | -------- | --------- |
| `prose/m/identical`     | 320    | 0       | 83.8ms   | 57.8ms  | 0.71ms | 0.0ms    | 27.4ms    |
| `prose/m/light`         | 320    | 22      | 248.0ms  | 197.1ms | 1.46ms | 69.1ms   | 62.7ms    |
| `prose/m/structural`    | 320    | 145     | 255.4ms  | 88.3ms  | 1.83ms | 76.1ms   | 36.4ms    |
| `prose/m/rewrite`       | 320    | 320     | 1885.7ms | 245.0ms | 1.58ms | 2403.7ms | 235.6ms   |
| `lists/m/identical`     | 320    | 0       | 171.9ms  | 124.9ms | 1.19ms | 0.0ms    | 94.8ms    |
| `lists/m/rewrite`       | 320    | 320     | 688.8ms  | 71.3ms  | 0.96ms | 452.1ms  | 63.6ms    |
| `tables/m/identical`    | 217    | 0       | 60.0ms   | 51.9ms  | 1.22ms | 0.0ms    | 30.4ms    |
| `tables/m/rewrite`      | 217    | 217     | 247.3ms  | 46.3ms  | 0.88ms | 110.5ms  | 46.5ms    |
| `notes/m/identical`     | 320    | 0       | 270.0ms  | 124.5ms | 0.71ms | 0.0ms    | 48.3ms    |
| `notes/m/rewrite`       | 320    | 320     | 655.2ms  | 93.4ms  | 0.71ms | 348.9ms  | 160.4ms   |
| `graphics/m/rewrite`    | 320    | 320     | 224.8ms  | 61.3ms  | 0.74ms | 122.7ms  | 38.4ms    |
| `fields/m/rewrite`      | 320    | 320     | 237.3ms  | 54.0ms  | 0.66ms | 97.7ms   | 32.0ms    |
| `sections/m/rewrite`    | 320    | 320     | 334.7ms  | 43.9ms  | 0.70ms | 270.0ms  | 38.2ms    |
| `multiscript/m/rewrite` | 320    | 320     | 242.7ms  | 26.0ms  | 0.50ms | 125.8ms  | 30.4ms    |
| `revised/m/rewrite`     | 320    | 320     | 333.6ms  | 42.2ms  | 0.63ms | 214.0ms  | 34.6ms    |

Right-to-left and CJK text (`multiscript`) is the cheapest class, not the most
expensive: the word diff pays nothing extra for bidirectional runs.

## Size `l` (2,200 blocks), class `prose`

| Configuration       | Wall      | parse    | align  | apply    | serialize |
| ------------------- | --------- | -------- | ------ | -------- | --------- |
| `prose/l/identical` | 2407.6ms  | 2075.6ms | 28.1ms | 0.0ms    | 1041.4ms  |
| `prose/l/light`     | 2987.4ms  | 4724.2ms | 35.3ms | 3381.6ms | 577.9ms   |
| `prose/l/heavy`     | 15872.9ms | 1114.0ms | 14.8ms | 3989.9ms | 244.2ms   |
| `prose/l/churn`     | 4541.5ms  | 1492.0ms | 34.9ms | 3513.4ms | 239.3ms   |
| `prose/l/reorder`   | 1863.5ms  | 829.5ms  | 21.5ms | 561.7ms  | 113.8ms   |

The target is a large comparison well under one second. It is not close.

## Scaling

`identical` (parse plus serialize, no edit work at all):

| Class         | `s`            | `m`              | Blocks | Time |
| ------------- | -------------- | ---------------- | ------ | ---- |
| `prose`       | 40 blk, 9.5ms  | 320 blk, 83.8ms  | x8.0   | x8.8 |
| `lists`       | 40 blk, 37.2ms | 320 blk, 171.9ms | x8.0   | x4.6 |
| `tables`      | 28 blk, 15.9ms | 217 blk, 60.0ms  | x7.8   | x3.8 |
| `notes`       | 40 blk, 36.6ms | 320 blk, 270.0ms | x8.0   | x7.4 |
| `graphics`    | 40 blk, 17.0ms | 320 blk, 83.6ms  | x8.0   | x4.9 |
| `fields`      | 40 blk, 24.1ms | 320 blk, 99.1ms  | x8.0   | x4.1 |
| `sections`    | 40 blk, 9.0ms  | 320 blk, 66.1ms  | x8.0   | x7.3 |
| `multiscript` | 40 blk, 10.6ms | 320 blk, 50.5ms  | x8.0   | x4.8 |
| `revised`     | 40 blk, 17.9ms | 320 blk, 75.5ms  | x8.0   | x4.2 |

Linear or better from `s` to `m`. From `m` to `l` it is not: `prose/identical`
goes 320 blocks at 83.8ms to 2,200 blocks at 2,407.6ms, which is 6.9 times the
blocks for 28.7 times the time. Something in parse or serialize is superlinear
above roughly a thousand blocks, and that is the first thing to profile.

## What to fix, in order

1. ~~**Find the superlinear term above ~1,000 blocks.**~~ Done; see below.
2. ~~**A comparison that found nothing should not rewrite the document.**~~
   Done; see below.
3. **Align on the parsed model, convert only what changed.** Parse is 35.5% of
   the total and is paid in full on both sides even when three paragraphs
   differ. The alignment needs block text and container coordinates, not a
   ProseMirror document.
4. **Apply is the largest single share (41.7%) and superlinear in change
   count.** `prose/m/rewrite` spends 2,404ms applying 320 changes, against
   452ms for the same count in `lists/m/rewrite`; the applier re-resolves
   snapshot anchors against the live document per operation.

Alignment needs no work. It is already 0.4%.

## Changes measured against this baseline

### A comparison that found nothing returns the base package

`serializeComparison` hands back the arriving bytes when no story yielded an
operation and the base carried no revisions of its own.

| Configuration       | serialize before | serialize after |
| ------------------- | ---------------- | --------------- |
| `prose/l/identical` | 1041.4ms         | 0.0ms           |
| `prose/m/identical` | 27.4ms           | 0.0ms           |
| `lists/m/identical` | 94.8ms           | 0.0ms           |
| `notes/m/identical` | 48.3ms           | 0.0ms           |

Read the stage column, not the wall column: the two runs were taken under
different host load, so wall time is not comparable between them, while the
serialize stage going to zero is the change itself.

`revised/*` still serializes, and should: its base carries someone else's
revisions, so the compared base is the accepted view and the arriving bytes are
a different document. The benchmark shows the distinction directly —
`revised/s/identical` keeps a non-zero serialize stage while every other class's
`identical` case drops to zero.

Digests moved for 18 configurations (every clean-base `identical`, plus the two
`notes/*/notes` cases, which also plan nothing) because their product is now
the input package rather than a re-serialization of it. Every `revised/*`
digest moved as well: the revision id seed is now read from the package as it
arrived, above any ids a previous reviewer used, rather than after resolution
when none remain.

The benchmark caught the first attempt at this: the short-circuit sat in
`compareDocx` rather than in the stage, so the harness's own composition of the
stages disagreed with the shipped one and `byte-determinism` failed on all 18
`identical` cases within one run. The property suite then caught the second
attempt, where the base carried revisions.

### The corpus stopped changing every two seconds

`--check` reported 18 drifted digests on an unmodified tree, and different ones
depending on the minute it ran in. The generator pins a fixed date on every
part it writes, but JSZip synthesizes a folder entry per directory in a part
name and stamps that one with `new Date()`; DOS timestamps have two-second
granularity, so two runs in one bucket agreed and two runs a minute apart did
not. It surfaced only once an identical pair started returning the base package
as it arrived: every other product is restamped on the way out.

Every digest moved, because no generated package carries folder entries any
more. `--check` now reproduces all 127 across runs, which it did not before.
Until it did, the digest baseline could not prove an optimization changed
nothing, which is the whole reason it exists.

### A paragraph appended after a nested table lands at body level

`tables/m/structural` failed the round-trip check. The generated tables carry a
nested table in one cell, so the story's last block is two levels deep, and the
target's appended body paragraph had no anchor outside a table. A block
insertion escapes the table it is anchored in, but it escaped only the
innermost one, leaving the paragraph in the outer cell.

| Configuration         | before                      | after               |
| --------------------- | --------------------------- | ------------------- |
| `tables/m/structural` | `CompareDocxRoundTripError` | 8 changes, 49.3ms   |
| `tables/l/structural` | `CompareDocxRoundTripError` | 20 changes, 488.0ms |

The fix is in folio's insertion semantics, not in the compare:
`findOutermostTableBoundary` replaces `findEnclosingTableBoundary` for
`insertBeforeBlock`, `insertAfterBlock`, and `insertSignatureTable`. One digest
was added and no other moved, which is the evidence that the change reached
only the case it was aimed at. `probes.test.ts` carries the minimal
reproduction as `append_after_nested_table`.

### The superlinear term above ~1,000 blocks was one `doc.resolve` per block

`createFolioAIEditSnapshot` asked the document where each block sat, twice:
once for the hidden-row check and once for the table coordinates.
`doc.resolve` re-descends from the root and finds each level's child by
scanning that level's fragment from index 0, so on a flat document it costs
O(blocks) per block and the snapshot costs O(blocks^2). The snapshot runs on
both sides, again per reviewed view, so parse and apply both carried it.
`formatStoryStateForLLM` had the same shape with `doc.nodeAt` per block.

A depth-first walk already visits every ancestor before the block, so the
snapshot now carries the path it is on and resolves nothing.

Standalone, one document class, over a growing paragraph count (ms, median of
one run each; the shape is the point, not the milliseconds):

| Paragraphs | snapshot before | snapshot after | final view before | final view after |
| ---------- | --------------- | -------------- | ----------------- | ---------------- |
| 250        | 12.3            | 7.8            | 16.7              | 7.3              |
| 1,000      | 16.2            | 7.6            | 52.9              | 21.6             |
| 2,000      | 49.4            | 13.3           | 130.9             | 35.3             |
| 4,000      | 124.0           | 24.2           | 381.4             | 79.2             |

Four times the blocks cost 7.7x before and 3.2x after: quadratic to linear.

In the compare benchmark, `prose/l`, both runs back to back on an idle machine
(the baseline table above was taken under load 21.6 and its absolute
milliseconds are not comparable with these):

| Configuration        | parse before | parse after | wall before | wall after |
| -------------------- | ------------ | ----------- | ----------- | ---------- |
| `prose/l/identical`  | 273.8ms      | 171.9ms     | 292.2ms     | 196.5ms    |
| `prose/l/light`      | 346.2ms      | 176.4ms     | 618.9ms     | 394.5ms    |
| `prose/l/heavy`      | 388.4ms      | 182.4ms     | 1794.3ms    | 1382.6ms   |
| `prose/l/churn`      | 391.4ms      | 188.2ms     | 1288.7ms    | 1019.5ms   |
| `prose/l/reorder`    | 325.4ms      | 200.3ms     | 722.6ms     | 428.2ms    |
| `prose/l/structural` | 395.3ms      | 213.7ms     | 9490.6ms    | 8923.3ms   |

Scaling of `identical`'s parse stage, same two runs:

| Blocks | before  | after   |
| ------ | ------- | ------- |
| 40     | 3.9ms   | 4.0ms   |
| 320    | 29.0ms  | 28.3ms  |
| 2,200  | 273.8ms | 171.9ms |

6.9 times the blocks from `m` to `l` cost 9.4x before and 6.1x after. The
superlinear term is gone; what is left tracks block count.

Every one of the 128 digests is unchanged, which is the proof that the
optimization changed nothing.

The guard is an invariant, not a stopwatch: `snapshot.test.ts` hands the
snapshot a document that throws if asked to resolve a position. Writing that
test surfaced a second defect: the hidden-row check consulted the block's
NEAREST row, so a table nested inside a hidden row published its text after
all. The walk now skips a hidden row's whole subtree.

## Correctness gaps the baseline surfaced

Three configurations failed, and each named a real gap rather than a flake.

- **`notes/s/notes`, `notes/m/notes` — `difference-is-reported` fails.** The
  pair differs only in the footnote and endnote stories. The comparison reports
  nothing and lists `secondary-story` in `unsupported`, so a caller who reads
  only `changes` is told two different documents agree. This is the documented
  main-story-only limitation, now with a number on it: 30 of 127 configurations
  carry an unreported story. Closing it needs per-story revision id ranges that
  do not collide, which the operation result does not currently expose.
- ~~**`tables/m/structural` — `CompareDocxRoundTripError`.**~~ Fixed; see "A
  paragraph appended after a nested table" above.
- **`change_list_level`** is not in the benchmark because the harness cannot
  build a target for it that the engine can see at all; it is covered by a
  probe in `packages/core/src/compare/probes.test.ts` instead.

Everything else passes every invariant: the round-trip algebra in both
directions, self-comparison, and byte determinism, across all nine classes at
both sizes.
