# Public DOCX corpus gate

Property tests cover the inputs someone thought to describe. This gate covers
the ones nobody did: thousands of real `.docx` packages from public test
suites, run through folio's entry points and checked against invariants rather
than expected output. A file needs no oracle to be useful, so any package that
opens is evidence.

## Running it

```sh
bun run corpus:fetch          # download the pinned corpus into the cache, relock
bun run corpus:gate           # run every invariant, then ratchet against the baseline
bun run corpus:check          # verify the cached corpus against corpus/sources.lock.json
```

`corpus:gate` takes `--concurrency N` (default 4), `--timeout MS` (per file,
default 300000), `--tiers 1,2` (default 1, see below), `--invariant-budget MS`,
`--file-budget MS`, `--shard k/n`, `--only ID[,ID...]` and `--out FILE`.
`report <census.json...>` prints a census without ratcheting it. CI shards four
ways and merges the censuses before the ratchet, because each shard sees only a
subset:

```sh
bun scripts/corpus-gate.ts run --shard 1/4 --out census-1.json
bun scripts/corpus-gate.ts check census-1.json census-2.json census-3.json census-4.json
```

`--only` narrows a run to named files, matching a `<source-id>/<path>` file id
in full or in part, so working on one signature costs seconds rather than the
hour a census takes:

```sh
bun scripts/corpus-gate.ts run --only apache-poi/test-data/document/55733.docx --concurrency 2
```

A pattern that matches no file is an error, not an empty run. A subset census
carries neither the rest of the corpus's signatures nor a baseline entry's file
counts, so `--only` refuses `--check` and `--shard`, and a baseline is never
written from one: measure the subset before and after a fix, then shrink the
affected family's baseline by the difference.

The gate is nightly (`.github/workflows/nightly-corpus-gate.yml`) and on
`workflow_dispatch`. It is not part of PR CI: it downloads a hundred megabytes
and takes the better part of an hour. Tier 1 is what the baseline ratchet runs;
tier 2 runs in a job of its own that keeps its census in the job.

## The invariants

| Invariant           | What must hold                                                                                                                                                          |
| ------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `parse`             | A package a reader can open parses.                                                                                                                                     |
| `fixed-point`       | parse → ProseMirror → repack → parse preserves the visible text and the text-block count. Same equality as `packages/core/src/docx/__tests__/corpusFixedPoint.test.ts`. |
| `repack-validates`  | The package folio just wrote satisfies folio's own package validator.                                                                                                   |
| `style-set-rebuild` | Extracting a document's style set and building a package from it does not panic.                                                                                        |
| `completes`         | The worker answered for the file: it did not die, and it did not answer out of order.                                                                                   |

Each file runs in a child process with a deadline, so an abort is a recorded
finding rather than a lost run. The deadline expiring is recorded under
`performance` instead: a watchdog cannot tell a hung worker from a slow
machine.

Those five keep `corpus/baseline.json`. The rest own one baseline file each under
`corpus/baselines/`, so re-measuring one never rewrites another's findings —
except `performance`, which is measured and reported but never ratcheted:

| Invariant             | What must hold                                                                                                                                                                                                                     |
| --------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `reserialize`         | With every rebuildable capture removed, so the real serializers run for every block, the saved package parses back to the same model. A difference here is a serializer defect verbatim replay hides.                              |
| `editor-round-trip`   | Document → `toProseDoc` → `fromProseDoc` → save → parse preserves the whole normalised model, not only the visible text and block count `fixed-point` checks.                                                                      |
| `edit-locality`       | One character inserted in the first non-empty body paragraph changes that paragraph and nothing else: no other block's model, no part outside the body.                                                                            |
| `save-idempotence`    | Saving is a fixed point after the first normalising save. Every part is byte-stable from the second save on.                                                                                                                       |
| `schema-validity`     | A part folio rebuilds gains no schema violation it did not arrive with, against `specifications/generated/docx-transitional-schema.gen.json`.                                                                                      |
| `pipeline-totality`   | Layout, display list, PDF, markdown, the agents snapshot and the comparison engine as self-diff each run without throwing, and `compare(x, x)` reports no changes.                                                                 |
| `kernel-differential` | The Rust kernel (`crates/docx-kernel` through `@stll/docx-core/projection`) and the TypeScript parser agree on the facts they both produce.                                                                                        |
| `performance`         | Report-only. Records which files cost more than ten times what the corpus costs at their size, in parse time or in peak resident set, and which stages overran their per-file budget. See [The cost baseline](#the-cost-baseline). |

### Gating and report-only families

Every family is classified `gating` or `report-only` in `CORPUS_FAMILY_GATING`
(`scripts/lib/corpus-invariants/contract.ts`), a total map over the family
union, so a family added later cannot arrive without that decision.

`performance` is report-only. Its verdicts are wall-clock comparisons, and the
ratchet is exact in both directions, so a baseline measured on a loaded machine
fails a quiet run and a baseline measured idle fails a busy one. Its findings
are still measured, written to the census and printed in the report with the
slowest files per stage; they own no baseline file and are never compared. The
per-file watchdog reports under `performance` for the same reason: its expiry
cannot tell a hung worker from a slow machine.

Timing is recorded, not ratcheted. Deterministic performance guards are
separate work.

### Truncated files

A per-file budget can stop a file's run part-way, and which invariants got to
run then depends on how busy the machine was. Such a file is `truncated`: it
contributes **no** gating findings at all, including the ones it produced
before the budget ran out, because a slower run would have stopped sooner and
reported fewer. Its timing findings are kept, it is counted, and it is listed
in the report under its own heading with the stage it stopped at.

The rule is structural rather than remembered: `CensusBuilder.add` takes a
`CorpusFileResult` (`complete` | `truncated` | `report-only` | `not-a-docx`)
and reads it with an exhaustive switch, so a result kind added later has to
state whether its findings gate.

Whether a file truncates depends on how loaded the machine was, so truncation
may not decide what is compared. Any file that stops at a budget without being
named in `corpus/report-only-files.json` makes the whole run degraded: it is
not compared, `write-baseline` refuses it, and the message names the file and
the stage. There is no tolerated share.

### Report-only files

`corpus/report-only-files.json` names the files whose findings are measured and
never gated, each with the sha256 the corpus lock records and a reason. A listed
file is still fetched, run and timed — its cost stays in the `performance`
family and in the report — but it contributes no gating finding whether it
finished or stopped at a budget. That is the only way out of the gating set, so
the set of files the ratchet compares is committed data rather than a function
of the clock.

An entry must name a file the lock still carries, with those bytes, or the gate
fails: an exemption is granted for a file, not for a path, and a repin that
replaces the content is reviewed again. Both the core baseline and each family
baseline record the digest of the list they were measured under; while it
differs, a signature that shrank or vanished is reported as kept rather than
ratcheted down, since a listing can only remove evidence. Growth still fails.

The nightly raises the budgets far above what the corpus needs (120s per
invariant, 600s per file, a 900s worker deadline) so that on CI the only files
that can reach a budget are the listed ones.

### The cost baseline

Milliseconds per megabyte assumes parse cost passes through the origin, and it
does not: opening a package, reading its styles and its theme, and building an
empty document cost the same tens of milliseconds whether the body is one
paragraph or ten thousand. Dividing that fixed cost by a small file's
megabytes produces an enormous rate, so a single median over the corpus mixes
two populations — small files priced mostly by overhead, large files priced
mostly by content — and settles between them, too high to catch a quadratic
path in a large file and too low to leave a small one alone.

The baseline is therefore affine, `parseMs ≈ intercept + slope · bytes`, fitted
by the median of pairwise slopes so the outliers being hunted cannot drag it
towards themselves. A file's verdict is how many times its own prediction it
cost, which no longer depends on its size. Peak resident set is fitted and
judged the same way but separately, because memory amplification and slowness
are different defects: a package that parses at corpus speed while leaving the
worker holding a gigabyte is a finding the time rule would never report.

### Why `reserialize` exists

folio replays captured bytes rather than re-serializing what nobody edited: a
paragraph's `w:pPr`, a table's properties, a header's whole part, a drawing, a
content control's properties. Each has its own capture slot and its own
fingerprint gate. So a round trip over an untouched document exercises the
capture machinery, not the serializers, and a serializer that writes `left`
where the source said `start` passes every other invariant here. Every edited
document takes the serializer path, so those defects are live for users and
invisible to a gate that only round-trips.

Only slots the model can rebuild are stripped. A `preserveOnly` drawing and a
shape's fill or outline markup have no model behind them: their captured XML is
the content, and removing it would test deletion.

### A duplication to remove

The model projection the new invariants compare against
(`scripts/lib/corpus-invariants/model-equality.ts`) restates the one
`packages/core/src/docx/saveEquivalence.property.test.ts` defines, because that
one is test-local and not importable. Two hand-maintained copies of an equality
are a drift hazard: a normalisation added to one and not the other turns into a
census of phantom defects. The projection wants a single owner in
`packages/core` that both the property test and the gate import. That is a
change to a published package, so it is named here rather than smuggled in.

### Budgets

`--invariant-budget MS` (default 30000) and `--file-budget MS` (default 120000)
are advisory. Nothing can interrupt a synchronous serializer mid-call, so an
overrun is recorded as a `performance` finding after the fact and the file's
remaining invariants are skipped rather than silently passing. The worker
deadline (`--timeout`) is still the only hard stop, and its expiry truncates
the file for the same reason a budget does: it cannot tell a hung worker from
a slow machine.

### Producers

Every failure signature reports which producers trigger it, read from the
extended-properties part's `Application` and `AppVersion` plus two structural
tells: Word Online names the main part `word/document2.xml`, and a package
without that part at all was built by a library from nothing. Nothing else in
`docProps` is read; the core properties carry author names and the gate has no
business looking at them. A signature confined to one producer names that
producer's quirk; one spread across Word, LibreOffice and a generator names
something folio gets wrong about the format.

## What counts as a `.docx`

Public suites carry files that are `.docx` only by name. The gate classifies
those out before any invariant runs, and it does so independently of folio: it
reads the container, finds the main part through the package relationship (not
by the conventional `word/document.xml` path — Word Online writes
`word/document2.xml`), and requires a well-formed `w:document` root in a
Transitional or Strict WordprocessingML namespace. Reasons are reported
separately (`encrypted-package`, `unreadable-archive`, `not-an-opc-package`,
`not-a-wordprocessing-package`, `malformed-document-xml`, `not-a-zip`,
`ole-compound-file`) with up to three example files each, so a misclassification
is visible in the census instead of quietly shrinking the corpus.

Everything else is in scope, including packages a strict validator would
reject. Word accepts those, so folio must too.

## Signatures and the baseline

A failure's signature is the invariant, the failure message with every per-file
particular erased (paths, identifiers, counts), and the innermost folio stack
frame. Two files with the same signature are the same defect.

`corpus/baseline.json` records how many files each known signature affects, and
may only shrink. The gate fails on a new signature, on a signature that gained
files, on a signature that lost files without the baseline being rewritten, and
on a baseline entry nothing reproduces any more. It is bound to the lock digest
of the corpus it was measured over, so repinning a source forces an explicit
refresh:

```sh
bun scripts/corpus-gate.ts write-baseline census.json
```

## Expected refusals

Not every failing file is a defect. Some are packages folio deliberately
declines, as typed errors with a stated reason: a part nested past the depth
bound, markup the resource preflight cannot scan safely. Those signatures live
in `corpus/expected-refusals.json`, each with the reason folio refuses it, and
the baseline is then a list of defects alone.

The list is hand-written, because a reason is a decision a person makes.
`write-baseline` refreshes the counts and touches nothing else, and those
counts ratchet the way the baseline does: a refusal that spreads to more files
is a finding, and one that stops reproducing must be removed. Promoting a
signature means adding its `signature` and `reason` to the file and rerunning
`write-baseline`.

Caveat: a `panic()` in tail position has no folio frame in the stack — the
engine eliminates the call — so those signatures carry `-` and are identified by
their message alone.

## Turning a failure into a synthetic seed

The corpus finds defects; it does not own regressions. A corpus file is
third-party content and cannot be committed, so the workflow is:

1. Minimise the file. Delta debugging over package parts, then over XML
   elements, keeping only what still fails with the same signature:

   ```sh
   bun run corpus:minimize apache-poi/test-data/document/55733.docx \
     --invariant style-set-rebuild
   ```

   Output goes to the cache (`minimized.docx`, `document.xml`, `summary.json`),
   never the repository.

2. Read `document.xml` and `summary.json` and name the construct: a reserved
   value, an element Word accepts that folio rejects, a part that had to be
   absent.

3. Write a **synthetic** `fast-check` property in the owning package that
   generates that construct, and a synthetic fixture if the property cannot
   express it. `packages/core/src/docx/styleNumberingReferences.property.test.ts`
   is the shape to copy.

4. Fix the defect, rerun the gate, and rewrite the baseline down.

## Licensing and the never-commit rule

**No corpus content is ever committed.** Files are fetched into
`~/.cache/folio-corpus` (override with `FOLIO_CORPUS_CACHE`), which must resolve
outside the repository or the tooling refuses to run. Minimised reproductions
are still third-party content and go to the cache too.

`corpus/sources.json` pins each source to a repository, a commit, a tree, the
`*.docx` sub-paths to take, and an SPDX licence that was read and confirmed.
Every source declares `redistribution: "cache-only"`; a source whose
`auditStatus` is not `reviewed` is refused at load. `corpus/sources.lock.json`
is committed and carries relative paths, SHA-256 digests and byte counts only —
no content — which is what makes a run reproducible and what CI keys its cache
on.

### Tiers

A tier records where a source runs. Nothing is redistributed from any tier.

`--tiers 1` is the default and the only selection pull-request CI runs.
`--tiers 1,2` runs locally and in the nightly workflow. The tier selection is
part of the digest a baseline is bound to, so a tier-2 run can never be compared
against the tier-1 baseline by accident, and adding a tier-2 source leaves the
tier-1 baseline valid.

**Tier 1.** Pull-request CI and the nightly run.

| Source                              | Licence             |
| ----------------------------------- | ------------------- |
| `apache/poi`                        | Apache-2.0          |
| `apache/tika`                       | Apache-2.0          |
| `sergey-tihon/Clippit`              | MIT                 |
| `VolodymyrBaydalka/docxjs`          | Apache-2.0          |
| `guigrpa/docx-templates`            | MIT                 |
| `ShayHill/docx2python`              | MIT                 |
| `plutext/docx4j`                    | Apache-2.0          |
| `open-xml-templating/docxtemplater` | MIT OR GPL-3.0-only |
| `dolanmiu/docx`                     | MIT                 |
| `mwilliamson/mammoth.js`            | BSD-2-Clause        |
| `nolze/msoffcrypto-tool`            | MIT                 |
| `nissl-lab/NPOI`                    | Apache-2.0          |
| `harshankur/officeParser`           | MIT                 |
| `OfficeDev/Open-Xml-PowerTools`     | MIT                 |
| `dotnet/Open-XML-SDK`               | MIT                 |
| `python-openxml/python-docx`        | MIT                 |

**Tier 2.** The nightly run only.

| Source                          | Licence          |
| ------------------------------- | ---------------- |
| `LibreOffice/core`              | MPL-2.0          |
| `jgm/pandoc`                    | GPL-2.0-or-later |
| `elapouya/python-docx-template` | LGPL-2.1-only    |

`LibreOffice/core` is by a wide margin the richest set of real-world quirk
reproductions that exists, and the most producer-diverse: it more than doubles
the corpus on its own.

**Tier 3, large crawl-derived, local sampling only.** Nothing qualifies. The
Apache Tika regression corpus (`corpora.tika.apache.org`, the documented
Common Crawl and govdocs1-derived set) no longer resolves: the DNS record is
gone, confirmed against a public resolver and over DNS-over-HTTPS with working
controls alongside. It is not a fetch worth retrying, and no other
crawl-derived set was scraped.

Not included: ONLYOFFICE (four `.docx` between its repositories),
`CollaboraOnline/online` (no `.docx`; the mirror that has them is a LibreOffice
copy and so redundant), `apache/openoffice` (no `.docx`; it predates OOXML test
corpora), and the Aspose and GroupDocs sample repositories.

## Adding a source

1. Read the SPDX identifier from the repository's own licence file rather than
   from GitHub's detected label, which disagreed with the file in four of the
   repositories surveyed for this corpus.
2. Add an entry to `corpus/sources.json`, sorted by id, with the commit and tree
   object IDs, `*.docx` sub-path patterns, a `tier` and a `tierReason` in your
   own words.
3. `bun run corpus:fetch`, then rerun the gate and
   `write-baseline`: new files change the lock digest, so the baseline must be
   re-measured.
