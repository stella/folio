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
default 300000), `--shard k/n` and `--out FILE`. CI shards four ways and merges
the censuses before the ratchet, because each shard sees only a subset:

```sh
bun scripts/corpus-gate.ts run --shard 1/4 --out census-1.json
bun scripts/corpus-gate.ts check census-1.json census-2.json census-3.json census-4.json
```

The gate is nightly (`.github/workflows/nightly-corpus-gate.yml`) and on
`workflow_dispatch`. It is not part of PR CI: it downloads a hundred megabytes
and takes minutes.

## The invariants

| Invariant           | What must hold                                                                                                                                                          |
| ------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `parse`             | A package a reader can open parses.                                                                                                                                     |
| `fixed-point`       | parse → ProseMirror → repack → parse preserves the visible text and the text-block count. Same equality as `packages/core/src/docx/__tests__/corpusFixedPoint.test.ts`. |
| `repack-validates`  | The package folio just wrote satisfies folio's own package validator.                                                                                                   |
| `style-set-rebuild` | Extracting a document's style set and building a package from it does not panic.                                                                                        |
| `completes`         | The file produced a verdict: no hang, no process abort.                                                                                                                 |

Each file runs in a child process with a deadline, so a hang or an abort is a
recorded finding rather than a lost run.

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

| Source                              | Licence             |
| ----------------------------------- | ------------------- |
| `apache/poi`                        | Apache-2.0          |
| `apache/tika`                       | Apache-2.0          |
| `sergey-tihon/Clippit`              | MIT                 |
| `VolodymyrBaydalka/docxjs`          | Apache-2.0          |
| `ShayHill/docx2python`              | MIT                 |
| `plutext/docx4j`                    | Apache-2.0          |
| `open-xml-templating/docxtemplater` | MIT OR GPL-3.0-only |
| `mwilliamson/mammoth.js`            | BSD-2-Clause        |
| `nolze/msoffcrypto-tool`            | MIT                 |
| `OfficeDev/Open-Xml-PowerTools`     | MIT                 |
| `dotnet/Open-XML-SDK`               | MIT                 |
| `python-openxml/python-docx`        | MIT                 |

LibreOffice's `sw/qa` test documents were considered and dropped: the
repository's `COPYING` is GPL-3.0 and the `.docx` files there are largely bug
attachments carrying no licence grant of their own, so the licence could not be
confirmed for the data.

## Adding a source

1. Confirm the licence by reading the repository's own licence file, and check
   whether it covers the test data specifically.
2. Add an entry to `corpus/sources.json`, sorted by id, with the commit and tree
   object IDs and `*.docx` sub-path patterns.
3. `bun run corpus:fetch`, then rerun the gate and
   `write-baseline`: new files change the lock digest, so the baseline must be
   re-measured.
