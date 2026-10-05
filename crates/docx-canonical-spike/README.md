# Canonical document experiment

Standalone, unpublished crate; no production package imports it. The TypeScript
canonical engine is the document, inverse, refusal, touched and revision oracle.
This is an incomplete experiment, not a replacement engine.

`canonical-spike` accepts JSON-lines `{document, ops}` requests. Browser
`apply(documentJson, operationsJson)` uses wasm-bindgen and the same core.
Unsupported dimensions have their own result, separate from semantic refusals.
`replaceText` is an editor intent: tests use the existing TypeScript compiler's
primitive batch rather than introducing an operation kind.

The implemented subset covers plain text/slices, run and paragraph patches,
split/join, paragraph block edits, simple table/row edits and their restoration
operations, plus tracked plain-text insertion/deletion. Complex inline records,
identity freshening and full own-undefined application remain unsupported.
Comment semantics come from schema-10 commit
`617c0f4703a6cf43adb7ecda34366c29b93cc60f`; main-story plain anchors and ordered
relationship state are ported. Main's schema-9 operation oracle has no comments.
The ordinary JSON boundary cannot represent JavaScript Maps.

S2 uses a **test-only** tagged codec through the native runner's `harness` request.
It preserves Maps, Dates, bytes, array holes and owned undefined fields; capture
symbols and shared references are counted separately. Roundtrip tests use existing
model generators. Any unported sidecar movement is reported as unsupported,
never silently projected away. Product wire schemas remain unchanged.

S3 uses one Playwright Chromium for TypeScript and WASM, plus self-timed native
Rust. All arms run on one machine with rotating interleaved repetitions; input
parsing/output JSON and the browser WASM boundary are inside operation timings.
Native startup/IPC are excluded. Each fixture's forward, undo and redo are
compared before timing. JSON model load and replay-save are measured separately;
ZIP/XML parsing, DOCX serialization and pagination are not measured by those rows.
Native memory is OS process peak RSS; browser memory is observed JS heap and
WASM linear capacity, not browser process peak RSS. Unsupported operations have
no latency claim. Raw/gzip WASM sizes are recorded.

Prepare with `bench/prepare.ts`, then `bench/validate.ts`; readiness is written
only after fixture verification. `bench/run.ts` accepts explicit `validate`,
`ci` or `quiet-window` mode and checks source/artifact hashes. Local cargo is
serialized, nice 10 and gated below load 5. Quiet measurements require one-minute
load below 3 at start and before every sample, with no retry. CI records runner
model/load for the relative cross-check. No WASI, workers or shared memory.
