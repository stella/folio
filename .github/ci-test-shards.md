# CI test scheduling

`ci.yml` runs lint/source guards, typecheck, build/API checks, published-package
validation and four test shards in parallel. The same path scopes and fast/full
suite depths apply. `ci-result` requires each selected job to succeed, including
all four shards. Failed merge-group shards retain separate `fuzz-log-*` artifacts
for the existing replay workflow.

The shard runner discovers tests from every `packages/*` workspace's test command,
then adds the root `scripts` and `benchmarks/compare` suites. It preserves workspace
working directories, Bun configuration and explicit Vue preloads. Unsupported
workspace commands fail instead of silently dropping checks. The focused inventory
is the original fast-suite file list; it runs on every code-selected event.

`scripts/ci-test-timings.json` records file durations above one second from successful
merge-group run [36987441670](https://github.com/stella/folio/actions/runs/36987441670).
Durations are milliseconds between file-group headers in each test process. Files
without a measurement receive a 100 ms scheduling estimate and still run. Scheduling
sorts by descending duration and path, then places each file in the least-loaded
shard (ties choose the first shard). Discovery and scheduling are guarded by
`scripts/ci-plan.test.ts`; aggregate failure handling by `scripts/ci-result.test.mjs`.

## Baseline and estimate

Three recent successful merge-group runs were measured, reading each job list once:

| Run         | Full suite | Typecheck | Build | Reproducibility |  API | Dist validation |
| ----------- | ---------: | --------: | ----: | --------------: | ---: | --------------: |
| 36989762795 |     1149 s |      56 s |  68 s |           134 s | 42 s |           357 s |
| 36987441670 |      945 s |      42 s |  52 s |           104 s | 32 s |           276 s |
| 36987440048 |     1138 s |      58 s |  67 s |           136 s | 41 s |           360 s |

The combined job took about 28–34 minutes in these runs. Four balanced test shards
estimate around four minutes each before overhead. Moving dist validation and
build/API checks off that path should put the split jobs below the existing
interaction job, targeting a critical path near seven minutes. This is an estimate;
full-depth elapsed time requires a merge-group run. Pull requests retain fast depth.

The split jobs share a Bun download-cache key. Frozen installs still run in each
isolated runner; measured baseline installs took 3–5 seconds. No dependency tree is
transferred between runners.
