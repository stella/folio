# Legal parity corpus: provenance and licensing

Files in this directory are public legal-document fixtures retained for
explicit, on-demand interoperability checks. They are not discovered by the
default parity, visual, differential, or unit-test suites.

## UK Model Services Contract schedules v2.2A

| Field       | Value                                                                                                                                                                    |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| File        | `uk-msc-schedules-v2.2a.docx`                                                                                                                                            |
| Publisher   | UK Cabinet Office                                                                                                                                                        |
| Publication | [Model Services Contract: schedules (England and Wales)](https://www.gov.uk/government/publications/the-model-services-contract-schedules-england-wales)                 |
| Source      | [Consolidated schedules v2.2A (2025)](https://assets.publishing.service.gov.uk/media/68af2675960e2d135b4c8eb6/MSC_-_Consolidated_Schedules_-_E_W_-_Word_v2.2A_2025.docx) |
| Retrieved   | 2026-08-29                                                                                                                                                               |
| SHA-256     | `49f6edd53f0d7987bb21c403c750004c5515fcafb5eeed093acaab559ba3fb17`                                                                                                       |
| License     | [Open Government Licence v3.0](http://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/)                                                                |

Attribution required by the document:

> Contains public sector information licensed under the Open Government Licence v3.0

The fixture is retained for document-interoperability testing. Its inclusion
does not imply endorsement by the Cabinet Office or the UK Government.

This is a deliberately demanding legal-document sample: approximately 1.1 MB
compressed, with 127 tables, 312 numbering instances, 23 explicit numbering
restarts, 3,781 bookmark pairs, thousands of fields, 44 sections, document
grids, drawings, and tracked changes. Some bookmarks span paragraph and table
boundaries, which makes the document useful for structural round-trip work.

Verify the pinned source before running a comparison:

```sh
shasum -a 256 parity/corpus/legal/uk-msc-schedules-v2.2a.docx
```

Run parity and differential analysis explicitly:

```sh
bun parity/cli.ts parity/corpus/legal/uk-msc-schedules-v2.2a.docx --reference word

bun run parity:line-endpoints capture \
  parity/corpus/legal/uk-msc-schedules-v2.2a.docx \
  --output /tmp/uk-msc.word-lines.json

bun packages/core/scripts/differential/diff.ts \
  parity/corpus/legal/uk-msc-schedules-v2.2a.docx
```

The complete document can require substantial time and memory to lay out. Keep
generated PDFs, page images, and line-endpoint manifests out of git; they are
reproducible local artifacts and may contain most of the document text.
