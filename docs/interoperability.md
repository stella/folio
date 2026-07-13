# Interoperability references

Folio is standards-first and implementation-tested. OOXML specifications and
published implementation notes define the intended document semantics;
independent libraries and renderers expose interoperability gaps that a single
implementation cannot.

No external application is a conformance oracle. Matching a renderer proves
agreement with that version, font environment, and export path. It does not
prove that either implementation satisfies every applicable OOXML requirement.

## Evidence layers

| Layer                         | References                                                                                                                                     | What it tells us                                                        |
| ----------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------- |
| Normative format              | [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/), ISO/IEC 29500                                       | Package, schema, and document-semantics requirements                    |
| Published extensions          | [Microsoft implementation notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/bd9e8289-844a-42e2-9809-66c7005bd9e2) | Documented producer-specific behavior and deviations                    |
| Structural differential tests | [python-docx](https://github.com/python-openxml/python-docx), [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)                           | Whether independent parsers project the same document structure         |
| Rendering comparisons         | [LibreOffice](https://www.libreoffice.org/), optionally Microsoft Word                                                                         | Pagination, line breaking, geometry, and font-environment differences   |
| Round-trip and behavior tests | Folio's unit, property, differential, and Playwright suites                                                                                    | Whether editing preserves document invariants and user-visible behavior |

## Projects evaluated

The licensing notes below concern the narrow testing architecture described
here: Folio invokes a separately installed executable or service and consumes
its output. Copying source, linking a library into a distributed package,
modifying an external project, or redistributing its binaries requires a
separate compliance review.

| Project                                                                   | Licence                                                              | Technically useful for                                                        | Decision                                                                                                                                                                                                                |
| ------------------------------------------------------------------------- | -------------------------------------------------------------------- | ----------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [LibreOffice](https://www.libreoffice.org/licenses/)                      | MPL-2.0, with components under other compatible open-source licences | Independent DOCX-to-PDF layout rendering                                      | **Integrated and default.** Invoke the user-installed `soffice` executable; do not bundle it with Folio. LibreOffice explicitly permits business use.                                                                   |
| [ONLYOFFICE Docs Community](https://github.com/ONLYOFFICE/DocumentServer) | AGPL-3.0                                                             | Independent DOCX-to-PDF conversion through its self-hosted Conversion Service | **Usable with deployment constraints; adapter pending.** Keep it an external, pinned service. Do not bundle or modify it without satisfying AGPL source, licence, and network-use obligations for the deployed version. |
| Microsoft Word                                                            | Proprietary, agreement-specific                                      | Widely deployed DOCX behavior and implementation-specific compatibility       | **Optional adapter.** Never describe it as ground truth. The operator must confirm that automated compatibility use is allowed by the licence covering the installed copy.                                              |
| [Apache OpenOffice](https://www.openoffice.org/license.html)              | Apache-2.0                                                           | Headless DOCX-to-PDF rendering                                                | **Permitted but lower priority.** Its shared ancestry with LibreOffice provides less independent evidence than ONLYOFFICE.                                                                                              |
| [python-docx](https://github.com/python-openxml/python-docx)              | MIT                                                                  | Structural parsing and round-trip projections                                 | **Integrated in differential CI.** It does not paginate or provide a layout oracle.                                                                                                                                     |
| [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)                    | MIT                                                                  | Structural parsing, package/schema validation, and projections                | **Integrated in differential CI.** It does not render pages.                                                                                                                                                            |
| [docx4j](https://github.com/plutext/docx4j)                               | Apache-2.0                                                           | Structural checks, transformations, and optional PDF export                   | **Permitted, not currently needed.** Its Java/PDF pipeline could provide another projection, but it should not be presented as a full-fidelity office renderer.                                                         |
| [Apache POI XWPF](https://poi.apache.org/components/document/index.html)  | Apache-2.0                                                           | Text extraction and partial DOCX structure checks                             | **Permitted, weak fit for layout.** Apache describes XWPF as moderately functional and support varies by feature.                                                                                                       |
| [Mammoth](https://github.com/mwilliamson/mammoth.js)                      | BSD-2-Clause                                                         | Semantic DOCX-to-HTML conversion                                              | **Permitted, unsuitable for layout parity.** Mammoth intentionally ignores many visual details.                                                                                                                         |
| [Pandoc](https://pandoc.org/)                                             | GPL-2.0-or-later                                                     | Semantic format conversion through an external executable                     | **Permitted as a separate tool, unsuitable for layout parity.** PDF output depends on another typesetting backend rather than a DOCX page-layout engine.                                                                |

Using an unmodified copyleft executable as a separate local process does not by
itself make Folio a derivative of that executable. Distribution, linking,
modification, or network deployment can change the obligations; pin the exact
release and re-check its licence before adding it to CI or a hosted service.

## Harness rules

- LibreOffice is the default rendering reference. Proprietary references are
  opt-in and identified by name and version in every report.
- Comparisons use synthetic, public, or explicitly authorized fixtures. Do not
  upload customer or legal documents to third-party services.
- Record renderer version, conversion path, fonts, and document hash. A result
  without provenance is not reproducible evidence.
- Treat font substitutions and platform-specific metrics as environment data,
  not automatically as Folio defects.
- Resolve disagreements from specifications and documented extensions first.
  Multi-renderer agreement is supporting evidence, not a vote that changes the
  standard.
- Use project names factually for interoperability identification; do not use
  logos or imply endorsement.

The local rendering harness and its report format are documented in
[`parity/README.md`](../parity/README.md).
