# Corpus fixtures — provenance & licensing

Most `.docx` files in this directory were **synthesised for this test suite** by
`packages/core/scripts/build-corpus-fixtures.ts`. They are hand-written OOXML
packages wrapped in `JSZip`; no third-party templates were copied.

Because the content was authored from scratch as a documentation artefact for
testing the parser/serializer, the fixtures are released under the same
**Apache-2.0** license as the rest of the package (see the repository
`LICENSE` file).

To regenerate those fixtures after editing the script:

```sh
bun packages/core/scripts/build-corpus-fixtures.ts
```

Four baseline fixtures are copied from the archived upstream
[test fixture directory](https://github.com/mhurhangee/docx-editor/tree/7ce79b8e6c93a2bd4befafaad43719e3d3ce6c5e/e2e/fixtures)
at commit `7ce79b8e6c93a2bd4befafaad43719e3d3ce6c5e`. They were generated from
hand-written OOXML by that repository's fixture generator and are covered by
its Apache-2.0 license.

| File                           | SHA-256                                                            |
| ------------------------------ | ------------------------------------------------------------------ |
| `upstream-empty.docx`          | `a580137f03c6f27a8396bdb3777e5ea3272ea11355a05b2c4d6a847fe9e4ef32` |
| `upstream-styled-content.docx` | `6830e894e1168c172e2933db4b7e6d7e679f38d53d939356e67703cffbe82e34` |
| `upstream-with-tables.docx`    | `8d2a6ab28dfaff8b8e5999eab26b95931dded7c01c0d0efc63f7a307ca1d1f3f` |
| `upstream-complex-styles.docx` | `cf529767f921698fc68ca68cda69894d886efeca8ab7c1389f7bb10ef8026338` |

## Inventory

| File                                  | What it exercises                                                                                        |
| ------------------------------------- | -------------------------------------------------------------------------------------------------------- |
| `block-sdt-richtext.docx`             | Block-level `<w:sdt>` wrapping a paragraph; alias, tag, lock                                             |
| `inline-sdt-dropdown.docx`            | Inline `<w:sdt>` with `<w:dropDownList>` (three `w:listItem`s)                                           |
| `inline-sdt-checkbox.docx`            | Inline `<w:sdt>` with `<w14:checkbox>` in the checked state                                              |
| `inline-sdt-mixed-rpr.docx`           | Nested inline SDTs whose content carries mixed `<w:rPr>` (bold/italic)                                   |
| `alt-prefix-sdt.docx`                 | All `w:` elements rebound to an `x:` prefix; date SDT                                                    |
| `nested-block-sdt.docx`               | Outer block SDT wrapping an inner block SDT wrapping a paragraph                                         |
| `sdt-rpr-placeholder.docx`            | `<w:sdtPr>` containing a `<w:rPr>` (color + bold) and an alias/tag                                       |
| `empty-sdt-content.docx`              | Inline SDT with an empty `<w:sdtContent/>`; tests round-trip idempotence                                 |
| `authored-empty-paragraph.docx`       | Block SDT whose content is an authored `<w:p/>` empty paragraph                                          |
| `date-fractional-seconds.docx`        | `w:fullDate="2026-06-02T00:00:00.000Z"` with millisecond precision                                       |
| `dropdown-empty-value.docx`           | Dropdown whose first `w:listItem` has `w:value=""`                                                       |
| `lock-sdt-locked.docx`                | `<w:lock w:val="sdtLocked"/>`                                                                            |
| `lock-content-locked.docx`            | `<w:lock w:val="contentLocked"/>`                                                                        |
| `lock-sdt-content-locked.docx`        | `<w:lock w:val="sdtContentLocked"/>`                                                                     |
| `repeating-section.docx`              | `<w15:repeatingSection/>` marker inside `<w:sdtPr>`                                                      |
| `checkbox-val-true.docx`              | `<w14:checked w14:val="true"/>` (boolean form of ST_OnOff)                                               |
| `checkbox-val-false.docx`             | `<w14:checked w14:val="false"/>` (boolean form of ST_OnOff)                                              |
| `placeholder-docpart.docx`            | `<w:placeholder><w:docPart w:val="DefaultText"/></w:placeholder>`                                        |
| `datahash-sdt.docx`                   | `<w16sdtdh:dataHash w16sdtdh:val="..."/>` marker inside `<w:sdtPr>`                                      |
| `extraspec-bookmark-sibling.docx`     | `<w:bookmarkStart/>` / `<w:bookmarkEnd/>` as direct siblings of `<w:sdtContent>` under `<w:sdt>`         |
| `extraspec-commentrange-sibling.docx` | `<w:commentRangeStart/>` / `<w:commentRangeEnd/>` as direct siblings of `<w:sdtContent>` under `<w:sdt>` |
| `opendope-encoded-ampersand-tag.docx` | OpenDoPE binding tag containing an encoded `&amp;`; round-trip must re-escape exactly once               |
| `sdt-without-sdtpr.docx`              | `<w:sdt>` with no `<w:sdtPr>` child; default-property tolerance                                          |
| `sdt-self-closing.docx`               | Self-closing `<w:sdt/>` with no `<w:sdtPr>` and no content                                               |
| `placeholder-without-docpart.docx`    | `<w:placeholder/>` with no `<w:docPart>` child; spec-non-conformant tolerance                            |
| `upstream-empty.docx`                 | Empty-document mutation and package-preservation baseline                                                |
| `upstream-styled-content.docx`        | Mixed run formatting, font sizes, and paragraph alignment                                                |
| `upstream-with-tables.docx`           | Body text and a three-by-three table                                                                     |
| `upstream-complex-styles.docx`        | Paragraph styles, font families, colors, and highlighting                                                |

All fixtures are under 30 KB.

## Source citations (extra-spec / tolerance fixtures)

| Fixture                               | Source                                                                                                                                                            |
| ------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `extraspec-bookmark-sibling.docx`     | [MS-OE376 §2.5.2.30 — Word's additional sdt children](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/1aee8ae1-fe3a-4ca4-96f4-e416cf43e461) |
| `extraspec-commentrange-sibling.docx` | [MS-OE376 §2.5.2.30 — Word's additional sdt children](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/1aee8ae1-fe3a-4ca4-96f4-e416cf43e461) |
| `opendope-encoded-ampersand-tag.docx` | [OpenDoPE conventions v2.3](https://www.opendope.org/opendope_conventions_v2.3.html)                                                                              |
| `sdt-without-sdtpr.docx`              | Word 2010–2024 lenient parse behaviour (no formal spec)                                                                                                           |
| `sdt-self-closing.docx`               | Word 2010–2024 lenient parse behaviour (no formal spec)                                                                                                           |
| `placeholder-without-docpart.docx`    | [ECMA-376 §17.5.2.27 (c-rex mirror)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_Structured_topic_ID0ERWJS.html) + observed Word tolerance             |

## Modelled-coverage skips

A few of the proposed edge cases would have required model surfaces that
folio's `SdtProperties` does not currently project. The fixtures are still
included so the upgrade path lands on green; the assertions are reduced to
what the model can express today, and the gap is recorded here.

- **BlockSdt property preservation** — block `<w:sdt>` wrappers (the
  `block-sdt-richtext.docx`, `nested-block-sdt.docx`, `repeating-section.docx`
  and `authored-empty-paragraph.docx` fixtures) are unwrapped on parse, so the
  outer alias/tag/lock are not assertable. Tests cover content fidelity only.
- **`rawPropertiesXml` for unknown sdtPr children** — there is no escape hatch
  on `SdtProperties` to carry through unknown markers verbatim, so the
  `sdt-rpr-placeholder.docx`, `repeating-section.docx`, and `datahash-sdt.docx`
  fixtures assert "sdtType is not mis-classified" and "alias/tag round-trip"
  instead of asserting the raw marker survives.
- **Dropdown last selected value (`w:val` on `w:sdtPr` / sdt content
  binding)** — `SdtProperties` exposes `listItems` but not a separate
  "currently selected" field, so the duplicate-displayText scenario from the
  task brief was not added as a dedicated fixture; it would have nothing to
  assert beyond what `inline-sdt-dropdown.docx` already covers.
- **`dateValueISO`** — folio stores the raw `w:fullDate` ISO string in
  `dateFormat`. The `date-fractional-seconds.docx` fixture asserts the raw
  string survives there rather than on a dedicated field.

## Out of scope here

`w15:repeatingSection` is covered as a regression-only fixture; modelling it
as its own `sdtType` is tracked separately.
