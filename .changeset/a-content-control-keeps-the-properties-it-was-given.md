---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Read `w:sdtPr` through the shared child dispatcher, so a content control keeps every property its author wrote. `SdtProperties.preserved` holds the children folio does not model at their `CT_SdtPr` ordinal, and one writer serialises block, inline, row and cell controls from the model rather than replaying the source's bytes. `SdtProperties.rawPropertiesXml` is gone; `SdtProperties.lock` no longer reports `unlocked` for a value the reader refuses.
