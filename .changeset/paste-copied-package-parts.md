---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

A chart, an embedded object or kept block markup copied from one document now pastes into another, directly or as a suggestion: the parts it names (the chart with its workbook and styles, the embedding, its preview picture) are copied in under fresh names and relationship ids and are written on save. An embedded object drawn by a picture now pastes as the object rather than as the picture alone. The package model gains `copiedParts`, the package resources operation carries them, and `ZipArchive.readDirectory` reads an archive's directory at once.
