---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

`ExhaustiveFields<Source, Classified>`, the compile-time gate that turns a model field added without a decision into a build failure, now has one owner and is exported from `@stll/docx-core/model`. The paragraph, text, and border serializers each carried a verbatim copy.
