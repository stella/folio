---
"@stll/folio-core": minor
---

Add `ensureParaIds` to the server entry: a headless, in-place `w14:paraId` backfill for `.docx` buffers (document body, headers, footers, footnotes, endnotes; table-cell and text-box paragraphs included). Deterministic, idempotent, and namespace-aware, so hosts can normalize documents once at ingest and block anchors never fall back to positional `seq-` ids.
