---
"@stll/folio-core": minor
---

Carry a run's own record through the editor on one mark.

An authored `w:r` had no record on the editor side, so a round trip rebuilt it from formatting alone: its attribute remainder — the `w:rsid*` family Word writes on nearly every run — and the `w:rPr` children no reader took a value from were both dropped, and two adjacent runs the formatting could not tell apart were written back as one.

`runIdentity` is that record. It is one mark rather than three, because all three payloads are facts about one element and the save leg asks that element a single question: where does this run begin and end. It is in the key adjacent leaves are grouped by, so a change of identity is a run boundary; it is minted only when a run holds a page break, a remainder or a sink, so a fully modelled document pays nothing. It replaces `pageBreakRunOwner`, whose job was the same one level narrower.

Two rules follow from what an rsid means. It names an editing session registered in the package's own `settings.xml`, which folio neither writes nor merges, so folio writes no rsid it did not read: text typed inside an authored run becomes a run of its own that states none, and the paragraph's `w:rsidRDefault` answers for it. And a pasted span carries the id alone, so a copy keys differently from its source and becomes its own run. Splitting a run, by contrast, manufactures nothing, so both halves keep what the run was authored with.

Collaboration snapshots move to attr schema 4. The rename is a value rewrite rather than an additive change, because mark attrs live in the shared text's delta under the mark's own name and an unknown name costs the text beneath it, not just the mark. The step rewrites the delta attribute before anything reads the fragment, and the marker turns an older build meeting a newer snapshot into a refusal. Payloads are not backfilled: a snapshot has no access to the package it was seeded from, so an existing room keeps today's behaviour until it is reseeded.
