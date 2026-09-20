---
"@stll/folio-core": patch
---

Type the paragraph attr patches a style or a list command produces. `paragraphAttrsFromResolvedStyle`, `listAttrsFromResolvedStyle`, `listAttrsFromNumbering` and `listLevelAttrPatch` returned `Record<string, unknown>`, so every attr they wrote was unchecked: a `numPr` taken straight off a model object assigned as readily as one minted by `paragraphNumberingAttr`, which is the crossing the `ParagraphNumberingAttr` brand exists to refuse, and a misspelled attr key was a silent no-op.

They now return `ParagraphAttrsPatch`, derived from `ParagraphAttrs` rather than hand-listed: every key carries its own attr type or the absent state the node spec's default stores, and an attr added to the spec is writable without a second edit. A compile-time proof beside the producers pins that the model's union, a hand-assembled record, and an undeclared key are all rejected.
