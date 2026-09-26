---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

An operation offset or range end that falls inside a character is refused instead of cutting it. A `splitBlock`, `replaceRange`, `formatRange` or `commentOnRange` offset between the two UTF-16 halves of an emoji (or any character outside the Basic Multilingual Plane) left a lone surrogate the save cannot write, so the character disappeared even for an operation that only formats or comments. Such an operation is now skipped before anything is applied with the new skip reason and issue code `splitsCharacter` (`retryable: true`, `recovery: "changeTarget"`) and a `message` naming the character and the offsets on either side of it. Operations that change text or break a paragraph (`replaceRange`, `replaceInBlock` matches, `splitBlock`) are refused inside a grapheme cluster as well: between a letter and its combining marks, inside a joined emoji sequence or a flag. `createFolioAITextRangeHandle` returns `null` for a range that cuts a surrogate pair.
