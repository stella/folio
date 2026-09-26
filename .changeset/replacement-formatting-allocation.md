---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

A text replacement now formats its new text by one rule in every mode, so applying it directly and accepting it as a tracked change leave the same result. Replacing `Supplier agrees` with `Provider performs` where only `Supplier` was linked or bold left `perform` inside the link and the final `s` outside it when applied directly, and the whole phrase inside it once the tracked change was accepted: the direct plan kept the letter the two words happen to share. A change across several words now keeps only whole words it shares, a change within one word still keeps the letters it shares, and the tracked redline writes each inserted character with the formatting the direct edit gives it; a field the direct edit keeps is no longer rewritten as plain text by the redline. A replacement that rewrites text of more than one formatting, link or comment reports it in the result's `normalizations` with the new code `uniformReplacementFormatting`, and `suggest_changes` explains it.
