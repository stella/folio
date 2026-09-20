---
"@stll/folio-core": patch
---

Stop the block alignment from reporting a crossing as a rewrite.

Anchoring is monotone, so two paragraphs that swap places produce two exact-text correspondences that cross and only one can be anchored. The leftover ends were then handed to the positional fallback, which pairs by offset and knows nothing of what it pairs: a heading and an unrelated paragraph came back as a replacement, taking with it the removal and the arrival the move pass exists to pair.

An exact correspondence the monotone pass had to drop is still evidence, and it is evidence against fusing that block with something else. The fallback now declines a pair when each side's text stands unpaired on the other side, so the relocation is reported as one.

Leaving the two ends unpaired reaches a shape the terminal-carrier repair did not recognise: the story's last paragraph removed and another written where it stood. The rotation that carries an inserted mark at a container's end turns on the paragraph the run was appended after, so a paragraph that is itself deleted leaves the addition nowhere to go, and the comparison refused its own round trip. That repair now covers it.
