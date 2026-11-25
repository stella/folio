---
"@stll/folio-core": minor
---

A template plan can write a block's automatic label as text (`labelAsText` on a source item and on a nested placement), for a label no list of the template prints. The label goes in front of the block's text with the tab or space its list puts after it, and the paragraph it leads joins no list. Only a label that reads the same wherever the file is opened is written (numbers, letters, roman numerals); a bullet or a number a language spells is refused with `label-not-writable`. The report names the label written for each block (`labelAsText`) and counts them, and `labelLost` counts only labels that are neither rendered nor written.
