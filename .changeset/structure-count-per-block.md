---
"@stll/folio-core": patch
---

Typing in a long document no longer recounts the whole document's structure on every keystroke. The change tracker keeps each block's structure for the block, so a keystroke recounts only the block it changed (about 3.5 ms down to 0.1 ms per transaction on a 100-article contract).
