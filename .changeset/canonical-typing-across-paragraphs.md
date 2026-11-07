---
"@stll/folio-core": patch
---

In the experimental canonical session, typing over a selection that spans paragraphs, over the whole document after select-all, or over selected cells, now reaches the session. Text across paragraphs and the whole document is replaced as typing and tracked in suggesting mode. Typing over selected cells still leaves them unchanged, now with the session's own reason. Before, every such key was refused with nothing typed.
