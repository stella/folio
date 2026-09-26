---
"@stll/folio-core": patch
---

Place a word diff's lone insertion or deletion where it reads best. Among the positions an insertion or deletion can take without changing either string, the word diff now picks the one whose edges fall at the string's edge, a line break, the gap after a sentence mark or a space, and never inside an untouched word. An appended sentence is marked with its own full stop rather than the preceding sentence's, and an insertion and the deletion that undoes it mark the same text.
