---
"@stll/folio-core": patch
---

A jump to an unpainted position paints its page and neighbours before scrolling and moves instantly when the target is more than two viewport heights away, and virtualization reads painted-page retention from an intersection observer, rooted at the editor scroll root, instead of measuring every painted page.
