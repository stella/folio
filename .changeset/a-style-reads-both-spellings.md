---
"@stll/folio-core": patch
---

Read a style's indent and cell margins in either spelling of the name.

ECMA-376 Part 1 spells a horizontal edge `start`/`end` where Part 4 spells it `left`/`right`. The document readers took both, and the style reader — which has its own copies of the same parsers — took only the Transitional one, so a style, a document default or a conditional table region written by a Strict producer lost its indent and its cell margins on the way in. Both now read through the generated rename table, which also covers `w:tblStylePr` and `w:docDefaults`, and a property test over every slot the table names holds every reading site to it.
