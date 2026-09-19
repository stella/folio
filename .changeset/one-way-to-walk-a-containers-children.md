---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Add the ordered verbatim sink and the shared child dispatcher, and put `w:comment` bodies on them.

`PreservedMarkup` holds a container's unmodelled children with their position relative to its modelled ones, plus an ordered attribute remainder, so the serializer puts them back between the same siblings rather than at the end. `dispatchChildren` walks a container with a handler map the compiler makes total over the children the schema declares for it, and routes anything undeclared — a foreign namespace, an `mc:` construct, an element a later OOXML revision adds — to the sink by default.

A comment body may hold everything a document body can. folio modelled only `w:p`, so a table, an equation, a content control, a bookmark or a range marker in a reviewer's comment disappeared on save; `Comment.preserved` now keeps them.
