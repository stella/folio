---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

A canonical session reports pending changes from its own history: `hasPendingChanges` is true once any change is committed in any story after the document opened or was last saved, and false after a save. `FolioEditor` gains `markCanonicalSaved` and `hasCanonicalPendingChanges`.
