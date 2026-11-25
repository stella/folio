---
"@stll/folio-core": patch
---

A canonical `FolioDocxReviewer` resolves staged suggestions with the other changes: a final view (`readReviewedStory`, `toResolvedDocument("accept")`) reads them accepted, `resolveReviewedStory` resolves a story's suggestions with its other changes in one commit, and a comparison projection accepts them with the rest. A suggestion reaching several stories still resolves only whole, through the suggestion methods.
