---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

Add `FolioDocxReviewer.save()`, which reports whether the package was written by the selective patch or a full repack (with the reason) and, with `repack: "refuse"`, declines a full repack instead of rewriting every part. Comments a stamped batch creates take the stamp's date, `replyTo` accepts a `date`, and `createReviewerBridge` accepts a `revisionStamp` that it applies to its batches and replies.
