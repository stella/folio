---
"@stll/folio-core": patch
---

Bound the SmartArt preview a parse generates. Only the VML preview was charged against a package-wide budget, and it was recognized by constants in a different file from the ones its producer wrote, so a rename would have silently stopped the charge. Both previews now come from one table that the producer builds from and the budget matches against, and each kind carries its own per-package allowance. The SmartArt cap is set above the public corpus maximum (51.3 MB of preview data URL, from a package under a megabyte), so no corpus file loses a preview that it keeps today.
