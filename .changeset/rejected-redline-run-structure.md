---
"@stll/folio-core": patch
---

Resolving tracked changes rejoins the pieces a revision cut out of one run: after rejecting (or accepting) a minimal redline, adjacent runs that write the same `w:rPr` and run attributes are one `w:r` again instead of several, and a rejected `w:rPrChange` restores the run's original properties. Runs a saved redline splits around a `w:ins`, `w:del` or `w:rPrChange` read back as pieces of one run, and a `w:rPr` holding only a `w:rPrChange` is no longer read as an authored empty `w:rPr`.
