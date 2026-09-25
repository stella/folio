---
"@stll/folio-cli": minor
---

Add the `folio` write commands: `suggest`, `comment`, `reply`, `resolve`, `accept`, `reject`, and `compare -o`. Each change is one transaction: `--in-place` (with a backup) or `-o`, a version precondition re-checked under a cooperative write lease (`locked`, exit 10, unless `--force`), a staged and validated package renamed into place, selective saves unless `--allow-repack`, author from `--author`/`FOLIO_AUTHOR`/git with no default, one `--date`-overridable timestamp, a `.folio/journal.jsonl` line, `--tx-id` replay, and recovery of a stage a crash left behind.
