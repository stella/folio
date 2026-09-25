---
"@stll/folio-cli": minor
---

Add the `folio` write commands: `suggest`, `comment`, `reply`, `resolve`, `accept`, `reject`, and `compare -o`. Each change is one transaction: `--in-place` or `-o` onto a plain `.docx` (replacing an existing file needs `--overwrite` and its version), a required `--expect-version` (waived only by `--no-expect-version`) re-checked under a fenced cooperative write lease (`locked`, exit 10, unless `--force`), a backup of whatever is replaced, a staged and validated package renamed into place, selective saves unless `--allow-repack`, author from `--author`/`FOLIO_AUTHOR`/git with no default, one `--date`-overridable timestamp, a `.folio/journal.jsonl` line, `--tx-id` replay, and checked recovery of a stage a crash left behind. Sidecar files are never used through symlinks or hard links.
