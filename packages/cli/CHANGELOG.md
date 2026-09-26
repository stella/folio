# @stll/folio-cli

## 0.2.0

### Minor Changes

- [#1072](https://github.com/stella/folio/pull/1072) [`ff14d26`](https://github.com/stella/folio/commit/ff14d260555f483090a43a3026416567b5c2c985) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `folio save`, which commits a package a live editor serialized with the same lease, backup, and journal as a tool call. A write that finds an editor holding the lease with unsaved edits now asks it to save and release first, then applies on the saved version instead of failing with `locked`.

### Patch Changes

- [#1070](https://github.com/stella/folio/pull/1070) [`d875deb`](https://github.com/stella/folio/commit/d875deb4d89361501f696c7deff124167d62a5c0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Scale `folio serve`'s preview page down to fit a browser window narrower than the document, instead of clipping the left margin and scrolling horizontally. Fix the herdr plugin's `npx` fallback so it no longer fails when the plugin runs from inside a folio checkout, and resolve a relative clicked `.docx` path against the pane it was clicked in.
- Updated dependencies [[`169fdad`](https://github.com/stella/folio/commit/169fdad184f0242a7145ce7373adfcdb8733c351), [`169fdad`](https://github.com/stella/folio/commit/169fdad184f0242a7145ce7373adfcdb8733c351), [`169fdad`](https://github.com/stella/folio/commit/169fdad184f0242a7145ce7373adfcdb8733c351), [`1fe6f3b`](https://github.com/stella/folio/commit/1fe6f3bfe4f2afc22cbed6bc3fe34ed61128e9a3), [`7d622f9`](https://github.com/stella/folio/commit/7d622f944365c283b0595d0a65b9e5d974af974c), [`169fdad`](https://github.com/stella/folio/commit/169fdad184f0242a7145ce7373adfcdb8733c351)]:
  - @stll/folio-core@0.51.0
  - @stll/folio-agents@0.14.1

## 0.1.1

### Patch Changes

- [#1068](https://github.com/stella/folio/pull/1068) [`a16deb9`](https://github.com/stella/folio/commit/a16deb9c4137c348d022e467c0e0f1d67323ae81) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The VS Code extension shows the stella logo in the Marketplace and the Extensions view.

## 0.1.0

### Minor Changes

- [#1058](https://github.com/stella/folio/pull/1058) [`3bb6ac0`](https://github.com/stella/folio/commit/3bb6ac008d9c0b0092e6f7a39748b1ce1d3fb5cd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `folio mcp`, a stdio MCP server over the same tool registry: each tool takes its folio-agents arguments plus a file envelope (`path`, `fileVersion` required on changes, `destination`, `overwrite`, `txId`, `allowRepack`, `mode`), every path must resolve inside an allowed `--root`, reads are bounded by block count, match count, and response size with version-bound cursors, and the operation schema and server rules are served as resources.

- [#1056](https://github.com/stella/folio/pull/1056) [`c9776b0`](https://github.com/stella/folio/commit/c9776b0694d2193d307999b788e1b72fb3aae80b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `@stll/folio-cli`, the `folio` command line over the folio-agents tools for `.docx` files on disk. This release reads: `read`, `outline`, `section`, `stories`, `story`, `find`, `comments`, and `changes`, with flags generated from each tool's schema, a `{ ok, data | error }` envelope, SHA-256 `fileVersion` preconditions, labelled synthetic block ids, and version-bound read cursors.

- [#1062](https://github.com/stella/folio/pull/1062) [`53f76ba`](https://github.com/stella/folio/commit/53f76ba8d0ae866d060dc68b15f19346e41ac14a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `folio render <file> -o <out.pdf|out.png|out.html> [--page n]`, which paints the document with folio's own layout and PDF or DOM backend (PNG through an optional `playwright-core` Chromium), and `folio serve <file> [--port n]`, a read-only live preview on 127.0.0.1 behind a random token that re-renders when the file's version changes.

- [#1057](https://github.com/stella/folio/pull/1057) [`1473902`](https://github.com/stella/folio/commit/1473902334d44918f4287bf0deef51da89ba0073) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add the `folio` write commands: `suggest`, `comment`, `reply`, `resolve`, `accept`, `reject`, and `compare -o`. Each change is one transaction: `--in-place` or `-o` onto a plain `.docx` (replacing an existing file needs `--overwrite` and its version), a required `--expect-version` (waived only by `--no-expect-version`) re-checked under a fenced cooperative write lease (`locked`, exit 10, unless `--force`), a backup of whatever is replaced, a staged and validated package renamed into place, selective saves unless `--allow-repack`, author from `--author`/`FOLIO_AUTHOR`/git with no default, one `--date`-overridable timestamp, a `.folio/journal.jsonl` line, `--tx-id` replay, and checked recovery of a stage a crash left behind. Sidecar files are never used through symlinks or hard links.

### Patch Changes

- Updated dependencies [[`a37cdab`](https://github.com/stella/folio/commit/a37cdab4c079d0d57bbda6bf44d15670b99241da), [`8450922`](https://github.com/stella/folio/commit/8450922f8afc1b5fe20a6b0ec7f9c75361a89bea)]:
  - @stll/folio-core@0.50.0
  - @stll/folio-agents@0.14.0
