# @stll/utf8-boundaries

## 0.1.0

### Minor Changes

- [#1705](https://github.com/stella/folio/pull/1705) [`7fbd519`](https://github.com/stella/folio/commit/7fbd519345dd2e3ca3c09eeb0c116e059064aecb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Saving could corrupt astral characters such as emoji at 16 KiB boundaries inside document XML parts. Upgrade to keep them intact.
