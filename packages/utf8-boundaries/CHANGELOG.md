# @stll/utf8-boundaries

## 0.1.1

### Patch Changes

- [#1997](https://github.com/stella/folio/pull/1997) [`20ee7f4`](https://github.com/stella/folio/commit/20ee7f4e19b0d4e3f47a8d72351ea9d288a5dcb2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Declare exported text generators without inferred literal unions.

## 0.1.0

### Minor Changes

- [#1705](https://github.com/stella/folio/pull/1705) [`7fbd519`](https://github.com/stella/folio/commit/7fbd519345dd2e3ca3c09eeb0c116e059064aecb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Saving could corrupt astral characters such as emoji at 16 KiB boundaries inside document XML parts. Upgrade to keep them intact.
