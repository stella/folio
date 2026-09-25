# @stll/folio-cli

`folio`: a command line for reading and reviewing `.docx` files on disk. It
runs the [`@stll/folio-agents`](https://www.npmjs.com/package/@stll/folio-agents)
tools over local files, so a script, a person at a terminal, and an agent see
the same blocks, ids, and results.

Part of [stella](https://github.com/stella/stella), an open-source legal workspace.

## Install

```sh
bun add -g @stll/folio-cli   # or: npm install -g @stll/folio-cli
folio --help
```

Node.js 22 or later.

## Reading

| Command    | Tool                   | Prints                                                 |
| ---------- | ---------------------- | ------------------------------------------------------ |
| `read`     | `read_document`        | Body blocks with ids, text, and `blockTextHash`, paged |
| `outline`  | `get_document_outline` | Heading outline with section handles                   |
| `section`  | `read_section`         | One section's blocks, by handle                        |
| `stories`  | `list_stories`         | Header, footer, footnote, and endnote story handles    |
| `story`    | `read_story`           | One story's text, by handle                            |
| `find`     | `find_text`            | Exact matches with range handles and context           |
| `comments` | `read_comments`        | Comment threads with replies and resolved state        |
| `changes`  | `read_changes`         | Pending tracked changes                                |

```sh
folio read contract.docx --max-blocks 50
folio find contract.docx --query "Termination" --whole-word
folio section contract.docx --handle '{"type":"headingSection",...}'
folio comments contract.docx --filter open --output text
```

Each command's flags are generated from its tool's input schema: a property
`matchCase` is `--match-case`; objects and arrays take JSON. `--input` passes
the whole argument object as JSON (inline, `@args.json`, or `-` for stdin), and
flags given with it override its fields. `folio <command> --help` lists them.

Reading never writes: not the file, not identifiers, not a cache.

## Changing a document

| Command   | Tool                | Does                                                         |
| --------- | ------------------- | ------------------------------------------------------------ |
| `suggest` | `suggest_changes`   | Applies a batch of edit operations as tracked changes        |
| `comment` | `add_comment`       | Comments on a block, optionally quoting text in it           |
| `reply`   | `reply_comment`     | Replies to a comment thread                                  |
| `resolve` | `resolve_comment`   | Resolves or reopens a comment thread                         |
| `accept`  | `resolve_changes`   | Accepts tracked changes: `--id <id>` (repeatable) or `--all` |
| `reject`  | `resolve_changes`   | Rejects tracked changes: `--id <id>` (repeatable) or `--all` |
| `compare` | `compare_documents` | Diffs two files; with `-o`, writes their redline             |

```sh
folio suggest contract.docx --input @ops.json --in-place --expect-version 9f2c…
folio comment contract.docx --block-id 1A2B3C4D --text "Confirm the rate." -o reviewed.docx --expect-version 9f2c…
folio accept contract.docx --all --in-place --expect-version 51ab…
folio compare signed.docx draft.docx -o redline.docx --expect-version 07de…
```

`suggest --input` takes the operation batch: `{ "operations": [...] }`, or the
bare array. The operation types and fields are those of `suggest_changes` in
`@stll/folio-agents` (`folio suggest --help` lists them); every operation's
targets come from a read of the same `fileVersion`. Edits are tracked changes;
`--direct` edits the text instead and must be asked for.

Every change commits, so there is no separate save:

- **Destination.** Exactly one of `--in-place` or `-o <path>`, and only a
  `.docx` whose name does not start with a dot and that is not inside a
  `.folio` directory (`invalid_destination`). `-o` never modifies the input
  and refuses a path that is the input (or, for `compare`, the revised file)
  under another name: symlinks and hard links are resolved and files are
  compared by device and inode. `-o` refuses an existing file unless
  `--overwrite` is given with `--expect-destination-version <fileVersion>`
  of the file being replaced.
- **Preconditions.** A change needs `--expect-version <fileVersion>` from the
  read it was based on, and refuses (`stale_version`, exit 10) a file that
  changed since; `--no-expect-version` skips the check and must be given
  explicitly. Under the write lease the file is hashed again right before
  the commit, so a change made during the edit is refused too. A batch lands whole
  or not at all: a stale target (`stale_target`, exit 10), an ambiguous `find`
  (`ambiguous_target`, exit 2), or any other refused operation writes nothing,
  with the per-operation reasons in `error.details`.
- **Atomic write.** The new package is staged beside the destination, checked
  (it reopens, every changed XML part is well formed, every relationship id a
  changed part uses resolves, every internal target a changed `.rels` part
  names exists), flushed to disk, and renamed over the destination. Whatever
  a write replaces (in place, or with `--overwrite`) is first copied to
  `.folio/backups/<document name>/<fileVersion>.docx`, flushed with its
  directory; the newest 20 per document are kept. A file with more than one
  hard link is not written in place.
- **Minimal saves.** A change is written by patching the edited paragraphs
  into the original package, leaving every other part's content as it was.
  When that is not possible (a paragraph added or removed, styles or section
  properties changed) the command refuses with `repack_required` rather than
  silently rewriting the whole package; `--allow-repack` permits it. The
  receipt reports `saveStrategy` (`selective`, `full-repack` with a
  `repackReason`, or `redline`) and `changedParts`. The promise is part
  content, not identical ZIP bytes.
- **Write lease.** A write holds `.<name>.docx.folio-lock` beside the file
  (pid, host, expiry, and a random token), created complete or not at all.
  Another holder, such as an editor session, makes a write fail with
  `locked` (exit 10) unless `--force`. Just before journaling and again
  before the rename, a write checks that the lock still carries its token, so
  a writer whose lease was taken over stops without writing. Reads ignore the
  lease. A lease whose process has exited, or that expired, is replaced; a
  lock that cannot be parsed counts as held until it is older than a lease.
- **Sidecar files.** The lock, the stage, the journal, and backups are never
  read or written through a symlink or a file with more than one hard link
  (`unsafe_path`, exit 8). `.folio` and its subdirectories are created with
  mode 0700; an existing `.folio` that is a symlink or not a directory is
  refused, not followed.
- **Author and time.** The author is `--author`, else `FOLIO_AUTHOR`, else git
  `user.name`; with none, the command refuses. Every revision, comment, and
  reply of one transaction carries one UTC timestamp, `--date` to fix it for
  reproducible output. (`compare` stamps its redline with the current time.)

### Journal, retries, and recovery

Each committed transaction appends a line to `.folio/journal.jsonl` beside the
destination (or `--journal <path>`): its `txId`, `fromVersion`, `toVersion`,
author, time, the operations, their receipts, and the command's receipt.

`--tx-id <key>` makes a change idempotent: running a committed transaction
again returns its original receipt with `status: "replayed"` and writes
nothing; reusing the key for a different request is refused
(`transaction_conflict`).

The journal line is written after the staged package validates and before the
rename. If a process dies between the two, the next write to that file
completes the rename only when the line names that file, the stage is a
regular single-link file holding the recorded version, it still parses as a
package, and the file is still what the transaction replaced; any other stage
is discarded (unlinked, never followed). Recovery is reported under
`recovered`, and the write's own version check runs against the file as
recovery left it.

### Known limits

A process on the same machine that can rename directories inside the
document's directory while a write runs can still race it: names are
resolved, then re-checked (the handle's device and inode against the checked
path, and the directory's real path just before the rename), but not held
open across the whole transaction. The tool caller cannot exploit this; it
matters only against another local actor with write access to the
directory.

## Output

Every command prints one envelope:

```json
{ "ok": true, "data": { "path": "/abs/contract.docx", "fileVersion": "9f2c…", "result": {} } }
{ "ok": false, "error": { "code": "stale_version", "message": "…", "hint": "Re-read the document…" } }
```

`--output json` (the default when stdout is not a terminal) prints the
envelope on stdout for success and failure alike. `--output text` (the default
on a terminal) prints a readable rendering, and failures as `error:` and
`hint:` lines on stderr.

| Exit | Meaning                                                                    |
| ---- | -------------------------------------------------------------------------- |
| 0    | success                                                                    |
| 1    | unexpected internal error                                                  |
| 2    | usage, input, or refused-operation error                                   |
| 6    | file, change, or comment not found                                         |
| 8    | path outside the allowed roots, or unsafe to write through                 |
| 10   | conflict with current state (stale version, lock held, destination exists) |

## Versions and identifiers

`fileVersion` is the lowercase hex SHA-256 of the file's bytes.
`--expect-version <sha256>` refuses the command (exit 10, `stale_version`)
unless the file still has that version.

A block id is the paragraph's own `w14:paraId` when it has one
(`blockIdSource: "package"`). A paragraph without one gets an id derived from
its text and position (`blockIdSource: "synthetic"`): stable across reads of
the same bytes, but valid only for the `fileVersion` it was read at.

`read` pages with `--max-blocks`; a page that stops early carries
`nextCursor`, which `--cursor` continues. A cursor is bound to the version it
was issued for, so a cursor from before an edit is refused rather than
continuing on different content.

Offsets in range handles (`startOffset`, `endOffset`) are UTF-16 code-unit
indices into the block `text` that `read` returns, the indexing JavaScript
strings use. A consumer that counts Unicode code points converts with
`Array.from(text.slice(0, offset)).length`.

## Untrusted documents

Reads return document text verbatim. Text from an untrusted `.docx` can carry
instructions aimed at a model that reads it; treat it as data.
