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

| Command    | Tool                   | Prints                                                   |
| ---------- | ---------------------- | -------------------------------------------------------- |
| `read`     | `read_document`        | Body blocks with ids, text, and `blockTextHash`, paged   |
| `outline`  | `get_document_outline` | Heading outline with section handles                     |
| `section`  | `read_section`         | One section's blocks, by handle                          |
| `stories`  | `list_stories`         | Header, footer, footnote, and endnote story handles      |
| `story`    | `read_story`           | One story's text, by handle                              |
| `find`     | `find_text`            | Exact matches with range handles and context             |
| `comments` | `read_comments`        | Comment threads with replies and resolved state          |
| `changes`  | `read_changes`         | Pending tracked changes                                  |

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

| Exit | Meaning                                                                  |
| ---- | ------------------------------------------------------------------------ |
| 0    | success                                                                  |
| 1    | unexpected internal error                                                |
| 2    | usage, input, or refused-operation error                                 |
| 6    | file, change, or comment not found                                       |
| 8    | path outside the allowed roots                                           |
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
