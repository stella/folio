# Folio DOCX for VS Code

- **Preview `.docx` files** with folio's own page layout, read-only. Right-click
  a `.docx` file and choose **Open with Folio DOCX Preview**, or pick
  **Folio DOCX Preview** from **Reopen Editor With...**. The preview shows the
  page count, keeps its place when the file changes on disk, and says why when
  a file cannot be read or laid out. It never writes the document.
- **Give agent mode folio's `.docx` tools.** The extension registers the folio
  MCP server, so the agent can read documents, add comments, and redline them
  with tracked changes. Its edits show up in an open preview a moment later.

The preview is not the default editor for `.docx`: it opens only when you
choose it.

## Try a build locally

From the repository root:

```sh
bun install
cd plugins/vscode
bun install
bun run package          # builds dist/ and writes folio.vsix
code --install-extension folio.vsix
```

Requires VS Code 1.101 or later. Nothing else needs to be installed: the
extension ships the folio CLI and runs it with VS Code's own Node.js, so it
works offline.

## Settings

- `folio.author` (default: empty): the name recorded on every tracked change,
  comment, and reply the MCP server makes. When it is empty, the server uses
  `FOLIO_AUTHOR`, then `git config user.name`, and refuses changes if neither
  is set.

## MCP server

The server appears as **Folio** in the MCP server list. It runs
`folio mcp` with each workspace folder on disk as an allowed root, so the
agent can only read and write `.docx` files inside the workspace. It is
offered only in a trusted workspace. When the workspace folders or
`folio.author` change, the extension updates the server's arguments; restart
the server from the MCP server list to apply them.

## Development

- `bun run build`: bundle the extension, the webview script, and the CLI from
  `packages/cli` into `dist/`
- `bun run test`: unit tests for the MCP launch, the render process, and the
  webview protocol
- `bun run typecheck`

The version follows `@stll/folio-cli`: `scripts/sync-plugin-cli-version.ts`
keeps them equal.
