# Folio DOCX for VS Code

<!-- Demo GIF: opening a .docx, typing, switching to tracked changes, saving. -->

Open a `.docx` file in VS Code and it opens as pages you can type into.
Save writes back to the same file.

- **Edit `.docx` files.** Folio's editor is the default for `.docx`: type,
  press Enter, format, comment, and save with Ctrl+S / Cmd+S. Undo and redo
  are VS Code's own. The toolbar switches between editing directly and
  recording tracked changes (`folio.editor.trackChanges` sets which one a
  document opens in).
- **Saves you can undo on disk.** Every save keeps the previous version in
  `.folio/backups/<name>/` beside the file and logs a line in
  `.folio/journal.jsonl`. A save that has to rewrite the whole package
  instead of only the paragraphs you changed asks first, once per session.
- **Agents and you on the same file.** The extension registers the folio MCP
  server, so agent mode can read documents, add comments, and redline them
  with tracked changes. When an agent writes to a document you have open with
  unsaved edits, the editor saves your edits first, then shows the agent's
  change. With no unsaved edits it just reloads and says who changed it.
- **Read-only when you want it.** **Folio: Open Read-Only** (command palette,
  or right-click a `.docx`) opens the same pages with editing off.

## Try a build locally

From the repository root:

```sh
bun install
cd plugins/vscode
bun install
bun run package          # builds dist/ and writes folio.vsix
code --install-extension folio.vsix
```

Every pull request that touches the extension also uploads the `.vsix` as the
`folio-vscode-vsix` artifact of its **VS Code extension** check.

Requires VS Code 1.101 or later. Nothing else needs to be installed: the
extension ships the folio CLI and runs it with VS Code's own Node.js, so it
works offline.

## How saving works

- A save goes through `folio save`: it is refused when the file changed on
  disk since you opened it (you choose to overwrite it, reload it, or
  cancel), and it replaces the file atomically.
- While a document has unsaved edits, the editor holds folio's write lease on
  it (`.<name>.docx.folio-lock`). A folio write that finds it asks the editor
  to save and let go instead of failing, then applies its change to the saved
  version. That save never stops to ask; if it had to rewrite the whole
  package, a notice afterwards says so and where the backup is.
- A change on disk that did not come through folio (another program saved the
  file) while you have unsaved edits shows a notice: reload theirs, keep
  yours, or save yours as a copy.
- Hot exit keeps unsaved edits across restarts; they come back as unsaved.
- Files on a remote or virtual file system open read-only; Save As still
  works.

## Settings

- `folio.author` (default: empty): the name recorded on tracked changes,
  comments, and replies. When empty, the editor uses `git config user.name`,
  then your OS account name. The MCP server uses `FOLIO_AUTHOR`, then
  `git config user.name`, and refuses changes if neither is set.
- `folio.editor.trackChanges` (`off` or `on`, default `off`): whether a
  document opens editing directly or recording tracked changes.
- `folio.editor.confirmRewrite` (default: `true`): ask before the first save
  of a session that rewrites the whole package. **Always** in that dialog
  turns it off.

## MCP server

The server appears as **Folio** in the MCP server list. It runs
`folio mcp` with each workspace folder on disk as an allowed root, so the
agent can only read and write `.docx` files inside the workspace. It is
offered only in a trusted workspace. When the workspace folders or
`folio.author` change, the extension updates the server's arguments; restart
the server from the MCP server list to apply them.

## Development

- `bun run build`: bundle the extension and the CLI from `packages/cli` into
  `dist/`, and copy the editor bundle from `packages/editor-web/dist/vscode`
  into `dist/editor` (building it first when it is missing)
- `bun run test`: unit tests for saving, the lease, backups, the webview
  protocol, and the MCP launch
- `bun run test:smoke`: after a build, download VS Code into `.vscode-test/`,
  open a `.docx` in it with this extension, type, save, and revert (CI runs it
  under `xvfb-run`)
- `bun run typecheck`

The version follows `@stll/folio-cli`: `scripts/sync-plugin-cli-version.ts`
keeps them equal, and each release that changes the CLI's version publishes
the extension.
