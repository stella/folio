# folio plugin for herdr

Ctrl-click a `.docx` link in a pane and the document opens in a split to the
right, without taking focus:

- with [terminal-browser](https://terminal-browser.sh) on `PATH`, a live
  preview from `folio serve`: the page re-renders whenever the file changes,
  including after `folio suggest`, `folio accept`, or an MCP edit;
- otherwise one rendered page from `folio render`, shown through the kitty
  graphics protocol (`n`/`p` to change page, `r` to re-render, `q` to quit).

Neither path writes the document. The same action is in the pane's context
menu as **Preview .docx with folio**.

## Install

```sh
herdr plugin install stella/folio/plugins/herdr
```

Requires herdr 0.8.2 or later and Node.js 22 or later. The plugin runs
`folio` from `PATH` when it is installed, otherwise
`npx -y @stll/folio-cli@<version>`; set `FOLIO_BIN` to run another build, for
example a checkout:

```sh
export FOLIO_BIN="bun /path/to/folio/packages/cli/src/bin.ts"
herdr plugin link /path/to/folio/plugins/herdr
```

The PNG fallback needs `playwright-core` and its Chromium next to the CLI
(see the `folio render` section of the CLI README) and a terminal that
speaks the kitty graphics protocol.

## Links

The handler takes `file://` links and paths ending in `.docx`, for example
the hyperlinks `ls --hyperlink=auto` prints. A `file://` link to a host other
than this machine is refused. Web links are left to other plugins.

## Files

- `herdr-plugin.toml`: the `open-docx` action, the `preview` pane, and the
  `docx` link handler
- `bin/open-docx.sh`: resolves the clicked link and opens the preview pane
- `bin/preview.sh`: runs `folio serve` in terminal-browser, or renders a page
- `bin/folio-command.sh`: picks the `folio` command
