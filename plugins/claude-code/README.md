# folio plugin for Claude Code

Registers the [`folio`](../../packages/cli) MCP server, which reads and changes
`.docx` files in the project directory, and a `docx-review` skill describing
how to use it: read, locate with `find_text`, then propose tracked changes that
name the `fileVersion` they were read at.

## Install

```sh
claude plugin marketplace add stella/folio
claude plugin install folio@folio
```

When the plugin is enabled, Claude Code asks for an **Author**: the name
recorded on every tracked change, comment, and reply.

The server runs `npx -y @stll/folio-cli@<version> mcp --root <project
directory>` and needs Node.js 22 or later. Every path a tool call reads or
writes must be inside the project directory.

## Local development

Until `@stll/folio-cli` is published, or to run a checkout, register the server
from source instead of installing the plugin's server:

```sh
claude mcp add folio --env FOLIO_AUTHOR="Jane Doe" -- \
  bun /path/to/folio/packages/cli/src/bin.ts mcp --root "$PWD"
```

and load the skill from the checkout:

```sh
claude --plugin-dir /path/to/folio/plugins/claude-code
```

Check the plugin and the marketplace with:

```sh
claude plugin validate plugins/claude-code
claude plugin validate .
```

## Versions

The plugin's `version` and the `@stll/folio-cli` version its server runs are
the CLI package's version; `bun scripts/sync-plugin-cli-version.ts --write`
updates both when the CLI is versioned.
