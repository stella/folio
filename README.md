<p align="center">
  <img src=".github/assets/banner.png" alt="@stll/folio" width="100%" />
</p>

<p align="center">
  <strong>Browser editor and framework-neutral engine for OOXML <code>.docx</code> documents.</strong>
</p>

<p align="center">
  English &middot; <a href="./README.zh-CN.md">简体中文</a> &middot; <a href="./README.pt-BR.md">Português (Brasil)</a>
</p>

<p align="center">
  <a href="https://github.com/stella/stella">stella</a> &middot;
  <a href="https://www.npmjs.com/package/@stll/folio-core">npm</a> &middot;
  <a href="https://github.com/stella/folio/issues">Issues</a> &middot;
  <a href="https://discord.gg/8dZjmVFjTK">Discord</a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@stll/folio-core"><img src="https://img.shields.io/npm/v/@stll/folio-core?label=%40stll%2Ffolio-core" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@stll/folio-core"><img src="https://img.shields.io/npm/dm/%40stll%2Ffolio-core" alt="npm downloads per month" /></a>
  <a href="https://github.com/stella/folio/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-blue" alt="License" /></a>
  <a href="https://github.com/stella/folio/issues"><img src="https://img.shields.io/github/issues/stella/folio" alt="Issues" /></a>
  <a href="https://discord.gg/8dZjmVFjTK"><img src="https://img.shields.io/badge/discord-join%20chat-5865F2?logo=discord&logoColor=white" alt="Discord" /></a>
</p>

# folio

Folio is an embeddable Word-document editor for web applications. Pass it a
`.docx` as a `File`, `Blob`, `ArrayBuffer`, or `Uint8Array`; it renders editable,
paginated content in the browser and returns a `.docx` when the user saves.

Use the React, Vue, or Nuxt editor in an application, or use `folio-core`
directly for parsing, editing, layout, and document review without a UI.

<p align="center">
  <img src=".github/assets/folio-showcase.gif" alt="Folio editing a five-page DOCX with tables, a chart, comments, and tracked changes" width="100%" />
</p>

## Install

```sh
bun add @stll/folio-react react react-dom use-intl
```

`@stll/folio-core` is installed with the React editor.

## React quick start

```tsx
import { useState } from "react";
import { IntlProvider } from "use-intl";
import { DocxEditor } from "@stll/folio-react";
import { getFolioMessages } from "@stll/folio-react/messages";
import "@stll/folio-react/standalone.css";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const downloadDocx = (buffer: ArrayBuffer) => {
  const url = URL.createObjectURL(new Blob([buffer], { type: DOCX_MIME }));
  const link = document.createElement("a");
  link.href = url;
  link.download = "edited.docx";
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

export function Editor() {
  const [file, setFile] = useState<File | null>(null);

  return (
    <IntlProvider locale="en" messages={getFolioMessages("en")}>
      <input
        type="file"
        accept=".docx"
        onChange={(event) => setFile(event.currentTarget.files?.[0] ?? null)}
      />
      {file && <DocxEditor documentBuffer={file} author="Editor" onSave={downloadDocx} />}
    </IntlProvider>
  );
}
```

The editor toolbar calls `onSave` with the updated DOCX bytes. You can also hold
a `DocxEditorRef` and call `await editorRef.current?.save()` yourself.

## What it handles

- Paginated text, formatting, lists, tables, images, sections, headers, and
  footers
- Comments, footnotes, and tracked changes that Microsoft Word can review,
  accept, or reject
- DOCX round trips that preserve untouched package parts and unsupported OOXML
- Editing modes, find, page setup, document outline, and save hooks
- Bundled messages for 17 locales, including right-to-left editor chrome

## Choose a package

| Package                                   | Choose it when you need                                                 |
| ----------------------------------------- | ----------------------------------------------------------------------- |
| [`@stll/folio-react`](./packages/react)   | The complete editor as a React component                                |
| [`@stll/folio-vue`](./packages/vue)       | The complete editor as a Vue 3 component                                |
| [`@stll/folio-nuxt`](./packages/nuxt)     | SSR-safe registration of the Vue editor in Nuxt 3 or 4                  |
| [`@stll/folio-core`](./packages/core)     | DOCX parsing, ProseMirror editing, page layout, review, or redline APIs |
| [`@stll/docx-core`](./packages/docx-core) | The typed OOXML model, validation, serialization, and projection kernel |
| [`@stll/folio-agents`](./packages/agents) | Tools that read documents and propose comments or tracked changes       |

Install the Vue editor with `bun add @stll/folio-vue vue`, the Nuxt module with
`bun add @stll/folio-nuxt`, or the framework-neutral engine with
`bun add @stll/folio-core`.

## Create a Word redline without an editor

Compare two DOCX files and write their differences as native tracked changes:

```ts
import { generateRedlineDocx } from "@stll/folio-core/redline";

const result = await generateRedlineDocx(originalDocx, revisedDocx, {
  author: "Reviewer",
});

await store(result.buffer);
```

For deterministic changes to one document, use `FolioDocxReviewer`. See the
[`folio-core` review APIs](./packages/core/README.md#native-word-redlines).

## Integration notes

- Use `standalone.css` by itself. Applications that already use Tailwind can
  instead follow the [`editor.css` setup](./packages/react/README.md#exports).
- The editor requires the DOM. In an SSR application, load it from a client-only
  component or dynamic import; Nuxt users can use `@stll/folio-nuxt`.
- Architecture and test methodology live in the
  [DOCX platform boundary](./docs/docx-platform.md) and
  [interoperability guide](./docs/interoperability.md).

## Development

```sh
bun install
bun run build
bun run typecheck
bun run test
bun run lint
bun run validate-dist
```

Published-package source changes require a [Changeset](https://github.com/changesets/changesets).

## Acknowledgements

Folio originated as a fork of [Eigenpal](https://eigenpal.com)'s
[docx-editor](https://github.com/eigenpal/docx-editor) by
[Jedr Blaszyk](https://github.com/jedrazb) and is independently maintained as
part of [stella](https://github.com/stella/stella). The original license and
copyright are preserved in [`NOTICE.md`](./NOTICE.md).

## License

[Apache-2.0](./LICENSE)
