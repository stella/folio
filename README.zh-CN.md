<p align="center">
  <img src=".github/assets/banner.png" alt="@stll/folio" width="100%" />
</p>

<p align="center">
  <strong>用于 OOXML <code>.docx</code> 文档的浏览器编辑器和框架无关引擎。</strong>
</p>

<p align="center">
  <a href="./README.md">English</a> &middot; 简体中文 &middot; <a href="./README.pt-BR.md">Português (Brasil)</a>
</p>

<p align="center">
  <a href="https://github.com/stella/stella">stella</a> &middot;
  <a href="https://www.npmjs.com/package/@stll/folio-core">npm</a> &middot;
  <a href="https://github.com/stella/folio/issues">问题反馈</a> &middot;
  <a href="https://discord.gg/8dZjmVFjTK">Discord</a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@stll/folio-core"><img src="https://img.shields.io/npm/v/@stll/folio-core?label=%40stll%2Ffolio-core" alt="npm 版本" /></a>
  <a href="https://www.npmjs.com/package/@stll/folio-core"><img src="https://img.shields.io/npm/dm/%40stll%2Ffolio-core" alt="npm 月下载量" /></a>
  <a href="https://github.com/stella/folio/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-blue" alt="许可证" /></a>
  <a href="https://github.com/stella/folio/issues"><img src="https://img.shields.io/github/issues/stella/folio" alt="问题" /></a>
  <a href="https://discord.gg/8dZjmVFjTK"><img src="https://img.shields.io/badge/discord-join%20chat-5865F2?logo=discord&logoColor=white" alt="Discord" /></a>
</p>

# folio

Folio 是可嵌入 Web 应用的 Word 文档编辑器。传入 `File`、`Blob`、`ArrayBuffer` 或 `Uint8Array` 格式的
`.docx`，即可在浏览器中渲染可编辑的分页内容，并在用户保存时返回 `.docx`。

在应用中使用 React、Vue 或 Nuxt 编辑器，也可以直接使用 `folio-core`，在没有 UI 的情况下完成解析、编辑、布局和文档审阅。

<p align="center">
  <img src=".github/assets/folio-showcase.gif" alt="Folio 编辑包含表格、图表、批注和修订的五页 DOCX 文档" width="100%" />
</p>

## 安装

```sh
bun add @stll/folio-react react react-dom use-intl
```

`@stll/folio-core` 会随 React 编辑器一同安装。

## React 快速开始

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

编辑器工具栏会通过 `onSave` 提供更新后的 DOCX 字节。也可以持有 `DocxEditorRef`，自行调用 `await editorRef.current?.save()`。

## 支持的功能

- 分页文本、格式、列表、表格、图片、节、页眉和页脚
- Microsoft Word 可以审阅、接受或拒绝的批注、脚注和修订
- 保留未修改文档部分和不支持 OOXML 的 DOCX 往返处理
- 编辑模式、查找、页面设置、文档大纲和保存钩子
- 内置 17 种语言的消息，包括从右向左的编辑器界面

## 选择软件包

| 软件包                                    | 适用场景                                              |
| ----------------------------------------- | ----------------------------------------------------- |
| [`@stll/folio-react`](./packages/react)   | 完整的 React 编辑器组件                               |
| [`@stll/folio-vue`](./packages/vue)       | 完整的 Vue 3 编辑器组件                               |
| [`@stll/folio-nuxt`](./packages/nuxt)     | 在 Nuxt 3 或 4 中以 SSR 安全方式注册 Vue 编辑器       |
| [`@stll/folio-core`](./packages/core)     | DOCX 解析、ProseMirror 编辑、页面布局、审阅或修订 API |
| [`@stll/docx-core`](./packages/docx-core) | 类型化 OOXML 模型、验证、序列化和投影内核             |
| [`@stll/folio-agents`](./packages/agents) | 读取文档并提出批注或修订建议的工具                    |

使用 `bun add @stll/folio-vue vue` 安装 Vue 编辑器，使用 `bun add @stll/folio-nuxt` 安装 Nuxt 模块，或使用 `bun add @stll/folio-core` 安装框架无关引擎。

## 无需编辑器即可创建 DOCX 修订

比较两个 DOCX 文件，并将差异写入修订：

```ts
import { compareDocx } from "@stll/folio-core";

const result = await compareDocx(originalDocx, revisedDocx, {
  author: "Reviewer",
  timestamp: "2024-03-01T00:00:00.000Z",
});
if (result.isErr()) throw result.error;

await store(result.value.buffer);
```

如需对单个文档执行确定性修改，请使用 `FolioDocxReviewer`。参阅 [`folio-core` 审阅 API](./packages/core/README.md#native-docx-redlines)。

## 集成说明

- 单独使用 `standalone.css`。已使用 Tailwind 的应用可以改用 [`editor.css` 配置](./packages/react/README.md#exports)。
- 编辑器需要 DOM。在 SSR 应用中，请从仅客户端组件或动态导入中加载；Nuxt 用户可以使用 `@stll/folio-nuxt`。
- 架构和测试方法请参阅 [DOCX 平台边界](./docs/docx-platform.md) 与 [互操作性指南](./docs/interoperability.md)。

## 开发

```sh
bun install
bun run build
bun run typecheck
bun run test
bun run lint
bun run validate-dist
```

发布软件包源代码的修改需要添加 [Changeset](https://github.com/changesets/changesets)。

## 致谢

Folio 源自 [Eigenpal](https://eigenpal.com) 的 [docx-editor](https://github.com/eigenpal/docx-editor)，原作者为
[Jedr Blaszyk](https://github.com/jedrazb)，现作为 [stella](https://github.com/stella/stella) 的一部分独立维护。
原始许可证和版权声明保留在 [`NOTICE.md`](./NOTICE.md) 中。

## 许可证

[Apache-2.0](./LICENSE)
