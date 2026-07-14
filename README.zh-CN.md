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
  <a href="https://github.com/stella/folio/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-blue" alt="许可证" /></a>
  <a href="https://github.com/stella/folio/issues"><img src="https://img.shields.io/github/issues/stella/folio" alt="问题" /></a>
  <a href="https://discord.gg/8dZjmVFjTK"><img src="https://img.shields.io/badge/discord-join%20chat-5865F2?logo=discord&logoColor=white" alt="Discord" /></a>
</p>

# folio

用于 `.docx` 文件的浏览器编辑器和框架无关引擎。它可以打开、编辑和写入
OOXML 文档，同时保留分页、表格、页眉和页脚、修订以及脚注。

核心包与框架无关。React、Vue、Nuxt 和文档审阅包均构建于其上。

本项目是开源法律工作区 [stella](https://github.com/stella/stella) 的一部分。

请参阅 [DOCX 平台边界](./docs/docx-platform.md)，了解 folio 的职责范围，以及编辑器、
无头工具、agent 和宿主如何共享同一套文档模型和操作约定。

## 标准优先的互操作性

Folio 通过已发布的标准、差异解析、往返测试、交互测试，以及不同独立实现之间可复现的
布局比较，来实现可互操作的 OOXML 行为。

比较报告会记录参考实现、版本和相关渲染环境，使结果保持明确且可复现。

完整的测试方法和参考矩阵请参阅[互操作性参考](./docs/interoperability.md)。

## 软件包

这是一个使用 [Bun](https://bun.sh) 的工作区，包含以下已发布软件包：

| 软件包                                    | 用途                                                    |
| ----------------------------------------- | ------------------------------------------------------- |
| [`@stll/folio-core`](./packages/core)     | OOXML 解析、文档模型、ProseMirror 集成和页面布局        |
| [`@stll/folio-react`](./packages/react)   | 基于 `@stll/folio-core` 构建的 React 编辑器 UI          |
| [`@stll/folio-vue`](./packages/vue)       | Vue 3 编辑器和组合式函数                                |
| [`@stll/folio-nuxt`](./packages/nuxt)     | 为 Vue 编辑器提供 Nuxt 3/4 注册                         |
| [`@stll/folio-agents`](./packages/agents) | 用于读取 `.docx` 文件并提出批注或修改建议的文档审阅工具 |

## 安装

```sh
# React 编辑器（会安装 @stll/folio-core）
bun add @stll/folio-react react react-dom use-intl

# Vue 编辑器
bun add @stll/folio-vue vue

# Nuxt 集成
bun add @stll/folio-nuxt

# agent/审阅工具
bun add @stll/folio-agents

# 或仅安装无头引擎
bun add @stll/folio-core
```

## 快速开始

```tsx
import { DocxEditor } from "@stll/folio-react";
import "@stll/folio-react/standalone.css";

export function Editor({ docx }: { docx: ArrayBuffer }) {
  return <DocxEditor documentBuffer={docx} onSave={(out) => download(out)} />;
}
```

在 SSR 应用中，请通过仅客户端加载或动态导入来加载编辑器。

## 样式

请选择一种样式表。

如果应用未使用 Tailwind，或者希望隔离 folio 的样式，请使用 `standalone.css`：

```tsx
import "@stll/folio-react/standalone.css";
```

可以在 `.folio-root` 上覆盖设计令牌：

```css
.folio-root {
  --background: #fdfdfc;
  --foreground: #1c1c1a;
  --primary: #3b5bdb;
  /* ……只覆盖需要修改的令牌…… */
}
```

如需深色模式，请为 `<html>` 等祖先元素添加 `.dark`。

如果应用已经使用 Tailwind，请使用 `editor.css`。先将 folio 发布的 JavaScript 文件添加到
Tailwind 的扫描源中，然后导入样式表：

```css
/* 应用的 Tailwind 入口文件 */
@import "tailwindcss";
@source "../node_modules/@stll/folio-react/dist/**/*.js";
```

```tsx
import "@stll/folio-react/editor.css";
```

请勿同时导入这两种样式表。`standalone.css` 已经包含 `editor.css` 的全部内容。

## 国际化

编辑器使用 [`use-intl`](https://github.com/amannn/use-intl)。请将编辑器包装在
`IntlProvider` 中，并传入 folio 内置的消息：

```tsx
import { IntlProvider } from "use-intl";
import { DocxEditor } from "@stll/folio-react";
import { FOLIO_LOCALES, getFolioMessages } from "@stll/folio-react/messages";
import "@stll/folio-react/editor.css";

export function Editor({ docx, locale }: { docx: ArrayBuffer; locale: string }) {
  return (
    <IntlProvider locale={locale} messages={getFolioMessages(locale)}>
      <DocxEditor documentBuffer={docx} />
    </IntlProvider>
  );
}
```

`@stll/folio-react/messages` 导出以下内容：

- `getFolioMessages(locale: string): FolioMessages`
- `FOLIO_LOCALES`
- `FolioLocale`
- `isFolioLocale(locale: string): locale is FolioLocale`

内置区域设置：`en`、`de`、`fr`、`es`、`cs`、`ar`、`et`、`he`、`hi`、`hu`、
`lt`、`lv`、`pl`、`pt-BR`、`sk`、`tr`、`zh-CN`。阿拉伯语（`ar`）和希伯来语
（`he`）从右向左书写；请在编辑器外层容器上设置 `dir="rtl"`。

如需将 folio 消息与应用消息合并，请将 folio 保留在自己的 `folio.*` 命名空间中：

```tsx
const messages = { ...getFolioMessages(locale), ...appMessages[locale] };
```

请勿将 folio 的 `folio.*` 键复制到应用的消息目录中。

## 开发

```sh
bun install
bun run build
bun run typecheck
bun run test
bun run lint
bun run validate-dist
```

## 发布

发布流程使用 [Changesets](https://github.com/changesets/changesets)。凡是修改了已发布软件包
`packages/{core,react,agents,vue,nuxt}/src` 下源代码的 PR，都应添加一个 changeset：

```sh
bunx changeset
```

对于无需发布版本的源代码修改，请运行：

```sh
bunx changeset --empty
```

CI 会通过 `bun run changeset:check` 检查此项。合并自动生成的 **Version Packages** PR
后，`publish.yml` 会发布有变更的软件包。

## 致谢

folio 最初是 [Eigenpal](https://eigenpal.com) 的
[docx-editor](https://github.com/eigenpal/docx-editor) 的私有分支，原作者为
[Jedr Blaszyk](https://github.com/jedrazb)。此后，该代码得到扩展，主要用于满足
[stella](https://github.com/stella/stella) 的需求。上游仓库下线后，我们将 folio 分支
作为一个独立维护的延续版本公开发布。原始许可证和版权声明保留在
[`NOTICE.md`](./NOTICE.md) 中。

## 许可证

[Apache-2.0](./LICENSE)
