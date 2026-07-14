<p align="center">
  <img src=".github/assets/banner.png" alt="@stll/folio" width="100%" />
</p>

<p align="center">
  <strong>Editor para navegador e mecanismo independente de framework para documentos OOXML <code>.docx</code>.</strong>
</p>

<p align="center">
  <a href="./README.md">English</a> &middot; <a href="./README.zh-CN.md">简体中文</a> &middot; Português (Brasil)
</p>

<p align="center">
  <a href="https://github.com/stella/stella">stella</a> &middot;
  <a href="https://www.npmjs.com/package/@stll/folio-core">npm</a> &middot;
  <a href="https://github.com/stella/folio/issues">Issues</a> &middot;
  <a href="https://discord.gg/8dZjmVFjTK">Discord</a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@stll/folio-core"><img src="https://img.shields.io/npm/v/@stll/folio-core?label=%40stll%2Ffolio-core" alt="versão no npm" /></a>
  <a href="https://github.com/stella/folio/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-blue" alt="licença" /></a>
  <a href="https://github.com/stella/folio/issues"><img src="https://img.shields.io/github/issues/stella/folio" alt="issues" /></a>
  <a href="https://discord.gg/8dZjmVFjTK"><img src="https://img.shields.io/badge/discord-join%20chat-5865F2?logo=discord&logoColor=white" alt="Discord" /></a>
</p>

# folio

Editor para navegador e mecanismo independente de framework para arquivos `.docx`.
Ele abre, edita e grava documentos OOXML, preservando paginação, tabelas, cabeçalhos
e rodapés, controle de alterações e notas de rodapé.

O pacote principal é independente de framework. Os pacotes para React, Vue, Nuxt e
revisão de documentos são construídos sobre ele.

Parte do [stella](https://github.com/stella/stella), um workspace jurídico de código aberto.

Consulte os [limites da plataforma DOCX](./docs/docx-platform.md) para saber o que pertence
ao folio e como editores, ferramentas headless, agentes e hosts compartilham o mesmo modelo
de documento e contrato de operações.

## Interoperabilidade orientada por padrões

O Folio busca um comportamento OOXML interoperável por meio de padrões publicados,
análise diferencial, testes de ida e volta e de interação, além de comparações de layout
reproduzíveis entre implementações independentes.

Os relatórios de comparação registram a implementação de referência, a versão e o ambiente
de renderização relevante para que os resultados permaneçam explícitos e reproduzíveis.

Consulte as [referências de interoperabilidade](./docs/interoperability.md) para ver a
metodologia de testes completa e a matriz de referências.

## Pacotes

Este é um workspace [Bun](https://bun.sh) com os seguintes pacotes publicados:

| Pacote                                    | Uso                                                                                  |
| ----------------------------------------- | ------------------------------------------------------------------------------------ |
| [`@stll/folio-core`](./packages/core)     | Análise OOXML, modelo de documento, integração com ProseMirror e layout de página    |
| [`@stll/folio-react`](./packages/react)   | Interface de edição em React criada sobre o `@stll/folio-core`                       |
| [`@stll/folio-vue`](./packages/vue)       | Editor para Vue 3 e composables                                                      |
| [`@stll/folio-nuxt`](./packages/nuxt)     | Registro do editor Vue para Nuxt 3/4                                                 |
| [`@stll/folio-agents`](./packages/agents) | Ferramentas de revisão que leem arquivos `.docx` e propõem comentários ou alterações |

## Instalação

```sh
# editor React (instala também o @stll/folio-core)
bun add @stll/folio-react react react-dom use-intl

# editor Vue
bun add @stll/folio-vue vue

# integração com Nuxt
bun add @stll/folio-nuxt

# ferramentas de agente/revisão
bun add @stll/folio-agents

# ou apenas o mecanismo headless
bun add @stll/folio-core
```

## Início rápido

```tsx
import { DocxEditor } from "@stll/folio-react";
import "@stll/folio-react/standalone.css";

export function Editor({ docx }: { docx: ArrayBuffer }) {
  return <DocxEditor documentBuffer={docx} onSave={(out) => download(out)} />;
}
```

Em aplicações com SSR, carregue o editor somente no cliente ou por meio de uma importação dinâmica.

## Estilos

Escolha uma folha de estilos.

Use `standalone.css` quando sua aplicação não utilizar Tailwind ou quando quiser manter os
estilos do folio isolados:

```tsx
import "@stll/folio-react/standalone.css";
```

Sobrescreva os tokens em `.folio-root`:

```css
.folio-root {
  --background: #fdfdfc;
  --foreground: #1c1c1a;
  --primary: #3b5bdb;
  /* ...apenas os tokens que você quiser alterar... */
}
```

Para o modo escuro, adicione `.dark` a um elemento ancestral, como `<html>`.

Use `editor.css` quando sua aplicação já utilizar Tailwind. Adicione o JavaScript distribuído
pelo folio às fontes verificadas pelo Tailwind e, em seguida, importe a folha de estilos:

```css
/* arquivo de entrada do Tailwind na sua aplicação */
@import "tailwindcss";
@source "../node_modules/@stll/folio-react/dist/**/*.js";
```

```tsx
import "@stll/folio-react/editor.css";
```

Não importe as duas folhas de estilos. `standalone.css` já inclui tudo o que existe em `editor.css`.

## Internacionalização

O editor usa [`use-intl`](https://github.com/amannn/use-intl). Envolva-o em um `IntlProvider`
e passe as mensagens incluídas no folio:

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

`@stll/folio-react/messages` exporta:

- `getFolioMessages(locale: string): FolioMessages`
- `FOLIO_LOCALES`
- `FolioLocale`
- `isFolioLocale(locale: string): locale is FolioLocale`

Locales incluídos: `en`, `de`, `fr`, `es`, `cs`, `ar`, `et`, `he`, `hi`, `hu`,
`lt`, `lv`, `pl`, `pt-BR`, `sk`, `tr`, `zh-CN`. Árabe (`ar`) e hebraico (`he`)
são escritos da direita para a esquerda; defina `dir="rtl"` em um contêiner ao redor
do editor para esses locales.

Para combinar as mensagens do folio com as mensagens da aplicação, mantenha o folio em seu
próprio namespace `folio.*`:

```tsx
const messages = { ...getFolioMessages(locale), ...appMessages[locale] };
```

Não copie as chaves `folio.*` do folio para o catálogo da sua aplicação.

## Desenvolvimento

```sh
bun install
bun run build
bun run typecheck
bun run test
bun run lint
bun run validate-dist
```

## Publicação

As publicações usam [Changesets](https://github.com/changesets/changesets). Adicione um changeset
a todo PR que editar o código-fonte de um pacote publicado em
`packages/{core,react,agents,vue,nuxt}/src`:

```sh
bunx changeset
```

Para uma alteração no código-fonte que não precise de uma nova versão, use:

```sh
bunx changeset --empty
```

O CI verifica isso com `bun run changeset:check`. O merge do PR **Version Packages** gerado
publica os pacotes alterados por meio de `publish.yml`.

## Agradecimentos

O folio começou como um fork privado do
[docx-editor](https://github.com/eigenpal/docx-editor), da [Eigenpal](https://eigenpal.com),
criado por [Jedr Blaszyk](https://github.com/jedrazb). Desde então, o código foi ampliado,
principalmente para atender às necessidades do [stella](https://github.com/stella/stella).
Depois que o repositório original foi retirado do ar, passamos a publicar o fork folio como uma
continuação mantida de forma independente. A licença e os direitos autorais originais foram
preservados em [`NOTICE.md`](./NOTICE.md).

## Licença

[Apache-2.0](./LICENSE)
