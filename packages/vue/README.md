# @stll/folio-vue

Vue 3 adapter for the folio document editor.

## Install

```sh
bun add @stll/folio-vue vue@^3.3.0 \
  prosemirror-commands@^1.7.1 prosemirror-dropcursor@^1.8.2 prosemirror-gapcursor@^1.4.1 prosemirror-history@^1.5.0 \
  prosemirror-inputrules@^1.5.0 prosemirror-keymap@^1.2.3 prosemirror-model@^1.25.9 prosemirror-state@^1.4.4 \
  prosemirror-tables@^1.8.5 prosemirror-transform@^1.12.0 prosemirror-view@^1.42.6 \
  y-protocols@^1.0.6 yjs@^13.6.31
```

ProseMirror and Yjs packages are required peers, including the collaboration
package `y-protocols`. Install one compatible copy of each in the host
application so schemas, plugins, editor states, and collaborative documents share
the same runtime. The command above includes the complete editor peer closure.

Import `DocxEditor` from `@stll/folio-vue` and load
`@stll/folio-vue/editor.css` once in your application.

Every edit is a typed operation on the OOXML document model, the same one the
headless engine and the agent tools apply. An edit the engine cannot make yet is
refused with a notice instead of being applied partly.

Keyboard shortcuts default to the editor root. Set `keyboardShortcuts="document"`
only when this editor should own shortcuts across the host page.
