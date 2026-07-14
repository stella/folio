<!--
  Vue parity playground — the @stll/folio-vue mirror of packages/playground's
  React App. Mounts the Vue `DocxEditor`, loads `?file=<name>` fixtures through
  the shared serveFixtures middleware, and installs the cross-adapter
  `window.__folioParity` bridge the `tests/parity` specs drive. The bridge is a
  byte-for-byte match of the React playground's, built only on `DocxEditorRef`
  members classified `paired` in scripts/parity/parity.contract.json.
-->
<template>
  <div class="pg-vue-shell">
    <main class="pg-vue-editor">
      <DocxEditor
        ref="editorRef"
        :document="documentBuffer ? null : currentDocument"
        :document-buffer="documentBuffer"
        author="Folio User"
        :show-toolbar="true"
        :show-ruler="true"
        :initial-zoom="1"
      />
    </main>
    <p v-if="status" class="pg-vue-status">{{ status }}</p>
  </div>
</template>

<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, shallowRef } from "vue";

import { DocxEditor, createEmptyDocument, createStellaStyleDocumentPreset } from "@stll/folio-vue";
import type { DocxEditorRef, Document as FolioDocument } from "@stll/folio-vue";

import type { FolioParityBridge } from "./parityBridge";
import { buildParityBridge } from "./parityBridge";

declare global {
  // eslint-disable-next-line no-var
  var __folioParity: FolioParityBridge | undefined;
}

const editorRef = ref<DocxEditorRef | null>(null);
const documentBuffer = shallowRef<ArrayBuffer | null>(null);
const currentDocument = shallowRef<FolioDocument | null>(null);
const status = ref("");

onMounted(() => {
  void loadFromQuery();
  globalThis.__folioParity = buildParityBridge(() => editorRef.value);
});

onBeforeUnmount(() => {
  globalThis.__folioParity = undefined;
});

async function loadFromQuery(): Promise<void> {
  const fixtureFile = new URLSearchParams(window.location.search).get("file");
  if (!fixtureFile) {
    currentDocument.value = createEmptyDocument({
      preset: createStellaStyleDocumentPreset(),
    });
    return;
  }
  try {
    status.value = "Loading fixture...";
    const response = await fetch(`/fixtures/${fixtureFile}`);
    if (!response.ok) {
      status.value = `Fixture not found: ${fixtureFile}`;
      return;
    }
    documentBuffer.value = await response.arrayBuffer();
    status.value = "";
  } catch {
    status.value = "Error loading fixture";
  }
}
</script>

<style scoped>
.pg-vue-shell {
  display: flex;
  flex-direction: column;
  height: 100dvh;
}
.pg-vue-editor {
  flex: 1 1 auto;
  min-height: 0;
}
.pg-vue-status {
  flex: 0 0 auto;
  padding: 8px 12px;
  opacity: 0.7;
}
</style>
