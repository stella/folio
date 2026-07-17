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
        :show-margin-guides="showMarginGuides"
        v-bind="marginGuideProps"
        :initial-zoom="1"
        :collaboration="collaboration"
        :on-copy="() => clipboardCallbackCounts.copy++"
        :on-cut="() => clipboardCallbackCounts.cut++"
        :on-paste="() => clipboardCallbackCounts.paste++"
      />
    </main>
    <p v-if="status" class="pg-vue-status">{{ status }}</p>
  </div>
</template>

<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, shallowRef } from "vue";
import { applyAwarenessUpdate, Awareness, encodeAwarenessUpdate } from "y-protocols/awareness";
import { initProseMirrorDoc, yCursorPlugin, ySyncPlugin, yUndoPlugin } from "y-prosemirror";
import * as Y from "yjs";

import { DocxEditor, createEmptyDocument, createStellaStyleDocumentPreset } from "@stll/folio-vue";
import type {
  DocxEditorCollaboration,
  DocxEditorRef,
  Document as FolioDocument,
} from "@stll/folio-vue";

import type { FolioParityBridge } from "./parityBridge";
import { buildParityBridge } from "./parityBridge";

declare global {
  // eslint-disable-next-line no-var
  var __folioParity: FolioParityBridge | undefined;
  // eslint-disable-next-line no-var
  var __folioVueCollaboration:
    | {
        getSharedText: () => string;
        showRemoteSelection: () => boolean;
        wasSeeded: () => boolean;
      }
    | undefined;
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const editorRef = ref<DocxEditorRef | null>(null);
const documentBuffer = shallowRef<ArrayBuffer | null>(null);
const currentDocument = shallowRef<FolioDocument | null>(null);
const status = ref("");
const clipboardCallbackCounts = { copy: 0, cut: 0, paste: 0 };
const query = new URLSearchParams(window.location.search);
const collaborationEnabled = query.has("collaboration");
const showMarginGuides = query.has("marginGuides");
const marginGuideColor = query.get("marginGuideColor") ?? undefined;
const marginGuideProps = marginGuideColor === undefined ? {} : { marginGuideColor };
const collaborationDocument = collaborationEnabled ? new Y.Doc() : null;
const collaborationAwareness = collaborationDocument ? new Awareness(collaborationDocument) : null;
let collaborationWasSeeded = false;
const remoteCollaborationDocuments: Y.Doc[] = [];
const remoteCollaborationAwareness: Awareness[] = [];
const collaboration: DocxEditorCollaboration | undefined =
  collaborationDocument && collaborationAwareness
    ? {
        awareness: collaborationAwareness,
        onSeeded: () => {
          collaborationWasSeeded = true;
        },
        plugins: [
          ySyncPlugin(collaborationDocument.getXmlFragment("prosemirror")),
          yCursorPlugin(collaborationAwareness),
          yUndoPlugin(),
        ],
        shouldSeed: true,
        yXmlFragment: collaborationDocument.getXmlFragment("prosemirror"),
      }
    : undefined;

collaborationAwareness?.setLocalStateField("user", {
  color: "#2563eb",
  name: "Vue collaborator",
});

onMounted(() => {
  void loadFromQuery();
  globalThis.__folioParity = buildParityBridge(
    () => editorRef.value,
    (kind) => clipboardCallbackCounts[kind],
  );
  if (collaboration) {
    globalThis.__folioVueCollaboration = {
      getSharedText: () => {
        const view = editorRef.value?.getEditor()?.getView();
        if (!view) {
          return "";
        }
        return initProseMirrorDoc(collaboration.yXmlFragment, view.state.schema).doc.textContent;
      },
      showRemoteSelection: () => {
        if (!globalThis.__folioParity?.selectFirstWord() || !collaborationAwareness) {
          return false;
        }
        const localState = collaborationAwareness.getLocalState();
        if (!isRecord(localState) || !isRecord(localState["cursor"])) {
          return false;
        }

        const remoteDocument = new Y.Doc();
        const remoteAwareness = new Awareness(remoteDocument);
        remoteAwareness.setLocalState({
          cursor: localState["cursor"],
          user: { color: "#dc2626", name: "Remote collaborator" },
        });
        applyAwarenessUpdate(
          collaborationAwareness,
          encodeAwarenessUpdate(remoteAwareness, [remoteAwareness.clientID]),
          "parity-test",
        );
        remoteCollaborationDocuments.push(remoteDocument);
        remoteCollaborationAwareness.push(remoteAwareness);
        return true;
      },
      wasSeeded: () => collaborationWasSeeded,
    };
  }
});

onBeforeUnmount(() => {
  globalThis.__folioParity = undefined;
  globalThis.__folioVueCollaboration = undefined;
  collaborationAwareness?.destroy();
  collaborationDocument?.destroy();
  for (const awareness of remoteCollaborationAwareness) {
    awareness.destroy();
  }
  for (const document of remoteCollaborationDocuments) {
    document.destroy();
  }
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
