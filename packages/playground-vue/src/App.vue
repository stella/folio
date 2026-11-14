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
        v-bind="optionalEditorProps"
        :mode="editorMode"
        :on-mode-change="(mode: EditorMode) => (editorMode = mode)"
        :show-toolbar="true"
        :show-ruler="true"
        :show-margin-guides="showMarginGuides"
        :initial-zoom="1"
        :collaboration="collaboration"
        :on-copy="() => clipboardCallbackCounts.copy++"
        :on-cut="() => clipboardCallbackCounts.cut++"
        :on-paste="() => clipboardCallbackCounts.paste++"
        :onFindResultChange="(result) => findResultNotifications.push(result)"
        :on-error="(error) => (status = error.message)"
        :on-editor-view-ready="scrollParityHost.onViewReady"
        @ready="scrollParityHost.onReady"
        :preserve-document-while-loading="
          query.has('readyScroll') || query.has('readyEventScroll')
        "
      />
    </main>
    <p v-if="status" class="pg-vue-status">{{ status }}</p>
  </div>
</template>

<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, shallowRef } from "vue";

import {
  createEmptyDocument,
  createStellaStyleDocumentPreset,
  DocxEditor,
} from "@stll/folio-vue";
import type {
  Document as FolioDocument,
  DocxEditorCollaboration,
  DocxEditorRef,
  EditorMode,
  FindResult,
} from "@stll/folio-vue";

import { buildCanonicalBridge } from "../../../tests/parity/canonicalBridge";
import { buildCanonicalRoomBridge } from "../../../tests/parity/canonicalRoomBridge";
import {
  createCanonicalRoomDemo,
  readCanonicalRoomSeed,
} from "../../../tests/parity/canonicalRoomDemo";
import { buildFindNotificationBridge } from "./findNotificationBridge";
import type { FindNotificationBridge } from "./findNotificationBridge";
import type { FolioParityBridge } from "./parityBridge";
import { buildParityBridge } from "./parityBridge";
import { buildScrollParityBridge } from "./scrollParityBridge";
import type { ScrollParityBridge } from "./scrollParityBridge";

declare global {
  // Test hook: browser specs read live editor state through this, as in the
  // React playground.
  // eslint-disable-next-line no-var
  var __folioPlayground:
    | {
        getEditorRef: () => DocxEditorRef | null;
      }
    | undefined;
  // eslint-disable-next-line no-var
  var __folioParity: FolioParityBridge | undefined;
  var __folioScrollParity: ScrollParityBridge | undefined;
  var __folioVueFindNotifications: FindNotificationBridge | undefined;
}

const editorRef = ref<DocxEditorRef | null>(null);
const scrollParityHost = buildScrollParityBridge(() => editorRef.value);
const documentBuffer = shallowRef<ArrayBuffer | null>(null);
const currentDocument = shallowRef<FolioDocument | null>(null);
const status = ref("");
const clipboardCallbackCounts = { copy: 0, cut: 0, paste: 0 };
const findResultNotifications: (FindResult | null)[] = [];
const query = new URLSearchParams(window.location.search);
const collaborationEnabled = query.has("collaboration");
// The canonical session is the editor's default; `?session=legacy` selects the
// ProseMirror-owned engine and `?session=canonical` states the default. The
// collaboration demo runs a canonical room, which only the canonical session
// joins, so it keeps the default whatever `?session=` names.
const requestedSession = query.get("session");
const experimentalSession =
  !collaborationEnabled &&
  (requestedSession === "legacy" || requestedSession === "canonical")
    ? requestedSession
    : undefined;
const editorMode = ref<EditorMode>(
  query.get("mode") === "suggesting" ? "suggesting" : "editing",
);
const showMarginGuides = query.has("marginGuides");
const marginGuideColor = query.get("marginGuideColor") ?? undefined;
const optionalEditorProps = {
  ...(experimentalSession === undefined
    ? {}
    : ({ experimentalSession } as const)),
  ...(marginGuideColor === undefined ? {} : { marginGuideColor }),
};
// `?collaboration=1`: a canonical room shared by this origin's tabs that
// open the same URL hash and fixture, as in the React playground.
const collaboration = shallowRef<DocxEditorCollaboration | undefined>(
  undefined,
);

onMounted(() => {
  void loadFromQuery();
  globalThis.__folioPlayground = { getEditorRef: () => editorRef.value };
  globalThis.__folioCanonical = buildCanonicalBridge(() => editorRef.value);
  globalThis.__folioCanonicalReady = true;
  globalThis.__folioScrollParity = scrollParityHost;
  globalThis.__folioVueFindNotifications = buildFindNotificationBridge(
    () => editorRef.value,
    findResultNotifications,
  );
  globalThis.__folioParity = buildParityBridge(
    () => editorRef.value,
    (kind) => clipboardCallbackCounts[kind],
  );
});

onBeforeUnmount(() => {
  globalThis.__folioPlayground = undefined;
  globalThis.__folioCanonicalReady = false;
  globalThis.__folioCanonical = undefined;
  globalThis.__folioParity = undefined;
  globalThis.__folioScrollParity = undefined;
  globalThis.__folioCanonicalRoom = undefined;
  globalThis.__folioVueFindNotifications = undefined;
});

const roomName = (): string => {
  const hash = window.location.hash.replace(/^#/u, "");
  if (hash) return hash;
  const room = `room-${crypto.randomUUID().slice(0, 8)}`;
  window.history.replaceState(null, "", `${window.location.search}#${room}`);
  return room;
};

async function openRoom(fixtureFile: string): Promise<void> {
  status.value = "Opening the room...";
  const bytes = await readCanonicalRoomSeed(fixtureFile);
  if (bytes === undefined) {
    status.value = `Fixture not found: ${fixtureFile}`;
    return;
  }
  const demo = await createCanonicalRoomDemo({
    room: roomName(),
    bytes,
    user: { name: "Vue collaborator", color: "#2563eb" },
  });
  globalThis.__folioCanonicalRoom = buildCanonicalRoomBridge(
    () => editorRef.value?.getEditor() ?? null,
    demo.status,
  );
  // The room and the bytes arrive together, so the first session joins it.
  collaboration.value = demo.collaboration;
  documentBuffer.value = bytes;
  const showNotice = () => {
    status.value = demo.getState().notice ?? "";
  };
  demo.subscribe(showNotice);
  showNotice();
}

async function loadFromQuery(): Promise<void> {
  const fixtureFile = new URLSearchParams(window.location.search).get("file");
  if (collaborationEnabled) {
    try {
      await openRoom(fixtureFile ?? "sample.docx");
    } catch (error) {
      status.value = error instanceof Error ? error.message : String(error);
    }
    return;
  }
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
