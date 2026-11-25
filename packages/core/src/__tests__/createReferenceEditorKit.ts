import type { Schema } from "prosemirror-model";

import { createExtensionRuntime } from "../internal/extensionRuntime";
import {
  createStarterKitExtensions,
  type StarterKitOptions,
} from "../internal/starterKitExtensions";
import { createExtension } from "../prosemirror/extensions/create";
import type { AnyExtension } from "../prosemirror/extensions/types";
import { HistoryExtension } from "./reference-editor/HistoryExtension";
import { ParagraphChangeTrackerExtension } from "./reference-editor/ParagraphChangeTrackerExtension";
import { ParaIdAllocatorExtension } from "./reference-editor/ParaIdAllocatorExtension";
import { createReferenceTablePlugins } from "./reference-editor/tablePlugins";

const TableGeometryExtension = createExtension({
  name: "tableGeometry",
  onSchemaReady: () => ({ plugins: createReferenceTablePlugins() }),
});

type ReferenceEditorKitOptions = Omit<
  StarterKitOptions,
  "historyDepth" | "historyNewGroupDelay"
> & {
  historyDepth?: number;
  historyNewGroupDelay?: number;
};

const historyExtension = ({
  historyDepth,
  historyNewGroupDelay,
  historyShortcuts,
}: ReferenceEditorKitOptions) =>
  HistoryExtension({
    ...(historyDepth === undefined ? {} : { depth: historyDepth }),
    ...(historyNewGroupDelay === undefined
      ? {}
      : { newGroupDelay: historyNewGroupDelay }),
    ...(historyShortcuts === undefined ? {} : { shortcuts: historyShortcuts }),
  });

const paragraphExtensions = (disabled: readonly string[] = []) => [
  // Allocate identity before recording the paragraphs touched by the edit.
  ...(!disabled.includes("paraIdAllocator")
    ? [ParaIdAllocatorExtension()]
    : []),
  ...(!disabled.includes("paragraphChangeTracker")
    ? [ParagraphChangeTrackerExtension()]
    : []),
];

/** The full ProseMirror reference runtime for tests of editor behavior. */
export const createReferenceEditorKit = (
  options: ReferenceEditorKitOptions = {},
): AnyExtension[] => {
  const paragraphPlugins = paragraphExtensions(options.disable);
  const extensions: AnyExtension[] = [];
  for (const extension of createStarterKitExtensions(options)) {
    switch (extension.config.name) {
      case "history":
        extensions.push(historyExtension(options));
        break;
      case "tablePlugin":
        extensions.push(extension, TableGeometryExtension());
        break;
      case "selectionTracker":
        extensions.push(extension, ...paragraphPlugins);
        break;
      default:
        extensions.push(extension);
    }
  }
  if (options.disable?.includes("selectionTracker"))
    extensions.push(...paragraphPlugins);
  return extensions;
};

/** Build plugins on the same schema as a projected document. */
export const createReferenceEditorRuntime = (schema: Schema) =>
  createExtensionRuntime(
    createReferenceEditorKit().sort(
      (a, b) => a.config.priority - b.config.priority,
    ),
    schema,
  );
