/**
 * The editor webview's protocol, from `packages/editor-web`: the bundle the
 * webview runs speaks it, and the extension checks every message against it.
 * The module imports nothing from the editor, so this bundles only the checks.
 */

import type { FolioEditorSaveStrategy } from "../../../packages/editor-web/src/host";

export type {
  FolioEditorDocument,
  FolioEditorSaveStrategy,
} from "../../../packages/editor-web/src/host";
export type { FolioEditingMode } from "../../../packages/editor-web/src/modes";
export {
  isEditorMessage,
  isHostMessage,
  type EditorMessage,
  type HostMessage,
} from "../../../packages/editor-web/src/protocol";

/** `folio save --save-strategy` for how the editor serialized. */
export const saveStrategyFlag = (strategy: FolioEditorSaveStrategy): "selective" | "full-repack" =>
  strategy.type === "full-repack" ? "full-repack" : "selective";
