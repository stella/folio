/**
 * The messages a VS Code custom-editor webview running this bundle exchanges
 * with its extension. Both sides validate what they receive: a webview
 * message crosses a process boundary, so neither trusts the other's shape.
 * The extension imports this module to type and check its side.
 */

import type { FolioEditorDocument, FolioEditorRepackReason, FolioEditorSaveStrategy } from "./host";
import type { FolioEditingMode } from "./modes";
import { isEditingMode } from "./modes";

/** Extension to webview. */
export type HostMessage =
  /** Open the document. Sent once, in answer to `ready`. */
  | {
      readonly type: "load";
      readonly document: FolioEditorDocument;
      readonly author: string;
      readonly mode: FolioEditingMode;
      readonly locale: string;
    }
  /** Replace the document (a revert, or the file changed on disk). */
  | { readonly type: "reload"; readonly document: FolioEditorDocument }
  /** Answer with `serialized` or `serializeFailed` carrying the same `requestId`. */
  | { readonly type: "serialize"; readonly requestId: number }
  | { readonly type: "undo" }
  | { readonly type: "redo" }
  | { readonly type: "setMode"; readonly mode: FolioEditingMode };

/** Webview to extension. */
export type EditorMessage =
  /** The script is listening; send `load`. */
  | { readonly type: "ready" }
  | { readonly type: "loaded"; readonly fileVersion: string }
  | { readonly type: "loadFailed"; readonly message: string }
  /** The user opened a new undo step: add one entry to the undo stack. */
  | { readonly type: "edit" }
  | { readonly type: "dirty"; readonly dirty: boolean }
  /** The user switched modes in the editor's toolbar. */
  | { readonly type: "modeChanged"; readonly mode: FolioEditingMode }
  | {
      readonly type: "serialized";
      readonly requestId: number;
      readonly bytes: Uint8Array;
      /** The `fileVersion` of the loaded document these bytes were edited from. */
      readonly fileVersion: string;
      readonly strategy: FolioEditorSaveStrategy;
    }
  | { readonly type: "serializeFailed"; readonly requestId: number; readonly message: string }
  | { readonly type: "error"; readonly message: string };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isString = (value: unknown): value is string => typeof value === "string";

const isRequestId = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;

const isEditorDocument = (value: unknown): value is FolioEditorDocument =>
  isRecord(value) &&
  value["bytes"] instanceof Uint8Array &&
  isString(value["fileVersion"]) &&
  isString(value["fileName"]);

const REPACK_REASONS = {
  structuralChange: true,
  untrackedChange: true,
  noBodyView: true,
} as const satisfies Record<FolioEditorRepackReason, true>;

const isSaveStrategy = (value: unknown): value is FolioEditorSaveStrategy => {
  if (!isRecord(value)) return false;
  switch (value["type"]) {
    case "selective-first":
      return true;
    case "full-repack":
      return isString(value["reason"]) && Object.hasOwn(REPACK_REASONS, value["reason"]);
    default:
      return false;
  }
};

export const isHostMessage = (value: unknown): value is HostMessage => {
  if (!isRecord(value)) return false;
  switch (value["type"]) {
    case "load":
      return (
        isEditorDocument(value["document"]) &&
        isString(value["author"]) &&
        isEditingMode(value["mode"]) &&
        isString(value["locale"])
      );
    case "reload":
      return isEditorDocument(value["document"]);
    case "serialize":
      return isRequestId(value["requestId"]);
    case "undo":
    case "redo":
      return true;
    case "setMode":
      return isEditingMode(value["mode"]);
    default:
      return false;
  }
};

export const isEditorMessage = (value: unknown): value is EditorMessage => {
  if (!isRecord(value)) return false;
  switch (value["type"]) {
    case "ready":
    case "edit":
      return true;
    case "loaded":
      return isString(value["fileVersion"]);
    case "loadFailed":
    case "error":
      return isString(value["message"]);
    case "dirty":
      return typeof value["dirty"] === "boolean";
    case "modeChanged":
      return isEditingMode(value["mode"]);
    case "serialized":
      return (
        isRequestId(value["requestId"]) &&
        value["bytes"] instanceof Uint8Array &&
        isString(value["fileVersion"]) &&
        isSaveStrategy(value["strategy"])
      );
    case "serializeFailed":
      return isRequestId(value["requestId"]) && isString(value["message"]);
    default:
      return false;
  }
};
