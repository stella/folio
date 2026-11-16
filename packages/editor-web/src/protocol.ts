/**
 * The messages a VS Code custom-editor webview running this bundle exchanges
 * with its extension. Both sides validate what they receive: a webview
 * message crosses a process boundary, so neither trusts the other's shape.
 * The extension imports this module to type and check its side.
 */

import type { SelectiveSaveDeclineReason } from "@stll/folio-core/types/selectiveSaveDecline";

import type { FolioEditorDocument, FolioEditorSaveStrategy } from "./host";
import type { FolioEditingMode } from "./modes";
import { isEditingMode } from "./modes";

/** Extension to webview. */
export type HostMessage =
  /** Open the document, in answer to `ready`. A later `load` replaces it. */
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
  | {
      readonly type: "serializeFailed";
      readonly requestId: number;
      readonly message: string;
    }
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

const DECLINE_REASONS = {
  structuralUnreadableBody: true,
  structuralParagraphIds: true,
  structuralBodyBlock: true,
  structuralDependentContent: true,
  structuralBarrierChanged: true,
  structuralReorder: true,
  structuralFragment: true,
  structuralRangeSplit: true,
  structuralSpliceRefused: true,
  untrackedChanges: true,
  tooLarge: true,
  unmaterializedBodyResource: true,
  unmaterializedStoryResource: true,
  unmaterializedHeaderFooter: true,
  pictureWatermark: true,
  orphanCommentRanges: true,
  unsynthesizedReplyRanges: true,
  invalidModel: true,
  unmaterializedHyperlink: true,
  unsafePackagePath: true,
  missingDocumentPart: true,
  addedNumbering: true,
  unroutedParagraphs: true,
  paragraphPatch: true,
  commentsExtended: true,
  numberingPart: true,
  styleMaterialization: true,
  strictConformance: true,
  packageError: true,
} as const satisfies Record<SelectiveSaveDeclineReason, true>;

const isSaveStrategy = (value: unknown): value is FolioEditorSaveStrategy => {
  if (!isRecord(value)) return false;
  switch (value["type"]) {
    case "selective":
      return true;
    case "full-repack": {
      const decline = value["decline"];
      if (decline === undefined) return true;
      if (
        !isRecord(decline) ||
        !isString(decline["reason"]) ||
        !Object.hasOwn(DECLINE_REASONS, decline["reason"])
      )
        return false;
      switch (decline["scope"]) {
        case "package":
          return (
            decline["part"] === undefined && decline["story"] === undefined
          );
        case "part":
          return (
            isString(decline["part"]) &&
            decline["part"].length > 0 &&
            decline["story"] === undefined
          );
        case "story": {
          const story = decline["story"];
          return (
            decline["part"] === undefined &&
            isRecord(story) &&
            (story["kind"] === "header" || story["kind"] === "footer") &&
            isString(story["rId"]) &&
            story["rId"].length > 0
          );
        }
        default:
          return false;
      }
    }
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
