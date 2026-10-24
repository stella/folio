import { hashFolioAIBlockText } from "@stll/folio-core/ai-edits";
import type { FolioDocumentStory } from "@stll/folio-core/ai-edits";

export { getFolioDocumentOutline } from "@stll/folio-core/ai-edits";
export type { FolioDocumentStoryHandle } from "@stll/folio-core/ai-edits";
export type ImportedStory = import("@stll/folio-core/ai-edits").FolioDocumentStory;
export type FacadeModule = typeof import("@stll/folio-core/ai-edits");
export * from "@stll/folio-core/ai-edits";
export const loadFacade = () => import("@stll/folio-core/ai-edits");
import * as facade from "@stll/folio-core/ai-edits";
import type * as FacadeTypes from "@stll/folio-core/ai-edits";
export type IndexedStory = import("@stll/folio-core/ai-edits/index").FolioDocumentStory;
export type { FolioDocumentStoryHandle as IndexedStoryHandle } from "@stll/folio-core/ai-edits/index";
export const loadIndexedFacade = () => import("@stll/folio-core/ai-edits/index");
export const requireFacade = () => require("@stll/folio-core/ai-edits");
export const requireIndexedFacade = () => require("@stll/folio-core/ai-edits/index");

export { facade, hashFolioAIBlockText };
export type { FacadeTypes, FolioDocumentStory };
