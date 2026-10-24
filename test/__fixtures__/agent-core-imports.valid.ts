import type { FolioUndefinedReferencePolicy } from "@stll/folio-core/ai-edits/apply";
import type { FolioCommentAnchor } from "@stll/folio-core/ai-edits/read";
import type { FolioDocumentStory } from "@stll/folio-core/server";

export type { FolioDocumentStoryHandle } from "@stll/folio-core/server";
export type { Unrelated } from "@stll/folio-core/ai-edits-extra";
export type { UnrelatedIndexed } from "@stll/folio-core/ai-edits/index-extra";
export type ImportedStory = import("@stll/folio-core/server").FolioDocumentStory;
export const loadServer = () => import("@stll/folio-core/server");
export const requireLeaf = () => require("@stll/folio-core/ai-edits/read");
export type { FolioCommentAnchor, FolioDocumentStory, FolioUndefinedReferencePolicy };
