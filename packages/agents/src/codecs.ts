import type {
  FolioAITextRangeHandle,
  FolioDocumentSectionHandle,
  FolioDocumentStoryHandle,
} from "@stll/folio-core/server";

import { FOLIO_TEXT_RANGE_JSON_SCHEMA } from "./operation-schema";
import type { FolioAgentJsonObjectSchema } from "./types";

const NORMALIZED_TEXT_HASH_PATTERN = /^h[0-9a-z]+$/;

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isNonEmptyString = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0;

export const FOLIO_STORY_HANDLE_JSON_SCHEMA = {
  type: "object",
  properties: {
    type: { type: "string", enum: ["main", "header", "footer", "footnote", "endnote"] },
    relationshipId: { type: "string" },
    noteId: { type: "integer" },
  },
  required: ["type"],
  additionalProperties: false,
} as const satisfies FolioAgentJsonObjectSchema;

export const FOLIO_SECTION_HANDLE_JSON_SCHEMA = {
  type: "object",
  properties: {
    type: { type: "string", enum: ["headingSection"] },
    story: { type: "string", enum: ["main"] },
    headingBlockId: { type: "string" },
    headingTextHash: { type: "string" },
    headingLevel: { type: "integer", minimum: 1, maximum: 9 },
  },
  required: ["type", "story", "headingBlockId", "headingTextHash", "headingLevel"],
  additionalProperties: false,
} as const satisfies FolioAgentJsonObjectSchema;

export const FOLIO_SCOPED_HANDLE_JSON_SCHEMA = {
  type: "object",
  properties: {
    type: {
      type: "string",
      enum: ["main", "header", "footer", "footnote", "endnote", "headingSection"],
    },
    story: { type: "string", enum: ["main"] },
    relationshipId: { type: "string" },
    noteId: { type: "integer" },
    headingBlockId: { type: "string" },
    headingTextHash: { type: "string" },
    headingLevel: { type: "integer", minimum: 1, maximum: 9 },
  },
  required: ["type"],
  additionalProperties: false,
} as const satisfies FolioAgentJsonObjectSchema;

export const FOLIO_COMMENT_ID_JSON_SCHEMA = {
  type: "string",
  description: "The comment id string from `read_comments`.",
} as const;

export type FolioAgentCommentId = {
  raw: string;
  numeric: number;
};

export const decodeSectionHandle = (value: unknown): FolioDocumentSectionHandle | null => {
  if (!isPlainObject(value)) {
    return null;
  }
  const headingBlockId = value["headingBlockId"];
  const headingTextHash = value["headingTextHash"];
  const headingLevel = value["headingLevel"];
  if (
    value["type"] !== "headingSection" ||
    value["story"] !== "main" ||
    !isNonEmptyString(headingBlockId) ||
    !isNonEmptyString(headingTextHash) ||
    !NORMALIZED_TEXT_HASH_PATTERN.test(headingTextHash) ||
    typeof headingLevel !== "number" ||
    !Number.isInteger(headingLevel) ||
    headingLevel < 1 ||
    headingLevel > 9
  ) {
    return null;
  }
  return {
    type: "headingSection",
    story: "main",
    headingBlockId,
    headingTextHash,
    headingLevel,
  };
};

export const decodeStoryHandle = (value: unknown): FolioDocumentStoryHandle | null => {
  if (!isPlainObject(value)) {
    return null;
  }
  const type = value["type"];
  if (type === "main") {
    return { type };
  }
  if (type === "header" || type === "footer") {
    const relationshipId = value["relationshipId"];
    if (!isNonEmptyString(relationshipId)) {
      return null;
    }
    return { type, relationshipId };
  }
  if (type === "footnote" || type === "endnote") {
    const noteId = value["noteId"];
    if (typeof noteId !== "number" || !Number.isInteger(noteId)) {
      return null;
    }
    return { type, noteId };
  }
  return null;
};

export const decodeMainStoryTextRangeHandle = (value: unknown): FolioAITextRangeHandle | null => {
  if (!isPlainObject(value)) {
    return null;
  }
  const blockId = value["blockId"];
  const startOffset = value["startOffset"];
  const endOffset = value["endOffset"];
  const selectedTextHash = value["selectedTextHash"];
  if (
    value["type"] !== "textRange" ||
    value["story"] !== "main" ||
    !isNonEmptyString(blockId) ||
    typeof startOffset !== "number" ||
    !Number.isInteger(startOffset) ||
    startOffset < 0 ||
    typeof endOffset !== "number" ||
    !Number.isInteger(endOffset) ||
    endOffset <= startOffset ||
    !isNonEmptyString(selectedTextHash) ||
    !NORMALIZED_TEXT_HASH_PATTERN.test(selectedTextHash)
  ) {
    return null;
  }
  return {
    type: "textRange",
    story: "main",
    blockId,
    startOffset,
    endOffset,
    selectedTextHash,
  };
};

export const decodeCommentId = (value: unknown): FolioAgentCommentId | null => {
  if (typeof value !== "string" || !/^\d+$/u.test(value)) {
    return null;
  }
  const numeric = Number(value);
  if (!Number.isSafeInteger(numeric)) {
    return null;
  }
  return { raw: value, numeric };
};

export { FOLIO_TEXT_RANGE_JSON_SCHEMA };
