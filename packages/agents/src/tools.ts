import {
  FOLIO_DOCUMENT_OPERATION_TYPES,
  type FolioDocumentOperationType,
} from "@stll/folio-core/server";

import { FOLIO_TEXT_RANGE_JSON_SCHEMA } from "./operation-schema";
import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type { FolioAgentToolDefinition } from "./types";

const storyHandleSchema = {
  type: "object",
  properties: {
    type: { type: "string", enum: ["main", "header", "footer", "footnote", "endnote"] },
    relationshipId: { type: "string" },
    noteId: { type: "integer" },
  },
  required: ["type"],
  additionalProperties: false,
} as const;

const sectionHandleSchema = {
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
} as const;

const scopedHandleSchema = {
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
} as const;

/**
 * `suggest_changes` deliberately narrows the full document-operation contract
 * (see `FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA` in `operation-schema.ts`):
 * - excluded types: `formatRange`, `insertSignatureTable`, `insertTableRow`, and
 *   `deleteTableRow` (direct-only, not representable as tracked changes for human review) and
 *   `commentOnBlock` (covered by the dedicated `add_comment` tool);
 * - `id` is optional here (auto-generated `op-1`, `op-2`, … by `parse.ts`)
 *   where the contract requires it;
 * - `comment` is a plain string here; `parse.ts` wraps it into the contract's
 *   `{ text }` object;
 * - the review metadata (`severity`, `area`), `precondition` guard, and the
 *   insert/replace formatting knobs (`inheritFormatting`, `pageBreakBefore`,
 *   `preserveFormatting`, `styleId`, `position`, `parties`, `quote`,
 *   `formatting`) are not exposed to the model.
 */
const SUGGEST_CHANGES_EXCLUDED_OPERATION_TYPES: ReadonlySet<FolioDocumentOperationType> = new Set([
  "formatRange",
  "commentOnBlock",
  "insertSignatureTable",
  "insertTableRow",
  "deleteTableRow",
]);

/**
 * The contract operation types `suggest_changes` exposes to the model,
 * derived from the shared contract constant minus the exclusions above.
 * Exported (module-level only, not from the package index) so
 * `parse.test.ts` can assert this enum stays in lockstep with what
 * `parseSuggestChangesInput` actually accepts.
 */
export const SUGGEST_CHANGES_OPERATION_TYPES: readonly FolioDocumentOperationType[] =
  FOLIO_DOCUMENT_OPERATION_TYPES.filter(
    (type) => !SUGGEST_CHANGES_EXCLUDED_OPERATION_TYPES.has(type),
  );

const suggestChangesOperationSchema = {
  type: "object",
  properties: {
    id: {
      type: "string",
      description:
        "Optional caller-supplied operation id, echoed back in `applied`/`skipped`. Auto-generated (op-1, op-2, …) when omitted.",
    },
    type: {
      type: "string",
      enum: SUGGEST_CHANGES_OPERATION_TYPES,
      description: "The kind of edit to make.",
    },
    blockId: {
      type: "string",
      description: "The block to edit, from `read_document` or `find_text`.",
    },
    range: {
      ...FOLIO_TEXT_RANGE_JSON_SCHEMA,
      description:
        "Required for `replaceRange` and `commentOnRange`: copy the range object returned by `find_text`.",
    },
    find: {
      type: "string",
      description:
        "Required for `replaceInBlock`: the exact text to find within the block, up to 100,000 characters.",
    },
    replace: {
      type: "string",
      description:
        "Required for `replaceInBlock` and `replaceRange`: replacement text, up to 100,000 characters.",
    },
    text: {
      type: "string",
      description:
        "Required for `insertAfterBlock` / `insertBeforeBlock` / `replaceBlock`: the text to insert or replace the block with, up to 100,000 characters.",
    },
    comment: {
      type: "string",
      description:
        "Optional comment explaining this edit, attached to the affected text, up to 100,000 characters.",
    },
  },
  required: ["type"],
  additionalProperties: false,
} as const;

/**
 * The tools this package exposes, described for an LLM. Every tool that reads
 * or mutates the document expects `blockId` values that came from
 * `read_document` or `find_text` in THIS conversation — block ids are not
 * guessable and change whenever the document's structure changes. Every
 * mutation (`add_comment`, `suggest_changes`) becomes a tracked change or
 * comment pending human review; nothing is silently finalized.
 */
export const FOLIO_AGENT_TOOLS: FolioAgentToolDefinition[] = [
  {
    name: FOLIO_AGENT_TOOL_NAMES.readDocument,
    description:
      "Read the full document body as a list of blocks (paragraphs, headings, list items). Call this first, " +
      "or whenever you need fresh block ids after a mutation — block ids from a stale read may no longer resolve.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.getDocumentOutline,
    description:
      "Read a lightweight heading outline before opening document content. Returns stable section handles, " +
      "heading hierarchy, and real rendered page numbers when a live paginated editor is available.",
    inputSchema: {
      type: "object",
      properties: {
        maxDepth: {
          type: "integer",
          minimum: 1,
          maximum: 9,
          description: "Deepest heading level to return. Defaults to 3.",
        },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.readSection,
    description:
      "Read one logical heading section using a handle from get_document_outline. Content is block-bounded " +
      "and paginated with an afterBlockId cursor, avoiding a full-document read.",
    inputSchema: {
      type: "object",
      properties: {
        handle: sectionHandleSchema,
        maxBlocks: {
          type: "integer",
          minimum: 1,
          maximum: 200,
          description: "Maximum blocks to return. Defaults to 100.",
        },
        afterBlockId: {
          type: "string",
          description: "Continue after this block id from the preceding read_section response.",
        },
      },
      required: ["handle"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.listStories,
    description: "List readable document stories and their typed handles.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.readStory,
    description: "Read one document story using a handle returned by `list_stories`.",
    inputSchema: {
      type: "object",
      properties: {
        handle: storyHandleSchema,
      },
      required: ["handle"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.findText,
    description:
      "Search a document, section, rendered page, or story and return `{ matches, truncated, totalMatches }`. " +
      "Main-story matches include a stable block and exact range; other stories return story-relative offsets. " +
      "Every match includes surrounding context. `matches` is " +
      "capped at 200 entries; `truncated` is true and `totalMatches` reports the real count when there were " +
      "more — narrow the query or scope instead of assuming you saw every hit.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Non-empty text to search for, up to 1,000 characters.",
        },
        matchCase: {
          type: "boolean",
          description: "Case-sensitive match. Defaults to false (case-insensitive).",
        },
        wholeWord: {
          type: "boolean",
          description: "Match only Unicode word boundaries. Defaults to false.",
        },
        scope: {
          type: "object",
          properties: {
            type: { type: "string", enum: ["document", "section", "page", "story"] },
            handle: scopedHandleSchema,
            page: { type: "integer", minimum: 1 },
          },
          required: ["type"],
          additionalProperties: false,
          description:
            "Limit search to the main document, a section, a real rendered page, or a story.",
        },
      },
      required: ["query"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.readComments,
    description:
      "Read comment threads in the document, each with its author, text, resolved status, anchored block, and " +
      'replies. Filter to unresolved ("open") comments to see what still needs attention.',
    inputSchema: {
      type: "object",
      properties: {
        filter: {
          type: "string",
          enum: ["all", "open", "resolved"],
          description: 'Which comments to return. Defaults to "all".',
        },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.readChanges,
    description:
      "Read pending tracked changes (insertions and deletions) awaiting human review. Use this to see the effect " +
      "of edits already suggested via `suggest_changes` before proposing more.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.addComment,
    description:
      "Attach a comment to a block, optionally quoting the specific text it is about. The comment is added " +
      "immediately (comments are not tracked changes) but the underlying text is left untouched — use this for " +
      "notes/questions, and `suggest_changes` for edits. `text` and `quote` are each capped at 100,000 characters.",
    inputSchema: {
      type: "object",
      properties: {
        blockId: {
          type: "string",
          description: "The block to comment on, from `read_document` or `find_text`.",
        },
        quote: {
          type: "string",
          description:
            "Optional exact text within the block this comment is about, up to 100,000 characters.",
        },
        text: { type: "string", description: "The comment body, up to 100,000 characters." },
      },
      required: ["blockId", "text"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.suggestChanges,
    description:
      "Propose one or more edits as tracked changes for a human to accept or reject — nothing is applied " +
      "directly to the visible text. Each operation needs a blockId from `read_document` or `find_text`; if the " +
      "document changed since that read, re-read it and retry with fresh ids (a skip reason will say so). At " +
      "most 50 operations per call — batch larger edits across multiple calls. Each `find` / `replace` / " +
      "`text` / `comment` string is capped at 100,000 characters.",
    inputSchema: {
      type: "object",
      properties: {
        operations: {
          type: "array",
          description: "The edits to propose, applied in order. At most 50 per call.",
          items: suggestChangesOperationSchema,
        },
      },
      required: ["operations"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.replyComment,
    description:
      "Reply to an existing comment thread, referenced by the id from `read_comments`. `text` is capped at " +
      "100,000 characters.",
    inputSchema: {
      type: "object",
      properties: {
        commentId: { type: "string", description: "The comment id from `read_comments`." },
        text: { type: "string", description: "The reply body, up to 100,000 characters." },
      },
      required: ["commentId", "text"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.resolveComment,
    description:
      "Mark a comment thread resolved, or pass `reopen: true` to reopen a previously resolved one.",
    inputSchema: {
      type: "object",
      properties: {
        commentId: { type: "string", description: "The comment id from `read_comments`." },
        reopen: {
          type: "boolean",
          description: "Reopen an already-resolved thread instead of resolving it.",
        },
      },
      required: ["commentId"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.readPage,
    description:
      "Read the plain text of one page (1-based) as currently paginated in the live editor. Only available when " +
      "the document is open in a live, paginated editor surface — not on a headless document.",
    inputSchema: {
      type: "object",
      properties: {
        page: { type: "number", description: "1-based page number." },
      },
      required: ["page"],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.readSelection,
    description:
      "Read the user's current text selection in the live editor, as plain text. Only available on a live editor " +
      "surface with an active selection.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.showInDocument,
    description:
      "Reveal and select a block or exact text range in the live editor. Pass either blockId or the range " +
      "returned by find_text; stale or missing targets return false.",
    inputSchema: {
      type: "object",
      properties: {
        blockId: { type: "string", description: "A main-story block id." },
        range: FOLIO_TEXT_RANGE_JSON_SCHEMA,
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: FOLIO_AGENT_TOOL_NAMES.scrollToBlock,
    description:
      "Scroll the live editor to the given block and select it, so the user can see what you are discussing. " +
      "Only available on a live editor surface.",
    inputSchema: {
      type: "object",
      properties: {
        blockId: {
          type: "string",
          description: "The block to scroll to, from `read_document` or `find_text`.",
        },
      },
      required: ["blockId"],
      additionalProperties: false,
    },
  },
];

/** The tool definitions this package exposes. Same array as {@link FOLIO_AGENT_TOOLS}, as a function for symmetry with the other providers. */
export const getFolioToolDefinitions = (): FolioAgentToolDefinition[] => FOLIO_AGENT_TOOLS;
