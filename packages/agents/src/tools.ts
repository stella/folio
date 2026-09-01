import {
  FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE,
  type FolioDocumentOperationType,
} from "@stll/folio-core/server";

import {
  FOLIO_COMMENT_ID_JSON_SCHEMA,
  FOLIO_SCOPED_HANDLE_JSON_SCHEMA,
  FOLIO_SECTION_HANDLE_JSON_SCHEMA,
  FOLIO_STORY_HANDLE_JSON_SCHEMA,
  FOLIO_TEXT_RANGE_JSON_SCHEMA,
} from "./codecs";
import { FOLIO_PRECONDITION_JSON_SCHEMA } from "./operation-schema";
import {
  DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES,
  resolveSuggestChangesOptions,
  type FolioAgentToolOptions,
  type FolioSuggestChangesOptions,
  type ResolvedFolioSuggestChangesOptions,
} from "./suggest-changes-options";
import { defineFolioAgentToolDefinition } from "./tool-contract";
import type { FolioAgentToolInputByName, FolioAgentToolOutputByName } from "./tool-contract";
import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type {
  FolioAgentJsonObjectSchema,
  FolioAgentToolDefinition,
  FolioAgentToolName,
  FolioAgentTypedToolDefinition,
} from "./types";

/**
 * `suggest_changes` is a host-configurable projection of the document-operation
 * contract (see `FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA` in `operation-schema.ts`):
 * the allowed operation types, the review-metadata policy, the per-call cap,
 * and an optional document-version pin all come from
 * {@link FolioSuggestChangesOptions}. The schema is derived from the resolved
 * options and from core's per-type key map, so the properties it advertises
 * are exactly the ones the contract parser accepts for the allowed types.
 * Two model-facing conveniences differ from the contract on the wire:
 * `id` is optional (the parser mints unique ids) and `comment` may be a plain
 * string (the parser wraps it into `{ text }`). `suggestionId` is host-side
 * grouping state and is never exposed to the model.
 */
export const SUGGEST_CHANGES_OPERATION_TYPES = DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES;

/** One-line meaning of each contract operation type, for schemas and prompts. */
const OPERATION_TYPE_SUMMARIES = {
  replaceInBlock: "replace an exact text match inside one block (`find` -> `replace`)",
  replaceRange: "replace the text covered by a `range` copied from find_text",
  commentOnRange: "attach a comment to a `range` copied from find_text",
  formatRange: "toggle bold, italic, or underline on a `range` copied from find_text",
  insertAfterBlock: "insert a new paragraph after a block",
  insertBeforeBlock: "insert a new paragraph before a block",
  replaceBlock: "replace one block's entire text",
  deleteBlock: "delete one block",
  commentOnBlock: "attach a comment to one block, optionally quoting text within it",
  insertSignatureTable: "insert a side-by-side signature table for the given `parties`",
  insertTableRow: "insert a table row next to the row containing a cell block",
  deleteTableRow: "delete the table row containing a cell block",
  insertTableColumn: "insert a table column next to the column containing a cell block",
  deleteTableColumn: "delete the table column containing a cell block",
  mergeTableCells:
    "merge table cells from a cell block to `endBlockId`, or `rowCount` rows downward",
  splitTableCell: "split a previously merged table cell",
} as const satisfies Record<FolioDocumentOperationType, string>;

/**
 * Model-facing schema for every operation property the contract knows,
 * keyed by wire name. A `suggest_changes` schema includes a property only
 * when at least one allowed operation type accepts it (per
 * `FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE`), so a three-type host surface
 * advertises three types' worth of fields.
 */
const OPERATION_PROPERTY_SCHEMAS = {
  id: {
    type: "string",
    description:
      "Optional caller-supplied operation id, echoed back in `applied` / `queued` / `skipped`. Generated when omitted.",
  },
  blockId: {
    type: "string",
    description: "The block to edit, from `read_document` or `find_text`.",
  },
  severity: {
    type: "string",
    enum: ["low", "medium", "high"],
    description: "Review severity of this edit, used to sort a review queue.",
  },
  area: {
    type: "string",
    description: 'Short review area label (e.g. "Payment terms"), used to group a review queue.',
  },
  precondition: FOLIO_PRECONDITION_JSON_SCHEMA,
  endBlockId: {
    type: "string",
    minLength: 1,
    description: "For cell merging, a block in the opposite corner cell.",
  },
  rowCount: {
    type: "integer",
    minimum: 2,
    description: "For vertical cell merging, the number of grid rows to merge downward.",
  },
  range: {
    ...FOLIO_TEXT_RANGE_JSON_SCHEMA,
    description:
      "Required for `replaceRange`, `commentOnRange`, and `formatRange`: copy the range object returned by `find_text`.",
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
  quote: {
    type: "string",
    description:
      "For `commentOnBlock`: optional exact text within the block the comment is about, up to 100,000 characters.",
  },
  styleId: {
    type: "string",
    description:
      "For inserts and `replaceBlock`: paragraph style id to apply (e.g. a clause-heading style from the document).",
  },
  pageBreakBefore: {
    type: "boolean",
    description: "For inserts: start the inserted paragraph on a new page.",
  },
  inheritFormatting: {
    type: "boolean",
    description: "For inserts: inherit the anchor block's formatting for the inserted paragraph.",
  },
  preserveFormatting: {
    type: "boolean",
    description:
      "For `replaceBlock`: keep the block's existing formatting for the replacement text.",
  },
  position: {
    type: "string",
    enum: ["after", "before"],
    description:
      "For row, column, or signature-table insertion, place the new structure after the anchor (default) or before it.",
  },
  parties: {
    type: "array",
    description: "For `insertSignatureTable`: the signing parties, one table cell per party.",
    items: {
      type: "object",
      properties: {
        name: { type: "string", description: "Party name (rendered bold)." },
        signatory: { type: "string", description: "Name of the person signing." },
        title: { type: "string", description: "Signatory title (rendered in italics)." },
      },
      required: ["name"],
      additionalProperties: false,
    },
  },
  cellTexts: {
    type: "array",
    description:
      "For row or column insertion, initial text for new physical cells in source order, up to 100,000 characters per cell.",
    maxItems: 256,
    items: { type: "string" },
  },
  formatting: {
    type: "object",
    description:
      "Required for `formatRange`: set one or more inline properties to enable or disable.",
    properties: {
      bold: { type: "boolean" },
      italic: { type: "boolean" },
      underline: { type: "boolean" },
    },
    minProperties: 1,
    additionalProperties: false,
  },
  comment: {
    type: "string",
    description:
      "Optional comment explaining this edit, attached to the affected text, up to 100,000 characters. Required for comment operations.",
  },
} as const;

/** Wire keys the model never sees (`type` is emitted separately with its enum). */
const HIDDEN_OPERATION_KEYS: ReadonlySet<string> = new Set(["suggestionId", "type"]);

const operationPropertyKeys = (types: readonly FolioDocumentOperationType[]): Set<string> => {
  const keys = new Set<string>();
  for (const type of types) {
    for (const key of FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE[type]) {
      if (!HIDDEN_OPERATION_KEYS.has(key)) {
        keys.add(key);
      }
    }
  }
  return keys;
};

const summarizeOperationTypes = (types: readonly FolioDocumentOperationType[]): string =>
  types.map((type) => `${type} (${OPERATION_TYPE_SUMMARIES[type]})`).join("; ");

const buildSuggestChangesOperationSchema = (
  resolved: ResolvedFolioSuggestChangesOptions,
): FolioAgentJsonObjectSchema => {
  const keys = operationPropertyKeys(resolved.operationTypes);
  const properties: Record<string, unknown> = {
    type: {
      type: "string",
      enum: resolved.operationTypes,
      description: `The kind of edit: ${summarizeOperationTypes(resolved.operationTypes)}.`,
    },
  };
  for (const [key, schema] of Object.entries(OPERATION_PROPERTY_SCHEMAS)) {
    if (keys.has(key)) {
      properties[key] = schema;
    }
  }
  return {
    type: "object",
    properties,
    required: resolved.reviewMeta === "required" ? ["type", "severity", "area"] : ["type"],
    additionalProperties: false,
  };
};

const hasAnyType = (
  resolved: ResolvedFolioSuggestChangesOptions,
  types: readonly FolioDocumentOperationType[],
): boolean => types.some((type) => resolved.operationTypes.includes(type));

/**
 * Plain-language capability statement for a configured `suggest_changes`
 * surface: the allowed operations, the structural knobs they unlock, the
 * limits, and the review-metadata and document-version requirements. Used
 * inside the tool description and exported so a host can paste the same
 * text into its system prompt instead of hand-maintaining a copy that drifts.
 */
export const describeSuggestChangesCapabilities = (options?: FolioSuggestChangesOptions): string =>
  describeResolvedCapabilities(resolveSuggestChangesOptions(options));

const describeResolvedCapabilities = (resolved: ResolvedFolioSuggestChangesOptions): string => {
  const lines = [`Supported operations: ${summarizeOperationTypes(resolved.operationTypes)}.`];
  const styleIdTargets = [
    ...(hasAnyType(resolved, ["insertAfterBlock", "insertBeforeBlock"]) ? ["an insert"] : []),
    ...(hasAnyType(resolved, ["replaceBlock"]) ? ["replaceBlock"] : []),
  ];
  if (styleIdTargets.length > 0) {
    const pageBreak = hasAnyType(resolved, ["insertAfterBlock", "insertBeforeBlock"])
      ? "set `pageBreakBefore: true` on an insert to start it on a new page; "
      : "";
    lines.push(
      `Structural edits: ${pageBreak}set \`styleId\` on ${styleIdTargets.join(" or ")} to apply a paragraph style such as a clause heading. Never emit directive markers or markdown syntax as paragraph text.`,
    );
  }
  if (hasAnyType(resolved, ["insertSignatureTable"])) {
    lines.push(
      "Signature blocks: use insertSignatureTable with `parties`; never draw one out of paragraphs.",
    );
  }
  if (!hasAnyType(resolved, ["formatRange"])) {
    lines.push(
      "This surface cannot change run formatting (fonts, bold/italic/underline, size, colour, alignment, spacing, list style); do not promise formatting changes.",
    );
  }
  lines.push(
    `Limits: at most ${resolved.maxOperations} operations per call (batch larger edits across calls); each text field at most 100,000 characters.`,
  );
  lines.push(
    resolved.reviewMeta === "required"
      ? "Every operation must set `severity` (low, medium, or high) and a short `area` label; the review queue sorts and groups by them."
      : "Set `severity` and `area` on operations when the edit is part of a structured review.",
  );
  if (resolved.documentVersion !== null) {
    lines.push(
      "Pass `documentVersion` exactly as given in this tool's schema; the batch is skipped as a whole when the document has moved on since.",
    );
  }
  return lines.join("\n");
};

const SUGGEST_CHANGES_BASE_DESCRIPTION =
  "Propose one or more edits as tracked changes for a human to accept or reject — nothing is applied " +
  "directly to the visible text. Each operation needs a blockId from `read_document` or `find_text`; if the " +
  "document changed since that read, re-read it and retry with fresh ids (a skip reason will say so). Pass " +
  "the `blockTextHash` from that read as `precondition.blockTextHash` to guard against the document " +
  "changing between the read and this call.";

const buildSuggestChangesToolDefinition = (
  resolved: ResolvedFolioSuggestChangesOptions,
): FolioAgentTypedToolDefinition<
  typeof FOLIO_AGENT_TOOL_NAMES.suggestChanges,
  FolioAgentToolInputByName[typeof FOLIO_AGENT_TOOL_NAMES.suggestChanges],
  FolioAgentToolOutputByName[typeof FOLIO_AGENT_TOOL_NAMES.suggestChanges]
> =>
  defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.suggestChanges,
    description: `${SUGGEST_CHANGES_BASE_DESCRIPTION}\n${describeResolvedCapabilities(resolved)}`,
    inputSchema: {
      type: "object",
      properties: {
        ...(resolved.documentVersion !== null && {
          documentVersion: {
            type: "string",
            enum: [resolved.documentVersion.current],
            description:
              "The current document version shown here. Copy it exactly; the batch is skipped if the document changes before it applies.",
          },
        }),
        operations: {
          type: "array",
          description: `The edits to propose, applied in order. At most ${resolved.maxOperations} per call.`,
          minItems: 1,
          maxItems: resolved.maxOperations,
          items: buildSuggestChangesOperationSchema(resolved),
        },
      },
      required:
        resolved.documentVersion !== null ? ["documentVersion", "operations"] : ["operations"],
      additionalProperties: false,
    },
  });

type FolioAgentToolRegistry = {
  [Name in FolioAgentToolName]: FolioAgentTypedToolDefinition<
    Name,
    FolioAgentToolInputByName[Name],
    FolioAgentToolOutputByName[Name]
  >;
};

const definitionsFromRegistry = (registry: FolioAgentToolRegistry): FolioAgentToolDefinition[] =>
  Object.values(registry);

/**
 * The tools this package exposes, described for an LLM. Every tool that reads
 * or mutates the document expects `blockId` values that came from
 * `read_document` or `find_text` in THIS conversation — block ids are not
 * guessable and change whenever the document's structure changes. Every
 * mutation (`add_comment`, `suggest_changes`) becomes a tracked change or
 * comment pending human review; nothing is silently finalized.
 */
export const FOLIO_AGENT_TOOL_REGISTRY = {
  [FOLIO_AGENT_TOOL_NAMES.readDocument]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.readDocument,
    description:
      "Read the full document body as a list of blocks (paragraphs, headings, list items). Call this first, " +
      "or whenever you need fresh block ids after a mutation — block ids from a stale read may no longer " +
      "resolve. Each block includes a `blockTextHash`; echo it as `precondition.blockTextHash` on a " +
      "suggest_changes / add_comment operation to guard against the block changing before that call runs.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  }),
  [FOLIO_AGENT_TOOL_NAMES.getDocumentOutline]: defineFolioAgentToolDefinition({
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
  }),
  [FOLIO_AGENT_TOOL_NAMES.readSection]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.readSection,
    description:
      "Read one logical heading section using a handle from get_document_outline. Content is block-bounded " +
      "and paginated with an afterBlockId cursor, avoiding a full-document read. Each block includes a " +
      "`blockTextHash`; echo it as `precondition.blockTextHash` on a suggest_changes / add_comment operation.",
    inputSchema: {
      type: "object",
      properties: {
        handle: FOLIO_SECTION_HANDLE_JSON_SCHEMA,
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
  }),
  [FOLIO_AGENT_TOOL_NAMES.listStories]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.listStories,
    description: "List readable document stories and their typed handles.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  }),
  [FOLIO_AGENT_TOOL_NAMES.readStory]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.readStory,
    description: "Read one document story using a handle returned by `list_stories`.",
    inputSchema: {
      type: "object",
      properties: {
        handle: FOLIO_STORY_HANDLE_JSON_SCHEMA,
      },
      required: ["handle"],
      additionalProperties: false,
    },
  }),
  [FOLIO_AGENT_TOOL_NAMES.findText]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.findText,
    description:
      "Search a document, section, rendered page, or story and return `{ matches, truncated, totalMatches }`. " +
      "Main-story matches include a stable block, exact range, and `blockTextHash` (echo it as " +
      "`precondition.blockTextHash` on a later suggest_changes / add_comment operation); other stories return " +
      "story-relative offsets. Every match includes surrounding context. `matches` is " +
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
            handle: FOLIO_SCOPED_HANDLE_JSON_SCHEMA,
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
  }),
  [FOLIO_AGENT_TOOL_NAMES.readComments]: defineFolioAgentToolDefinition({
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
  }),
  [FOLIO_AGENT_TOOL_NAMES.readChanges]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.readChanges,
    description:
      "Read pending tracked changes awaiting human review. Use this to see the effect " +
      "of edits already suggested via `suggest_changes` before proposing more.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  }),
  [FOLIO_AGENT_TOOL_NAMES.addComment]: defineFolioAgentToolDefinition({
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
        precondition: FOLIO_PRECONDITION_JSON_SCHEMA,
      },
      required: ["blockId", "text"],
      additionalProperties: false,
    },
  }),
  [FOLIO_AGENT_TOOL_NAMES.suggestChanges]: buildSuggestChangesToolDefinition(
    resolveSuggestChangesOptions(),
  ),
  [FOLIO_AGENT_TOOL_NAMES.replyComment]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.replyComment,
    description:
      "Reply to an existing comment thread, referenced by the id from `read_comments`. `text` is capped at " +
      "100,000 characters.",
    inputSchema: {
      type: "object",
      properties: {
        commentId: FOLIO_COMMENT_ID_JSON_SCHEMA,
        text: { type: "string", description: "The reply body, up to 100,000 characters." },
      },
      required: ["commentId", "text"],
      additionalProperties: false,
    },
  }),
  [FOLIO_AGENT_TOOL_NAMES.resolveComment]: defineFolioAgentToolDefinition({
    name: FOLIO_AGENT_TOOL_NAMES.resolveComment,
    description:
      "Mark a comment thread resolved, or pass `reopen: true` to reopen a previously resolved one.",
    inputSchema: {
      type: "object",
      properties: {
        commentId: FOLIO_COMMENT_ID_JSON_SCHEMA,
        reopen: {
          type: "boolean",
          description: "Reopen an already-resolved thread instead of resolving it.",
        },
      },
      required: ["commentId"],
      additionalProperties: false,
    },
  }),
  [FOLIO_AGENT_TOOL_NAMES.readPage]: defineFolioAgentToolDefinition({
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
  }),
  [FOLIO_AGENT_TOOL_NAMES.readSelection]: defineFolioAgentToolDefinition({
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
  }),
  [FOLIO_AGENT_TOOL_NAMES.showInDocument]: defineFolioAgentToolDefinition({
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
  }),
  [FOLIO_AGENT_TOOL_NAMES.scrollToBlock]: defineFolioAgentToolDefinition({
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
  }),
} as const satisfies FolioAgentToolRegistry;

export const FOLIO_AGENT_TOOLS = definitionsFromRegistry(FOLIO_AGENT_TOOL_REGISTRY);

/**
 * The tool definitions this package exposes. Without options this is
 * {@link FOLIO_AGENT_TOOLS}; with `suggestChanges` options the
 * `suggest_changes` definition is rebuilt for that surface. Pass the same
 * options to `executeFolioToolCall` so the parser enforces what the schema
 * advertises.
 */
export const getFolioToolDefinitions = (
  options: FolioAgentToolOptions = {},
): FolioAgentToolDefinition[] => {
  if (options.suggestChanges === undefined) {
    return FOLIO_AGENT_TOOLS;
  }
  const suggestChanges = buildSuggestChangesToolDefinition(
    resolveSuggestChangesOptions(options.suggestChanges),
  );
  return FOLIO_AGENT_TOOLS.map((definition) =>
    definition.name === FOLIO_AGENT_TOOL_NAMES.suggestChanges ? suggestChanges : definition,
  );
};
