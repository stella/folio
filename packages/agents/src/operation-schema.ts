/**
 * Reusable projections of folio's versioned document-operation contract for
 * tool builders: a JSON Schema for one operation, a JSON Schema for the
 * versioned batch envelope, and a Standard Schema V1 wrapper around
 * `parseFolioDocumentOperationBatch` so downstream tool definitions (e.g.
 * TanStack AI tools) can consume the contract directly instead of
 * hand-mirroring it in valibot/zod.
 *
 * The strict throwing parser in `@stll/folio-core` stays the single source of
 * truth for semantics; everything here is a projection of it. Constraints
 * JSON Schema cannot express (`endOffset > startOffset`, unique operation
 * ids within a batch, per-type mode support) are noted in `description`s and
 * enforced by the parser.
 */

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_MODES,
  InvalidFolioDocumentOperationBatchError,
  parseFolioDocumentOperationBatch,
  UnsupportedFolioDocumentOperationVersionError,
  type FolioDocumentOperationBatch,
} from "@stll/folio-core/server";

/**
 * Pattern for the normalized text hashes the contract uses to detect stale
 * targets (`selectedTextHash`, `precondition.blockTextHash`). Mirrors the
 * parser's check in `@stll/folio-core`'s `document-operations.ts`.
 */
const NORMALIZED_TEXT_HASH_PATTERN = "^h[0-9a-z]+$";

/**
 * JSON Schema for the contract's `textRange` handle (a serializable range
 * over the visible text of one main-story block, as returned by `find_text`
 * style tools). Shared between {@link FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA}
 * and the `suggest_changes` tool schema in `tools.ts` so the two cannot
 * drift. `endOffset` must be greater than `startOffset`; JSON Schema cannot
 * express that relation, the parser enforces it.
 */
export const FOLIO_TEXT_RANGE_JSON_SCHEMA = {
  type: "object",
  description:
    "A range over the visible text of one main-story block. Offsets are zero-based UTF-16 " +
    "boundaries; `endOffset` must be greater than `startOffset`. Copy range objects verbatim " +
    "from the tool that produced them — `selectedTextHash` makes a shifted or edited " +
    "selection fail as stale instead of hitting the wrong text.",
  properties: {
    type: { type: "string", enum: ["textRange"] },
    story: { type: "string", enum: ["main"] },
    blockId: { type: "string", minLength: 1 },
    startOffset: { type: "integer", minimum: 0 },
    endOffset: {
      type: "integer",
      minimum: 1,
      description: "Exclusive end offset; must be greater than `startOffset`.",
    },
    selectedTextHash: {
      type: "string",
      pattern: NORMALIZED_TEXT_HASH_PATTERN,
      description: "Normalized hash of the selected text, used to detect stale ranges.",
    },
  },
  required: ["type", "story", "blockId", "startOffset", "endOffset", "selectedTextHash"],
  additionalProperties: false,
} as const;

const preconditionJsonSchema = {
  type: "object",
  description:
    "Optional guard: the operation is skipped (reason `preconditionFailed`) unless the target " +
    "block's normalized text hash still matches.",
  properties: {
    blockTextHash: {
      type: "string",
      pattern: NORMALIZED_TEXT_HASH_PATTERN,
      description: "Normalized hash of the target block's text.",
    },
  },
  required: ["blockTextHash"],
  additionalProperties: false,
} as const;

const commentJsonSchema = {
  type: "object",
  description: "A comment attached to the text affected by this operation.",
  properties: {
    text: { type: "string", description: "The comment body." },
  },
  required: ["text"],
  additionalProperties: false,
} as const;

/**
 * Properties every operation variant accepts: the required `id` plus the
 * optional review metadata (`severity`, `area`) and `precondition` guard.
 */
const operationMetaProperties = {
  id: {
    type: "string",
    description:
      "Caller-supplied operation id, echoed back in `applied` / `skipped` results. Must be " +
      "unique within a batch (enforced by the parser).",
  },
  severity: {
    type: "string",
    enum: ["low", "medium", "high"],
    description: "Optional review severity for structured-review workflows.",
  },
  area: {
    type: "string",
    description: 'Optional review area label (e.g. "Penalty") for structured-review workflows.',
  },
  precondition: preconditionJsonSchema,
} as const;

const blockIdProperty = {
  type: "string",
  description: "Id of the target block, from a prior document read.",
} as const;

/**
 * JSON Schema (draft-07 compatible) for ONE document operation: the full
 * thirteen-variant union accepted by `parseFolioDocumentOperationBatch` in
 * `@stll/folio-core`, one `oneOf` variant per entry in
 * `FOLIO_DOCUMENT_OPERATION_TYPES`. Intended for LLM tool definitions and
 * other consumers that need the contract's wire shape without re-declaring
 * it; note that the `suggest_changes` tool in `tools.ts` deliberately
 * narrows this union (see the comment there).
 */
export const FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA = {
  description:
    "One document operation, discriminated by `type`. Mirrors the operation union accepted by " +
    "`parseFolioDocumentOperationBatch` in @stll/folio-core.",
  oneOf: [
    {
      type: "object",
      description: "Replace an exact text match inside one block.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["replaceInBlock"] },
        blockId: blockIdProperty,
        find: { type: "string", description: "The exact text to find within the block." },
        replace: { type: "string", description: "The replacement text." },
        comment: commentJsonSchema,
      },
      required: ["id", "type", "blockId", "find", "replace"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Replace the text covered by a range handle.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["replaceRange"] },
        range: FOLIO_TEXT_RANGE_JSON_SCHEMA,
        replace: { type: "string", description: "The replacement text." },
        comment: commentJsonSchema,
      },
      required: ["id", "type", "range", "replace"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Attach a comment to the text covered by a range handle.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["commentOnRange"] },
        range: FOLIO_TEXT_RANGE_JSON_SCHEMA,
        comment: commentJsonSchema,
      },
      required: ["id", "type", "range", "comment"],
      additionalProperties: false,
    },
    {
      type: "object",
      description:
        "Apply inline formatting to the text covered by a range handle. Direct mode only " +
        "(formatting is not representable as a tracked change).",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["formatRange"] },
        range: FOLIO_TEXT_RANGE_JSON_SCHEMA,
        formatting: {
          type: "object",
          description: "Inline formatting to apply; at least one property is required.",
          properties: {
            bold: { type: "boolean" },
            italic: { type: "boolean" },
            underline: { type: "boolean" },
          },
          minProperties: 1,
          additionalProperties: false,
        },
      },
      required: ["id", "type", "range", "formatting"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Insert a new paragraph after the anchor block.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["insertAfterBlock"] },
        blockId: blockIdProperty,
        text: { type: "string", description: "The paragraph text to insert." },
        inheritFormatting: {
          type: "boolean",
          description: "Inherit the anchor block's formatting for the inserted paragraph.",
        },
        pageBreakBefore: {
          type: "boolean",
          description: "Start the inserted paragraph on a new page (`pageBreakBefore`).",
        },
        styleId: {
          type: "string",
          description: 'Paragraph style id for the inserted block (e.g. "ClauseHeading1").',
        },
        comment: commentJsonSchema,
      },
      required: ["id", "type", "blockId", "text"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Insert a new paragraph before the anchor block.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["insertBeforeBlock"] },
        blockId: blockIdProperty,
        text: { type: "string", description: "The paragraph text to insert." },
        inheritFormatting: {
          type: "boolean",
          description: "Inherit the anchor block's formatting for the inserted paragraph.",
        },
        pageBreakBefore: {
          type: "boolean",
          description: "Start the inserted paragraph on a new page (`pageBreakBefore`).",
        },
        styleId: {
          type: "string",
          description: 'Paragraph style id for the inserted block (e.g. "ClauseHeading1").',
        },
        comment: commentJsonSchema,
      },
      required: ["id", "type", "blockId", "text"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Replace one block's entire text.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["replaceBlock"] },
        blockId: blockIdProperty,
        text: { type: "string", description: "The new block text." },
        preserveFormatting: {
          type: "boolean",
          description: "Keep the block's existing formatting for the replacement text.",
        },
        styleId: {
          type: "string",
          description: "Paragraph style id to set on the replaced block.",
        },
        comment: commentJsonSchema,
      },
      required: ["id", "type", "blockId", "text"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Delete one block.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["deleteBlock"] },
        blockId: blockIdProperty,
        comment: commentJsonSchema,
      },
      required: ["id", "type", "blockId"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Attach a comment to one block, optionally quoting text within it.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["commentOnBlock"] },
        blockId: blockIdProperty,
        quote: {
          type: "string",
          description: "Exact text within the block the comment is about.",
        },
        comment: commentJsonSchema,
      },
      required: ["id", "type", "blockId", "comment"],
      additionalProperties: false,
    },
    {
      type: "object",
      description:
        "Insert a signature table next to the anchor block. Direct mode only (table insertion " +
        "is not representable as a tracked change).",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["insertSignatureTable"] },
        blockId: blockIdProperty,
        position: {
          type: "string",
          enum: ["after", "before"],
          description: 'Insert after the anchor block (default) or before it. Defaults to "after".',
        },
        parties: {
          type: "array",
          description: "The signing parties, one table cell per party.",
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
        comment: commentJsonSchema,
      },
      required: ["id", "type", "blockId", "parties"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Insert a row next to the row containing the anchor block. Direct mode only.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["insertTableRow"] },
        blockId: blockIdProperty,
        position: {
          type: "string",
          enum: ["after", "before"],
          description: 'Insert after the anchor row (default) or before it. Defaults to "after".',
        },
        cellTexts: {
          type: "array",
          description: "Initial text for physical cells in source order; omitted cells stay empty.",
          items: { type: "string" },
        },
      },
      required: ["id", "type", "blockId"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Delete the row containing the anchor block. Direct mode only.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["deleteTableRow"] },
        blockId: blockIdProperty,
      },
      required: ["id", "type", "blockId"],
      additionalProperties: false,
    },
    {
      type: "object",
      description:
        "Insert a column next to the column containing the anchor block. Direct mode only.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["insertTableColumn"] },
        blockId: blockIdProperty,
        position: {
          type: "string",
          enum: ["after", "before"],
          description:
            'Insert after the anchor column (default) or before it. Defaults to "after".',
        },
        cellTexts: {
          type: "array",
          description: "Initial text for newly created physical cells in row order.",
          items: { type: "string" },
        },
      },
      required: ["id", "type", "blockId"],
      additionalProperties: false,
    },
    {
      type: "object",
      description: "Delete the column containing the anchor block. Direct mode only.",
      properties: {
        ...operationMetaProperties,
        type: { type: "string", enum: ["deleteTableColumn"] },
        blockId: blockIdProperty,
      },
      required: ["id", "type", "blockId"],
      additionalProperties: false,
    },
  ],
} as const;

/**
 * JSON Schema (draft-07 compatible) for the versioned batch envelope accepted
 * by `parseFolioDocumentOperationBatch`: `version` (always
 * `FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION`), `operations`, and the
 * optional `mode` / `atomic` / `dryRun` flags — the exact wire shape of
 * `FolioDocumentOperationBatch`. Hand this to an LLM tool definition (or any
 * JSON Schema consumer) instead of re-declaring the contract.
 */
export const FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA = {
  type: "object",
  description:
    "A versioned batch of document operations. Operation ids must be unique within the batch " +
    "(enforced by the parser, not expressible in JSON Schema).",
  properties: {
    version: {
      type: "integer",
      enum: [FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION],
      description: "Document-operation contract version.",
    },
    operations: {
      type: "array",
      description: "The operations to apply, in order.",
      items: FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA,
    },
    mode: {
      type: "string",
      enum: FOLIO_DOCUMENT_OPERATION_MODES,
      description:
        'How edits land: "tracked-changes" (default) proposes revisions for human review, ' +
        '"direct" applies immediately. `formatRange`, `insertSignatureTable`, `insertTableRow`, ' +
        '`deleteTableRow`, `insertTableColumn`, and `deleteTableColumn` support "direct" only.',
    },
    atomic: {
      type: "boolean",
      description: "Reject the whole batch when any operation would be skipped.",
    },
    dryRun: {
      type: "boolean",
      description: "Preview the batch without mutating the document.",
    },
  },
  required: ["version", "operations"],
  additionalProperties: false,
} as const;

/**
 * Minimal structural subset of the Standard Schema V1 interface
 * (https://standardschema.dev). Declared locally instead of depending on
 * `@standard-schema/spec` — the spec is a tiny type-only interface designed
 * to be inlined, and this keeps the package dependency-free. Consumers that
 * accept the spec type (TanStack AI, tRPC, etc.) match it structurally.
 */
type StandardSchemaV1Issue = {
  readonly message: string;
  readonly path?: readonly PropertyKey[] | undefined;
};

type StandardSchemaV1Result<Output> =
  | { readonly value: Output; readonly issues?: undefined }
  | { readonly issues: readonly StandardSchemaV1Issue[] };

type StandardSchemaV1<Input = unknown, Output = Input> = {
  readonly "~standard": {
    readonly version: 1;
    readonly vendor: string;
    readonly validate: (value: unknown) => StandardSchemaV1Result<Output>;
    readonly types?: { readonly input: Input; readonly output: Output } | undefined;
  };
};

type FolioDocumentOperationBatchSchema = StandardSchemaV1<unknown, FolioDocumentOperationBatch> & {
  /**
   * The batch envelope's JSON Schema
   * ({@link FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA}), attached so tool
   * builders get the runtime validator and the LLM-facing schema from one
   * import.
   */
  readonly jsonSchema: typeof FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA;
};

/**
 * Convert the parser's JSONPath-style location (`$.operations[0].blockId`)
 * into a Standard Schema issue path (`["operations", 0, "blockId"]`).
 * Returns `undefined` for the root path so root-level issues carry no path.
 */
const toStandardSchemaPath = (path: string): readonly PropertyKey[] | undefined => {
  const segments: PropertyKey[] = [];
  for (const match of path.matchAll(/\.(?<key>[^.[\]]+)|\[(?<index>\d+)\]/gu)) {
    const { key, index } = match.groups ?? {};
    if (key !== undefined) {
      segments.push(key);
    } else if (index !== undefined) {
      segments.push(Number(index));
    }
  }
  return segments.length > 0 ? segments : undefined;
};

const toStandardSchemaIssue = (error: unknown): StandardSchemaV1Issue => {
  if (error instanceof InvalidFolioDocumentOperationBatchError) {
    const path = toStandardSchemaPath(error.path);
    return { message: error.message, ...(path !== undefined && { path }) };
  }
  if (error instanceof UnsupportedFolioDocumentOperationVersionError) {
    return { message: error.message, path: ["version"] };
  }
  return { message: error instanceof Error ? error.message : String(error) };
};

/**
 * Standard Schema V1 (https://standardschema.dev) validator for a document
 * operation batch, for spec-aware consumers (TanStack AI tool `inputSchema`,
 * tRPC input, etc.). Delegates to `parseFolioDocumentOperationBatch` from
 * `@stll/folio-core`, so the strict parser stays the single source of truth:
 * `validate` never throws, returning `{ value }` with the parser's exact
 * output on success and `{ issues }` (message plus a key path when the parser
 * reported one) on failure. The matching LLM-facing JSON schema is attached
 * as {@link FolioDocumentOperationBatchSchema.jsonSchema}.
 */
export const folioDocumentOperationBatchSchema: FolioDocumentOperationBatchSchema = {
  "~standard": {
    version: 1,
    vendor: "folio",
    validate: (value) => {
      // Boundary adapter: the Standard Schema spec mandates a non-throwing
      // result surface over the intentionally throwing core parser.
      try {
        return { value: parseFolioDocumentOperationBatch(value) };
      } catch (error) {
        return { issues: [toStandardSchemaIssue(error)] };
      }
    },
  },
  jsonSchema: FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA,
};
