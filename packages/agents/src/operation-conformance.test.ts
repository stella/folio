import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_TYPES,
  FolioDocxReviewer,
  InvalidFolioDocumentOperationBatchError,
  parseFolioDocumentOperationBatch,
  UnsupportedFolioDocumentOperationVersionError,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationType,
} from "@stll/folio-core/server";

import { createEditorRefBridge, type FolioAgentEditorRefLike } from "./bridges/editor-ref";
import { createReviewerBridge } from "./bridges/reviewer";
import {
  FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA,
  FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA,
  folioDocumentOperationBatchSchema,
} from "./operation-schema";

const AUTHOR = "Conformance";
const SOURCE_PATH = path.join(
  import.meta.dir,
  "../../core/src/docx/__tests__/__fixtures__/corpus/authored-empty-paragraph.docx",
);
const OPERATION_FIXTURE_PATH = path.join(
  import.meta.dir,
  "../../../tests/operations/tracked-replace.v1.json",
);

const readArrayBuffer = (filePath: string): ArrayBuffer => {
  const bytes = readFileSync(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

const readOperationFixture = (): FolioDocumentOperationBatch =>
  parseFolioDocumentOperationBatch(JSON.parse(readFileSync(OPERATION_FIXTURE_PATH, "utf8")));

type ConformanceSurface = {
  name: string;
  reviewer: FolioDocxReviewer;
  apply(batch: FolioDocumentOperationBatch): FolioDocumentOperationResult;
};

const createEditorRef = (reviewer: FolioDocxReviewer): FolioAgentEditorRefLike => ({
  createAIEditSnapshot: () => reviewer.snapshot(),
  applyAIEditOperations: ({ operations, mode, author }) =>
    reviewer.applyOperations(operations, { mode, author }),
  applyDocumentOperations: ({ batch }) => reviewer.applyDocumentOperations(batch),
  scrollToBlock: () => false,
  getTotalPages: () => 1,
  getTrackedChanges: () => reviewer.getChanges(),
});

const createSurfaces = async (): Promise<ConformanceSurface[]> => {
  const source = readArrayBuffer(SOURCE_PATH);
  const headlessReviewer = await FolioDocxReviewer.fromBuffer(source, { author: AUTHOR });
  const reviewerBridgeReviewer = await FolioDocxReviewer.fromBuffer(source, { author: AUTHOR });
  const editorBridgeReviewer = await FolioDocxReviewer.fromBuffer(source, { author: AUTHOR });
  const reviewerBridge = createReviewerBridge(reviewerBridgeReviewer);
  const editorBridge = createEditorRefBridge({
    ref: createEditorRef(editorBridgeReviewer),
    author: AUTHOR,
    getComments: () => [],
    setComments: () => {},
  });

  return [
    {
      name: "headless",
      reviewer: headlessReviewer,
      apply: (batch) => headlessReviewer.applyDocumentOperations(batch),
    },
    {
      name: "reviewer bridge",
      reviewer: reviewerBridgeReviewer,
      apply: (batch) => reviewerBridge.applyDocumentOperations(batch),
    },
    {
      name: "editor-ref bridge",
      reviewer: editorBridgeReviewer,
      apply: (batch) => editorBridge.applyDocumentOperations(batch),
    },
  ];
};

const normalizeResult = ({
  version,
  status,
  applied,
  skipped,
  issues,
  receipts,
}: FolioDocumentOperationResult) => ({
  version,
  status,
  applied: applied.map(({ id }) => ({ id })),
  skipped,
  issues,
  receipts,
});

const normalizeChanges = (reviewer: FolioDocxReviewer) =>
  reviewer.getChanges().map(({ type, author, text, blockId }) => ({
    type,
    author,
    text,
    blockId,
  }));

describe("document operation cross-surface conformance", () => {
  test("produces equivalent results and saved semantics", async () => {
    const batch = readOperationFixture();
    const outputs = [];

    for (const surface of await createSurfaces()) {
      const result = surface.apply(batch);
      const savedReviewer = await FolioDocxReviewer.fromBuffer(await surface.reviewer.toBuffer());
      outputs.push({
        name: surface.name,
        result: normalizeResult(result),
        content: surface.reviewer.getContentAsText(),
        changes: normalizeChanges(surface.reviewer),
        savedContent: savedReviewer.getContentAsText(),
        savedChanges: normalizeChanges(savedReviewer),
      });
    }

    const expected = {
      result: {
        version: 1,
        status: "committed",
        applied: [{ id: "replace-heading" }],
        skipped: [],
        issues: [],
        receipts: [
          {
            operationId: "replace-heading",
            operationIndex: 0,
            affected: [
              {
                type: "block",
                story: "main",
                blockId: "0304003A",
                effect: "updated",
              },
            ],
          },
        ],
      },
      content: "[0304003A] Intro paragraph.\n[32560014] Trailing paragraph.",
      changes: [
        {
          type: "deletion",
          author: AUTHOR,
          text: "Heading",
          blockId: "0304003A",
        },
        {
          type: "insertion",
          author: AUTHOR,
          text: "Intro",
          blockId: "0304003A",
        },
      ],
    };

    for (const output of outputs) {
      expect(output.result, output.name).toEqual(expected.result);
      expect(output.content, output.name).toBe(expected.content);
      expect(output.changes, output.name).toEqual(expected.changes);
      expect(output.savedContent, output.name).toBe(expected.content);
      expect(output.savedChanges, output.name).toEqual(expected.changes);
    }
  });

  test("previews equivalent results without mutating any surface", async () => {
    const batch = { ...readOperationFixture(), dryRun: true };
    const outputs = [];

    for (const surface of await createSurfaces()) {
      const contentBefore = surface.reviewer.getContentAsText();
      const result = surface.apply(batch);
      outputs.push({
        name: surface.name,
        result: normalizeResult(result),
        contentBefore,
        contentAfter: surface.reviewer.getContentAsText(),
        changes: normalizeChanges(surface.reviewer),
      });
    }

    for (const output of outputs) {
      expect(output.result, output.name).toEqual({
        version: 1,
        status: "previewed",
        applied: [{ id: "replace-heading" }],
        skipped: [],
        issues: [],
        receipts: [
          {
            operationId: "replace-heading",
            operationIndex: 0,
            affected: [
              {
                type: "block",
                story: "main",
                blockId: "0304003A",
                effect: "updated",
              },
            ],
          },
        ],
      });
      expect(output.contentAfter, output.name).toBe(output.contentBefore);
      expect(output.changes, output.name).toEqual([]);
    }
  });
});

/**
 * Targeted structural JSON Schema checker (no ajv): supports exactly the
 * keywords the exported contract schemas use, so a keyword the checker does
 * not know cannot silently pass.
 */
type JsonSchemaNode = {
  readonly type?: string;
  readonly enum?: readonly unknown[];
  readonly pattern?: string;
  readonly minimum?: number;
  readonly minLength?: number;
  readonly minProperties?: number;
  readonly properties?: Readonly<Record<string, JsonSchemaNode>>;
  readonly required?: readonly string[];
  readonly additionalProperties?: boolean;
  readonly items?: JsonSchemaNode;
  readonly oneOf?: readonly JsonSchemaNode[];
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const admitsObject = (schema: JsonSchemaNode, value: unknown): boolean => {
  if (!isRecord(value)) {
    return false;
  }
  if (schema.required?.some((key) => !(key in value))) {
    return false;
  }
  const keys = Object.keys(value);
  if (schema.minProperties !== undefined && keys.length < schema.minProperties) {
    return false;
  }
  const properties = schema.properties ?? {};
  return keys.every((key) => {
    const propertySchema = properties[key];
    if (propertySchema === undefined) {
      return schema.additionalProperties !== false;
    }
    return admits(propertySchema, value[key]);
  });
};

const admits = (schema: JsonSchemaNode, value: unknown): boolean => {
  if (schema.oneOf !== undefined) {
    // oneOf semantics: exactly one variant must admit the value, so the union
    // stays a real discriminated union.
    return schema.oneOf.filter((variant) => admits(variant, value)).length === 1;
  }
  if (schema.enum !== undefined && !schema.enum.includes(value)) {
    return false;
  }
  switch (schema.type) {
    case "string":
      return (
        typeof value === "string" &&
        (schema.minLength === undefined || value.length >= schema.minLength) &&
        (schema.pattern === undefined || new RegExp(schema.pattern, "u").test(value))
      );
    case "integer":
      return (
        typeof value === "number" &&
        Number.isInteger(value) &&
        (schema.minimum === undefined || value >= schema.minimum)
      );
    case "boolean":
      return typeof value === "boolean";
    case "array": {
      if (!Array.isArray(value)) {
        return false;
      }
      const { items } = schema;
      return items === undefined || value.every((item) => admits(items, item));
    }
    case "object":
      return admitsObject(schema, value);
    default:
      return true;
  }
};

const OPERATION_SCHEMA: JsonSchemaNode = FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA;
const BATCH_SCHEMA: JsonSchemaNode = FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA;

const CONTRACT_RANGE = {
  type: "textRange",
  story: "main",
  blockId: "0304003A",
  startOffset: 0,
  endOffset: 5,
  selectedTextHash: "h1a2b3",
} as const;

/**
 * One representative fixture per contract operation type, exercising every
 * optional field the parser accepts for that type at least once across the
 * set.
 */
const CONTRACT_OPERATION_FIXTURES: Record<FolioDocumentOperationType, Record<string, unknown>> = {
  replaceInBlock: {
    id: "op-replace-in-block",
    type: "replaceInBlock",
    blockId: "0304003A",
    find: "Heading",
    replace: "Intro",
    comment: { text: "Tighten the heading." },
    severity: "low",
    area: "Style",
    precondition: { blockTextHash: "h1a2b3" },
  },
  replaceRange: {
    id: "op-replace-range",
    type: "replaceRange",
    range: CONTRACT_RANGE,
    replace: "Intro",
    comment: { text: "Reworded." },
  },
  commentOnRange: {
    id: "op-comment-on-range",
    type: "commentOnRange",
    range: CONTRACT_RANGE,
    comment: { text: "Check this clause." },
  },
  formatRange: {
    id: "op-format-range",
    type: "formatRange",
    range: CONTRACT_RANGE,
    formatting: { bold: true, underline: false },
  },
  insertAfterBlock: {
    id: "op-insert-after",
    type: "insertAfterBlock",
    blockId: "0304003A",
    text: "New paragraph.",
    inheritFormatting: true,
    pageBreakBefore: true,
    styleId: "Heading1",
  },
  insertBeforeBlock: {
    id: "op-insert-before",
    type: "insertBeforeBlock",
    blockId: "0304003A",
    text: "New paragraph.",
  },
  replaceBlock: {
    id: "op-replace-block",
    type: "replaceBlock",
    blockId: "0304003A",
    text: "Replacement text.",
    preserveFormatting: true,
    styleId: "Normal",
  },
  deleteBlock: {
    id: "op-delete-block",
    type: "deleteBlock",
    blockId: "0304003A",
    comment: { text: "Redundant." },
  },
  commentOnBlock: {
    id: "op-comment-on-block",
    type: "commentOnBlock",
    blockId: "0304003A",
    quote: "Heading",
    comment: { text: "Rename this?" },
  },
  insertSignatureTable: {
    id: "op-signature-table",
    type: "insertSignatureTable",
    blockId: "0304003A",
    position: "before",
    parties: [{ name: "Acme s.r.o.", signatory: "Jane Doe", title: "CEO" }],
  },
  insertTableRow: {
    id: "op-insert-table-row",
    type: "insertTableRow",
    blockId: "0304003A",
    position: "before",
    cellTexts: ["First", "Second"],
  },
  deleteTableRow: {
    id: "op-delete-table-row",
    type: "deleteTableRow",
    blockId: "0304003A",
  },
  insertTableColumn: {
    id: "op-insert-table-column",
    type: "insertTableColumn",
    blockId: "0304003A",
    position: "before",
    cellTexts: ["First", "Second"],
  },
};

const variantTypeOf = (variant: JsonSchemaNode): unknown => variant.properties?.["type"]?.enum?.[0];

const variantForType = (type: FolioDocumentOperationType): JsonSchemaNode => {
  const variant = OPERATION_SCHEMA.oneOf?.find((candidate) => variantTypeOf(candidate) === type);
  if (variant === undefined) {
    throw new Error(`no schema variant for operation type "${type}"`);
  }
  return variant;
};

describe("document operation contract JSON schema conformance", () => {
  test("the schema union covers exactly the contract's operation types", () => {
    const variantTypes = (OPERATION_SCHEMA.oneOf ?? []).map(variantTypeOf);
    expect(variantTypes).toEqual([...FOLIO_DOCUMENT_OPERATION_TYPES]);
  });

  test("every operation type round-trips the parser and is admitted by the schema", () => {
    for (const type of FOLIO_DOCUMENT_OPERATION_TYPES) {
      const operation = CONTRACT_OPERATION_FIXTURES[type];
      const batch = {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [operation],
        mode: "direct",
        atomic: true,
        dryRun: true,
      };
      const parsed = parseFolioDocumentOperationBatch(batch);
      expect(JSON.parse(JSON.stringify(parsed)), type).toEqual(batch);
      expect(admits(OPERATION_SCHEMA, operation), type).toBe(true);
      expect(admits(BATCH_SCHEMA, batch), type).toBe(true);
    }
  });

  test("rejects a wrong contract version in both the parser and the schema", () => {
    const wrongVersion = { version: 2, operations: [] };
    expect(() => parseFolioDocumentOperationBatch(wrongVersion)).toThrow(
      UnsupportedFolioDocumentOperationVersionError,
    );
    expect(admits(BATCH_SCHEMA, wrongVersion)).toBe(false);
  });

  test("rejects an unknown operation type in both the parser and the schema", () => {
    const unknownType = { id: "op-unknown", type: "renameBlock", blockId: "0304003A" };
    expect(() =>
      parseFolioDocumentOperationBatch({ version: 1, operations: [unknownType] }),
    ).toThrow(InvalidFolioDocumentOperationBatchError);
    expect(admits(OPERATION_SCHEMA, unknownType)).toBe(false);
  });

  test("rejects an unexpected property in both the parser and the schema", () => {
    const extraProperty = { ...CONTRACT_OPERATION_FIXTURES.deleteBlock, bogus: true };
    expect(() =>
      parseFolioDocumentOperationBatch({ version: 1, operations: [extraProperty] }),
    ).toThrow(InvalidFolioDocumentOperationBatchError);
    expect(admits(OPERATION_SCHEMA, extraProperty)).toBe(false);
  });

  test("rejects each missing required field in both the parser and the schema", () => {
    for (const type of FOLIO_DOCUMENT_OPERATION_TYPES) {
      for (const missingKey of variantForType(type).required ?? []) {
        const { [missingKey]: _removed, ...broken } = CONTRACT_OPERATION_FIXTURES[type];
        const label = `${type} without ${missingKey}`;
        expect(
          () => parseFolioDocumentOperationBatch({ version: 1, operations: [broken] }),
          label,
        ).toThrow(InvalidFolioDocumentOperationBatchError);
        expect(admits(OPERATION_SCHEMA, broken), label).toBe(false);
      }
    }
  });
});

describe("document operation batch standard schema", () => {
  const validate = folioDocumentOperationBatchSchema["~standard"].validate;

  test("exposes the standard-schema surface and the attached JSON schema", () => {
    expect(folioDocumentOperationBatchSchema["~standard"].version).toBe(1);
    expect(folioDocumentOperationBatchSchema["~standard"].vendor).toBe("folio");
    expect(folioDocumentOperationBatchSchema.jsonSchema).toBe(
      FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA,
    );
  });

  test("returns the parser's exact output for a valid batch", () => {
    const batch = {
      version: 1,
      operations: [CONTRACT_OPERATION_FIXTURES.replaceInBlock],
      mode: "tracked-changes",
    };
    const result = validate(batch);
    if (result.issues !== undefined) {
      throw new Error("expected a success result");
    }
    expect(result.value).toEqual(parseFolioDocumentOperationBatch(batch));
  });

  test("returns issues with a path for an invalid operation instead of throwing", () => {
    const result = validate({
      version: 1,
      operations: [{ id: "op-1", type: "deleteBlock" }],
    });
    if (result.issues === undefined) {
      throw new Error("expected a failure result");
    }
    expect(result.issues).toEqual([
      {
        message: "Invalid document operation batch at $.operations[0].blockId: expected a string.",
        path: ["operations", 0, "blockId"],
      },
    ]);
  });

  test("returns a version issue for an unsupported contract version", () => {
    const result = validate({ version: 2, operations: [] });
    if (result.issues === undefined) {
      throw new Error("expected a failure result");
    }
    expect(result.issues).toEqual([
      { message: "Unsupported document operation contract version.", path: ["version"] },
    ]);
  });

  test("never throws, whatever the input", () => {
    const inputs: unknown[] = [
      undefined,
      null,
      42,
      "batch",
      [],
      {},
      { version: "1", operations: [] },
      { version: 1 },
      { version: 1, operations: [{}] },
      { version: 1, operations: [{ id: "a", type: "deleteBlock", blockId: "b" }], mode: "bulk" },
    ];
    for (const input of inputs) {
      const result = validate(input);
      expect(result.issues, JSON.stringify(input)).toBeDefined();
      const issue = result.issues?.at(0);
      expect(typeof issue?.message, JSON.stringify(input)).toBe("string");
      expect(issue?.message.length, JSON.stringify(input)).toBeGreaterThan(0);
    }
  });
});
