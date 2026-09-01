/**
 * `suggest_changes` host-configuration tests: {@link resolveSuggestChangesOptions}'s
 * defaults and validation, how a resolved surface reshapes the tool's JSON
 * Schema (`tools.ts`) and capability description, how `parseSuggestChangesInput`
 * enforces what the schema advertised, the parser's lenient decode, and the
 * executor's document-version pinning and host-queue paths against a real
 * `FolioDocxReviewer`.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE,
  FOLIO_DOCUMENT_OPERATION_TYPES,
  FolioDocxReviewer,
  getFolioDocumentOperationReceipts,
  type FolioDocumentOperationType,
} from "@stll/folio-core/server";

import type { FolioAgentBridge } from "./bridge";
import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCall } from "./execute";
import { parseSuggestChangesInput } from "./parse";
import {
  DEFAULT_MAX_OPERATIONS_PER_CALL,
  DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES,
  InvalidFolioSuggestChangesOptionsError,
  resolveSuggestChangesOptions,
  type FolioSuggestChangesOptions,
} from "./suggest-changes-options";
import {
  describeSuggestChangesCapabilities,
  FOLIO_AGENT_TOOLS,
  getFolioToolDefinitions,
} from "./tools";
import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type { FolioAgentToolDefinition } from "./types";

const ALL_TYPES = FOLIO_DOCUMENT_OPERATION_TYPES;

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const asSchemaObject = (value: unknown): Record<string, unknown> => {
  if (!isPlainObject(value)) {
    throw new Error("expected a JSON-Schema object");
  }
  return value;
};

/** Read a nested object-valued schema property (`schema.properties[key]`). */
const propertyOf = (schema: unknown, key: string): Record<string, unknown> => {
  const properties = asSchemaObject(asSchemaObject(schema)["properties"]);
  if (properties[key] === undefined) {
    throw new Error(`expected a \`${key}\` property on the schema`);
  }
  return asSchemaObject(properties[key]);
};

const suggestChangesDefinitionFor = (
  options?: FolioSuggestChangesOptions,
): FolioAgentToolDefinition => {
  const definitions = getFolioToolDefinitions(
    options === undefined ? {} : { suggestChanges: options },
  );
  const definition = definitions.find(
    (entry) => entry.name === FOLIO_AGENT_TOOL_NAMES.suggestChanges,
  );
  if (definition === undefined) {
    throw new Error("expected a suggest_changes tool definition");
  }
  return definition;
};

const operationItemSchemaOf = (definition: FolioAgentToolDefinition): Record<string, unknown> =>
  asSchemaObject(propertyOf(definition.inputSchema, "operations")["items"]);

/** Ids the parser mints for unnamed operations: a per-call nonce plus the 1-based index. */
const MINTED_ID_PATTERN = /^op-[0-9a-f-]{36}-1$/;

// ---------------------------------------------------------------------------
// resolveSuggestChangesOptions
// ---------------------------------------------------------------------------

describe("resolveSuggestChangesOptions", () => {
  test("defaults to every contract type except commentOnBlock and insertSignatureTable", () => {
    const resolved = resolveSuggestChangesOptions();
    expect(resolved.operationTypes).toEqual(DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES);
    expect(resolved.operationTypes).toHaveLength(14);
    expect(resolved.operationTypes).not.toContain("commentOnBlock");
    expect(resolved.operationTypes).not.toContain("insertSignatureTable");
    expect(resolved.reviewMeta).toBe("optional");
    expect(resolved.maxOperations).toBe(DEFAULT_MAX_OPERATIONS_PER_CALL);
    expect(resolved.maxOperations).toBe(50);
    expect(resolved.documentVersion).toBeNull();
  });

  test("requested operation types are deduplicated and returned in contract order", () => {
    const resolved = resolveSuggestChangesOptions({
      operationTypes: ["deleteBlock", "replaceInBlock", "deleteBlock"],
    });
    // Contract order lists replaceInBlock before deleteBlock; the input
    // order (deleteBlock first, plus a repeat) must not leak through.
    expect(resolved.operationTypes).toEqual(["replaceInBlock", "deleteBlock"]);
  });

  const INVALID_OPTIONS_CASES: { label: string; options: FolioSuggestChangesOptions }[] = [
    { label: "an empty operationTypes array", options: { operationTypes: [] } },
    {
      label: "an unknown operation type",
      // SAFETY: the runtime guard exists for callers outside TypeScript (a
      // JSON config, a JS host); the cast is the only way to hand it an
      // off-contract string from typed test code.
      options: { operationTypes: ["notAContractType" as FolioDocumentOperationType] },
    },
    { label: "maxOperations 0", options: { maxOperations: 0 } },
    { label: "maxOperations 201", options: { maxOperations: 201 } },
    { label: "maxOperations 2.5", options: { maxOperations: 2.5 } },
    { label: "an empty documentVersion.current", options: { documentVersion: { current: "" } } },
  ];

  test.each(INVALID_OPTIONS_CASES)(
    "$label throws InvalidFolioSuggestChangesOptionsError",
    ({ options }) => {
      expect(() => resolveSuggestChangesOptions(options)).toThrow(
        InvalidFolioSuggestChangesOptionsError,
      );
    },
  );
});

// ---------------------------------------------------------------------------
// Schema derivation (property-style over operation-type subsets)
// ---------------------------------------------------------------------------

/** Small deterministic PRNG so the pseudo-random subsets are reproducible across runs. */
const createLcg = (seed: number): (() => number) => {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 0x1_0000_0000;
  };
};

const randomOperationTypeSubset = (rng: () => number): FolioDocumentOperationType[] => {
  const chosen = ALL_TYPES.filter(() => rng() < 0.5);
  if (chosen.length > 0) {
    return chosen;
  }
  // Every subset must be non-empty (an empty operationTypes array is
  // invalid input); fall back to one randomly chosen type.
  const index = Math.floor(rng() * ALL_TYPES.length);
  const fallback = ALL_TYPES.at(index) ?? ALL_TYPES.at(0);
  if (fallback === undefined) {
    throw new Error("expected at least one contract operation type");
  }
  return [fallback];
};

const RANDOM_SUBSET_RNG = createLcg(1337);
const RANDOM_SUBSETS: FolioDocumentOperationType[][] = Array.from({ length: 20 }, () =>
  randomOperationTypeSubset(RANDOM_SUBSET_RNG),
);

const SCHEMA_SUBSETS: { label: string; types: readonly FolioDocumentOperationType[] }[] = [
  ...ALL_TYPES.map((type) => ({ label: `single type: ${type}`, types: [type] })),
  { label: "the full 16-type set", types: ALL_TYPES },
  { label: "the default type set", types: DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES },
  ...RANDOM_SUBSETS.map((types, index) => ({
    label: `random subset #${index + 1} of ${types.length} types`,
    types,
  })),
];

describe("suggest_changes schema + capability description follow operationTypes", () => {
  test.each(SCHEMA_SUBSETS)("$label", ({ types }) => {
    const resolved = resolveSuggestChangesOptions({ operationTypes: types });
    const definition = suggestChangesDefinitionFor({ operationTypes: types });
    const itemSchema = operationItemSchemaOf(definition);

    const typeSchema = propertyOf(itemSchema, "type");
    expect(typeSchema["enum"]).toEqual(resolved.operationTypes);

    const expectedKeys = new Set<string>();
    for (const type of resolved.operationTypes) {
      for (const key of FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE[type]) {
        if (key !== "type" && key !== "suggestionId") {
          expectedKeys.add(key);
        }
      }
    }
    const itemProperties = asSchemaObject(itemSchema["properties"]);
    const actualKeys = new Set(Object.keys(itemProperties).filter((key) => key !== "type"));
    expect([...actualKeys].sort()).toEqual([...expectedKeys].sort());

    const description = describeSuggestChangesCapabilities({ operationTypes: types });
    for (const type of resolved.operationTypes) {
      expect(
        new RegExp(`\\b${type}\\b`).test(description),
        `expected "${type}" to be mentioned`,
      ).toBe(true);
    }
    for (const type of ALL_TYPES) {
      if (resolved.operationTypes.includes(type)) {
        continue;
      }
      expect(
        new RegExp(`\\b${type}\\b`).test(description),
        `expected "${type}" not to be mentioned`,
      ).toBe(false);
    }
  });

  test("reviewMeta: required tightens the operation schema's required list", () => {
    const definition = suggestChangesDefinitionFor({ reviewMeta: "required" });
    const itemSchema = operationItemSchemaOf(definition);
    expect(itemSchema["required"]).toEqual(["type", "severity", "area"]);
  });

  test("documentVersion pins a top-level enum and marks documentVersion required", () => {
    const definition = suggestChangesDefinitionFor({ documentVersion: { current: "v7" } });
    const documentVersionSchema = propertyOf(definition.inputSchema, "documentVersion");
    expect(documentVersionSchema["enum"]).toEqual(["v7"]);
    expect(definition.inputSchema["required"]).toEqual(["documentVersion", "operations"]);
  });

  test("maxOperations sets the operations array's maxItems", () => {
    const definition = suggestChangesDefinitionFor({ maxOperations: 3 });
    const operationsSchema = propertyOf(definition.inputSchema, "operations");
    expect(operationsSchema["maxItems"]).toBe(3);
  });

  test("without options, getFolioToolDefinitions returns the same array instance as FOLIO_AGENT_TOOLS", () => {
    expect(getFolioToolDefinitions()).toBe(FOLIO_AGENT_TOOLS);
  });
});

// ---------------------------------------------------------------------------
// Parser follows the resolved options
// ---------------------------------------------------------------------------

describe("parseSuggestChangesInput follows suggestChanges options", () => {
  test("a type outside the configured subset is rejected, listing the allowed types", () => {
    const result = parseSuggestChangesInput(
      { operations: [{ type: "deleteBlock", blockId: "b1" }] },
      { operationTypes: ["replaceInBlock", "replaceBlock"] },
    );
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("replaceInBlock");
    expect(result.error).toContain("replaceBlock");
  });

  test("reviewMeta: required rejects a missing severity and a missing or empty area", () => {
    const options: FolioSuggestChangesOptions = { reviewMeta: "required" };

    const missingSeverity = parseSuggestChangesInput(
      { operations: [{ type: "deleteBlock", blockId: "b1", area: "Payment terms" }] },
      options,
    );
    expect(missingSeverity.ok).toBe(false);
    if (missingSeverity.ok) {
      throw new Error("expected ok:false");
    }
    expect(missingSeverity.error).toContain("severity");

    const missingArea = parseSuggestChangesInput(
      { operations: [{ type: "deleteBlock", blockId: "b1", severity: "low" }] },
      options,
    );
    expect(missingArea.ok).toBe(false);
    if (missingArea.ok) {
      throw new Error("expected ok:false");
    }
    expect(missingArea.error).toContain("area");

    const emptyArea = parseSuggestChangesInput(
      { operations: [{ type: "deleteBlock", blockId: "b1", severity: "low", area: "" }] },
      options,
    );
    expect(emptyArea.ok).toBe(false);
    if (emptyArea.ok) {
      throw new Error("expected ok:false");
    }
    expect(emptyArea.error).toContain("area");
  });

  test("severity and area pass through onto parsed operations under both review-meta policies", () => {
    const policyOptions: (FolioSuggestChangesOptions | undefined)[] = [
      undefined,
      { reviewMeta: "required" },
    ];
    for (const options of policyOptions) {
      const result = parseSuggestChangesInput(
        {
          operations: [
            { type: "deleteBlock", blockId: "b1", severity: "high", area: "Payment terms" },
          ],
        },
        options,
      );
      expect(result.ok).toBe(true);
      if (!result.ok) {
        throw new Error("expected ok:true");
      }
      expect(result.operations[0]).toMatchObject({ severity: "high", area: "Payment terms" });
    }
  });

  test("maxOperations: 2 rejects a 3-operation batch with the configured limit in the message", () => {
    const operations = Array.from({ length: 3 }, (_, index) => ({
      type: "deleteBlock",
      blockId: `b${index}`,
    }));
    const result = parseSuggestChangesInput({ operations }, { maxOperations: 2 });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("2-operation limit");
  });

  test("the documentVersion option makes documentVersion required and lands as result.precondition", () => {
    const options: FolioSuggestChangesOptions = { documentVersion: { current: "v3" } };

    const missing = parseSuggestChangesInput(
      { operations: [{ type: "deleteBlock", blockId: "b1" }] },
      options,
    );
    expect(missing.ok).toBe(false);
    if (missing.ok) {
      throw new Error("expected ok:false");
    }
    expect(missing.error).toContain("documentVersion");

    const provided = parseSuggestChangesInput(
      { operations: [{ type: "deleteBlock", blockId: "b1" }], documentVersion: "v3" },
      options,
    );
    expect(provided.ok).toBe(true);
    if (!provided.ok) {
      throw new Error("expected ok:true");
    }
    expect(provided.precondition).toEqual({ documentVersion: "v3" });
  });

  test("without the documentVersion option, a supplied documentVersion is ignored with a normalization", () => {
    const result = parseSuggestChangesInput({
      operations: [{ type: "deleteBlock", blockId: "b1" }],
      documentVersion: "v3",
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.precondition).toBeUndefined();
    expect(result.normalizations).toEqual([
      { path: "documentVersion", message: expect.stringContaining("ignored") },
    ]);
  });
});

// ---------------------------------------------------------------------------
// Lenient decode
// ---------------------------------------------------------------------------

describe("parseSuggestChangesInput lenient decode", () => {
  test("arguments given as a JSON string decode with one normalization at $", () => {
    const args = JSON.stringify({ operations: [{ type: "deleteBlock", blockId: "b1" }] });
    const result = parseSuggestChangesInput(args);
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.normalizations).toEqual([
      { path: "$", message: expect.stringContaining("JSON string") },
    ]);
  });

  test("operations given as a JSON string decodes with one normalization at operations", () => {
    const result = parseSuggestChangesInput({
      operations: JSON.stringify([{ type: "deleteBlock", blockId: "b1" }]),
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.normalizations).toEqual([
      { path: "operations", message: expect.stringContaining("JSON string") },
    ]);
  });

  test("one operation given as a JSON string decodes with a normalization at operations[0]", () => {
    const result = parseSuggestChangesInput({
      operations: [JSON.stringify({ type: "deleteBlock", blockId: "b1" })],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.normalizations).toEqual([
      { path: "operations[0]", message: expect.stringContaining("JSON string") },
    ]);
  });

  test("`kind` is read as `type` with a normalization at operations[0].kind", () => {
    const result = parseSuggestChangesInput({
      operations: [{ kind: "deleteBlock", blockId: "b1" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.normalizations).toEqual([
      { path: "operations[0].kind", message: expect.stringContaining("`kind`") },
    ]);
    expect(result.operations[0]).toMatchObject({ type: "deleteBlock" });
  });

  test("an unknown top-level key is ignored with a normalization", () => {
    const result = parseSuggestChangesInput({
      operations: [{ type: "deleteBlock", blockId: "b1" }],
      foo: "bar",
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.normalizations).toEqual([
      { path: "foo", message: expect.stringContaining("foo") },
    ]);
  });

  test("a key that does not apply to the operation type is dropped with a normalization", () => {
    const findOnDeleteBlock = parseSuggestChangesInput({
      operations: [{ type: "deleteBlock", blockId: "b1", find: "x" }],
    });
    expect(findOnDeleteBlock.ok).toBe(true);
    if (!findOnDeleteBlock.ok) {
      throw new Error("expected ok:true");
    }
    expect(findOnDeleteBlock.normalizations).toEqual([
      { path: "operations[0].find", message: expect.stringContaining("does not apply") },
    ]);

    const junkKey = parseSuggestChangesInput({
      operations: [{ type: "deleteBlock", blockId: "b1", foo: "bar" }],
    });
    expect(junkKey.ok).toBe(true);
    if (!junkKey.ok) {
      throw new Error("expected ok:true");
    }
    expect(junkKey.normalizations).toEqual([
      { path: "operations[0].foo", message: expect.stringContaining("does not apply") },
    ]);
  });

  test("a clean call has no normalizations", () => {
    const result = parseSuggestChangesInput({
      operations: [{ type: "deleteBlock", blockId: "b1" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.normalizations).toEqual([]);
  });

  test("two consecutive calls mint different ids for the same unnamed operation", () => {
    const input = { operations: [{ type: "deleteBlock", blockId: "b1" }] };
    const first = parseSuggestChangesInput(input);
    const second = parseSuggestChangesInput(input);
    if (!first.ok || !second.ok) {
      throw new Error("expected ok:true");
    }
    const firstId = first.operations.at(0)?.id;
    const secondId = second.operations.at(0)?.id;
    expect(firstId).toMatch(MINTED_ID_PATTERN);
    expect(secondId).toMatch(MINTED_ID_PATTERN);
    expect(firstId).not.toBe(secondId);
  });

  test("caller-supplied ids are kept", () => {
    const result = parseSuggestChangesInput({
      operations: [{ id: "custom-1", type: "deleteBlock", blockId: "b1" }],
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.operations.at(0)?.id).toBe("custom-1");
  });

  test("duplicate caller ids are rejected", () => {
    const result = parseSuggestChangesInput({
      operations: [
        { id: "dup", type: "deleteBlock", blockId: "b1" },
        { id: "dup", type: "deleteBlock", blockId: "b2" },
      ],
    });
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("id");
  });
});

// ---------------------------------------------------------------------------
// Executor
// ---------------------------------------------------------------------------

// Reuses the same corpus fixture `execute.test.ts` builds its reviewer
// round-trip tests against.
const FIXTURE = path.join(
  import.meta.dir,
  "../../core/src/docx/__tests__/__fixtures__/corpus/authored-empty-paragraph.docx",
);

const readFixture = (): ArrayBuffer => {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

describe("executeFolioToolCall: suggest_changes document-version pinning and queue bridges", () => {
  test("a version-pinned call fails when the bridge cannot report a document version", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const block = bridge.snapshot().blocks.at(0);
    if (block === undefined) {
      throw new Error("expected at least one block in the fixture");
    }

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { documentVersion: "v1", operations: [{ type: "deleteBlock", blockId: block.id }] },
      bridge,
      { suggestChanges: { documentVersion: { current: "v1" } } },
    );
    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("expected ok:false");
    }
    expect(result.error).toContain("documentVersion");
  });

  test("a document-version mismatch skips the whole batch with documentVersionMismatch", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const reviewerBridge = createReviewerBridge(reviewer);
    const block = reviewerBridge.snapshot().blocks.at(0);
    if (block === undefined) {
      throw new Error("expected at least one block in the fixture");
    }
    const bridge: FolioAgentBridge = { ...reviewerBridge, getDocumentVersion: () => "v2" };

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { documentVersion: "v1", operations: [{ type: "deleteBlock", blockId: block.id }] },
      bridge,
      { suggestChanges: { documentVersion: { current: "v1" } } },
    );
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.result.applied).toEqual([]);
    expect(result.result.skipped).toHaveLength(1);
    expect(result.result.skipped[0]?.reason).toContain("document changed");
    expect(result.result.issues).toHaveLength(1);
    expect(result.result.issues[0]?.code).toBe("documentVersionMismatch");
    expect(result.result.issues[0]?.recovery).toBe("refreshDocument");
  });

  test("a matching document version lets the batch apply", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const reviewerBridge = createReviewerBridge(reviewer);
    const block = reviewerBridge.snapshot().blocks.at(0);
    if (block === undefined) {
      throw new Error("expected at least one block in the fixture");
    }
    const bridge: FolioAgentBridge = { ...reviewerBridge, getDocumentVersion: () => "v1" };

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { documentVersion: "v1", operations: [{ type: "deleteBlock", blockId: block.id }] },
      bridge,
      { suggestChanges: { documentVersion: { current: "v1" } } },
    );
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.result.applied).toHaveLength(1);
    expect(result.result.queued).toEqual([]);
  });

  test("a hand-written queue bridge reports queued operations with receipts", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const block = reviewer.snapshot().blocks.at(0);
    if (block === undefined) {
      throw new Error("expected at least one block in the fixture");
    }

    const queueBridge: FolioAgentBridge = {
      snapshot: () => reviewer.snapshot(),
      applyDocumentOperations: (batch) => ({
        version: 1,
        status: "queued",
        applied: [],
        queued: batch.operations.map(({ id }) => ({ id })),
        skipped: [],
        issues: [],
        receipts: getFolioDocumentOperationReceipts(
          batch.operations,
          batch.operations.map(({ id }) => ({ id })),
        ),
        undoHandle: null,
      }),
      getComments: () => [],
      getChanges: () => [],
      replyToComment: () => false,
      resolveComment: () => false,
    };

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations: [{ type: "replaceInBlock", blockId: block.id, find: "x", replace: "y" }] },
      queueBridge,
    );
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.result.applied).toEqual([]);
    expect(result.result.queued).toHaveLength(1);
    expect(result.result.queued[0]?.id).toMatch(MINTED_ID_PATTERN);
    expect(result.result.receipts).toHaveLength(1);
    expect(result.result.receipts[0]?.affected[0]).toMatchObject({
      type: "block",
      blockId: block.id,
      effect: "updated",
    });
  });

  test("the summary carries the parser's normalizations", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(readFixture());
    const bridge = createReviewerBridge(reviewer);
    const block = bridge.snapshot().blocks.at(0);
    if (block === undefined) {
      throw new Error("expected at least one block in the fixture");
    }

    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      { operations: [{ kind: "deleteBlock", blockId: block.id }] },
      bridge,
    );
    expect(result.ok).toBe(true);
    if (!result.ok) {
      throw new Error("expected ok:true");
    }
    expect(result.result.normalizations).toEqual([
      { path: "operations[0].kind", message: expect.stringContaining("`kind`") },
    ]);
  });
});
