/**
 * Batch operations the file surface adds to `suggest_changes`, expanded into
 * contract operations against the document they are about to change:
 *
 * - `replaceAll {find, replace, matchCase?, wholeWord?}` becomes one
 *   `replaceRange` per `find_text` match in the body (tables included), so a
 *   rename needs no search call and keeps each run's formatting.
 * - `addComment {comment, blockId?, quote?}` becomes `commentOnBlock`; with
 *   only a quote, it anchors to the one block containing that text.
 *
 * Matching is `find_text`'s own; nothing here searches text itself.
 */

import { Result } from "better-result";

import type { FolioAgentBridge } from "@stll/folio-agents/bridge";
import { executeFolioToolCallUntyped } from "@stll/folio-agents/execute";
import {
  DEFAULT_MAX_OPERATIONS_PER_CALL,
  MAX_OPERATIONS_PER_CALL_LIMIT,
  type FolioAgentToolOptions,
} from "@stll/folio-agents/suggest-changes-options";
import { SUGGEST_CHANGES_OPERATION_TYPES } from "@stll/folio-agents/tools";
import { FOLIO_AGENT_TOOL_NAMES } from "@stll/folio-agents/types";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import type { JsonObjectSchema } from "./registry";

export const BATCH_OPERATION_TYPES = ["replaceAll", "addComment"] as const;

/** Fields each batch operation accepts, besides `type`. */
const BATCH_OPERATION_KEYS: Readonly<Record<string, ReadonlySet<string>>> = {
  replaceAll: new Set(["id", "find", "replace", "matchCase", "wholeWord", "severity", "area"]),
  addComment: new Set(["id", "blockId", "quote", "comment", "severity", "area"]),
};

const BATCH_PROPERTY_SCHEMAS = {
  matchCase: { type: "boolean", description: "For `replaceAll`: match case exactly." },
  wholeWord: { type: "boolean", description: "For `replaceAll`: match whole words only." },
} as const;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * `suggest_changes`' arguments with the batch operations added to the
 * operation schema: their types in the `type` enum and their extra fields.
 */
export const withBatchOperations = (schema: JsonObjectSchema): JsonObjectSchema => {
  const operations = schema.properties["operations"];
  if (!isRecord(operations) || !isRecord(operations["items"])) return schema;
  const items = operations["items"];
  const properties = isRecord(items["properties"]) ? items["properties"] : {};
  const type = isRecord(properties["type"]) ? properties["type"] : {};
  const types = Array.isArray(type["enum"]) ? type["enum"] : [];
  return {
    ...schema,
    properties: {
      ...schema.properties,
      operations: {
        ...operations,
        items: {
          ...items,
          properties: {
            ...properties,
            type: {
              ...type,
              enum: [...types, ...BATCH_OPERATION_TYPES],
              description:
                `${typeof type["description"] === "string" ? type["description"] : ""} ` +
                "replaceAll (replace every match of `find` in the body and tables); " +
                "addComment (comment on `blockId`, or on the one block containing `quote`).",
            },
            ...BATCH_PROPERTY_SCHEMAS,
          },
        },
      },
    },
  };
};

export type ExpandedBatch = {
  operations: unknown[];
  /** Options the expanded batch needs from `suggest_changes`. */
  options: FolioAgentToolOptions;
  /** How many replacements each `replaceAll`, in batch order, produced. */
  replaced: { find: string; count: number }[];
};

const invalidInput = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message, hint });

type Match = { blockId: string; range: unknown };

const findMatches = (
  bridge: FolioAgentBridge,
  query: string,
  matchCase: boolean,
  wholeWord: boolean,
): Result<{ matches: Match[]; truncated: boolean; total: number }, FolioCliError> => {
  const found = executeFolioToolCallUntyped(
    FOLIO_AGENT_TOOL_NAMES.findText,
    { query, matchCase, wholeWord },
    bridge,
  );
  if (!found.ok) return Result.err(invalidInput(found.error));
  const result = isRecord(found.result) ? found.result : {};
  const matches = (Array.isArray(result["matches"]) ? result["matches"] : []).flatMap(
    (match: unknown) =>
      isRecord(match) && typeof match["blockId"] === "string" && isRecord(match["range"])
        ? [{ blockId: match["blockId"], range: match["range"] }]
        : [],
  );
  return Result.ok({
    matches,
    truncated: result["truncated"] === true,
    total: typeof result["totalMatches"] === "number" ? result["totalMatches"] : matches.length,
  });
};

const checkKeys = (
  operation: Record<string, unknown>,
  type: string,
  index: number,
): FolioCliError | null => {
  const allowed = BATCH_OPERATION_KEYS[type] ?? new Set();
  const extra = Object.keys(operation).filter((key) => key !== "type" && !allowed.has(key));
  return extra.length === 0
    ? null
    : invalidInput(
        `operations[${index}] (${type}) does not take ${extra.join(", ")}.`,
        `${type} takes ${[...allowed].join(", ")}.`,
      );
};

const metaOf = (operation: Record<string, unknown>): Record<string, unknown> => ({
  ...(operation["severity"] !== undefined && { severity: operation["severity"] }),
  ...(operation["area"] !== undefined && { area: operation["area"] }),
});

const expandReplaceAll = (
  operation: Record<string, unknown>,
  index: number,
  bridge: FolioAgentBridge,
): Result<unknown[], FolioCliError> => {
  const { find, replace, id } = operation;
  if (typeof find !== "string" || find === "") {
    return Result.err(invalidInput(`operations[${index}] (replaceAll) needs a non-empty find.`));
  }
  if (typeof replace !== "string") {
    return Result.err(invalidInput(`operations[${index}] (replaceAll) needs replace.`));
  }
  const found = findMatches(
    bridge,
    find,
    operation["matchCase"] === true,
    operation["wholeWord"] === true,
  );
  if (found.isErr()) return found;
  const { matches, truncated, total } = found.value;
  if (matches.length === 0) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.notFound,
        message: `operations[${index}] (replaceAll) found no ${JSON.stringify(find)}; nothing was written.`,
        hint: "Check the spelling, matchCase and wholeWord.",
      }),
    );
  }
  if (truncated) {
    return Result.err(
      invalidInput(
        `operations[${index}] (replaceAll) matches ${total} times, more than one batch can change.`,
        "Narrow find, or split the rename across calls with replaceInBlock.",
      ),
    );
  }
  const meta = metaOf(operation);
  return Result.ok(
    matches.map((match, occurrence) => {
      const replacement: Record<string, unknown> = {
        type: "replaceRange",
        range: match.range,
        replace,
      };
      if (typeof id === "string") replacement["id"] = `${id}.${occurrence + 1}`;
      return Object.assign(replacement, meta);
    }),
  );
};

const expandAddComment = (
  operation: Record<string, unknown>,
  index: number,
  bridge: FolioAgentBridge,
): Result<unknown, FolioCliError> => {
  const { blockId, quote, comment, id } = operation;
  if (typeof comment !== "string" || comment === "") {
    return Result.err(invalidInput(`operations[${index}] (addComment) needs comment text.`));
  }
  if (quote !== undefined && (typeof quote !== "string" || quote === "")) {
    return Result.err(invalidInput(`operations[${index}] (addComment) quote must be text.`));
  }
  const anchored = (target: string): Result<unknown, FolioCliError> =>
    Result.ok({
      type: "commentOnBlock",
      blockId: target,
      comment,
      ...(quote !== undefined && { quote }),
      ...(typeof id === "string" && { id }),
      ...metaOf(operation),
    });
  if (typeof blockId === "string" && blockId !== "") return anchored(blockId);
  if (quote === undefined) {
    return Result.err(invalidInput(`operations[${index}] (addComment) needs blockId or quote.`));
  }
  const found = findMatches(bridge, quote, true, false);
  if (found.isErr()) return found;
  const blocks = [...new Set(found.value.matches.map((match) => match.blockId))];
  if (blocks.length === 0) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.notFound,
        message: `operations[${index}] (addComment) quote ${JSON.stringify(quote)} is not in the document; nothing was written.`,
        hint: "Quote exact text from read_document, or pass blockId.",
      }),
    );
  }
  const [only] = blocks;
  if (blocks.length > 1 || only === undefined) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.ambiguousTarget,
        message: `operations[${index}] (addComment) quote ${JSON.stringify(quote)} is in ${blocks.length} blocks; nothing was written.`,
        hint: "Pass blockId, or quote longer text.",
        details: { blockIds: blocks.slice(0, 20) },
      }),
    );
  }
  return anchored(only);
};

/**
 * Expand the batch operations of a `suggest_changes` call against the
 * document the bridge holds. Contract operations pass through unchanged.
 */
export const expandBatchOperations = (
  operations: unknown,
  bridge: FolioAgentBridge,
): Result<ExpandedBatch, FolioCliError> => {
  if (!Array.isArray(operations)) {
    return Result.ok({ operations: [], options: {}, replaced: [] });
  }
  const expanded: unknown[] = [];
  const replaced: ExpandedBatch["replaced"] = [];
  let comments = false;
  let batchOperations = false;
  for (const [index, operation] of operations.entries()) {
    const type = isRecord(operation) ? operation["type"] : undefined;
    if (!isRecord(operation) || (type !== "replaceAll" && type !== "addComment")) {
      expanded.push(operation);
      continue;
    }
    batchOperations = true;
    const keys = checkKeys(operation, type, index);
    if (keys !== null) return Result.err(keys);
    if (type === "replaceAll") {
      const replacements = expandReplaceAll(operation, index, bridge);
      if (replacements.isErr()) return replacements;
      expanded.push(...replacements.value);
      replaced.push({ find: String(operation["find"]), count: replacements.value.length });
    } else {
      const comment = expandAddComment(operation, index, bridge);
      if (comment.isErr()) return comment;
      expanded.push(comment.value);
      comments = true;
    }
  }
  if (!batchOperations) return Result.ok({ operations, options: {}, replaced });
  if (expanded.length > MAX_OPERATIONS_PER_CALL_LIMIT) {
    return Result.err(
      invalidInput(
        `The batch expands to ${expanded.length} edits, over the ${MAX_OPERATIONS_PER_CALL_LIMIT}-edit limit; nothing was written.`,
        "Split it across calls.",
      ),
    );
  }
  return Result.ok({
    operations: expanded,
    options: {
      suggestChanges: {
        maxOperations: Math.max(DEFAULT_MAX_OPERATIONS_PER_CALL, expanded.length),
        ...(comments && {
          operationTypes: [...SUGGEST_CHANGES_OPERATION_TYPES, "commentOnBlock"],
        }),
      },
    },
    replaced,
  });
};
