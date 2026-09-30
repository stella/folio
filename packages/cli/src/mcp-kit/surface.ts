/**
 * A tool surface with lazy discovery. The frequent tools are listed directly
 * with compact schemas; the rest stay out of `tools/list` and are reached
 * through three capability tools, so a turn pays only for what is listed:
 *
 * - `list_capabilities` pages the unlisted tools: id, summary, access.
 * - `describe_capability` returns one tool's full input schema and guidance
 *   (listed tools included, whose listed schema is the compact one).
 * - `invoke_capability` runs one by id with `input` checked against that
 *   full schema; `validate_only` checks without running.
 *
 * Every tool is also callable by its own name. Calls are read through
 * `readToolInput` and answered with `toCallResult`'s one envelope.
 */

import {
  closestNames,
  didYouMean,
  failure,
  KIT_ERROR_CODES,
  success,
  toCallResult,
  validationError,
} from "./envelope";
import { readToolInput } from "./input";
import { compactSchema } from "./schema";
import type {
  JsonSchema,
  ListedTool,
  ToolAccess,
  ToolCallResult,
  ToolDefinition,
  ToolOutcome,
} from "./types";

export const CAPABILITY_TOOL_NAMES = {
  list: "list_capabilities",
  describe: "describe_capability",
  invoke: "invoke_capability",
} as const;

const DEFAULT_LIST_LIMIT = 50;
const MAX_LIST_LIMIT = 100;

const ACCESS_FILTERS = ["all", "read", "write"] as const;

const LIST_SCHEMA = {
  type: "object",
  properties: {
    domain: { type: "string" },
    access: { type: "string", enum: ACCESS_FILTERS },
    cursor: { type: "string" },
    limit: { type: "integer", minimum: 1, maximum: MAX_LIST_LIMIT },
  },
} as const;

const DESCRIBE_SCHEMA = {
  type: "object",
  properties: { capability: { type: "string" } },
  required: ["capability"],
} as const;

const INVOKE_SCHEMA = {
  type: "object",
  properties: {
    capability: { type: "string" },
    input: { type: "object" },
    validate_only: { type: "boolean" },
  },
  required: ["capability"],
} as const;

const READ_ONLY = { readOnlyHint: true, destructiveHint: false, openWorldHint: false };

export type ToolSurfaceOptions<Context> = {
  readonly tools: readonly ToolDefinition<Context>[];
};

export type ToolSurface<Context> = {
  /** What `tools/list` returns: the direct tools, then the capability tools. */
  readonly listTools: () => ListedTool[];
  readonly callTool: (name: string, args: unknown, context: Context) => Promise<ToolCallResult>;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const destructiveOf = (tool: ToolDefinition<unknown>): boolean =>
  tool.destructive ?? tool.access === "write";

const listedSchema = (schema: JsonSchema): ListedTool["inputSchema"] => ({
  ...compactSchema(schema),
  type: "object",
});

const encodeCursor = (id: string): string => Buffer.from(id, "utf8").toString("base64url");

const decodeCursor = (cursor: string): string | null => {
  const decoded = Buffer.from(cursor, "base64url").toString("utf8");
  return decoded.length > 0 && encodeCursor(decoded) === cursor ? decoded : null;
};

const hasJsonType = (value: unknown, type: string): boolean => {
  if (type === "object") return isRecord(value);
  if (type === "integer") return Number.isInteger(value);
  return typeof value === type;
};

type MetaArgs = { ok: true; args: Record<string, unknown> } | { ok: false; outcome: ToolOutcome };

/**
 * The capability tools' own arguments, read strictly: a misspelt flag or a
 * string where a boolean belongs would otherwise change what runs.
 */
const readMetaArgs = (
  tool: string,
  schema: {
    properties: Record<string, { type: string; enum?: readonly string[] }>;
    required?: readonly string[];
  },
  value: unknown,
): MetaArgs => {
  if (value !== undefined && !isRecord(value)) {
    return {
      ok: false,
      outcome: validationError(`${tool} takes a JSON object.`, []),
    };
  }
  const args: Record<string, unknown> = {};
  const issues: { path: string; message: string }[] = [];
  for (const [key, entry] of Object.entries(value ?? {})) {
    const property = schema.properties[key];
    if (property === undefined) {
      issues.push({ path: key, message: `Unknown parameter: ${key}` });
      continue;
    }
    if (entry === null) continue;
    if (
      !hasJsonType(entry, property.type) ||
      (property.enum !== undefined && !property.enum.includes(String(entry)))
    ) {
      issues.push({
        path: key,
        message: `Expected ${property.enum === undefined ? `a JSON ${property.type}` : property.enum.join(" | ")}.`,
      });
      continue;
    }
    args[key] = entry;
  }
  for (const key of schema.required ?? []) {
    if (args[key] === undefined && !issues.some(({ path }) => path === key)) {
      issues.push({ path: key, message: `Missing parameter: ${key}` });
    }
  }
  return issues.length > 0
    ? {
        ok: false,
        outcome: validationError(
          `${tool} arguments need clarification.`,
          issues,
          `Accepted parameters: ${Object.keys(schema.properties).join(", ")}.`,
        ),
      }
    : { ok: true, args };
};

/** Build a surface over `tools`. Names must be unique and not a capability tool's. */
export const createToolSurface = <Context>({
  tools,
}: ToolSurfaceOptions<Context>): ToolSurface<Context> => {
  const byName = new Map<string, ToolDefinition<Context>>();
  const reserved = new Set<string>(Object.values(CAPABILITY_TOOL_NAMES));
  for (const tool of tools) {
    if (byName.has(tool.name) || reserved.has(tool.name)) {
      throw new Error(`Tool name ${tool.name} is taken.`);
    }
    byName.set(tool.name, tool);
  }
  const lazy = tools
    .filter((tool) => tool.direct === undefined)
    .toSorted((a, b) => a.name.localeCompare(b.name));
  const names = [...byName.keys()];

  const unknownCapability = (id: string): ToolOutcome =>
    failure({
      code: KIT_ERROR_CODES.notFound,
      message: `No capability with id "${id}".`,
      hint: `${didYouMean(closestNames(id, names))}Call ${CAPABILITY_TOOL_NAMES.list} to browse the ids.`,
    });

  const run = async (
    tool: ToolDefinition<Context>,
    value: unknown,
    context: Context,
    validateOnly = false,
  ): Promise<ToolCallResult> => {
    const read = readToolInput({
      schema: tool.inputSchema,
      value,
      access: tool.access,
      exactProperties: tool.exactProperties,
    });
    if (!read.ok) {
      return toCallResult(validationError(read.message, read.issues, read.hint));
    }
    if (validateOnly) {
      return toCallResult(success({ valid: true, input: read.value }), read.notes);
    }
    let outcome: ToolOutcome;
    try {
      outcome = await tool.run(read.value, context);
    } catch (error) {
      outcome = failure({
        code: KIT_ERROR_CODES.internal,
        message: error instanceof Error ? error.message : String(error),
      });
    }
    return toCallResult(outcome, read.notes);
  };

  const listCapabilities = (value: unknown): ToolOutcome => {
    const read = readMetaArgs(CAPABILITY_TOOL_NAMES.list, LIST_SCHEMA, value);
    if (!read.ok) return read.outcome;
    const { domain, access = "all", cursor, limit = DEFAULT_LIST_LIMIT } = read.args;
    if (typeof limit !== "number" || limit < 1 || limit > MAX_LIST_LIMIT) {
      return validationError("limit is out of range.", [
        { path: "limit", message: `Expected 1 to ${MAX_LIST_LIMIT}.` },
      ]);
    }
    const after = typeof cursor === "string" ? decodeCursor(cursor) : undefined;
    if (after === null) {
      return validationError("cursor is not a nextCursor from list_capabilities.", [
        { path: "cursor", message: "Unrecognised cursor." },
      ]);
    }
    const matching = lazy.filter(
      (tool) =>
        (domain === undefined || tool.domain === domain) &&
        (access === "all" || tool.access === (access as ToolAccess)) &&
        (after === undefined || tool.name > after),
    );
    const page = matching.slice(0, limit);
    const last = page.at(-1);
    return success({
      items: page.map((tool) => ({
        id: tool.name,
        summary: tool.summary,
        access: tool.access,
        ...(destructiveOf(tool) && { destructive: true }),
      })),
      nextCursor:
        last !== undefined && matching.length > page.length ? encodeCursor(last.name) : null,
    });
  };

  const describeCapability = (value: unknown): ToolOutcome => {
    const read = readMetaArgs(CAPABILITY_TOOL_NAMES.describe, DESCRIBE_SCHEMA, value);
    if (!read.ok) return read.outcome;
    const id = String(read.args["capability"]);
    const tool = byName.get(id);
    if (tool === undefined) return unknownCapability(id);
    return success({
      id: tool.name,
      description: tool.guide === undefined ? tool.summary : `${tool.summary}\n${tool.guide}`,
      access: tool.access,
      destructive: destructiveOf(tool),
      inputSchema: tool.inputSchema,
    });
  };

  const invokeCapability = async (value: unknown, context: Context): Promise<ToolCallResult> => {
    const read = readMetaArgs(CAPABILITY_TOOL_NAMES.invoke, INVOKE_SCHEMA, value);
    if (!read.ok) return toCallResult(read.outcome);
    const id = String(read.args["capability"]);
    const tool = byName.get(id);
    if (tool === undefined) return toCallResult(unknownCapability(id));
    return await run(tool, read.args["input"] ?? {}, context, read.args["validate_only"] === true);
  };

  const listTools = (): ListedTool[] => {
    const listed: ListedTool[] = tools.flatMap((tool) =>
      tool.direct === undefined
        ? []
        : [
            {
              name: tool.name,
              description: tool.summary,
              inputSchema: listedSchema(tool.direct.inputSchema),
              annotations: {
                readOnlyHint: tool.access === "read",
                destructiveHint: destructiveOf(tool),
                openWorldHint: false,
              },
            },
          ],
    );
    if (lazy.length === 0) return listed;
    return [
      ...listed,
      {
        name: CAPABILITY_TOOL_NAMES.list,
        description: `List tools not shown here (${lazy.map(({ name }) => name).join(", ")}).`,
        inputSchema: listedSchema(LIST_SCHEMA),
        annotations: READ_ONLY,
      },
      {
        name: CAPABILITY_TOOL_NAMES.describe,
        description: "Full input schema and guidance for any tool, by id.",
        inputSchema: listedSchema(DESCRIBE_SCHEMA),
        annotations: READ_ONLY,
      },
      {
        name: CAPABILITY_TOOL_NAMES.invoke,
        description: "Run a tool by id with its arguments as `input`.",
        inputSchema: listedSchema(INVOKE_SCHEMA),
        annotations: {
          readOnlyHint: false,
          destructiveHint: lazy.some(destructiveOf),
          openWorldHint: false,
        },
      },
    ];
  };

  const callTool = async (
    name: string,
    args: unknown,
    context: Context,
  ): Promise<ToolCallResult> => {
    switch (name) {
      case CAPABILITY_TOOL_NAMES.list:
        return toCallResult(listCapabilities(args));
      case CAPABILITY_TOOL_NAMES.describe:
        return toCallResult(describeCapability(args));
      case CAPABILITY_TOOL_NAMES.invoke:
        return await invokeCapability(args, context);
      default: {
        const tool = byName.get(name);
        if (tool !== undefined) return await run(tool, args, context);
        return toCallResult(
          failure({
            code: KIT_ERROR_CODES.unknownTool,
            message: `Unknown tool "${name}".`,
            hint: `${didYouMean(closestNames(name, names))}Call ${CAPABILITY_TOOL_NAMES.list} to browse the rest.`,
          }),
        );
      }
    }
  };

  return { listTools, callTool };
};
