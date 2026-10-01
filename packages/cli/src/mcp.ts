/**
 * `folio mcp`: the file tool registry as a Model Context Protocol server over
 * stdio. The protocol owns stdout; diagnostics go to stderr.
 *
 * A model pays for `tools/list` on every turn, so the server lists only the
 * frequent tools, with compact schemas, and reaches the rest through
 * `list_capabilities` / `describe_capability` / `invoke_capability`
 * (`mcp-kit`). Every tool takes the file envelope (`path`, `fileVersion`, and
 * for writes `destination`, `txId`, ...) beside its own arguments; the listed
 * schemas show the part a call usually needs and `describe_capability` shows
 * all of it. Every path must resolve inside an allowed root. Results carry
 * what the next call needs (the new `fileVersion`, ids, counts); a failure is
 * `{ error: { code, message, hint, retryable } }`.
 */

import { panic, Result } from "better-result";
import type { Readable, Writable } from "node:stream";

import { FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA } from "@stll/folio-agents/operation-schema";
import {
  describeSuggestChangesCapabilities,
  SUGGEST_CHANGES_OPERATION_TYPES,
} from "@stll/folio-agents/tools";
import { FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE } from "@stll/folio-core/server";
import {
  ProtocolError,
  ProtocolErrorCode,
  Server,
  type JSONObject,
  type JSONValue,
  type ReadResourceResult,
  type Tool,
} from "@modelcontextprotocol/server";
import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";

import packageJson from "../package.json" with { type: "json" };
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { executeReadTool, type FileToolCall, type FolioReadBounds } from "./execute-read";
import { executeWriteTool, type WriteDestination } from "./execute-write";
import {
  compactSchema,
  createToolSurface,
  failure,
  hoistRepeatedSchemas,
  success,
  type JsonSchema,
  type ToolDefinition,
  type ToolOutcome,
} from "./mcp-kit";
import {
  FOLIO_FILE_TOOLS,
  toolAccess,
  type FolioFileToolSpec,
  type JsonObjectSchema,
} from "./registry";
import { checkWithinRoots, type AllowedRoots } from "./roots";

/** Limits on what one MCP tool call returns. */
export const MCP_READ_BOUNDS: FolioReadBounds = {
  defaultMaxBlocks: 200,
  maxMatches: 100,
  maxResponseBytes: 256 * 1024,
};

export type FolioMcpServerOptions = {
  roots: AllowedRoots;
  /** Author for every change; `undefined` refuses writes with `author_required`. */
  author: string | undefined;
  bounds?: FolioReadBounds;
  now?: () => Date;
};

const ABOUT_URI = "folio://about";
const OPERATIONS_SCHEMA_URI = "folio://schema/operations";

const instructionsFor = (author: string | undefined): string =>
  "Reads and edits .docx files. read_document returns the fileVersion and a `[blockId] text` line " +
  "per block; a write needs the latest fileVersion and returns the next. Block ids stay valid " +
  "across edits, so chain calls without re-reading. Edits are tracked changes. One suggest_changes " +
  "batch can do a whole review: replaceAll {find, replace} renames everywhere, tables included; " +
  "addComment {comment, quote or blockId} comments alongside. A successful write is authoritative " +
  "(fileVersion, author, applied counts): no verification read or shell check is needed. " +
  (author === undefined
    ? "No author is configured, so writes are refused. "
    : `Every change and comment is recorded under the author ${JSON.stringify(author)}. `) +
  "Failures are { error: { code, message, hint, retryable } }.";

const PATH_PROPERTY = {
  type: "string",
  description: "The .docx file, absolute or relative to the first allowed root.",
};

const FILE_VERSION_PATTERN = "^[0-9a-f]{64}$";

const envelopeProperties = (tool: FolioFileToolSpec): Record<string, unknown> => {
  const access = toolAccess(tool);
  const properties: Record<string, unknown> = {
    path: PATH_PROPERTY,
    fileVersion: {
      type: "string",
      pattern: FILE_VERSION_PATTERN,
      description: {
        write:
          "The file's fileVersion (SHA-256) from your latest read or write. Required: a change to a file that moved on is refused.",
        readOrWrite:
          "The file's fileVersion (SHA-256) from your latest read; required with `destination`.",
        read: "When given, refuse unless the file still has this fileVersion.",
      }[access],
    },
  };
  if (access === "read") return properties;
  Object.assign(properties, {
    destination: {
      type: "string",
      description:
        access === "write"
          ? "Write the result to this .docx instead of changing `path` in place. It must not start with a dot or sit inside .folio."
          : "Write the redline to this .docx. Without it, only the differences are returned.",
    },
    overwrite: {
      type: "boolean",
      description:
        "Let `destination` replace an existing .docx; needs `expectedDestinationVersion`. The replaced file is backed up.",
    },
    expectedDestinationVersion: {
      type: "string",
      pattern: FILE_VERSION_PATTERN,
      description: "The fileVersion of the existing file `destination` replaces.",
    },
    txId: {
      type: "string",
      pattern: "^[\\w.-]{1,128}$",
      description: "Idempotency key: repeating a committed call returns its original result.",
    },
  });
  if (access === "write") {
    Object.assign(properties, {
      allowRepack: {
        type: "boolean",
        description:
          "Allow rewriting the whole package when the change cannot be patched in (e.g. paragraphs added or removed).",
      },
    });
  }
  if (tool.type === "agentWrite" && tool.editMode === "tracked-or-direct") {
    Object.assign(properties, {
      mode: {
        type: "string",
        enum: ["tracked", "direct"],
        description: "tracked (default) records tracked changes; direct edits the text.",
      },
    });
  }
  return properties;
};

/** Arguments only the MCP surface takes, beside the registry's. */
const MCP_ONLY_PROPERTIES: Readonly<Record<string, Readonly<Record<string, unknown>>>> = {
  read_document: {
    formatting: {
      type: "boolean",
      description:
        "Return each block's fields (kind, displayLabel, headingLevel, listLevel, blockTextHash, blockIdSource, tableCell) instead of `[id] text` lines.",
    },
  },
};

/** The tool's full MCP input schema: its own arguments plus the file envelope. */
export const mcpInputSchema = (tool: FolioFileToolSpec): JsonObjectSchema => {
  const envelope = envelopeProperties(tool);
  for (const key of Object.keys(envelope)) {
    if (key in tool.argsSchema.properties) {
      panic(`${tool.name} declares ${key}, which the file envelope also uses`);
    }
  }
  const access = toolAccess(tool);
  return {
    type: "object",
    properties: {
      ...envelope,
      ...tool.argsSchema.properties,
      ...MCP_ONLY_PROPERTIES[tool.name],
    },
    required: ["path", ...(access === "write" ? ["fileVersion"] : []), ...tool.argsSchema.required],
    additionalProperties: false,
  };
};

/** Switches that widen what a write may do: they count only as JSON `true`. */
const EXACT_PROPERTIES = ["overwrite", "allowRepack"] as const;

const STRING = { type: "string" } as const;
const BOOLEAN = { type: "boolean" } as const;

type DirectTool = {
  summary: string;
  properties: Readonly<Record<string, unknown>>;
  required: readonly string[];
};

/**
 * The tools listed on every turn, with the schema they are listed with. Each
 * accepts its full schema too; `describe_capability` shows it.
 */
const DIRECT_TOOLS: Readonly<Record<string, DirectTool>> = {
  read_document: {
    summary:
      "Read blocks as `[id] text` lines (a table row as `| [id] cell | [id] cell |`), with the fileVersion writes need. When `nextCursor` is set, pass it as `cursor` for the next page.",
    properties: {
      path: STRING,
      cursor: STRING,
      maxBlocks: { type: "integer" },
      formatting: BOOLEAN,
    },
    required: ["path"],
  },
  find_text: {
    summary:
      "Find text. Each match has its blockId, a `range` to pass to replaceRange or commentOnRange, and context.",
    properties: { path: STRING, query: STRING, matchCase: BOOLEAN, wholeWord: BOOLEAN },
    required: ["path", "query"],
  },
  suggest_changes: {
    summary:
      "Apply edits as tracked changes in one batch that lands whole or not at all; returns the new " +
      "fileVersion. Operations: replaceAll {find, replace, matchCase?, wholeWord?} (every match); " +
      "replaceInBlock {blockId, find, replace}; replaceRange {range, replace}; replaceBlock " +
      "{blockId, text}; insertAfterBlock / insertBeforeBlock {blockId, text}; deleteBlock {blockId}; " +
      "addComment {comment, quote or blockId}; commentOnRange {range, comment}. Others: describe_capability.",
    properties: {
      path: STRING,
      fileVersion: STRING,
      operations: {
        type: "array",
        items: { type: "object", properties: { type: STRING }, required: ["type"] },
      },
    },
    required: ["path", "fileVersion", "operations"],
  },
  add_comment: {
    summary:
      "Comment on a block; `quote` anchors it to exact text in the block. Returns the new fileVersion.",
    properties: { path: STRING, fileVersion: STRING, blockId: STRING, text: STRING, quote: STRING },
    required: ["path", "fileVersion", "blockId", "text"],
  },
  read_changes: {
    summary: "List pending tracked changes.",
    properties: { path: STRING },
    required: ["path"],
  },
  read_comments: {
    summary: "List comment threads.",
    properties: { path: STRING, filter: { type: "string", enum: ["all", "open", "resolved"] } },
    required: ["path"],
  },
};

/** One-line summaries for the tools reached through the capability tools. */
const LAZY_SUMMARIES: Readonly<Record<string, { summary: string; domain: string }>> = {
  get_document_outline: { summary: "Heading outline with section handles.", domain: "read" },
  read_section: { summary: "Read one section by its outline handle.", domain: "read" },
  list_stories: { summary: "List header, footer and note stories.", domain: "read" },
  read_story: { summary: "Read one story by its handle.", domain: "read" },
  reply_comment: { summary: "Reply to a comment thread.", domain: "comments" },
  resolve_comment: { summary: "Resolve or reopen a comment thread.", domain: "comments" },
  resolve_changes: { summary: "Accept or reject tracked changes.", domain: "changes" },
  compare_documents: {
    summary: "Diff two .docx files, or write their redline to a destination.",
    domain: "compare",
  },
};

const ENVELOPE_KEYS: ReadonlySet<string> = new Set([
  "path",
  "fileVersion",
  "destination",
  "overwrite",
  "expectedDestinationVersion",
  "txId",
  "allowRepack",
  "mode",
  "formatting",
]);

type ToolCallContext = FolioMcpServerOptions & { bounds: FolioReadBounds; now: () => Date };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const invalidInput = (message: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message });

const optionalString = (value: unknown, name: string): Result<string | undefined, FolioCliError> =>
  value === undefined || typeof value === "string"
    ? Result.ok(value)
    : Result.err(invalidInput(`${name} must be a string.`));

const toRevisionDate = (date: Date): string => date.toISOString().replace(/\.\d{3}Z$/u, "Z");

const runTool = async (
  tool: FolioFileToolSpec,
  input: Readonly<Record<string, unknown>>,
  context: ToolCallContext,
): Promise<Result<unknown, FolioCliError>> => {
  const args: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(input)) {
    if (!ENVELOPE_KEYS.has(key)) args[key] = value;
  }
  const rawPath = input["path"];
  if (typeof rawPath !== "string" || rawPath === "") {
    return Result.err(invalidInput("path is required."));
  }
  const filePath = await checkWithinRoots({
    roots: context.roots,
    target: rawPath,
    argument: "path",
  });
  if (filePath.isErr()) return filePath;
  const fileVersion = optionalString(input["fileVersion"], "fileVersion");
  if (fileVersion.isErr()) return fileVersion;
  const destination = optionalString(input["destination"], "destination");
  if (destination.isErr()) return destination;
  const txId = optionalString(input["txId"], "txId");
  if (txId.isErr()) return txId;

  if (tool.type === "compare") {
    const revised = args["revisedPath"];
    if (typeof revised !== "string") return Result.err(invalidInput("revisedPath is required."));
    const revisedPath = await checkWithinRoots({
      roots: context.roots,
      target: revised,
      argument: "revisedPath",
    });
    if (revisedPath.isErr()) return revisedPath;
    args["revisedPath"] = revisedPath.value;
  }

  const call: FileToolCall = { path: filePath.value, fileVersion: fileVersion.value, args };
  const access = toolAccess(tool);
  if (access === "read" || (access === "readOrWrite" && destination.value === undefined)) {
    return await executeReadTool(tool, call, context.bounds);
  }
  if (context.author === undefined) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.authorRequired,
        message: "This server has no author configured for changes.",
        hint: "Restart folio mcp with --author, FOLIO_AUTHOR, or a git user.name.",
      }),
    );
  }
  const expectedDestinationVersion = optionalString(
    input["expectedDestinationVersion"],
    "expectedDestinationVersion",
  );
  if (expectedDestinationVersion.isErr()) return expectedDestinationVersion;
  let target: WriteDestination = { type: "inPlace" };
  if (destination.value !== undefined) {
    const checked = await checkWithinRoots({
      roots: context.roots,
      target: destination.value,
      argument: "destination",
    });
    if (checked.isErr()) return checked;
    target = {
      type: "file",
      path: checked.value,
      overwrite: input["overwrite"] === true,
      expectedVersion: expectedDestinationVersion.value,
    };
  }
  // Every MCP write names the version it read: the precondition is never waived.
  return await executeWriteTool(tool, call, {
    sourcePrecondition: "required",
    destination: target,
    author: context.author,
    date: toRevisionDate(context.now()),
    repack: input["allowRepack"] === true ? "allow" : "refuse",
    force: false,
    txId: txId.value,
    journalPath: undefined,
    mode: input["mode"] === "direct" ? "direct" : "tracked-changes",
  });
};

// --- compact results ---------------------------------------------------------

const arrayOf = (value: unknown): unknown[] => (Array.isArray(value) ? value : []);

const recordOf = (value: unknown): Record<string, unknown> => (isRecord(value) ? value : {});

/**
 * A number or bullet worth showing. An unnumbered heading's label is its
 * style id (`Heading1`, `Title`), which says nothing the `(hN)` mark does
 * not; every other label (`iii.`, `Article 1`, `•`) is shown.
 */
const isNumberLabel = (label: unknown, heading: boolean): label is string =>
  typeof label === "string" && !(heading && /^\p{L}{3,}\d*$/u.test(label));

/** One block as the line a model reads: `[id] (h2) 2.1 text`. */
const blockLine = (block: unknown): string => {
  const row = recordOf(block);
  const label = row["displayLabel"];
  const isHeading = typeof row["headingLevel"] === "number";
  const heading = isHeading ? `(h${String(row["headingLevel"])}) ` : "";
  const number = isNumberLabel(label, isHeading) ? `${label} ` : "";
  return `[${String(row["blockId"])}] ${heading}${number}${String(row["text"] ?? "")}`;
};

/**
 * Blocks as lines, a table row on one line: `| [id] cell | [id] cell |`,
 * a cell's paragraphs side by side. Empty cells hold no block and show none.
 */
export const blockLines = (blocks: readonly unknown[]): string => {
  const lines: string[] = [];
  let row: { key: string; cells: string[][]; cell: unknown } | undefined;
  const flush = (): void => {
    if (row !== undefined) lines.push(`| ${row.cells.map((cell) => cell.join(" ")).join(" | ")} |`);
    row = undefined;
  };
  for (const block of blocks) {
    const location = recordOf(recordOf(block)["tableCell"]);
    if (typeof location["table"] !== "number") {
      flush();
      lines.push(blockLine(block));
      continue;
    }
    const key = `${location["table"]}:${String(location["row"])}`;
    if (row?.key !== key) {
      flush();
      row = { key, cells: [], cell: undefined };
    }
    if (row.cell !== location["cell"] || row.cells.length === 0) {
      row.cells.push([]);
      row.cell = location["cell"];
    }
    row.cells.at(-1)?.push(blockLine(block));
  }
  flush();
  return lines.join("\n");
};

/** A main-story range without the fields the server fills back in. */
const compactRange = (range: unknown): unknown => {
  const { type, story, ...rest } = recordOf(range);
  return type === "textRange" && story === "main" ? rest : range;
};

const compactMatch = (match: unknown): unknown => {
  const row = recordOf(match);
  if (isRecord(row["range"])) {
    return { range: compactRange(row["range"]), context: row["context"] };
  }
  const { blockTextHash: _hash, type: _type, ...rest } = row;
  return rest;
};

const CHANGE_LINE_KEYS: ReadonlySet<string> = new Set([
  "id",
  "type",
  "blockId",
  "text",
  "author",
  "date",
]);

/** One tracked change as a line: `12 insertion [blockId] "text" author`, then any other fields. */
const changeLine = (change: unknown): string => {
  const row = recordOf(change);
  const extra = Object.entries(row)
    .filter(([key]) => !CHANGE_LINE_KEYS.has(key))
    .map(([key, value]) => `${key}=${JSON.stringify(value)}`);
  return [
    String(row["id"]),
    String(row["type"]),
    ...(typeof row["blockId"] === "string" ? [`[${row["blockId"]}]`] : []),
    JSON.stringify(row["text"] ?? ""),
    ...(typeof row["author"] === "string" ? [row["author"]] : []),
    ...extra,
  ].join(" ");
};

const commentIdsOf = (receipts: unknown): string[] =>
  arrayOf(receipts).flatMap((receipt) =>
    arrayOf(recordOf(receipt)["affected"]).flatMap((affected) => {
      const entry = recordOf(affected);
      return entry["type"] === "comment" ? [String(entry["commentId"])] : [];
    }),
  );

type CompactOptions = { tool: FolioFileToolSpec; input: Readonly<Record<string, unknown>> };

/** A read's result, cut to what the next call needs. */
const compactRead = ({ tool, input }: CompactOptions, data: Record<string, unknown>): unknown => {
  const { fileVersion } = data;
  const result = data["result"];
  switch (tool.name) {
    case "read_document": {
      const page = recordOf(result);
      const blocks = arrayOf(page["blocks"]);
      return {
        fileVersion,
        blocks: input["formatting"] === true ? blocks : blockLines(blocks),
        ...(page["nextCursor"] !== undefined && {
          nextCursor: page["nextCursor"],
          totalBlocks: page["totalBlocks"],
        }),
      };
    }
    case "find_text": {
      const found = recordOf(result);
      return {
        fileVersion,
        matches: arrayOf(found["matches"]).map(compactMatch),
        ...(found["truncated"] === true && {
          truncated: true,
          totalMatches: found["totalMatches"],
        }),
      };
    }
    case "read_changes":
      return { fileVersion, changes: arrayOf(result).map(changeLine).join("\n") };
    case "compare_documents":
      return { fileVersion, revised: data["revised"], diff: recordOf(result)["text"] };
    default:
      return isRecord(result) ? { fileVersion, ...result } : { fileVersion, result };
  }
};

/** A write's receipt, cut to the new version and what the change produced. */
const compactWrite = (
  { tool, input }: CompactOptions,
  receipt: Record<string, unknown>,
): unknown => {
  const result = recordOf(receipt["result"]);
  const base = {
    fileVersion: receipt["fileVersion"],
    author: receipt["author"],
    ...(input["destination"] !== undefined && { path: receipt["path"] }),
    ...(receipt["status"] === "replayed" && { replayed: true }),
    ...(receipt["rebased"] !== undefined && { rebased: receipt["rebased"] }),
  };
  switch (tool.name) {
    case "suggest_changes": {
      const normalizations = arrayOf(result["normalizations"]);
      const replaced = arrayOf(result["replaced"]);
      const commentIds = commentIdsOf(result["receipts"]);
      return {
        ...base,
        applied: arrayOf(result["applied"]).length,
        ...(replaced.length > 0 && { replaced }),
        ...(commentIds.length > 0 && { commentIds }),
        ...(normalizations.length > 0 && { normalizations }),
      };
    }
    case "add_comment":
    case "reply_comment":
      return { ...base, commentId: commentIdsOf(result["receipts"]).at(0) };
    case "resolve_changes":
      return { ...base, resolved: result["resolved"], remaining: result["remaining"] };
    case "compare_documents":
      return { ...base, ...result };
    default:
      return base;
  }
};

const RETRYABLE_CODES: ReadonlySet<string> = new Set([
  FOLIO_CLI_ERROR_CODES.staleVersion,
  FOLIO_CLI_ERROR_CODES.staleTarget,
  FOLIO_CLI_ERROR_CODES.locked,
]);

/** Refusal details, without the parts that only restate the message. */
const compactDetails = (details: unknown): unknown => {
  if (!isRecord(details) || !Array.isArray(details["skipped"])) return details;
  return { skipped: details["skipped"] };
};

const toOutcome = (
  options: CompactOptions,
  result: Result<unknown, FolioCliError>,
): ToolOutcome => {
  if (result.isErr()) {
    const { code, message, hint, details } = result.error;
    return failure({
      code,
      message,
      hint,
      retryable: RETRYABLE_CODES.has(code),
      details: compactDetails(details),
    });
  }
  const data = recordOf(result.value);
  const wrote = "txId" in data;
  return success(wrote ? compactWrite(options, data) : compactRead(options, data));
};

// --- the surface -------------------------------------------------------------

/** Fill in a main-story range's `type` and `story`, which find_text leaves out. */
const withFullRanges = (operations: unknown): unknown =>
  Array.isArray(operations)
    ? operations.map((operation) => {
        if (!isRecord(operation) || !isRecord(operation["range"])) return operation;
        return {
          ...operation,
          range: { type: "textRange", story: "main", ...operation["range"] },
        };
      })
    : operations;

/** Operation keys the model never writes. */
const HIDDEN_OPERATION_KEYS: ReadonlySet<string> = new Set(["type", "suggestionId"]);

/**
 * The fields each operation type accepts, one line per type. The described
 * schema lists every field once for all types, without their descriptions.
 */
const OPERATION_FIELDS = [
  "replaceAll: id, find, replace, matchCase, wholeWord, severity, area",
  "addComment: id, comment, blockId, quote, severity, area",
  ...SUGGEST_CHANGES_OPERATION_TYPES.map(
    (type) =>
      `${type}: ${FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE[type]
        .filter((key) => !HIDDEN_OPERATION_KEYS.has(key))
        .join(", ")}`,
  ),
].join("\n");

/** What the compact `describe_capability` says about suggest_changes. */
const SUGGEST_CHANGES_BRIEF = [
  "Fields (* required):",
  "replaceAll {find*, replace*, matchCase, wholeWord}: every match in the body and tables; run formatting is kept",
  "replaceInBlock {blockId*, find*, replace*}: find must occur once in the block",
  "replaceRange {range*, replace*}: range from find_text",
  "replaceBlock {blockId*, text*, preserveFormatting}",
  "insertAfterBlock | insertBeforeBlock {blockId*, text*, styleId}",
  "deleteBlock {blockId*}",
  "addComment {comment*, blockId | quote}: a quote alone anchors to the one block containing it",
  "commentOnRange {range*, comment*}",
  "formatRange {range*, formatting* {bold, italic, underline}}",
  'Edits take an optional comment. mode "direct" edits without tracking. Split, merge and table operations: detail "full".',
].join("\n");

const SUGGEST_CHANGES_EXAMPLE = {
  path: "contract.docx",
  fileVersion: "<from read_document>",
  operations: [
    { type: "replaceAll", find: "Supplier", replace: "Provider", matchCase: true, wholeWord: true },
    { type: "addComment", quote: "total liability", comment: "Is this cap acceptable?" },
  ],
};

const guideOf = (tool: FolioFileToolSpec): string =>
  tool.name === "suggest_changes"
    ? `${describeSuggestChangesCapabilities()}\nFields by operation type:\n${OPERATION_FIELDS}`
    : tool.description;

const toKitTool = (tool: FolioFileToolSpec): ToolDefinition<ToolCallContext> => {
  const direct = DIRECT_TOOLS[tool.name];
  const lazy = LAZY_SUMMARIES[tool.name];
  const summary = direct?.summary ?? lazy?.summary ?? panic(`${tool.name} has no MCP summary`);
  const access = toolAccess(tool);
  return {
    name: tool.name,
    summary,
    guide: guideOf(tool),
    ...(tool.name === "suggest_changes" && {
      brief: SUGGEST_CHANGES_BRIEF,
      example: SUGGEST_CHANGES_EXAMPLE,
    }),
    access: access === "read" ? "read" : "write",
    destructive: access !== "read",
    ...(lazy !== undefined && { domain: lazy.domain }),
    inputSchema: mcpInputSchema(tool),
    describedSchema: hoistRepeatedSchemas(
      compactSchema(mcpInputSchema(tool), { describedDepth: 1 }),
    ),
    ...(direct !== undefined && {
      direct: {
        inputSchema: {
          type: "object",
          properties: direct.properties,
          required: [...direct.required],
        } satisfies JsonSchema,
      },
    }),
    exactProperties: EXACT_PROPERTIES,
    run: async (args, context) => {
      const input =
        tool.name === "suggest_changes"
          ? { ...args, operations: withFullRanges(args["operations"]) }
          : args;
      const result = await Result.tryPromise({
        try: () => runTool(tool, input, context),
        catch: (error) =>
          cliError({
            code: FOLIO_CLI_ERROR_CODES.internal,
            message: error instanceof Error ? error.message : String(error),
          }),
      });
      return toOutcome({ tool, input }, result.isOk() ? result.value : result);
    },
  };
};

const surface = createToolSurface({ tools: FOLIO_FILE_TOOLS.map(toKitTool) });

const isJsonValue = (value: unknown): value is JSONValue => {
  if (value === null || typeof value === "string" || typeof value === "boolean") return true;
  if (typeof value === "number") return Number.isFinite(value);
  if (Array.isArray(value)) return value.every(isJsonValue);
  return typeof value === "object" && Object.values(value).every(isJsonValue);
};

/** The schemas are JSON by construction; a value that is not is a registry bug. */
const toJsonObject = (properties: Readonly<Record<string, unknown>>): JSONObject => {
  const object: JSONObject = {};
  for (const [key, value] of Object.entries(properties)) {
    object[key] = isJsonValue(value) ? value : panic(`Schema property ${key} is not JSON`);
  }
  return object;
};

const toMcpInputSchema = (schema: Readonly<Record<string, unknown>>): Tool["inputSchema"] => {
  const required = arrayOf(schema["required"]).filter((key) => typeof key === "string");
  return {
    type: "object",
    properties: toJsonObject(recordOf(schema["properties"])),
    ...(required.length > 0 && { required }),
  };
};

/** What `tools/list` returns. */
export const listMcpTools = (): Tool[] =>
  surface.listTools().map(({ name, description, inputSchema, annotations }) => ({
    name,
    description,
    inputSchema: toMcpInputSchema(inputSchema),
    annotations,
  }));

const aboutText = (roots: AllowedRoots): string =>
  [
    `# folio ${packageJson.version}`,
    "",
    `Allowed roots: ${roots.join(", ")}`,
    "",
    "- fileVersion is the SHA-256 of the file's bytes; changes require the one you read.",
    "- Block ids are the paragraph's w14:paraId. A paragraph without one is given one derived from",
    "  the file's bytes, which the first change writes into the file, so ids survive edits.",
    "- Offsets in range handles are UTF-16 code units into the block text read_document returns.",
    "- A change writes atomically in place or to destination (a plain .docx, never a dotfile or inside",
    "  .folio); replacing an existing file needs overwrite plus expectedDestinationVersion. Whatever is",
    "  replaced is backed up in .folio/backups/<name>/. Changes are journaled in .folio/journal.jsonl and",
    "  refuse a full package repack unless allowRepack is true.",
    "- A batch lands whole or not at all: stale_target, ambiguous_target, and other refusals write nothing.",
    "",
  ].join("\n");

const readResource = (uri: string, roots: AllowedRoots): ReadResourceResult | null => {
  if (uri === ABOUT_URI) {
    return { contents: [{ uri, mimeType: "text/markdown", text: aboutText(roots) }] };
  }
  if (uri === OPERATIONS_SCHEMA_URI) {
    return {
      contents: [
        {
          uri,
          mimeType: "application/schema+json",
          text: JSON.stringify(FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA),
        },
      ],
    };
  }
  return null;
};

/** A low-level MCP server exposing the file tools; connect it to a transport. */
export const createFolioMcpServer = (options: FolioMcpServerOptions): Server => {
  const context: ToolCallContext = {
    ...options,
    bounds: options.bounds ?? MCP_READ_BOUNDS,
    now: options.now ?? (() => new Date()),
  };
  const server = new Server(
    { name: "folio", version: packageJson.version },
    { capabilities: { tools: {}, resources: {} }, instructions: instructionsFor(options.author) },
  );
  server.setRequestHandler("tools/list", () => ({ tools: listMcpTools() }));
  server.setRequestHandler("tools/call", (request) =>
    surface.callTool(request.params.name, request.params.arguments, context),
  );
  server.setRequestHandler("resources/list", () => ({
    resources: [
      { uri: ABOUT_URI, name: "about", mimeType: "text/markdown" },
      {
        uri: OPERATIONS_SCHEMA_URI,
        name: "operations-schema",
        mimeType: "application/schema+json",
      },
    ],
  }));
  server.setRequestHandler("resources/read", (request) => {
    const resource = readResource(request.params.uri, context.roots);
    if (resource === null) {
      // The protocol's own error channel: the SDK answers with a JSON-RPC error.
      throw new ProtocolError(
        ProtocolErrorCode.InvalidParams,
        `Unknown resource: ${request.params.uri}`,
      );
    }
    return resource;
  });
  return server;
};

type ServeOptions = FolioMcpServerOptions & { input: Readable; output: Writable };

/** Serve the file tools over stdio until the client disconnects. */
export const serveFolioMcp = async ({ input, output, ...options }: ServeOptions): Promise<void> => {
  const server = createFolioMcpServer(options);
  const closed = new Promise<void>((resolve) => {
    server.onclose = () => resolve();
  });
  await server.connect(new StdioServerTransport(input, output));
  await closed;
};
