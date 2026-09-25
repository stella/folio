/**
 * `folio mcp`: the file tool registry as a Model Context Protocol server over
 * stdio. The protocol owns stdout; diagnostics go to stderr. Every tool takes
 * the file envelope (`path`, `fileVersion`, and for writes `destination`,
 * `txId`, ...) beside its own arguments, every path must resolve inside an
 * allowed root, and results are the same `{ ok, data | error }` envelope the
 * command line prints.
 */

import { panic, Result } from "better-result";
import type { Readable, Writable } from "node:stream";

import { FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA } from "@stll/folio-agents/operation-schema";
import {
  ProtocolError,
  ProtocolErrorCode,
  Server,
  type CallToolResult,
  type ReadResourceResult,
  type Tool,
} from "@modelcontextprotocol/server";
import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";

import packageJson from "../package.json" with { type: "json" };
import { isFileVersion } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { executeReadTool, type FileToolCall, type FolioReadBounds } from "./execute-read";
import { executeWriteTool } from "./execute-write";
import { failureEnvelope, successEnvelope } from "./output";
import {
  findFileTool,
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

const INSTRUCTIONS =
  "Tools read and change .docx files on disk. Start with get_document_outline or read_document to get " +
  "block ids and the file's fileVersion; find_text returns exact range handles. Mutating tools require " +
  "the fileVersion you read and refuse (stale_version) when the file changed; re-read and retry. " +
  "Edits are tracked changes unless mode is direct. Without destination a change is written in place " +
  "with a backup; with destination it goes to that new file. Results are { ok, data } or " +
  "{ ok: false, error: { code, message, hint } }.";

const PATH_PROPERTY = {
  type: "string",
  description: "The .docx file, absolute or relative to the first allowed root.",
};

const envelopeProperties = (tool: FolioFileToolSpec): Record<string, unknown> => {
  const access = toolAccess(tool);
  const properties: Record<string, unknown> = {
    path: PATH_PROPERTY,
    fileVersion: {
      type: "string",
      pattern: "^[0-9a-f]{64}$",
      description:
        access === "write"
          ? "The file's fileVersion (SHA-256) from your latest read. Required: a change to a file that moved on is refused."
          : "When given, refuse unless the file still has this fileVersion.",
    },
  };
  if (access === "read") return properties;
  Object.assign(properties, {
    destination: {
      type: "string",
      description:
        access === "write"
          ? "Write the result to this new file instead of changing `path` in place."
          : "Write the redline to this new file. Without it, only the differences are returned.",
    },
    overwrite: { type: "boolean", description: "Let `destination` replace an existing file." },
    txId: {
      type: "string",
      pattern: "^[\\w.-]{1,128}$",
      description: "Idempotency key: repeating a committed call returns its original receipt.",
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

/** The tool's MCP input schema: its own arguments plus the file envelope. */
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
    properties: { ...envelope, ...tool.argsSchema.properties },
    required: ["path", ...(access === "write" ? ["fileVersion"] : []), ...tool.argsSchema.required],
    additionalProperties: false,
  };
};

export const listMcpTools = (): Tool[] =>
  FOLIO_FILE_TOOLS.map((tool) => ({
    name: tool.name,
    description: tool.description,
    inputSchema: mcpInputSchema(tool),
    annotations: {
      readOnlyHint: toolAccess(tool) === "read",
      destructiveHint: false,
      openWorldHint: false,
    },
  }));

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const invalidInput = (message: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message });

const ENVELOPE_KEYS: ReadonlySet<string> = new Set([
  "path",
  "fileVersion",
  "destination",
  "overwrite",
  "txId",
  "allowRepack",
  "mode",
]);

type ToolCallContext = FolioMcpServerOptions & { bounds: FolioReadBounds; now: () => Date };

const optionalString = (value: unknown, name: string): Result<string | undefined, FolioCliError> =>
  value === undefined || typeof value === "string"
    ? Result.ok(value)
    : Result.err(invalidInput(`${name} must be a string.`));

const toRevisionDate = (date: Date): string => date.toISOString().replace(/\.\d{3}Z$/u, "Z");

const runTool = async (
  tool: FolioFileToolSpec,
  rawArgs: unknown,
  context: ToolCallContext,
): Promise<Result<unknown, FolioCliError>> => {
  if (rawArgs !== undefined && !isRecord(rawArgs)) {
    return Result.err(invalidInput("Tool arguments must be an object."));
  }
  const input = rawArgs ?? {};
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
  if (access === "write" && !isFileVersion(fileVersion.value)) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidInput,
        message: "fileVersion is required for a change.",
        hint: "Pass the fileVersion from your latest read of this file.",
      }),
    );
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
  let target: { type: "inPlace" } | { type: "file"; path: string; overwrite: boolean } = {
    type: "inPlace",
  };
  if (destination.value !== undefined) {
    const checked = await checkWithinRoots({
      roots: context.roots,
      target: destination.value,
      argument: "destination",
    });
    if (checked.isErr()) return checked;
    target = { type: "file", path: checked.value, overwrite: input["overwrite"] === true };
  }
  return await executeWriteTool(tool, call, {
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

const toCallToolResult = (result: Result<unknown, FolioCliError>): CallToolResult => {
  const envelope = result.isOk() ? successEnvelope(result.value) : failureEnvelope(result.error);
  return {
    content: [{ type: "text", text: JSON.stringify(envelope) }],
    isError: !envelope.ok,
  };
};

const aboutText = (roots: AllowedRoots): string =>
  [
    `# folio ${packageJson.version}`,
    "",
    `Allowed roots: ${roots.join(", ")}`,
    "",
    "- fileVersion is the SHA-256 of the file's bytes; changes require the one you read.",
    "- Block ids are the paragraph's w14:paraId (blockIdSource: package) or derived from text and",
    "  position (synthetic), valid only for the fileVersion they were read at.",
    "- Offsets in range handles are UTF-16 code units into the block text read_document returns.",
    "- A change writes atomically in place (backup in .folio/backups) or to destination, is journaled",
    "  in .folio/journal.jsonl, and refuses a full package repack unless allowRepack is true.",
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
    { capabilities: { tools: {}, resources: {} }, instructions: INSTRUCTIONS },
  );
  server.setRequestHandler("tools/list", () => ({ tools: listMcpTools() }));
  server.setRequestHandler("tools/call", async (request) => {
    const tool = findFileTool(request.params.name);
    if (tool === undefined) {
      return toCallToolResult(
        Result.err(invalidInput(`Unknown tool ${JSON.stringify(request.params.name)}.`)),
      );
    }
    const result = await Result.tryPromise({
      try: () => runTool(tool, request.params.arguments, context),
      catch: (error) =>
        cliError({
          code: FOLIO_CLI_ERROR_CODES.internal,
          message: error instanceof Error ? error.message : String(error),
        }),
    });
    return toCallToolResult(result.isOk() ? result.value : result);
  });
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
