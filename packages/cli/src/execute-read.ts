import { panic, Result } from "better-result";

import { createReviewerBridge } from "@stll/folio-agents/bridges/reviewer";
import { compareDocxVersions, formatVersionDiffForLLM } from "@stll/folio-agents/compare";
import { executeFolioToolCallUntyped } from "@stll/folio-agents/execute";
import { FOLIO_AGENT_TOOL_NAMES } from "@stll/folio-agents/types";
import type { FolioAIBlock, FolioDocxReviewer } from "@stll/folio-core/server";

import { checkExpectedVersion, openReviewer, readDocumentFile } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { MAX_READ_BLOCKS, type FolioFileToolSpec } from "./registry";

/** One call against one file: the envelope plus the tool's own arguments. */
export type FileToolCall = {
  path: string;
  /** When present, the call is refused unless the file still has this version. */
  fileVersion?: string | undefined;
  args: Readonly<Record<string, unknown>>;
};

/** Per-surface limits on what one read returns. */
export type FolioReadBounds = {
  /** Page size when `read_document` passes none; `null` returns every block. */
  defaultMaxBlocks: number | null;
  /** Most `find_text` matches returned; the rest are counted, not listed. */
  maxMatches: number;
  /** Largest serialized result; `null` for no limit. */
  maxResponseBytes: number | null;
};

export const CLI_READ_BOUNDS: FolioReadBounds = {
  defaultMaxBlocks: null,
  maxMatches: 200,
  maxResponseBytes: null,
};

/** The data every file read returns: which file, at which version, and the result. */
export type FileReadData = {
  path: string;
  fileVersion: string;
  result: unknown;
};

/**
 * Where a block id comes from. `package` ids are the paragraph's own
 * `w14:paraId`; `synthetic` ids are derived from the paragraph's text and
 * position, valid only for the fileVersion they were read at.
 */
export type FolioBlockIdSource = "package" | "synthetic";

const idSourceOf = ({ idStability }: FolioAIBlock): FolioBlockIdSource =>
  idStability === "positional" ? "synthetic" : "package";

type BlockIdSources = ReadonlyMap<string, FolioBlockIdSource>;

const blockIdSources = (reviewer: FolioDocxReviewer): BlockIdSources =>
  new Map(reviewer.snapshot().blocks.map((block) => [block.id, idSourceOf(block)]));

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const blockIdOf = (block: unknown): string | undefined =>
  isRecord(block) && typeof block["blockId"] === "string" ? block["blockId"] : undefined;

/** Label each block of a read with where its id comes from. */
const labelBlocks = (blocks: readonly unknown[], sources: BlockIdSources): unknown[] =>
  blocks.map((block) => {
    const blockId = blockIdOf(block);
    if (blockId === undefined || !isRecord(block)) {
      return block;
    }
    return { ...block, blockIdSource: sources.get(blockId) ?? "synthetic" };
  });

type ReadCursor = { fileVersion: string; afterBlockId: string };

export const encodeReadCursor = ({ fileVersion, afterBlockId }: ReadCursor): string =>
  Buffer.from(JSON.stringify({ v: fileVersion, a: afterBlockId }), "utf8").toString("base64url");

const decodeReadCursor = (cursor: string): ReadCursor | null => {
  const parsed = Result.try((): unknown =>
    JSON.parse(Buffer.from(cursor, "base64url").toString("utf8")),
  );
  if (parsed.isErr() || !isRecord(parsed.value)) {
    return null;
  }
  const { v, a } = parsed.value;
  return typeof v === "string" && typeof a === "string"
    ? { fileVersion: v, afterBlockId: a }
    : null;
};

const invalidInput = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message, hint });

const byteLength = (value: unknown): number =>
  Buffer.byteLength(JSON.stringify(value) ?? "", "utf8");

type ReadDocumentPage = {
  blocks: unknown[];
  totalBlocks: number;
  truncated: boolean;
  nextCursor?: string;
};

type PageReadDocumentOptions = {
  blocks: readonly unknown[];
  sources: BlockIdSources;
  fileVersion: string;
  args: Readonly<Record<string, unknown>>;
  bounds: FolioReadBounds;
  /** Bytes the page's blocks may take; `null` for no limit. */
  pageByteBudget: number | null;
};

/** Room kept in a paged response for its envelope fields and `nextCursor`. */
const PAGE_ENVELOPE_RESERVE_BYTES = 512;

/**
 * Page `read_document`'s blocks by count and serialized size. A cursor names
 * the last block of the previous page and the version it was read at, so a
 * page from one version never continues on another.
 */
export const pageReadDocument = ({
  blocks,
  sources,
  fileVersion,
  args,
  bounds,
  pageByteBudget,
}: PageReadDocumentOptions): Result<ReadDocumentPage, FolioCliError> => {
  const { maxBlocks, cursor } = args;
  if (
    maxBlocks !== undefined &&
    (typeof maxBlocks !== "number" ||
      !Number.isInteger(maxBlocks) ||
      maxBlocks < 1 ||
      maxBlocks > MAX_READ_BLOCKS)
  ) {
    return Result.err(invalidInput(`maxBlocks must be an integer from 1 to ${MAX_READ_BLOCKS}.`));
  }
  let start = 0;
  if (cursor !== undefined) {
    const decoded = typeof cursor === "string" ? decodeReadCursor(cursor) : null;
    if (decoded === null) {
      return Result.err(invalidInput("cursor is not a nextCursor from read_document."));
    }
    if (decoded.fileVersion !== fileVersion) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.staleVersion,
          message: "The document changed after this cursor was issued.",
          hint: "Read again from the start, without a cursor.",
          details: { expected: decoded.fileVersion, actual: fileVersion },
        }),
      );
    }
    const index = blocks.findIndex((block) => blockIdOf(block) === decoded.afterBlockId);
    if (index === -1) {
      return Result.err(invalidInput("cursor names a block this version does not have."));
    }
    start = index + 1;
  }
  const limit = typeof maxBlocks === "number" ? maxBlocks : bounds.defaultMaxBlocks;
  const end = limit === null ? blocks.length : Math.min(blocks.length, start + limit);
  const labelled = labelBlocks(blocks.slice(start, end), sources);
  const page: unknown[] = [];
  let bytes = 0;
  for (const block of labelled) {
    bytes += byteLength(block);
    if (pageByteBudget !== null && bytes > pageByteBudget && page.length > 0) {
      break;
    }
    page.push(block);
  }
  // A cursor names the last block of the page, so a page that stops early
  // must end on a block with an id; drop trailing id-less blocks to the next page.
  while (
    start + page.length < blocks.length &&
    page.length > 0 &&
    blockIdOf(page.at(-1)) === undefined
  ) {
    page.pop();
  }
  const lastBlockId = blockIdOf(page.at(-1));
  const truncated = start + page.length < blocks.length;
  if (truncated && lastBlockId === undefined) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.internal,
        message: "read_document page has no block with an id; cannot issue a cursor.",
      }),
    );
  }
  return Result.ok({
    blocks: page,
    totalBlocks: blocks.length,
    truncated,
    ...(truncated &&
      lastBlockId !== undefined && {
        nextCursor: encodeReadCursor({ fileVersion, afterBlockId: lastBlockId }),
      }),
  });
};

type RunAgentReadOptions = {
  tool: AgentReadTool;
  reviewer: FolioDocxReviewer;
  fileVersion: string;
  args: Readonly<Record<string, unknown>>;
  bounds: FolioReadBounds;
  pageByteBudget: number | null;
};

const unexpectedShape = (tool: string): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.internal,
    message: `${tool} returned an unexpected result shape.`,
  });

const runAgentRead = ({
  tool,
  reviewer,
  fileVersion,
  args,
  bounds,
  pageByteBudget,
}: RunAgentReadOptions): Result<unknown, FolioCliError> => {
  const bridge = createReviewerBridge(reviewer);
  const isReadDocument = tool.agentTool === FOLIO_AGENT_TOOL_NAMES.readDocument;
  const agentArgs = isReadDocument ? {} : args;
  const executed = executeFolioToolCallUntyped(tool.agentTool, agentArgs, bridge);
  if (!executed.ok) {
    return Result.err(invalidInput(executed.error));
  }
  const { result } = executed;
  if (isReadDocument) {
    if (!Array.isArray(result)) return Result.err(unexpectedShape(tool.name));
    return pageReadDocument({
      blocks: result,
      sources: blockIdSources(reviewer),
      fileVersion,
      args,
      bounds,
      pageByteBudget,
    });
  }
  if (tool.agentTool === FOLIO_AGENT_TOOL_NAMES.readSection) {
    if (!isRecord(result) || !Array.isArray(result["blocks"])) {
      return Result.err(unexpectedShape(tool.name));
    }
    return Result.ok({
      ...result,
      blocks: labelBlocks(result["blocks"], blockIdSources(reviewer)),
    });
  }
  if (tool.agentTool === FOLIO_AGENT_TOOL_NAMES.findText) {
    if (!isRecord(result) || !Array.isArray(result["matches"])) {
      return Result.err(unexpectedShape(tool.name));
    }
    const matches: unknown[] = result["matches"];
    return Result.ok(
      matches.length <= bounds.maxMatches
        ? result
        : { ...result, matches: matches.slice(0, bounds.maxMatches), truncated: true },
    );
  }
  return Result.ok(result);
};

const tooLarge = (tool: string, limit: number): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.tooLarge,
    message: `The ${tool} result exceeds the ${limit}-byte response limit.`,
    hint: "Narrow the read: get_document_outline then read_section, a scoped find_text, or read_document paging.",
  });

type AgentReadTool = Extract<FolioFileToolSpec, { type: "agentRead" }>;
type CompareTool = Extract<FolioFileToolSpec, { type: "compare" }>;

/** Diff two files without writing: the structured diff plus a compact rendering. */
const compareFiles = async (
  tool: CompareTool,
  call: FileToolCall,
  bounds: FolioReadBounds,
): Promise<Result<FileReadData, FolioCliError>> => {
  const revisedPath = call.args["revisedPath"];
  if (typeof revisedPath !== "string" || revisedPath === "") {
    return Result.err(invalidInput("compare_documents needs revisedPath."));
  }
  const base = await readDocumentFile(call.path);
  if (base.isErr()) return Result.err(base.error);
  const baseVersion = checkExpectedVersion(base.value, call.fileVersion);
  if (baseVersion.isErr()) return Result.err(baseVersion.error);
  const revised = await readDocumentFile(revisedPath);
  if (revised.isErr()) return Result.err(revised.error);
  const expectedRevised = call.args["revisedFileVersion"];
  const revisedVersion = checkExpectedVersion(
    revised.value,
    typeof expectedRevised === "string" ? expectedRevised : undefined,
  );
  if (revisedVersion.isErr()) return Result.err(revisedVersion.error);

  const diff = await Result.tryPromise({
    try: () =>
      compareDocxVersions(base.value.bytes.slice().buffer, revised.value.bytes.slice().buffer),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `The documents could not be compared: ${error instanceof Error ? error.message : String(error)}`,
      }),
  });
  if (diff.isErr()) return Result.err(diff.error);
  const data = {
    path: base.value.path,
    fileVersion: base.value.fileVersion,
    revised: { path: revised.value.path, fileVersion: revised.value.fileVersion },
    result: { text: formatVersionDiffForLLM(diff.value), diff: diff.value },
  };
  return bounds.maxResponseBytes !== null && byteLength(data) > bounds.maxResponseBytes
    ? Result.err(tooLarge(tool.name, bounds.maxResponseBytes))
    : Result.ok(data);
};

const readWithAgentTool = async (
  tool: AgentReadTool,
  call: FileToolCall,
  bounds: FolioReadBounds,
): Promise<Result<FileReadData, FolioCliError>> => {
  const file = await readDocumentFile(call.path);
  if (file.isErr()) return Result.err(file.error);
  const version = checkExpectedVersion(file.value, call.fileVersion);
  if (version.isErr()) return Result.err(version.error);
  const reviewer = await openReviewer(file.value);
  if (reviewer.isErr()) return Result.err(reviewer.error);

  const { path, fileVersion } = file.value;
  const result = runAgentRead({
    tool,
    reviewer: reviewer.value,
    fileVersion,
    args: call.args,
    bounds,
    pageByteBudget:
      bounds.maxResponseBytes === null
        ? null
        : bounds.maxResponseBytes - byteLength({ path, fileVersion }) - PAGE_ENVELOPE_RESERVE_BYTES,
  });
  if (result.isErr()) return Result.err(result.error);

  const data = { path, fileVersion, result: result.value };
  if (bounds.maxResponseBytes !== null && byteLength(data) > bounds.maxResponseBytes) {
    return Result.err(tooLarge(tool.name, bounds.maxResponseBytes));
  }
  return Result.ok(data);
};

/** Run one read tool, or a compare without a destination. Never writes. */
export const executeReadTool = async (
  tool: FolioFileToolSpec,
  call: FileToolCall,
  bounds: FolioReadBounds,
): Promise<Result<FileReadData, FolioCliError>> => {
  switch (tool.type) {
    case "agentRead":
      return await readWithAgentTool(tool, call, bounds);
    case "compare":
      return await compareFiles(tool, call, bounds);
    case "agentWrite":
    case "resolveChanges":
      return panic("A write tool reached the read executor", { tool: tool.name });
    default: {
      const unreachable: never = tool;
      return panic("Unhandled tool type", { unreachable });
    }
  }
};
