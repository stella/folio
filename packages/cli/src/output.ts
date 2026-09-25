/**
 * The response envelope both surfaces share: `{ ok: true, data }` on success,
 * `{ ok: false, error: { code, message, hint?, details? } }` on an expected
 * failure. `json` prints the envelope to stdout, success or failure, so a
 * script parses one stream; `text` prints a readable rendering to stdout and
 * failures as `error:` / `hint:` lines on stderr.
 */

import type { FolioCliError, FolioCliErrorCode } from "./errors";

export const OUTPUT_FORMATS = ["json", "text"] as const;

export type OutputFormat = (typeof OUTPUT_FORMATS)[number];

export const isOutputFormat = (value: unknown): value is OutputFormat =>
  OUTPUT_FORMATS.some((format) => format === value);

export type FailureBody = {
  code: FolioCliErrorCode;
  message: string;
  hint?: string;
  details?: unknown;
};

export type Envelope = { ok: true; data: unknown } | { ok: false; error: FailureBody };

export const successEnvelope = (data: unknown): Envelope => ({ ok: true, data });

export const failureEnvelope = ({ code, message, hint, details }: FolioCliError): Envelope => ({
  ok: false,
  error: {
    code,
    message,
    ...(hint !== undefined && { hint }),
    ...(details !== undefined && { details }),
  },
});

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const text = (value: unknown): string => (typeof value === "string" ? value : "");

const renderBlocks = (blocks: readonly unknown[]): string[] =>
  blocks.filter(isRecord).map((block) => `[${text(block["blockId"])}] ${text(block["text"])}`);

const renderComments = (comments: readonly unknown[]): string[] => {
  const lines: string[] = [];
  for (const comment of comments.filter(isRecord)) {
    const status = comment["resolved"] === true ? " (resolved)" : "";
    lines.push(
      `#${text(comment["id"])} ${text(comment["author"])}${status}: ${text(comment["text"])}`,
    );
    const replies: unknown[] = Array.isArray(comment["replies"]) ? comment["replies"] : [];
    for (const reply of replies.filter(isRecord)) {
      lines.push(`  #${text(reply["id"])} ${text(reply["author"])}: ${text(reply["text"])}`);
    }
  }
  return lines;
};

const renderChanges = (changes: readonly unknown[]): string[] =>
  changes
    .filter(isRecord)
    .map(
      (change) =>
        `#${String(change["id"])} ${text(change["type"])} by ${text(change["author"])}: ${text(change["text"])}`,
    );

const renderMatches = (matches: readonly unknown[]): string[] =>
  matches
    .filter(isRecord)
    .map((match) => `[${text(match["blockId"]) || "story"}] ${text(match["context"])}`);

/** Readable lines for a tool result, or `null` to fall back to pretty JSON. */
const renderResult = (tool: string, result: unknown): string[] | null => {
  if (tool === "read_document" && isRecord(result) && Array.isArray(result["blocks"])) {
    const lines = renderBlocks(result["blocks"]);
    if (typeof result["nextCursor"] === "string") {
      lines.push(`(more: --cursor ${result["nextCursor"]})`);
    }
    return lines;
  }
  if (tool === "read_section" && isRecord(result) && Array.isArray(result["blocks"])) {
    return renderBlocks(result["blocks"]);
  }
  if (tool === "find_text" && isRecord(result) && Array.isArray(result["matches"])) {
    return [...renderMatches(result["matches"]), `(${String(result["totalMatches"])} matches)`];
  }
  if (tool === "read_comments" && Array.isArray(result)) {
    return renderComments(result);
  }
  if (tool === "read_changes" && Array.isArray(result)) {
    return renderChanges(result);
  }
  return null;
};

const renderData = (tool: string, data: unknown): string => {
  if (!isRecord(data)) {
    return `${JSON.stringify(data, null, 2)}\n`;
  }
  const header =
    typeof data["path"] === "string" && typeof data["fileVersion"] === "string"
      ? [`${data["path"]} @ ${data["fileVersion"]}`]
      : [];
  const lines = renderResult(tool, data["result"]);
  if (lines === null) {
    return `${JSON.stringify(data, null, 2)}\n`;
  }
  return `${[...header, ...lines].join("\n")}\n`;
};

export type RenderedOutput = { stdout: string; stderr: string };

type RenderEnvelopeOptions = {
  envelope: Envelope;
  format: OutputFormat;
  /** The tool that produced the data, for its text rendering. */
  tool: string;
};

export const renderEnvelope = ({
  envelope,
  format,
  tool,
}: RenderEnvelopeOptions): RenderedOutput => {
  if (format === "json") {
    return { stdout: `${JSON.stringify(envelope)}\n`, stderr: "" };
  }
  if (envelope.ok) {
    return { stdout: renderData(tool, envelope.data), stderr: "" };
  }
  const hint = envelope.error.hint === undefined ? "" : `hint: ${envelope.error.hint}\n`;
  return { stdout: "", stderr: `error: ${envelope.error.message}\n${hint}` };
};
